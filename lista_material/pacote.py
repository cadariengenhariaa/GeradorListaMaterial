"""Junta as abas da lista ao modelo da folha de rosto, direto no pacote .xlsx.

O modelo (MODELO_-FOLHA_DE_ROSTO.xlsx) usa recursos do Excel 365 que o openpyxl
não preserva: imagens dentro da célula (aba FOTOS, logos dos clientes) e o
vínculo com a relação de projetos. Por isso o modelo não é aberto pelo
openpyxl: as abas novas são montadas à parte e copiadas para dentro do pacote,
e só as células da folha de rosto que mudam são editadas no XML.

O resultado tem: F. DE ROSTO, as abas da lista e FOTOS (oculta). As abas de
cálculo que existirem no modelo (ex.: CHAMINÉS, PARÂMETROS) não entram.
"""
import copy
import datetime as dt
import io
import posixpath
import re
import zipfile

from lxml import etree

from . import config

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_XML = "http://www.w3.org/XML/1998/namespace"
REL_SHEET = NS_REL + "/worksheet"
REL_CALC = NS_REL + "/calcChain"
CT_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing"

M = lambda tag: f"{{{NS_MAIN}}}{tag}"     # noqa: E731


class ErroModelo(Exception):
    """Modelo da folha de rosto fora do esperado."""


def _xml(dados):
    return etree.fromstring(dados)


def _bytes(raiz):
    return etree.tostring(raiz, xml_declaration=True, encoding="UTF-8", standalone=True)


def _rels_de(parte):
    pasta, nome = posixpath.split(parte)
    return posixpath.join(pasta, "_rels", nome + ".rels")


def _resolver(base, alvo):
    if alvo.startswith("/"):
        return alvo.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base), alvo))


def _relativo(de, para):
    return posixpath.relpath(para, posixpath.dirname(de))


def _ler_rels(partes, parte):
    nome = _rels_de(parte)
    if nome not in partes:
        return []
    raiz = _xml(partes[nome])
    return [(r.get("Id"), r.get("Type"), r.get("Target"), r.get("TargetMode")) for r in raiz]


def _col_num(ref):
    letras = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letras:
        n = n * 26 + ord(ch) - 64
    return n


def _data_excel(d):
    if isinstance(d, dt.datetime):
        d = d.date()
    return (d - dt.date(1899, 12, 30)).days


# ---------------------------------------------------------------------------
# estilos
# ---------------------------------------------------------------------------
def _mesclar_estilos(estilos_modelo, estilos_temp):
    """Acrescenta ao styles.xml do modelo os estilos do arquivo temporário. Retorna (xml, deslocamento)."""
    t = _xml(estilos_modelo)
    s = _xml(estilos_temp)

    def secao(raiz, tag, criar_antes=None):
        el = raiz.find(M(tag))
        if el is None:
            el = etree.Element(M(tag))
            if criar_antes is not None and raiz.find(M(criar_antes)) is not None:
                raiz.find(M(criar_antes)).addprevious(el)
            else:
                raiz.insert(0, el)
        return el

    # formatos numéricos
    t_fmt = secao(t, "numFmts", "fonts")
    existentes = {nf.get("formatCode"): int(nf.get("numFmtId")) for nf in t_fmt}
    proximo = max([163] + list(existentes.values())) + 1
    mapa_fmt = {}
    s_fmt = s.find(M("numFmts"))
    if s_fmt is not None:
        for nf in s_fmt:
            codigo, ident = nf.get("formatCode"), int(nf.get("numFmtId"))
            if codigo in existentes:
                mapa_fmt[ident] = existentes[codigo]
            else:
                novo = copy.deepcopy(nf)
                novo.set("numFmtId", str(proximo))
                t_fmt.append(novo)
                existentes[codigo] = proximo
                mapa_fmt[ident] = proximo
                proximo += 1
    t_fmt.set("count", str(len(t_fmt)))
    if len(t_fmt) == 0:
        t.remove(t_fmt)

    desloc = {}
    for tag in ("fonts", "fills", "borders"):
        t_sec, s_sec = t.find(M(tag)), s.find(M(tag))
        desloc[tag] = len(t_sec)
        for el in s_sec:
            t_sec.append(copy.deepcopy(el))
        t_sec.set("count", str(len(t_sec)))
    t_xfs, s_xfs = t.find(M("cellXfs")), s.find(M("cellXfs"))
    desloc_xf = len(t_xfs)
    for xf in s_xfs:
        novo = copy.deepcopy(xf)
        novo.set("fontId", str(int(xf.get("fontId", "0")) + desloc["fonts"]))
        novo.set("fillId", str(int(xf.get("fillId", "0")) + desloc["fills"]))
        novo.set("borderId", str(int(xf.get("borderId", "0")) + desloc["borders"]))
        nf = int(xf.get("numFmtId", "0"))
        novo.set("numFmtId", str(mapa_fmt.get(nf, nf)))
        novo.set("xfId", "0")
        t_xfs.append(novo)
    t_xfs.set("count", str(len(t_xfs)))
    return _bytes(t), desloc_xf


# ---------------------------------------------------------------------------
# células
# ---------------------------------------------------------------------------
def _texto_inline(c, texto):
    for filho in list(c):
        if filho.tag in (M("v"), M("is"), M("f")):
            c.remove(filho)
    c.set("t", "inlineStr")
    is_ = etree.SubElement(c, M("is"))
    t = etree.SubElement(is_, M("t"))
    t.text = texto
    if texto != texto.strip():
        t.set(f"{{{NS_XML}}}space", "preserve")


def _valor_numero(c, numero):
    for filho in list(c):
        if filho.tag in (M("v"), M("is"), M("f")):
            c.remove(filho)
    if "t" in c.attrib:
        del c.attrib["t"]
    v = etree.SubElement(c, M("v"))
    v.text = repr(int(numero)) if float(numero).is_integer() else repr(float(numero))


def _limpar(c):
    for filho in list(c):
        if filho.tag in (M("v"), M("is"), M("f")):
            c.remove(filho)
    if "t" in c.attrib:
        del c.attrib["t"]


def _cache_formula(c, valor):
    v = c.find(M("v"))
    if v is None:
        v = etree.SubElement(c, M("v"))
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if "t" in c.attrib:
            del c.attrib["t"]
        v.text = repr(int(valor)) if float(valor).is_integer() else ("%.15g" % valor)
    else:
        c.set("t", "str")
        v.text = str(valor)


def _celula(raiz_aba, ref):
    """Acha (ou cria, na ordem certa) a célula ref no sheetData."""
    dados = raiz_aba.find(M("sheetData"))
    num_linha = int(re.search(r"\d+", ref).group(0))
    linha = None
    for row in dados.iter(M("row")):
        r = int(row.get("r"))
        if r == num_linha:
            linha = row
            break
        if r > num_linha:
            linha = etree.Element(M("row"), r=str(num_linha))
            row.addprevious(linha)
            break
    if linha is None:
        linha = etree.SubElement(dados, M("row"), r=str(num_linha))
    col = _col_num(ref)
    for c in linha.iter(M("c")):
        cc = _col_num(c.get("r"))
        if cc == col:
            return c
        if cc > col:
            nova = etree.Element(M("c"), r=ref)
            c.addprevious(nova)
            return nova
    return etree.SubElement(linha, M("c"), r=ref)


# ---------------------------------------------------------------------------
# geração
# ---------------------------------------------------------------------------
def gerar_pacote(caminho_modelo, bytes_listas, abas_geradas, capa, caminho_saida):
    with zipfile.ZipFile(caminho_modelo) as z:
        ordem = [i.filename for i in z.infolist()]
        partes = {n: z.read(n) for n in ordem}
    with zipfile.ZipFile(io.BytesIO(bytes_listas)) as z:
        temp = {n: z.read(n) for n in z.namelist()}

    wb = _xml(partes["xl/workbook.xml"])
    wb_rels = _xml(partes["xl/_rels/workbook.xml.rels"])
    rid_alvo = {r.get("Id"): _resolver("xl/workbook.xml", r.get("Target")) for r in wb_rels}
    folhas = wb.find(M("sheets"))
    abas_modelo = [(s.get("name"), s, rid_alvo[s.get(f"{{{NS_REL}}}id")]) for s in folhas]
    nomes_modelo = [n for n, _, _ in abas_modelo]
    if config.ABA_CAPA not in nomes_modelo:
        raise ErroModelo(f'O modelo não tem a aba "{config.ABA_CAPA}". Use o modelo oficial da folha de rosto.')
    manter = [config.ABA_CAPA] + ([config.ABA_FOTOS] if config.ABA_FOTOS in nomes_modelo else [])
    remover = [(n, el, p) for n, el, p in abas_modelo if n not in manter]

    # ---- remove abas que não fazem parte da lista (e o que só elas usam) ----
    excluidas = set()
    for nome, el, parte in remover:
        excluidas.add(parte)
        for _, tipo, alvo, modo in _ler_rels(partes, parte):
            if modo == "External":
                continue
            if any(x in tipo for x in ("/drawing", "/comments", "/vmlDrawing", "/printerSettings", "/table")):
                destino = _resolver(parte, alvo)
                excluidas.add(destino)
                excluidas.add(_rels_de(destino))
        excluidas.add(_rels_de(parte))
        folhas.remove(el)
    excluidas.add("xl/calcChain.xml")
    for r in list(wb_rels):
        alvo = rid_alvo.get(r.get("Id"))
        if alvo in excluidas or r.get("Type") == REL_CALC:
            wb_rels.remove(r)

    # ---- estilos ----
    estilos, desloc_xf = _mesclar_estilos(partes["xl/styles.xml"], temp["xl/styles.xml"])
    partes["xl/styles.xml"] = estilos

    # strings compartilhadas do arquivo temporário
    sst = []
    if "xl/sharedStrings.xml" in temp:
        for si in _xml(temp["xl/sharedStrings.xml"]).iter(M("si")):
            sst.append(si)

    # ---- abas novas ----
    tmp_wb = _xml(temp["xl/workbook.xml"])
    tmp_rels = _xml(temp["xl/_rels/workbook.xml.rels"])
    tmp_alvo = {r.get("Id"): _resolver("xl/workbook.xml", r.get("Target")) for r in tmp_rels}
    usados_ids = {int(s.get("sheetId")) for s in folhas}
    usados_rid = {r.get("Id") for r in wb_rels}
    n_parte = 1 + max([int(m.group(1)) for n in partes for m in [re.match(r"xl/worksheets/sheet(\d+)\.xml$", n)] if m]
                      + [0])
    novas_folhas, novos_ct = [], []
    proximo_sheet_id = max(usados_ids | {0}) + 1
    contador = 0
    por_titulo = {g.titulo: g for g in abas_geradas}

    def novo_rid():
        k = 1
        while f"rId{k}" in usados_rid:
            k += 1
        usados_rid.add(f"rId{k}")
        return f"rId{k}"

    for s in tmp_wb.find(M("sheets")):
        titulo = s.get("name")
        if titulo not in por_titulo:
            continue
        contador += 1
        parte_tmp = tmp_alvo[s.get(f"{{{NS_REL}}}id")]
        raiz = _xml(temp[parte_tmp])
        valores = por_titulo[titulo].valores_formulas
        for el in raiz.iter():
            if el.tag in (M("c"), M("row")) and el.get("s") is not None:
                el.set("s", str(int(el.get("s")) + desloc_xf))
            elif el.tag == M("col") and el.get("style") is not None:
                el.set("style", str(int(el.get("style")) + desloc_xf))
        for c in raiz.iter(M("c")):
            if c.get("t") == "s":
                v = c.find(M("v"))
                si = sst[int(v.text)]
                c.remove(v)
                c.set("t", "inlineStr")
                is_ = etree.SubElement(c, M("is"))
                for filho in si:
                    is_.append(copy.deepcopy(filho))
            f = c.find(M("f"))
            if f is not None and c.get("r") in valores and valores[c.get("r")] is not None:
                _cache_formula(c, valores[c.get("r")])
            elif f is not None:
                v = c.find(M("v"))
                if v is not None and not (v.text or "").strip():
                    c.remove(v)
        nova_parte = f"xl/worksheets/sheet{n_parte}.xml"
        n_parte += 1
        # relações da aba (comentários, desenho do logo)
        rels_tmp = _ler_rels(temp, parte_tmp)
        if rels_tmp:
            rels_raiz = etree.Element(f"{{{NS_PKG_REL}}}Relationships", nsmap={None: NS_PKG_REL})
            for rid, tipo, alvo, modo in rels_tmp:
                origem = _resolver(parte_tmp, alvo)
                if tipo.endswith("/comments"):
                    destino = f"xl/comments/comment_lista{contador}.xml"
                    partes[destino] = temp[origem]
                    novos_ct.append((destino, CT_COMMENTS))
                elif tipo.endswith("/vmlDrawing"):
                    destino = f"xl/drawings/vmlDrawing_lista{contador}.vml"
                    partes[destino] = temp[origem]
                elif tipo.endswith("/drawing"):
                    destino = f"xl/drawings/drawing_lista{contador}.xml"
                    partes[destino] = temp[origem]
                    novos_ct.append((destino, CT_DRAWING))
                    d_rels = _ler_rels(temp, origem)
                    if d_rels:
                        dr = etree.Element(f"{{{NS_PKG_REL}}}Relationships", nsmap={None: NS_PKG_REL})
                        for drid, dtipo, dalvo, dmodo in d_rels:
                            mid = _resolver(origem, dalvo)
                            ext = posixpath.splitext(mid)[1]
                            nova_midia = f"xl/media/logo_cadari_lista{contador}{ext}"
                            partes[nova_midia] = temp[mid]
                            etree.SubElement(dr, f"{{{NS_PKG_REL}}}Relationship", Id=drid, Type=dtipo,
                                             Target=_relativo(destino, nova_midia))
                        partes[_rels_de(destino)] = _bytes(dr)
                else:
                    continue
                etree.SubElement(rels_raiz, f"{{{NS_PKG_REL}}}Relationship", Id=rid, Type=tipo,
                                 Target=_relativo(nova_parte, destino))
            partes[_rels_de(nova_parte)] = _bytes(rels_raiz)
        partes[nova_parte] = _bytes(raiz)
        novos_ct.append((nova_parte, CT_SHEET))
        rid = novo_rid()
        etree.SubElement(wb_rels, f"{{{NS_PKG_REL}}}Relationship", Id=rid, Type=REL_SHEET,
                         Target=_relativo("xl/workbook.xml", nova_parte))
        el = etree.Element(M("sheet"), name=titulo, sheetId=str(proximo_sheet_id))
        el.set(f"{{{NS_REL}}}id", rid)
        proximo_sheet_id += 1
        novas_folhas.append(el)

    # ordem final: folha de rosto, abas da lista, FOTOS (oculta)
    capa_el = next(s for s in folhas if s.get("name") == config.ABA_CAPA)
    fotos_el = next((s for s in folhas if s.get("name") == config.ABA_FOTOS), None)
    for s in list(folhas):
        folhas.remove(s)
    folhas.append(capa_el)
    for el in novas_folhas:
        folhas.append(el)
    if fotos_el is not None:
        fotos_el.set("state", "hidden")
        folhas.append(fotos_el)
    nomes_finais = [s.get("name") for s in folhas]

    # ---- nomes definidos ----
    dn = wb.find(M("definedNames"))
    if dn is None:
        dn = etree.Element(M("definedNames"))
        ref = wb.find(M("calcPr"))
        (ref.addprevious(dn) if ref is not None else wb.append(dn))
    removidas = [n for n, _, _ in remover]
    for nome in list(dn):
        texto = nome.text or ""
        local = nome.get("localSheetId")
        if local is not None:
            antigo = nomes_modelo[int(local)] if int(local) < len(nomes_modelo) else None
            if antigo not in nomes_finais:
                dn.remove(nome)
                continue
            nome.set("localSheetId", str(nomes_finais.index(antigo)))
        if "#REF!" in texto or any(re.search(r"(^|[^A-Za-zÀ-ÿ])'?" + re.escape(r) + r"'?!", texto) for r in removidas):
            dn.remove(nome)
    for g in abas_geradas:
        if g.titulo not in nomes_finais:
            continue
        idx = nomes_finais.index(g.titulo)
        ult = _col_letra(g.ultima_coluna)
        q = "'" + g.titulo.replace("'", "''") + "'"
        a = etree.SubElement(dn, M("definedName"), name="_xlnm.Print_Area", localSheetId=str(idx))
        a.text = f"{q}!$A$1:${ult}${g.ultima_linha}"
        b = etree.SubElement(dn, M("definedName"), name="_xlnm.Print_Titles", localSheetId=str(idx))
        b.text = f"{q}!$1:$13"
    if len(dn) == 0:
        wb.remove(dn)

    # ---- cálculo e exibição ----
    calc = wb.find(M("calcPr"))
    if calc is None:
        calc = etree.SubElement(wb, M("calcPr"))
    calc.set("fullCalcOnLoad", "1")
    vista = wb.find(M("bookViews"))
    if vista is not None:
        for v in vista:
            v.set("activeTab", "0")
            if "firstSheet" in v.attrib:
                v.set("firstSheet", "0")

    # ---- folha de rosto ----
    parte_capa = rid_alvo[capa_el.get(f"{{{NS_REL}}}id")]
    capa_xml = _xml(partes[parte_capa])
    cel = config.CELULAS_CAPA
    _texto_inline(_celula(capa_xml, cel["ca"]), capa.ca)
    for campo in ("sistema", "etapa", "documento"):
        _texto_inline(_celula(capa_xml, cel[campo]), getattr(capa, campo))
    for campo, attr in (("rev", "rev"), ("te", "te"), ("descricao", "descricao_rev"),
                        ("elab", "elab"), ("verif", "verif"), ("resp", "resp")):
        valor = (getattr(capa, attr) or "").strip()
        c = _celula(capa_xml, cel[campo])
        if valor:
            _texto_inline(c, valor)
        else:
            _limpar(c)
    if capa.data:
        _valor_numero(_celula(capa_xml, cel["data"]), _data_excel(capa.data))
    obra = _celula(capa_xml, cel["obra"])
    if obra.find(M("f")) is not None:
        _cache_formula(obra, capa.obra or "")
    for c in capa_xml.iter(M("c")):
        f = c.find(M("f"))
        if f is not None and "CELL(" in (f.text or "") and "filename" in (f.text or ""):
            _cache_formula(c, capa.numero_documento)
    partes[parte_capa] = _bytes(capa_xml)

    partes["xl/workbook.xml"] = _bytes(wb)
    partes["xl/_rels/workbook.xml.rels"] = _bytes(wb_rels)

    # ---- tipos de conteúdo ----
    ct = _xml(partes["[Content_Types].xml"])
    for o in list(ct):
        if o.tag == f"{{{NS_CT}}}Override" and o.get("PartName").lstrip("/") in excluidas:
            ct.remove(o)
    exts = {d.get("Extension").lower() for d in ct if d.tag == f"{{{NS_CT}}}Default"}
    if "vml" not in exts:
        etree.SubElement(ct, f"{{{NS_CT}}}Default", Extension="vml", ContentType=CT_VML)
    if "png" not in exts:
        etree.SubElement(ct, f"{{{NS_CT}}}Default", Extension="png", ContentType="image/png")
    for parte, tipo in novos_ct:
        etree.SubElement(ct, f"{{{NS_CT}}}Override", PartName="/" + parte, ContentType=tipo)
    partes["[Content_Types].xml"] = _bytes(ct)

    # ---- propriedades (lista de abas) ----
    if "docProps/app.xml" in partes:
        partes["docProps/app.xml"] = _app_xml(nomes_finais)

    # ---- grava ----
    for p in excluidas:
        partes.pop(p, None)
    novas = [p for p in partes if p not in ordem]
    with zipfile.ZipFile(caminho_saida, "w", zipfile.ZIP_DEFLATED) as zo:
        for nome in ["[Content_Types].xml"] + [n for n in ordem if n != "[Content_Types].xml"] + novas:
            if nome in partes:
                zo.writestr(nome, partes[nome])
    return nomes_finais


def _col_letra(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _app_xml(abas):
    ns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    itens = "".join(f"<vt:lpstr>{_esc(a)}</vt:lpstr>" for a in abas)
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           f'<Properties xmlns="{ns}" xmlns:vt="{vt}"><Application>Microsoft Excel</Application>'
           f'<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs>'
           f'<vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Planilhas</vt:lpstr></vt:variant>'
           f'<vt:variant><vt:i4>{len(abas)}</vt:i4></vt:variant></vt:vector></HeadingPairs>'
           f'<TitlesOfParts><vt:vector size="{len(abas)}" baseType="lpstr">{itens}</vt:vector></TitlesOfParts>'
           f'<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
           f'<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>')
    return xml.encode("utf-8")


def _esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
