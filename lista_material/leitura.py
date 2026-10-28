"""Leitura do CSV por instância exportado do Revit (tabela "LISTA DE MATERIAL GERAL").

Formato esperado (o mesmo da exportação usada no projeto 1276):
    linha 1: título da tabela
    linha 2: cabeçalho  DESCRIÇÃO, CA_Local Atendido, CA_Localização, CA_Nível de Referência, UNID., QTD.
    depois: cabeçalho de grupo, cabeçalho de subgrupo, instâncias, linha de subtotal ...

O Revit exporta sem aspas: a descrição pode ter vírgulas e a quantidade usa
vírgula decimal. A leitura reconstrói os campos a partir das colunas do
cabeçalho, sem depender de aspas.
"""
import csv
import io
import re
import unicodedata
from dataclasses import dataclass, field

from . import config


def normalizar(txt):
    """Maiúsculas, sem acento e com espaços simples (só para comparar nomes)."""
    txt = unicodedata.normalize("NFKD", str(txt or ""))
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", txt).strip().upper()


@dataclass
class Instancia:
    linha: int
    grupo: str
    subgrupo: str
    descricao: str
    local: str
    localizacao: str
    nivel: str
    unidade: str
    qtd: float


@dataclass
class Problema:
    linha: int
    problemas: list
    descricao: str
    local: str
    nivel: str

    @property
    def texto(self):
        return "; ".join(self.problemas)


@dataclass
class Subtotal:
    linha: int
    grupo: str
    subgrupo: str
    valor: float


@dataclass
class LeituraCSV:
    caminho: str
    titulo: str = ""
    codificacao: str = ""
    delimitador: str = ","
    instancias: list = field(default_factory=list)
    problemas: list = field(default_factory=list)
    subtotais: list = field(default_factory=list)
    total_linhas: int = 0

    @property
    def linhas_com_problema(self):
        return {p.linha for p in self.problemas}

    @property
    def instancias_validas(self):
        ruins = self.linhas_com_problema
        return [i for i in self.instancias if i.linha not in ruins]


class ErroLeitura(Exception):
    """Arquivo que não segue o padrão de exportação."""


_NUM_INT = re.compile(r"^-?\d+$")
_NUM_DEC = re.compile(r"^-?\d{1,3}(\.\d{3})*([.,]\d+)?$|^-?\d+([.,]\d+)?$")


def _ler_texto(caminho):
    with open(caminho, "rb") as f:
        dados = f.read()
    if dados.startswith(b"\xff\xfe") or dados.startswith(b"\xfe\xff"):
        return dados.decode("utf-16"), "UTF-16"
    if dados.startswith(b"\xef\xbb\xbf"):
        return dados.decode("utf-8-sig"), "UTF-8"
    try:
        return dados.decode("utf-8"), "UTF-8"
    except UnicodeDecodeError:
        return dados.decode("cp1252"), "Windows-1252"


def _num(txt):
    """Converte '12', '0,31', '1.234,56' ou '0.31' em float. Retorna None se não for número."""
    t = (txt or "").strip()
    if not t:
        return None
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    try:
        return float(t)
    except ValueError:
        return None


def _achar_cabecalho(linhas):
    """Procura a linha de cabeçalho e o delimitador. Retorna (índice, delimitador, mapa de colunas)."""
    for idx, linha in enumerate(linhas[:30]):
        for delim in ("\t", ";", ","):
            campos = [normalizar(c) for c in linha.split(delim)]
            if len(campos) < 4:
                continue
            mapa = {}
            for pos, nome in enumerate(campos):
                if nome.startswith("DESCRI"):
                    mapa.setdefault("descricao", pos)
                elif "LOCAL ATENDIDO" in nome:
                    mapa.setdefault("local", pos)
                elif "LOCALIZA" in nome:
                    mapa.setdefault("localizacao", pos)
                elif "NIVEL" in nome:
                    mapa.setdefault("nivel", pos)
                elif nome.startswith("UNID"):
                    mapa.setdefault("unidade", pos)
                elif nome.startswith("QTD") or nome.startswith("QUANT"):
                    mapa.setdefault("qtd", pos)
            obrigatorias = {"descricao", "local", "nivel", "unidade", "qtd"}
            if obrigatorias <= set(mapa):
                return idx, delim, mapa, len(campos)
    raise ErroLeitura(
        "Não encontrei o cabeçalho da tabela. O CSV precisa ter as colunas DESCRIÇÃO, CA_Local Atendido, "
        "CA_Nível de Referência, UNID. e QTD. Exporte a tabela por instância do Revit com essas colunas.")


def _dividir(linha, delim):
    if delim == "," and ('",' in linha or linha.startswith('"')):
        return next(csv.reader(io.StringIO(linha), delimiter=",", quotechar='"'))
    return linha.split(delim)


def _resolver_campos(campos, mapa, n_cab, delim):
    """Reagrupa os campos quando a descrição tem vírgulas ou a quantidade tem vírgula decimal."""
    n = len(campos)
    if n == n_cab:
        return campos, True
    if n < n_cab:
        return campos + [""] * (n_cab - n), True
    extra = n - n_cab
    d, q = mapa["descricao"], mapa["qtd"]
    candidatos = []
    for extra_qtd in (1, 0):
        extra_desc = extra - extra_qtd
        if extra_desc < 0:
            continue
        novos = list(campos)
        # quantidade com vírgula decimal: junta dois campos na posição da quantidade
        pos_q = q + (extra_desc if q > d else 0)
        if extra_qtd:
            if pos_q + 1 >= len(novos):
                continue
            a, b = novos[pos_q].strip(), novos[pos_q + 1].strip()
            if not (_NUM_INT.match(a.replace(".", "")) and b.isdigit()):
                continue
            novos[pos_q:pos_q + 2] = [a + "," + b]
        if extra_desc:
            novos[d:d + extra_desc + 1] = [delim.join(novos[d:d + extra_desc + 1])]
        if len(novos) != n_cab:
            continue
        unid = novos[mapa["unidade"]].strip()
        qtd = novos[q].strip()
        ok = (not unid or _num(unid) is None) and (not qtd or _num(qtd) is not None)
        candidatos.append((ok, novos))
    for ok, novos in candidatos:
        if ok:
            return novos, True
    return (candidatos[0][1] if candidatos else campos[:n_cab]), False


def ler_csv(caminho, progresso=None):
    """Lê o CSV e devolve instâncias, subtotais e problemas (com o número da linha no arquivo)."""
    texto, codificacao = _ler_texto(caminho)
    linhas = texto.splitlines()
    idx_cab, delim, mapa, n_cab = _achar_cabecalho(linhas)
    res = LeituraCSV(caminho=caminho, codificacao=codificacao, delimitador=delim, total_linhas=len(linhas))
    if idx_cab > 0:
        res.titulo = linhas[0].split(delim)[0].strip()

    grupo, subgrupo = "", ""
    pendentes = []          # cabeçalhos de grupo consecutivos ainda não aplicados
    profundidade = None     # 1 ou 2 níveis de agrupamento na tabela
    marcador_nf = re.compile(r"(?<![0-9A-Za-zÀ-ÿ])" + re.escape(config.MARCADOR_NF) + r"(?![0-9A-Za-zÀ-ÿ])")

    total = len(linhas)
    for i in range(idx_cab + 1, total):
        if progresso and i % 2000 == 0:
            progresso(i / total)
        bruta = linhas[i]
        num_linha = i + 1
        if not bruta.strip(delim + " \t\r\n"):
            # linha só com separadores: subgrupo sem nome (ex.: ACESSÓRIOS)
            pendentes.append("")
            continue
        campos, legivel = _resolver_campos(_dividir(bruta, delim), mapa, n_cab, delim)
        val = {k: (campos[p].strip() if p < len(campos) else "") for k, p in mapa.items()}
        outros_vazios = all(not val[k] for k in ("local", "nivel", "unidade", "qtd") if k in val) and \
            not val.get("localizacao")
        if val["descricao"] and outros_vazios:
            pendentes.append(val["descricao"])
            continue
        so_qtd = not val["descricao"] and not val["local"] and not val["nivel"] and not val["unidade"] \
            and not val.get("localizacao") and val["qtd"]
        if so_qtd:
            valor = _num(val["qtd"])
            if valor is not None:
                res.subtotais.append(Subtotal(num_linha, grupo, subgrupo, valor))
            continue

        # linha de instância: aplica os cabeçalhos pendentes
        if pendentes:
            if profundidade is None:
                profundidade = 2 if len(pendentes) >= 2 else 1
            if profundidade == 1:
                grupo, subgrupo = pendentes[-1], ""
            elif len(pendentes) >= 2:
                grupo, subgrupo = pendentes[-2], pendentes[-1]
            else:
                subgrupo = pendentes[-1]
            pendentes = []

        problemas = []
        if not legivel:
            problemas.append("Linha fora do padrão de exportação")
        if not val["descricao"]:
            problemas.append("Descrição vazia")
        elif marcador_nf.search(val["descricao"]):
            problemas.append('Descrição marcada como "NF"')
        if not val["local"]:
            problemas.append("Sem CA_Local Atendido")
        if not val["nivel"]:
            problemas.append("Sem CA_Nível de Referência")
        if not val["unidade"]:
            problemas.append("Sem unidade")
        qtd = _num(val["qtd"])
        if qtd is None:
            problemas.append("Quantidade vazia ou inválida")
            qtd = 0.0

        res.instancias.append(Instancia(
            linha=num_linha, grupo=grupo, subgrupo=subgrupo, descricao=val["descricao"],
            local=val["local"], localizacao=val.get("localizacao", ""), nivel=val["nivel"],
            unidade=val["unidade"], qtd=qtd))
        if problemas:
            res.problemas.append(Problema(num_linha, problemas, val["descricao"], val["local"], val["nivel"]))

    if not res.instancias:
        raise ErroLeitura("O CSV não tem nenhuma instância. Confira se a tabela foi exportada por instância "
                          "(sem marcar 'Itemizar cada instância' desligado).")
    return res


def conferir_subtotais(leitura, ignoradas=None):
    """Compara cada linha de subtotal do CSV com a soma das instâncias lidas do mesmo subgrupo.

    'diferenca' mostra se a leitura bate com o Revit (só arredondamento de 2 casas é aceito);
    'fora' é o quanto ficou fora da lista por causa das linhas com problema.
    """
    ignoradas = ignoradas or set()
    soma, lista, n_dec = {}, {}, {}
    for inst in leitura.instancias:
        chave = (inst.grupo, inst.subgrupo)
        soma[chave] = soma.get(chave, 0.0) + inst.qtd
        if inst.linha not in ignoradas:
            lista[chave] = lista.get(chave, 0.0) + inst.qtd
        if abs(inst.qtd - round(inst.qtd)) > 1e-9:
            n_dec[chave] = n_dec.get(chave, 0) + 1
    resultado = []
    for st in leitura.subtotais:
        chave = (st.grupo, st.subgrupo)
        total = soma.get(chave, 0.0)
        diferenca = st.valor - total
        tolerancia = max(0.011, 0.005 * n_dec.get(chave, 0))      # arredondamento de 2 casas por instância
        resultado.append(dict(grupo=st.grupo, subgrupo=st.subgrupo, linha=st.linha, csv=st.valor,
                              soma=round(total, 4), lista=round(lista.get(chave, 0.0), 4),
                              diferenca=round(diferenca, 4), fora=round(total - lista.get(chave, 0.0), 4),
                              ok=abs(diferenca) <= tolerancia))
    return resultado
