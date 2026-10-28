"""Consulta à relação de projetos guardada no modelo da folha de rosto.

A folha de rosto busca a obra e a logo do cliente na planilha
Z:\\Publico\\RELAÇÃO DE PROJETOS.xlsm (vínculo externo). O Excel guarda uma cópia
desses dados dentro do próprio modelo; este módulo lê essa cópia para mostrar
no programa a obra e o cliente do código CA digitado.
"""
import re
import zipfile

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ABA_RELACAO = "RELAÇÃO DE PROJETOS"


def _col_idx(ref):
    letras = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letras:
        n = n * 26 + (ord(ch) - 64)
    return n


class RelacaoProjetos:
    def __init__(self, caminho_modelo):
        self.projetos = {}
        self.carregada = False
        try:
            self._carregar(caminho_modelo)
        except Exception:
            self.projetos = {}

    def _carregar(self, caminho):
        with zipfile.ZipFile(caminho) as z:
            nomes = [n for n in z.namelist() if n.startswith("xl/externalLinks/externalLink") and n.endswith(".xml")]
            for nome in nomes:
                raiz = etree.fromstring(z.read(nome))
                abas = [s.get("val") for s in raiz.iter(f"{{{NS}}}sheetName")]
                if ABA_RELACAO not in abas:
                    continue
                id_aba = str(abas.index(ABA_RELACAO))
                for sd in raiz.iter(f"{{{NS}}}sheetData"):
                    if sd.get("sheetId") != id_aba:
                        continue
                    for row in sd.iter(f"{{{NS}}}row"):
                        valores = {}
                        for c in row.iter(f"{{{NS}}}cell"):
                            v = c.findtext(f"{{{NS}}}v")
                            if v is not None:
                                valores[_col_idx(c.get("r"))] = v
                        codigo = (valores.get(2) or "").strip()      # coluna B
                        if codigo.upper().startswith("CA"):
                            self.projetos[codigo.upper()] = dict(
                                codigo=codigo,
                                obra=(valores.get(4) or "").strip(),          # coluna D (VLOOKUP ..., 3)
                                referencia=(valores.get(5) or "").strip(),    # coluna E (ex.: A056 - PLACE VENDÔME)
                                cliente=(valores.get(6) or "").strip(),       # coluna F (VLOOKUP ..., 5)
                            )
                    self.carregada = True

    @staticmethod
    def normalizar_codigo(txt):
        t = (txt or "").strip().upper().replace(" ", "")
        if re.fullmatch(r"\d+[A-Z]?", t):
            t = "CA-" + t
        elif re.fullmatch(r"CA\d+[A-Z]?", t):
            t = "CA-" + t[2:]
        return t

    def buscar(self, codigo):
        return self.projetos.get(self.normalizar_codigo(codigo))

    @staticmethod
    def prefixo_obra(projeto):
        """'A056 - PLACE VENDÔME' -> 'A056'."""
        if not projeto:
            return ""
        m = re.match(r"\s*([A-Z]\d{2,4}[A-Z]?)\b", projeto.get("referencia", ""))
        return m.group(1) if m else ""
