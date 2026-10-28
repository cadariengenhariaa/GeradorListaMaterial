"""Fluxo completo: CSV -> análise -> abas da lista -> pacote final com a folha de rosto."""
import os
import posixpath
import zipfile
from dataclasses import dataclass, field

from lxml import etree

from . import config
from .leitura import conferir_subtotais, ler_csv
from .modelo import Analise
from .pacote import gerar_pacote
from .planilha import DadosCapa, montar_listas
from .relacao import RelacaoProjetos

__all__ = ["carregar_csv", "nova_analise", "gerar", "DadosCapa", "RelacaoProjetos", "logo_do_modelo"]


@dataclass
class Resultado:
    caminho: str
    abas: list
    itens: int
    comentarios: int
    mesclagens: int
    ignoradas: int
    conferencia: list = field(default_factory=list)

    @property
    def diferencas(self):
        return [c for c in self.conferencia if not c["ok"]]


def carregar_csv(caminho, progresso=None):
    return ler_csv(caminho, progresso)


def nova_analise(leitura, realocar=True):
    return Analise(leitura.instancias_validas, realocar=realocar)


def logo_do_modelo(caminho_modelo, pasta_recursos=None):
    """Logo da Cadari: a imagem do topo da folha de rosto do modelo (ou recursos/marca_cadari.png)."""
    try:
        with zipfile.ZipFile(caminho_modelo) as z:
            ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            wb = etree.fromstring(z.read("xl/workbook.xml"))
            rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            alvo = {r.get("Id"): r.get("Target") for r in rels}
            for s in wb.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"):
                if s.get("name") != config.ABA_CAPA:
                    continue
                parte = posixpath.normpath(posixpath.join("xl", alvo[s.get(f"{{{ns_r}}}id")].lstrip("/")))
                srels = etree.fromstring(z.read(posixpath.join(posixpath.dirname(parte), "_rels",
                                                               posixpath.basename(parte) + ".rels")))
                for r in srels:
                    if r.get("Type").endswith("/drawing"):
                        dpart = posixpath.normpath(posixpath.join(posixpath.dirname(parte), r.get("Target")))
                        drels = etree.fromstring(z.read(posixpath.join(posixpath.dirname(dpart), "_rels",
                                                                       posixpath.basename(dpart) + ".rels")))
                        dalvo = {x.get("Id"): x.get("Target") for x in drels}
                        dxml = etree.fromstring(z.read(dpart))
                        ns_x = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                        ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
                        for anc in dxml:
                            linha = anc.find(f"{{{ns_x}}}from/{{{ns_x}}}row")
                            blip = anc.find(f".//{{{ns_a}}}blip")
                            if linha is not None and blip is not None and int(linha.text) < 4:
                                img = posixpath.normpath(posixpath.join(posixpath.dirname(dpart),
                                                                        dalvo[blip.get(f"{{{ns_r}}}embed")]))
                                return z.read(img)
    except Exception:
        pass
    if pasta_recursos:
        caminho = os.path.join(pasta_recursos, "marca_cadari.png")
        if os.path.exists(caminho):
            with open(caminho, "rb") as f:
                return f.read()
    return None


def gerar(leitura, analise, capa, caminho_modelo, caminho_saida, progresso=None, pasta_recursos=None):
    def etapa(inicio, fim):
        return (lambda x: progresso(inicio + (fim - inicio) * x)) if progresso else None
    if progresso:
        progresso(0.05)
    logo = logo_do_modelo(caminho_modelo, pasta_recursos)
    bytes_listas, abas = montar_listas(analise, capa, logo, etapa(0.05, 0.6))
    if not abas:
        raise ValueError("Nenhum item com quantitativo para lançar na lista.")
    if progresso:
        progresso(0.7)
    gerar_pacote(caminho_modelo, bytes_listas, abas, capa, caminho_saida)
    if progresso:
        progresso(1.0)
    ignoradas = leitura.linhas_com_problema
    return Resultado(
        caminho=caminho_saida, abas=[a.titulo for a in abas], itens=sum(a.itens for a in abas),
        comentarios=sum(a.comentarios for a in abas), mesclagens=sum(a.mesclagens for a in abas),
        ignoradas=len(ignoradas), conferencia=conferir_subtotais(leitura, ignoradas))
