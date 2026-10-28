"""Uso pela linha de comando (sem interface):

    python -m lista_material arquivo.csv --ca CA-1276 --documento A056-0000-CLEX-LM-TT-GER-LMT-R00
"""
import argparse
import datetime as dt
import os
import sys

from .gerador import DadosCapa, RelacaoProjetos, carregar_csv, gerar, nova_analise


def main():
    raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    p = argparse.ArgumentParser(description="Gera a lista de material no padrão Cadari a partir do CSV por instância.")
    p.add_argument("csv")
    p.add_argument("--modelo", default=os.path.join(raiz, "modelos", "MODELO_-FOLHA_DE_ROSTO.xlsx"))
    p.add_argument("--ca", default="")
    p.add_argument("--documento", default="")
    p.add_argument("--sistema", default="CLIMATIZAÇÃO")
    p.add_argument("--desenhos", default="")
    p.add_argument("--prefixo", default="CLI")
    p.add_argument("--saida", default="")
    p.add_argument("--ignorar-problemas", action="store_true",
                   help="gera mesmo com linhas com problema (elas ficam fora da lista)")
    a = p.parse_args()

    leitura = carregar_csv(a.csv)
    if leitura.problemas:
        print(f"{len(leitura.problemas)} linhas com problema no CSV:")
        for pr in leitura.problemas[:200]:
            print(f"  linha {pr.linha}: {pr.texto} | {pr.descricao[:60]} | {pr.local} | {pr.nivel}")
        if not a.ignorar_problemas:
            print("Corrija o modelo e exporte de novo, ou rode com --ignorar-problemas.")
            sys.exit(2)
    relacao = RelacaoProjetos(a.modelo)
    projeto = relacao.buscar(a.ca)
    ca = RelacaoProjetos.normalizar_codigo(a.ca)
    documento = a.documento or (os.path.splitext(os.path.basename(a.saida))[0] if a.saida else "LISTA_DE_MATERIAL")
    capa = DadosCapa(ca=ca, obra=(projeto or {}).get("obra", ""), sistema=a.sistema, numero_documento=documento,
                     data=dt.date.today(), desenhos=a.desenhos, prefixo=a.prefixo)
    saida = a.saida or os.path.join(os.path.dirname(os.path.abspath(a.csv)), documento + ".xlsx")
    res = gerar(leitura, nova_analise(leitura), capa, a.modelo, saida)
    print(f"Planilha gerada: {res.caminho}")
    print(f"Abas: {', '.join(res.abas)} | itens: {res.itens} | mesclagens: {res.mesclagens} | comentários: {res.comentarios}")
    for c in res.diferencas:
        print(f"  diferença no subtotal {c['grupo']} / {c['subgrupo']}: CSV {c['csv']} x lista {c['soma']}")


if __name__ == "__main__":
    main()
