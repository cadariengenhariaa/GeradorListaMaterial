# Gerador de lista de material (Cadari)

Lê o CSV por instância da tabela **0. LISTA DE MATERIAL GERAL** exportada do Revit e gera a lista de material no padrão Cadari: folha de rosto oficial, abas **CLI - ÁREAS DE USO COMUM** e **CLI - ÁREAS PRIVATIVAS** e a aba **FOTOS** oculta.

## Como usar

1. Rode `executar.bat`. Na primeira vez ele instala `openpyxl` e `lxml` (precisa do Python 3.10 ou mais novo).
2. **Selecionar CSV**: escolha a exportação por instância. Se houver linhas com problema, a lista aparece na hora, com o número da linha do CSV.
3. **Folha de rosto**: digite o código da obra (ex.: `CA-1276` ou só `1276`). O programa mostra a obra e o cliente pela relação de projetos salva no modelo e sugere o nº do documento.
4. **Níveis**: confira a aba de cada nível, o texto do cabeçalho, a ordem e as repetições.
5. **Ambientes**: renomeie ou mova colunas, se precisar.
6. **Gerar planilha**: o nome do arquivo é o nº do documento (a folha de rosto lê o nome do arquivo).

## Exportação no Revit

- Tabela por instância, com **Itemizar todas as instâncias** marcado.
- Colunas: `DESCRIÇÃO`, `CA_Local Atendido`, `CA_Localização`, `CA_Nível de Referência`, `UNID.`, `QTD.`.
- Pode manter cabeçalhos de grupo e linhas de subtotal: o programa usa os subtotais para conferir a leitura.
- Delimitador vírgula, ponto e vírgula ou tabulação. Não precisa de aspas.

## O que é verificado antes de gerar

| Problema | Exemplo |
|---|---|
| Descrição marcada como `NF` | `NF - BUCHA DE REDUÇÃO...` |
| Descrição vazia | linha com local, nível e quantidade, sem descrição |
| Sem `CA_Local Atendido` | |
| Sem `CA_Nível de Referência` | |
| Sem unidade ou quantidade inválida | |

Ao clicar em **Gerar planilha** com problemas, o programa pergunta antes: ver problemas, cancelar ou gerar mesmo assim. Gerando assim, essas linhas ficam fora da lista e a mensagem final informa quanto ficou de fora. A lista de problemas pode ser exportada em CSV.

## Regras de montagem

- Cada nível de referência vira um grupo de colunas e cada local atendido vira uma coluna.
- Local com vários ambientes (ex.: `SAT-MOTORISTA`, `ESPAÇO KIDS-MINIMARKET`): os ambientes ficam em colunas vizinhas e o quantitativo vai em células mescladas. Hífen com espaço (`APARTAMENTO FINAL 02 - OPÇÃO 1`) não separa ambientes.
- Se a linha já tiver valor próprio de um desses ambientes, o valor vai para um ambiente livre do grupo (ou é somado a um deles), com comentário na célula.
- Local atendido com o nome de um nível (ex.: `1º SUBSOLO`) ou com `GERAL` vira a coluna **GERAL** do nível.
- Até 3 peças com nível de referência diferente do ambiente atendido (ex.: condensadora no 2º subsolo atendendo a academia do 1º subsolo) vão para o ambiente atendido, com comentário. A opção pode ser desligada na tela.
- Itens com a mesma descrição e unidade na mesma subseção são somados (ex.: os dois grupos de chapa galvanizada).
- Tubos e conexões de cobre ficam em ordem de diâmetro; o restante segue a ordem do Revit.
- Repetições (linha 13) é informativo: os quantitativos já vêm somados do modelo.

## Folha de rosto

O modelo fica em `modelos/MODELO_-FOLHA_DE_ROSTO.xlsx` e é usado sem ser regravado. O programa preenche só o código CA (L4), sistema, etapa e documento (A6 a A8) e a linha de revisão (linha 39). A obra (A5) e a logo do cliente (A1) continuam vindo da relação de projetos pelo vínculo do modelo.

- Para atualizar o modelo (ex.: logo de cliente nova na aba FOTOS), troque o arquivo na pasta `modelos` ou use **Trocar modelo**. O modelo precisa ter as abas `F. DE ROSTO` e `FOTOS`.
- Abas de cálculo que existirem no modelo (ex.: CHAMINÉS, PARÂMETROS) não entram na lista.
- As logos em célula só aparecem no Excel 365. Ao abrir, o Excel pode pedir para atualizar o vínculo com a relação de projetos.

## Ajustes

As regras editáveis ficam em `lista_material/config.py`:

- Mapeamento dos grupos do Revit para as seções e subseções da lista.
- Textos fixos de elétrica, outros e notas.
- Termos que mandam um nível para áreas de uso comum.
- Células da folha de rosto e padrão do nº do documento.

## Executável

`gerar_executavel.bat` gera `dist\GeradorListaMaterial\GeradorListaMaterial.exe` (PyInstaller). Para distribuir, copie a pasta `dist\GeradorListaMaterial` inteira. As pastas `modelos` e `recursos` ao lado do .exe podem ser trocadas sem gerar de novo.

## Linha de comando

```
python -m lista_material arquivo.csv --ca CA-1276 --documento A056-0000-CLEX-LM-TT-GER-LMT-R00 --desenhos "A056-0001 @ 0011"
```

Com linhas com problema, o comando lista as linhas e para. Use `--ignorar-problemas` para gerar mesmo assim.

## Versões

- 1.1.0: títulos de seção com a listra para tópicos (check + faixa translúcida) e assinatura C✓ dourada no rodapé e no ícone, conforme o guia de identidade visual.
- 1.0.0: primeira versão.
