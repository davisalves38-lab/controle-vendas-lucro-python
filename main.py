from pathlib import Path

from openpyxl import Workbook
from calcular import calcular_lucro, calcular_custo, calcular_faturamento, calcular_margem_lucro
from formatacao import azul, branco, borda, alinhar
from openpyxl.utils import get_column_letter

planilha = Workbook()
aba = planilha.active
aba.title = "Relatório de Vendas"

pasta_projeto = Path(__file__).resolve().parent
pasta_exemplos = pasta_projeto / "exemplos"

pasta_exemplos.mkdir(parents=True, exist_ok=True)

aba.append(["Produto", "Quantidade", "Faturamento", "Custo Total", "Lucro", "Margem de Lucro (%)"])

for cell in aba[1]:
    cell.fill = azul
    cell.font = branco
    cell.alignment = alinhar

for numero in range(1, aba.max_column + 1):
    coluna = get_column_letter(numero)
    if coluna in ["A"]:
        aba.column_dimensions[coluna].width = 25
    elif coluna in ["F"]:
        aba.column_dimensions[coluna].width = 24
    else:
        aba.column_dimensions[coluna].width = 15

vendas = [
    ["Arroz 5kg", 20, 15.00, 22.00],
    ["Feijão 1 kg", 15, 5.50, 8.90],
    ["Café 500g", 10, 12.00, 18.50],
    ["Açúcar 1kg", 25, 3.20, 4.99],
    ["Leite 1L", 0, 3.50, 5.99],
    ["Macarrão 500g", 12, 4.00, -6.50],
    ["Óleo 900ml", 8, 0, 9.90]
]

quantidade_total = 0
faturamento_total = 0
custo_bruto = 0
lucro_bruto_total = 0


for produto, quantidade, preco_unitario, preco_venda  in vendas:
    if quantidade <= 0:
        print(f"\nProduto: {produto} - Quantidade inválida ({quantidade}).")
        continue

    elif preco_unitario <= 0:
        print(f"Produto: {produto} - Preço unitário inválido (R${preco_unitario:.2f}).")
        continue

    elif preco_venda <= 0:
        print(f"Produto: {produto} - Preço de venda inválido (R${preco_venda:.2f}).")
        continue

    faturamento = calcular_faturamento(quantidade, preco_venda )
    custo_total = calcular_custo(quantidade, preco_unitario)
    lucro = calcular_lucro(faturamento, custo_total)
    margem_lucro = calcular_margem_lucro(lucro, faturamento)

    aba.append([produto, quantidade, faturamento, custo_total, lucro, margem_lucro])

    print("\nProduto: ", produto)
    print("Quantidade vendida: ", quantidade)
    print(f"Faturamento: R${faturamento:.2f}")
    print(f"Custo Total: R${custo_total:.2f}")
    print(f"Lucro: R${lucro:.2f}")
    print(f"Margem de Lucro: {margem_lucro:.2f}%")

    quantidade_total += quantidade
    faturamento_total += faturamento
    custo_bruto += custo_total
    lucro_bruto_total += lucro

margem_lucro_total = calcular_margem_lucro(lucro_bruto_total, faturamento_total)

print("\n===== RESUMO GERAL =====")
print(f"\nQuantidade total vendida: {quantidade_total}")
print(f"Faturamento total: R${faturamento_total:.2f}")
print(f"Custo total: R${custo_bruto:.2f}")
print(f"Lucro total: R${lucro_bruto_total:.2f}")
print(f"Margem geral: {margem_lucro_total:.2f}%")

aba.append(["TOTAL", quantidade_total, faturamento_total, custo_bruto, lucro_bruto_total, margem_lucro_total])

for linha in range(2, aba.max_row + 1):
    for coluna in ["C", "D", "E"]:
        aba[f"{coluna}{linha}"].number_format = '"R$" #,##0.00'

    aba[f"F{linha}"].number_format = '0.00"%"'


for cell in aba[aba.max_row]:
    cell.fill = azul
    cell.font = branco

for linha in aba.iter_rows(min_row=1, max_row=aba.max_row, min_col=1, max_col=aba.max_column):
    for cell in linha:
        cell.border = borda

planilha.save(pasta_exemplos / "relatorio_vendas.xlsx")