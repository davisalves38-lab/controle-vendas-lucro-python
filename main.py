from pathlib import Path

from openpyxl import Workbook
from calcular import calcular_lucro, calcular_custo, calcular_faturamento, calcular_margem_lucro
from formatacao import azul, branco, borda, alinhar, vermelho_escuro
from openpyxl.utils import get_column_letter
from ler_vendas import carregar_vendas


planilha = Workbook()
aba = planilha.active
aba.title = "Relatório de Vendas"
aba_rejeitados = planilha.create_sheet("Registros Rejeitados")


pasta_projeto = Path(__file__).resolve().parent
pasta_exemplos = pasta_projeto / "exemplos"

pasta_exemplos.mkdir(parents=True, exist_ok=True)

aba.append(["Produto", "Quantidade", "Faturamento", "Custo Total", "Lucro", "Margem de Lucro (%)"])
aba_rejeitados.append(["Produto", "Motivo de rejeição"])

for cell in aba[1]:
    cell.fill = azul
    cell.font = branco
    cell.alignment = alinhar

for cell in aba_rejeitados[1]:
    cell.fill = vermelho_escuro
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

for numero in range(1, aba_rejeitados.max_column + 1):
    coluna = get_column_letter(numero)
    if coluna in ["A"]:
        aba_rejeitados.column_dimensions[coluna].width = 25
    else:
        aba_rejeitados.column_dimensions[coluna].width = 45

vendas = carregar_vendas(pasta_exemplos / "vendas.xlsx")

quantidade_total = 0
faturamento_total = 0
custo_bruto = 0
lucro_bruto_total = 0

campos = {
    "Quantidade": "invalida",
    "preco_unitario": "inválido",
    "preco_venda": "inválido"
}


for produto, quantidade, preco_unitario, preco_venda  in vendas:

    valores = [quantidade, preco_unitario, preco_venda]

    if not all(isinstance(valor, (int, float)) for valor in valores):
        for campo, valor in zip(campos, valores):
            if not isinstance(valor, (int, float)):
                print(f"\nProduto: {produto} - {campo} {campos[campo]} ({valor}).")
                aba_rejeitados.append([produto, f"{campo} {campos[campo]} ({valor})."])
        continue

    if quantidade <= 0:
        print(f"\nProduto: {produto} - Quantidade inválida ({quantidade}).")
        aba_rejeitados.append([produto, f"Quantidade inválida ({quantidade})"])
        continue

    elif preco_unitario <= 0:
        print(f"\nProduto: {produto} - Preço unitário inválid (R${preco_unitario:.2f}).")
        aba_rejeitados.append([produto, f"Preço unitário inválido (R${preco_unitario:.2f})"])
        continue

    elif preco_venda <= 0:
        print(f"\nProduto: {produto} - Preço de venda inválido (R${preco_venda:.2f}).")
        aba_rejeitados.append([produto, f"Preço de venda inválido (R${preco_venda:.2f})"])
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

for linha in aba_rejeitados.iter_rows(min_row=1, max_row=aba_rejeitados.max_row, min_col=1, max_col=aba_rejeitados.max_column):
    for cell in linha:
        cell.border = borda

planilha.save(pasta_exemplos / "relatorio_vendas.xlsx")