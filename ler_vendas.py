from openpyxl import load_workbook



def carregar_vendas(caminho):
    planilha = load_workbook(caminho, data_only=True)
    aba = planilha.active

    vendas = []

    
    for linha in aba.iter_rows(min_row=2, values_only=True):
        vendas.append(linha)
    return vendas

if __name__ == "__main__":
    resultados = carregar_vendas("exemplos/vendas.xlsx")
    print("Total de vendas: ", len(resultados))
    print("Primeira venda: ", resultados[0])

