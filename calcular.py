def calcular_faturamento(quantidade, preco_venda):
    return quantidade * preco_venda

def calcular_custo(quantidade, custo_unitario):
    return quantidade * custo_unitario

def calcular_lucro(faturamento, custo_total):
    return faturamento - custo_total

def calcular_margem_lucro(lucro, faturamento):
    if faturamento == 0:
        return 0
    return (lucro / faturamento) * 100