import pandas as pd

# === Dados de Estoque com fornecedores reais ===
estoque = pd.DataFrame({
    "Código": [201, 202, 203, 204, 205, 206, 207, 208],
    "Produto": [
        "Notebook Dell Inspiron", "Monitor LG 24''", "Mouse Logitech M170",
        "Teclado Microsoft", "Impressora HP DeskJet", "Cadeira Gamer ThunderX3",
        "Smartphone Samsung Galaxy A55", "HD Externo Seagate 1TB"
    ],
    "Categoria": [
        "Informática", "Informática", "Periféricos", "Periféricos",
        "Impressão", "Móveis", "Telefonia", "Armazenamento"
    ],
    "Fornecedor": [
        "Dell", "LG Electronics", "Logitech", "Microsoft",
        "HP", "Kabum", "Samsung", "Seagate"
    ],
    "Quantidade": [15, 20, 50, 35, 10, 5, 25, 12],
    "Preço Unitário": [3500.0, 950.0, 80.0, 120.0, 600.0, 1200.0, 2200.0, 400.0]
})
estoque["Valor Total"] = estoque["Quantidade"] * estoque["Preço Unitário"]

# === Dados de Caixa Diário com lançamentos ===
caixa = pd.DataFrame({
    "Data": pd.date_range("2025-11-01", periods=10, freq="D"),
    "Descrição": [
        "Venda Notebook Dell", "Compra Cadeiras", "Venda Monitor LG",
        "Venda Smartphone Samsung", "Compra Impressora HP", "Venda HD Seagate",
        "Venda Mouse Logitech", "Venda Teclado Microsoft", "Compra Estoque Kabum",
        "Venda Cadeira Gamer"
    ],
    "Entrada": [15000, 0, 1900, 4400, 0, 4800, 4000, 4200, 0, 6000],
    "Saída": [0, 6000, 0, 0, 1200, 0, 0, 0, 3500, 0]
})
caixa["Saldo"] = caixa["Entrada"] - caixa["Saída"]
caixa["Saldo Acumulado"] = caixa["Saldo"].cumsum()

# === Criar arquivo Excel ===
with pd.ExcelWriter("Gestao_Empresarial.xlsx", engine="xlsxwriter") as writer:
    estoque.to_excel(writer, sheet_name="Estoque", index=False)
    caixa.to_excel(writer, sheet_name="Caixa Diário", index=False)

    workbook = writer.book
    worksheet_estoque = writer.sheets["Estoque"]
    worksheet_caixa = writer.sheets["Caixa Diário"]

    # Transformar em tabelas estruturadas
    worksheet_estoque.add_table(0, 0, len(estoque), len(estoque.columns)-1,
        {"name": "TabelaEstoque", "columns": [{"header": col} for col in estoque.columns]})
    worksheet_caixa.add_table(0, 0, len(caixa), len(caixa.columns)-1,
        {"name": "TabelaCaixa", "columns": [{"header": col} for col in caixa.columns]})

    # === Dashboard ===
    dashboard = workbook.add_worksheet("Dashboard")
    dashboard.write("A1", "Indicadores")
    dashboard.write("A2", "Valor Total em Estoque")
    dashboard.write_formula("B2", "=SUM(TabelaEstoque[Valor Total])")
    dashboard.write("A3", "Saldo Final em Caixa")
    dashboard.write_formula("B3", "=LOOKUP(2,1/(TabelaCaixa[Saldo Acumulado]<>\"\"),TabelaCaixa[Saldo Acumulado])")
    dashboard.write("A4", "Produtos com Baixo Estoque")
    dashboard.write_formula("B4", "=COUNTIF(TabelaEstoque[Quantidade],\"<10\")")

    # Gráfico de barras - Estoque por Categoria
    chart1 = workbook.add_chart({"type": "column"})
    chart1.add_series({
        "categories": "=TabelaEstoque[Categoria]",
        "values": "=TabelaEstoque[Quantidade]",
        "name": "Estoque por Categoria"
    })
    chart1.set_title({"name": "Estoque por Categoria"})
    dashboard.insert_chart("D2", chart1)

    # Gráfico de linha - Evolução do Caixa
    chart2 = workbook.add_chart({"type": "line"})
    chart2.add_series({
        "categories": "=TabelaCaixa[Data]",
        "values": "=TabelaCaixa[Saldo Acumulado]",
        "name": "Evolução do Caixa"
    })
    chart2.set_title({"name": "Evolução do Caixa"})
    dashboard.insert_chart("D20", chart2)

    # Gráfico de pizza - Distribuição por Fornecedor
    chart3 = workbook.add_chart({"type": "pie"})
    chart3.add_series({
        "categories": "=TabelaEstoque[Fornecedor]",
        "values": "=TabelaEstoque[Valor Total]",
        "name": "Distribuição por Fornecedor"
    })
    chart3.set_title({"name": "Distribuição por Fornecedor"})
    dashboard.insert_chart("H2", chart3)

    # Gráfico de barras - Entradas vs Saídas
    chart4 = workbook.add_chart({"type": "column"})
    chart4.add_series({
        "name": "Entradas",
        "categories": "=TabelaCaixa[Data]",
        "values": "=TabelaCaixa[Entrada]"
    })
    chart4.add_series({
        "name": "Saídas",
        "categories": "=TabelaCaixa[Data]",
        "values": "=TabelaCaixa[Saída]"
    })
    chart4.set_title({"name": "Entradas vs Saídas"})
    dashboard.insert_chart("H20", chart4)

    # === Tabelas Dinâmicas ===
    pivot_sheet = workbook.add_worksheet("Tabelas Dinâmicas")
    pivot_sheet.write("A1", "Por Fornecedor")
    fornecedor_pivot = estoque.groupby("Fornecedor")["Valor Total"].sum().reset_index()
    for i, row in fornecedor_pivot.iterrows():
        pivot_sheet.write(i+2, 0, row["Fornecedor"])
        pivot_sheet.write(i+2, 1, row["Valor Total"])

    pivot_sheet.write("A10", "Por Categoria")
    categoria_pivot = estoque.groupby("Categoria")["Quantidade"].sum().reset_index()
    for i, row in categoria_pivot.iterrows():
        pivot_sheet.write(i+11, 0, row["Categoria"])
        pivot_sheet.write(i+11, 1, row["Quantidade"])
