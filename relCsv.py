import pandas as pd

# Lê o arquivo de texto linha por linha e divide cada linha em colunas
with open("./relatorio.txt", "r") as file:
    lines = file.readlines()
    data = [line.strip().split(", ") for line in lines]

# Cria um dataframe pandas com os dados
relatorio = pd.DataFrame(data, columns=["NOME", "CPF/CNPJ", "TELEFONE", "OCORRÊNCIA"])

# Salva o dataframe como um arquivo CSV sem incluir o índice
relatorio.to_csv("./relatorio.csv", index=False)

relatoriocsv = pd.read_csv("./relatorio.csv")

relatoriocsv.to_excel("./relatorio.xlsx")
