import pandas as pd

def pegarCpf():

    teste = pd.read_excel("relatorio (21).xlsx", sheet_name=0, header=None)
    coluna = teste.iloc[:, [4, 8]]
    coluna.to_csv("cobraAtrasoGTS/colunas_extraidas.csv", index=False, header=False)

pegarCpf()        