import pandas as pd
import requests

# Defina o caminho para sua planilha e a coluna que deseja ler
excel_file = 'integracao.xlsx'
coluna = 'A'  # Altere para a coluna que deseja ler (pode ser um número ou letra da coluna)
colunaB = 'B'  # Essa é padrão pra marcar o Check

# Leia a planilha
df = pd.read_excel(excel_file, usecols=[coluna])

# Pegar o Token da conta do Richard 118
token = "64a0beb0c17bf04c.37f980bcfeed7ba8.cc7e3fdee76c265" 

# Itere sobre cada linha na coluna especificada
try:
    for index, row in df.iterrows():
        print("Iniciando publicação!")
        dado = str(row[coluna])
        print(f"Produto: [ {dado} ] - Index [ {index} ]")
        
        checkColuna = "OK"
        
        response = 200

        # Defina os parâmetros e o corpo da requisição
        url = f"https://www.melistock.com/recycleserver/ControlFunctions?Context=Integracao&FunctionSel=PublicarIntegracoes&idEmpresa=118&VaaptUserId=294&strTokenOperation={token}"
        headers = {
            'Content-Type': 'application/json',
        }
        body = {
            "idPeca": dado,
            "idEmpresaCliente": 300,
            "idVaaptUserCliente": 708,
            "arrIntegracoes": [
                {
                "name":"OLX",
                "integracao":True,
                "price":0.0
                },
                {
                "name":"Shopee",
                "integracao":False,
                "price":0.0
                },
                {
                "name":"PecaVerde",
                "integracao":False,
                "price":0.0
                }
            ]
        }
        # Executando requisição do lambda
        response = requests.post(url, headers=headers, json=body)
        
        if response.status_code == 200:
            
            # Atualize o valor da celula atual na Coluna B
            df.at[index, colunaB] = checkColuna

            print(response)
            print("Produto Publicado")
        else:
            print(f'Falha para o valor: {dado} - Status: {response.status_code}')
            print(response)
            print("Produto falhou na publicacao")
    
    df.to_excel('integracao.xlsx', index=False)
except Exception as e:
    df.to_excel('integracao.xlsx', index=False)
