import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split, cross_val_score, StratifiedKFold
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.preprocessing import StandardScaler, LabelEncoder
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import classification_report, accuracy_score
from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.utils.class_weight import compute_class_weight
import joblib  # Para salvar e carregar o modelo
import requests
import time
import argparse
import os

parser = argparse.ArgumentParser(description='Processar datas de início e fim.')
parser.add_argument('--dia_inicial', type=str, required=True, help='Data inicial no formato AAAA-MM-DD')
parser.add_argument('--dia_final', type=str, required=True, help='Data final no formato AAAA-MM-DD')
parser.add_argument('--cookie', type=str, required=True, help='Cookies')
args = parser.parse_args()

dia_inicial = args.dia_inicial
dia_final = args.dia_final
cookie = args.cookie

if os.path.exists(r"produtos.xlsx"):
    os.remove(r"produtos.xlsx")
if os.path.exists(r"produtos2.xlsx"):
    os.remove(r"produtos2.xlsx")

headers = {
    "Cookie": cookie
}

urls = ["JFA", "JFA%20ELETRONICOS"]       

all_dados = pd.DataFrame()  # Inicializar o DataFrame all_dados

for i in urls:
    response = requests.get(f"https://corp.shoppingdeprecos.com.br/vendedores/exportar_vendas_marca?id={i}&ini={dia_inicial}&fim={dia_final}", headers=headers)

    if response.status_code == 200:  
        with open("produtos.xlsx", 'wb') as file:
            file.write(response.content)

    time.sleep(5)

    novos_dados = pd.read_excel("produtos.xlsx", engine='openpyxl')
    novos_dados['Preço Unitário'] = novos_dados['Preço Unitário'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
    # Carregar o pipeline treinado
    pipeline_carregado = joblib.load('modelo_treinado.pkl')

    # Fazer previsões nos novos dados
    previsoes = pipeline_carregado.predict(novos_dados)

    # Carregar o label encoder
    label_encoder_carregado = joblib.load('label_encoder.pkl')

    # Decodificar as previsões para obter os nomes das classes
    nomes_classes = label_encoder_carregado.inverse_transform(previsoes)
        
    
    # Adicionar as previsões ao DataFrame original
    novos_dados['Produto2'] = nomes_classes
    

    # Juntar os novos dados ao DataFrame original
    all_dados = pd.concat([all_dados, novos_dados])
    
all_dados = all_dados[["Vendedor", "Produto", "Marca", "Frete Grátis", "Qtde", "Preço Unitário", "Total", "Produto2"]]
all_dados.to_excel("modelos_jfa.xlsx", index=False)



df = pd.read_excel("modelos_jfa.xlsx")

for index, row in df.iterrows():  # Iterate through rows with index
    nome = row["Produto"].strip().lower()

    if "inversor" in nome:
        if "1000w" in nome and "24v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 1000W 24V SENOIDAL PURA'  
        elif "1000w" in nome and "12v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 1000W 12V SENOIDAL PURA'
        elif "1500w" in nome and "24v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 1500W 24V SENOIDAL PURA'  
        elif "1500w" in nome and "12v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 1500W 12V SENOIDAL PURA'
        elif "2000w" in nome and "24v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 2000W 24V SENOIDAL PURA'
            continue
        elif "2000w" in nome and "12v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 2000W 12V SENOIDAL PURA'
            continue
        elif "3500w" in nome and "12v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 3500W 12V SENOIDAL PURA'
            continue
        elif "3500w" in nome and "24v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 3500W 24V SENOIDAL PURA'
            continue
        elif "4000w" in nome and "24v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 4000W 24V SENOIDAL PURA'
            continue
        elif "4000w" in nome and "12v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 4000W 12V SENOIDAL PURA'
            continue
        elif "3000w" in nome and "12v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 3000W 12V SENOIDAL PURA'
            continue
        elif "3000w" in nome and "24v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 3000W 24V SENOIDAL PURA'
            continue
        elif "6000w" in nome and "12v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 6000W 12V SENOIDAL PURA'
            continue
        elif "6000w" in nome and "24v" in nome and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            df.loc[index, 'Produto2'] = 'INVERSOR 6000W 24V SENOIDAL PURA'
            continue

