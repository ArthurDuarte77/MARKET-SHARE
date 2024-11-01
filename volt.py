import argparse
from unidecode import unidecode
from selenium.webdriver.support.ui import Select
import threading
import subprocess
import os
import time
from tqdm import tqdm
import shutil
import json
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.common.exceptions import *
import re
import sys
import numpy as np
import cv2
import requests
from typing import Dict, List

items = []

titulo_arquivo = ""
# options.add_argument("--headless=new")

# Dicionário com produtos e seus preços para cada categoria
produtos = {
    "FONTE 40A": {"classico": 414.87, "premium": 445.99},
    "FONTE 60A": {"classico": 456.36, "premium": 487.48},
    "FONTE LITE 60A": {"classico": 375.9, "premium": 402.14},
    "FONTE 70A": {"classico": 508.22, "premium": 539.34},
    "FONTE LITE 70A": {"classico": 420.99, "premium": 447.46},
    "FONTE 120A": {"classico": 653.43, "premium": 694.92},
    "FONTE LITE 120A": {"classico": 552.93, "premium": 590.56},
    "FONTE 200A": {"classico": 829.76, "premium": 871.25},
    "FONTE LITE 200A": {"classico": 702.29, "premium": 738.22},
    "FONTE BOB 90A": {"classico": 435.62, "premium": 456.36},
    "FONTE BOB 120A": {"classico": 514.45, "premium": 555.93},
    "FONTE BOB 200A": {"classico": 643.06, "premium": 715.66},
    "FONTE 200A MONO": {"classico": 758.71, "premium": 798.64},
    "CONTROLE K1200": {"classico": 63.47, "premium": 68.66},
    "CONTROLE K600": {"classico": 60.29, "premium": 65.29},
    "CONTROLE REDLINE": {"classico": 94.16, "premium": 104.25},
    "CONTROLE ACQUA": {"classico": 81.84, "premium": 91.17}
}

def identificar_produto(tipo, preco):
    tolerancia = 0.05  # Tolerância de 1%
    for produto, precos in produtos.items():
        if tipo.lower() == "classico":
            preco_base = precos["classico"]
        elif tipo.lower() == "premium":
            preco_base = precos["premium"]
        else:
            return "Tipo inválido. Use 'classico' ou 'premium'."
        
        if preco_base * (1 - tolerancia) <= preco <= preco_base * (1 + tolerancia):
            return produto
    return "OUTROS"

if os.path.exists(r"produtos.xlsx"):
    os.remove(r"produtos.xlsx")
if os.path.exists(r"modelos_jfa.xlsx"):
    os.remove(r"modelos_jfa.xlsx")

    


def SelecionarFonte(item):
    nome = unidecode(item["Produto"].strip().lower())
    price = float(item["Preço Unitário"].replace(".", "").replace(",", "."))
    tipo = unidecode(item["Tipo de Anúncio"].strip().lower())
    total = float(item["Total"].replace(".", "").replace(",", "."))
    if "amplificador" in nome or "processador" in nome or "capa" in nome or "multimidia" in nome or "gerenciador" in nome or "suspensao" in nome or "stetsom" in nome or "central" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "OUTROS"})
        return

    if isinstance(item["Vendedor"], int):
        response = requests.get(f"https://api.mercadolibre.com/users/{item['Vendedor']}")
        if response.status_code == 200:
            data = response.json()
            item['Vendedor'] = data.get("nickname", item['Vendedor'])
    
    if "nobreak" in nome or "fonte nobreak" in nome:
        if "620w" in nome and ("48v" in nome or "48" in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 620W 48V 2U"})
            return

        if "250w" in nome and "24v" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 250W 24V 10A"})
            return

        if "200w" in nome and "24v" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 200W 24V 7A"})
            return

        if "200w" in nome and "12v" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 200W 12V 8A"})
            return

        if "200w" in nome and "48v" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 200W 48V 4A"})
            return

        if ("mini max" in nome or "max" in nome or "mini" in nome) and "24v" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MINI MAX 24V"})
            return

        if "mini max" in nome and "13.8v" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MINI MAX 13.8V"})
            return

        if "mini max" in nome and "12v" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MINI MAX 12V"})
            return
        
    items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "OUTROS"})
                

parser = argparse.ArgumentParser(description='Processar datas de início e fim.')
parser.add_argument('--dia_inicial', type=str, required=True, help='Data inicial no formato AAAA-MM-DD')
parser.add_argument('--dia_final', type=str, required=True, help='Data final no formato AAAA-MM-DD')
parser.add_argument('--cookie', type=str, required=True, help='Cookies')

args = parser.parse_args()

dia_inicial = args.dia_inicial
dia_final = args.dia_final
cookie = args.cookie

headers = {
    "Cookie": cookie
}

urls = ["VOLT"]             
for i in urls:
    response = requests.get(f"https://corp.shoppingdeprecos.com.br/vendedores/exportar_vendas_marca?id={i}&ini={dia_inicial}&fim={dia_final}", headers=headers)

    if response.status_code == 200:  
        print("resposta ok")
        time.sleep(20)
        with open("produtos.xlsx", 'wb') as file:

            file.write(response.content)

    time.sleep(5)



    db = pd.read_excel("produtos.xlsx", engine='openpyxl')
                    
    for index, item in db.iterrows():
        SelecionarFonte(item)

        
df = pd.DataFrame(items)


df.to_excel("modelos_volt.xlsx", index=False)