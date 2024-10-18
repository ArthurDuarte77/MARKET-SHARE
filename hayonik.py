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

items = []


titulo_arquivo = ""
# options.add_argument("--headless=new")

if os.path.exists(r"produtos.xlsx"):
    os.remove(r"produtos.xlsx")
if os.path.exists(r"produtos2.xlsx"):
    os.remove(r"produtos2.xlsx")
if os.path.exists(r"modelos_hayonik.xlsx"):
    os.remove(r"modelos_hayonik.xlsx")


def SelecionarFonte(item):
    nome = item["Produto"].strip().lower()
    price = float(item["Preço Unitário"].replace(".", "").replace(",", "."))
    tipo = unidecode(item["Tipo de Anúncio"].strip().lower())
    total = float(item["Total"].replace(".", "").replace(",", "."))
            
    if "kit" in nome or "controle" in nome or "truck"  in nome or "48v" in nome or "48 v" in nome or "fita led" in nome or "máquina" in nome or "fumaça" in nome or "vela" in nome or "refletor" in nome or "moving" in nome or "nauticlin" in nome or "nauticline" in nome or "nautic" in nome or "truck lin" in nome or "tru" in nome or "48v" in nome or "truck line" in nome or "truck" in nome or "fontes 48v" in nome or "fontes 24v" in nome or "32bv" in nome or "32a" in nome or "5v" in nome or "2a" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "OUTROS"})
        return
    if isinstance(item["Vendedor"], int):
        response = requests.get(f"https://api.mercadolibre.com/users/{item['Vendedor']}")
        if response.status_code == 200:
            data = response.json()
            item['Vendedor'] = data.get("nickname", item['Vendedor'])
        
    if "inversor":
        if "1500w" in nome or "1500" in nome or "1.500" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 1500W"})
            return
        elif "800w" in nome or "800" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 800W"})
            return
        elif "500w" in nome or "500" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 500W"})
            return
        elif "300w" in nome or "300" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 300W"})
            return
        elif "600w" in nome or "600" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 600W"})
            return    
        elif "1000w" in nome or "1000" in nome or "1.000" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 1000W"})
            return
        elif "1200w" in nome or "1200" in nome or "1.200" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 1200W"})
            return
        elif "1800w" in nome or "1800" in nome or "1.800" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 1800W"})
            return
        elif "2000w" in nome or "2000" in nome or "2.000" in nome or "200" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 2000W"})
            return
        elif "2500w" in nome or "2500" in nome or "2.500" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 2500W"})
            return
        elif "3000w" in nome or "3000" in nome or "3.000" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 3000W"})
            return
        elif "5000w" in nome or "5000" in nome or "5.000" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 5000W"})
            return
        else:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR"})
            return
    
    items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "OUTROS"})

parser = argparse.ArgumentParser(description='Processar datas de início e fim.')
parser.add_argument('--dia_inicial', type=str, required=True, help='Data inicial no formato AAAA-MM-DD')
parser.add_argument('--dia_final', type=str, required=True, help='Data final no formato AAAA-MM-DD')
parser.add_argument('--cookie', type=str, required=True, help='Cookie')

args = parser.parse_args()


dia_inicial = args.dia_inicial
dia_final = args.dia_final
cookie = args.cookie

headers = {
    "Cookie": cookie
}

urls = ["HAYONIK"]             
for i in urls:
    response = requests.get(f"https://corp.shoppingdeprecos.com.br/vendedores/exportar_vendas_marca?id={i}&ini={dia_inicial}&fim={dia_final}", headers=headers)
    if response.status_code == 200:  

        with open("produtos.xlsx", 'wb') as file:

            file.write(response.content)

    time.sleep(5)



    db = pd.read_excel("produtos.xlsx", engine='openpyxl')
                    
    for index, item in db.iterrows():
        SelecionarFonte(item)

        
df = pd.DataFrame(items)

# Exportar o DataFrame para um arquivo Excel
df.to_excel("modelos_hayonik.xlsx", index=False)