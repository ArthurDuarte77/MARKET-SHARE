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


def SelecionarFonte(item):
    nome = item["Produto"].strip().lower()
    price = float(item["Preço Unitário"].replace(".", "").replace(",", "."))
    tipo = unidecode(item["Tipo de Anúncio"].strip().lower())
    total = float(item["Total"].replace(".", "").replace(",", "."))
        
    if "inversor":
        if "500w" in nome and ("24v" in nome or "24" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 500W 24V SENOIDAL PURA"})
            return
        if "500w" in nome and ("12v" in nome or "12" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 500W 12V SENOIDAL PURA"})
            return
        if "1000w" in nome and ("24v" in nome or "24" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 1000W 24V SENOIDAL PURA"})
            return
        if "1000w" in nome and ("12v" in nome or "12" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 1000W 12V SENOIDAL PURA"})
            return
        if "1500w" in nome and ("24v" in nome or "24" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 1500W 24V SENOIDAL PURA"})
            return
        if "1500w" in nome and ("12v" in nome or "12" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 1500W 12V SENOIDAL PURA"})
            return
        if "2000w" in nome and ("24v" in nome or "24" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 2000W 24V SENOIDAL PURA"})
            return
        if "2000w" in nome and ("12v" in nome or "12" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 2000W 12V SENOIDAL PURA"})
            return
        if "3000w" in nome and ("24v" in nome or "24" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 3000W 24V SENOIDAL PURA"})
            return
        if "3000w" in nome and ("12v" in nome or "12" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 3000W 12V SENOIDAL PURA"})
            return
        if "4000w" in nome and ("24v" in nome or "24" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 4000W 24V SENOIDAL PURA"})
            return
        if "4000w" in nome and ("12v" in nome or "12" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 4000W 12V SENOIDAL PURA"})
            return
        if "5000w" in nome and ("24v" in nome or "24" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 5000W 24V SENOIDAL PURA"})
            return
        if "5000w" in nome and ("12v" in nome or "12" in nome) and ("senoidal" in nome or 'pura' in nome or 'onda sen' in nome):
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 5000W 12V SENOIDAL PURA"})
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

urls = ["EPEVER"]             
for i in urls:
    response = requests.get(f"https://corp.shoppingdeprecos.com.br/vendedores/exportar_vendas_marca?id={i}&ini={dia_inicial}&fim={dia_final}", headers=headers)
    if response.status_code == 200:  
        #print("alright")
        
        with open("produtos.xlsx", 'wb') as file:

            file.write(response.content)

    time.sleep(5)



    db = pd.read_excel("produtos.xlsx", engine='openpyxl')
                    
    for index, item in db.iterrows():
        SelecionarFonte(item)

        
df = pd.DataFrame(items)
# Exportar o DataFrame para um arquivo Excel
df.to_excel("modelos_knup.xlsx", index=False)