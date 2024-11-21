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
    "FONTE 40A": {"classico": 371.20, "premium": 399.04},
    "FONTE 60A": {"classico": 408.32, "premium": 436.16},
    "FONTE LITE 60A": {"classico": 336.33, "premium": 359.81},
    "FONTE 70A": {"classico": 454.73, "premium": 482.57},
    "FONTE LITE 70A": {"classico": 376.68, "premium": 400.36},
    "FONTE 120A": {"classico": 584.65, "premium": 621.77},
    "FONTE LITE 120A": {"classico": 494.21, "premium": 528.39},
    "FONTE 200A": {"classico": 742.41, "premium": 779.54},
    "FONTE LITE 200A": {"classico": 628.36, "premium": 660.51},
    "FONTE BOB 90A": {"classico": 389.76, "premium": 408.32},
    "FONTE BOB 120A": {"classico": 460.29, "premium": 497.41},
    "FONTE BOB 200A": {"classico": 575.37, "premium": 640.33},
    "FONTE 200A MONO": {"classico": 678.85, "premium": 714.11},
    "CONTROLE K1200": {"classico": 56.79, "premium": 61.43},
    "CONTROLE K600": {"classico": 53.95, "premium": 58.36},
    "CONTROLE REDLINE": {"classico": 84.25, "premium": 93.71},
    "CONTROLE ACQUA": {"classico": 73.22, "premium": 81.57}
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
    if "amplificador" in nome or "processador" in nome or "capa" in nome or "nobreak" in nome or "retificadora" in nome or "multimidia" in nome or "gerenciador" in nome or "suspensao" in nome or "stetsom" in nome or "central" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "OUTROS"})
        return
    
    if "inversor" in nome and ("3000w" in nome or "30" in nome):
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 3000W"})
        return
    if "inversor" in nome and ("1000w" in nome or "10" in nome):
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 1000W"})
        return
    if "inversor" in nome and ("2000w" in nome or "20" in nome):
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "INVERSOR 2000W"})
        return
    
    if ("k600" in nome or "k6" in nome) and "fonte" not in nome and "k1200" not in nome:
    
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "CONTROLE K600"})
        return
        
    if ("k1200" in nome or "k12" in nome) and "fonte" not in nome and "k600" not in nome:
    
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "CONTROLE K1200"})
        return
        
    if ("controle wr" in nome or "wr" in nome or "redline" in nome or "red line" in nome) and "fonte" not in nome:
    
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "CONTROLE REDLINE"})
        return
        
    if ("acqua" in nome or "aqua" in nome or "agua" in nome) and "fonte" not in nome:
    
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "CONTROLE ACQUA"})
        return
    
    
    if "controle" not in nome and "lite" not in nome and "light" not in nome:
        if "40" in nome or "40a" in nome or "40 amperes" in nome or "40amperes" in nome or "36a" in nome or "36" in nome or "36 amperes" in nome or "36amperes" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 40A"})
            return
        
    if "bob" not in nome and ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "50" in nome or "50a" in nome or "50 amperes" in nome or "50amperes" in nome or "50 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 50A"})
            return
            
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome:
        if "60" in nome or "60a" in nome or "60 amperes" in nome or "60amperes" in nome or "60 a" in nome or "-60" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 60A"})
            return
            
    if "bob" not in nome and ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "60" in nome or "60a" in nome or "60 amperes" in nome or "60amperes" in nome or "60 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 60A"})
            return
            
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome:
        if "70" in nome or "70a" in nome or "70 amperes" in nome or "70amperes" in nome or "70 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 70A"})
            return

    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "70" in nome or "70a" in nome or "70 amperes" in nome or "70amperes" in nome or "70 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 70A"})
            return
            
    if "bob" not in nome and  ("lite" in nome or "light" in nome) and "controle" not in nome:
        if "40" in nome or "40a" in nome or "40 amperes" in nome or "40amperes" in nome or "40 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 40A"})
            return
            
    if "lite" not in nome and "light" not in nome  and "controle" not in nome:
        if "90" in nome or "90a" in nome or "90 amperes" in nome or "90amperes" in nome or "90 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE BOB 90A"})
            return
            
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and "lit" not in nome:
        if "120" in nome or "120a" in nome or "120 amperes" in nome or "120amperes" in nome or "120 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 120A"})
            return

    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and "lit" not in nome:
        if "150" in nome or "150a" in nome or "150 amperes" in nome or "150amperes" in nome or "150 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 150A"})
            return
             
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "120" in nome or "120a" in nome or "120 amperes" in nome or "120amperes" in nome or "120 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 120A"})
            return
                
    if "bob" in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and "lit" not in nome:
        if "120" in nome or "120a" in nome or "120 amperes" in nome or "120amperes" in nome or "120 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE BOB 120A"})
            return
                
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and 'mono' not in nome and 'mono' not in nome and 'monovolt' not in nome and '220v' not in nome and "lit" not in nome:
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A"})
            return
                
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome and ("mono" in nome or "220v" in nome or "monovolt" in nome):
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 200A MONO"})
            return
        
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and ("mono" in nome or "220v" in nome or "monovolt" in nome):
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A MONO"})
            return
                
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 200A"})
            return
                
    if "bob" in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and 'mono' not in nome and 'mono' not in nome and 'monovolt' not in nome and '220v' not in nome:
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE BOB 200A"})
            return
        
        
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and 'mono' not in nome and 'mono' not in nome and 'monovolt' not in nome and '220v' not in nome and "lit" not in nome:
        if "20" in nome or "20a" in nome or "20 amperes" in nome or "20amperes" in nome or "20 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A"})
            return
                
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome and ("mono" in nome or "220v" in nome or "monovolt" in nome):
        if "20" in nome or "20a" in nome or "20 amperes" in nome or "20amperes" in nome or "20 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 200A MONO"})
            return
        
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and ("mono" in nome or "220v" in nome or "monovolt" in nome):
        if "20" in nome or "20a" in nome or "20 amperes" in nome or "20amperes" in nome or "20 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A MONO"})
            return
                
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "20" in nome or "20a" in nome or "20 amperes" in nome or "20amperes" in nome or "20 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 200A"})
            return
                
    if "bob" in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and 'mono' not in nome and 'mono' not in nome and 'monovolt' not in nome and '220v' not in nome:
        if "20" in nome or "20a" in nome or "20 amperes" in nome or "20amperes" in nome or "20 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE BOB 200A"})
            return
        
    
    
    if "inversor" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "OUTROS INVERSORES"})
        return
    
    if "fonte" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": identificar_produto(unidecode(item["Tipo de Anúncio"]).lower(), price)})
        return
    
    if "controle" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": identificar_produto(unidecode(item["Tipo de Anúncio"]).lower(), price)})
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

urls = ["JFA", "JFA%20ELETRONICOS"]             
for i in urls:
    response = requests.get(f"https://corp.shoppingdeprecos.com.br/vendedores/exportar_vendas_marca?id={i}&ini={dia_inicial}&fim={dia_final}", headers=headers)

    if response.status_code == 200:  
        #print("resposta ok")
        time.sleep(20)
        with open("produtos.xlsx", 'wb') as file:

            file.write(response.content)

    time.sleep(5)



    db = pd.read_excel("produtos.xlsx", engine='openpyxl')
                    
    for index, item in db.iterrows():
        SelecionarFonte(item)

        
df = pd.DataFrame(items)


df.to_excel("modelos_jfa.xlsx", index=False)