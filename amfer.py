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

if os.path.exists(r"produtos.xlsx"):
    os.remove(r"produtos.xlsx")
if os.path.exists(r"produtos2.xlsx"):
    os.remove(r"produtos2.xlsx")

    


def SelecionarFonte(item):
    nome = unidecode(item["Produto"].strip().lower())
    price = float(item["Preço Unitário"].replace(".", "").replace(",", "."))
    tipo = unidecode(item["Tipo de Anúncio"].strip().lower())
    total = float(item["Total"].replace(".", "").replace(",", "."))
    if "amplificador" in nome or "processador" in nome or "capa" in nome or "multimidia" in nome or "gerenciador" in nome or "suspensao" in nome or "stetsom" in nome or "central" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "OUTROS"})
        return
    
    if "nobreak" in nome:
        if "12v" in nome and "10a" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK FX250 12V/10A"})
            return
        elif "24v" in nome and "8a" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK FX250 24/8A"})
            return
        elif "48v" in nome and "4a" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK FX250 48v/4a"})
            return
        elif "12v" in nome and "15a" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK FX300 12V/15A"})
            return
        elif "24v" in nome and "10a" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK FX300 24/10A"})
            return
        elif "48v" in nome and "6a" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK FX300 48v/6a"})
            return
        elif "12v" in nome and "20a" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK FX500 12v/20a"})
            return
        elif "24v" in nome and "15a" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK FX500 24v/15a"})
            return
        elif "24v" in nome and "20a" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK FX500 24v/20a"})
            return
        elif "48v" in nome and "10a" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK FX500 48v/10a"})
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

urls = ["AMFER TELECOM"]             
for i in urls:
    response = requests.get(f"https://corp.shoppingdeprecos.com.br/vendedores/exportar_vendas_marca?id={i}&ini={dia_inicial}&fim={dia_final}", headers=headers)

    if response.status_code == 200:  
        # print("resposta ok")
        time.sleep(20)
        with open("produtos.xlsx", 'wb') as file:

            file.write(response.content)

    time.sleep(5)



    db = pd.read_excel("produtos.xlsx", engine='openpyxl')
                    
    for index, item in db.iterrows():
        SelecionarFonte(item)

        
df = pd.DataFrame(items)


df.to_excel("modelos_amfer.xlsx", index=False)