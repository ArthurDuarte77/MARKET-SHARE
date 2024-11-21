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
        if "200w" in nome and "rack" not in nome:
            if "12v" in nome and "8a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 200W 12v/8a"})
                return
            elif "24v" in nome and "7a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 200W 24v/7a"})
                return
            elif "48v" in nome and "4a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 200W 48v/4a"})
                return
            elif "-48v" in nome and "4a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 200W -48v/4a"})
                return
        elif "200w" in nome and ("rack" in nome or "1u" in nome):
            if "12v" in nome and "8a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 200W 12v/8a 1U"})
                return
            elif "24v" in nome and "7a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 200W 24v/7a 1U"})
                return
            elif "48v" in nome and "4a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 200W 48v/4a 1U"})
                return
            elif "-48v" in nome and "4a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 200W -48v/4a 1U"})
                return
        elif "250w" in nome:
            if "12v" in nome and "9a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 250W 12v/9a"})
                return
            elif "24v" in nome and ("9a" in nome or "10a" in nome):
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 250W 24v/9a"})
                return
            elif "48v" in nome and "4a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 250W 48v/4a"})
                return
            elif "-48v" in nome and "4a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 250W -48v/4a"})
                return
        elif "250w" in nome and "plus" in nome:
            if "12v" in nome and "9a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 250W 12v/9a PLUS"})
                return
            elif "24v" in nome and ("9a" in nome or "10a" in nome):
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 250W 24v/9a PLUS"})
                return
            elif "48v" in nome and "4a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 250W 48v/4a PLUS"})
                return
            elif "-48v" in nome and "4a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 250W -48v/4a PLUS"})
                return
        elif "380w" in nome:
            if "12v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 380W 12v/10a"})
                return
            elif "24v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 380W 24v/10a"})
                return
            elif "48v" in nome and "5a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 380W 48v/5a"})
                return
            elif "-48v" in nome and "5a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 380W -48v/5a"})
                return
        elif "380w" in nome and "gerenciavel" in nome:
            if "12v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 380W 12v/10a GERENCIAVEL"})
                return
            elif "24v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 380W 24v/10a GERENCIAVEL"})
                return
            elif "48v" in nome and "5a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 380W 48v/5a GERENCIAVEL"})
                return
            elif "-48v" in nome and "5a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 380W -48v/5a GERENCIAVEL"})
                return
        elif "520w" in nome and "gerenciavel" in nome:
            if "12v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 520W 12v/10a GERENCIAVEL"})
                return
            elif "24v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 520W 24v/10a GERENCIAVEL"})
                return
            elif "48v" in nome and "5a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 520W 48v/5a GERENCIAVEL"})
                return
            elif "-48v" in nome and "5a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 520W -48v/5a GERENCIAVEL"})
                return
        elif "520w" in nome:
            if "12v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 520W 12v/10a"})
                return
            elif "24v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 520W 24v/10a"})
                return
            elif "48v" in nome and "5a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 520W 48v/5a"})
                return
            elif "-48v" in nome and "5a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 520W -48v/5a"})
                return
        elif "620w" in nome:
            if "12v" in nome and "20a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 620W 12v/20a"})
                return
            elif "24v" in nome and "20a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 620W 24v/20a"})
                return
            elif "48v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 620W 48v/10a"})
                return
            elif "-48v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 620W -48v/10a"})
                return
        elif "620w" in nome and "2u" in nome:
            if "12v" in nome and "20a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 620W 12v/20a 2U"})
                return
            elif "24v" in nome and "20a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 620W 24v/20a 2U"})
                return
            elif "48v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 620W 48v/10a 2U"})
                return
            elif "-48v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 620W -48v/10a 2U"})
                return
        elif "1000w" in nome:
            if "12v" in nome and "30a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 1000W 12v/30a"})
                return
            elif "12v" in nome and "15a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 1000W 12v/15a"})
                return
            elif "12v" in nome and "45a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 1000W 12v/45a"})
                return
            elif "24v" in nome and "40a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 1000W 24v/40a"})
                return
            elif "24v" in nome and "30a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 1000W 24v/30a"})
                return
            elif "24v" in nome and "20a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 1000W 24v/20a"})
                return
            elif "24v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 1000W 24v/10a"})
                return
            elif "48v" in nome and "20a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 1000W 48v/20a"})
                return
            elif "48v" in nome and "15a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 1000W 48v/15a"})
                return
            elif "48v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 1000W 48v/10a"})
                return
            elif "48v" in nome and "5a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 1000W 48v/5a"})
                return
            elif "-48v" in nome and "15a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 1000W -48v/15a"})
                return
        elif "2000w" in nome:
            if "48v" in nome and "40a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 2000W 48v/40a"})
                return
            elif "48v" in nome and "30a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 2000W 48v/30a"})
                return
            elif "48v" in nome and "20a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 2000W 48v/20a"})
                return
            elif "48v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 2000W 48v/10a"})
                return
            elif "-48v" in nome and "30a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 2000W -48v/30a"})
                return
            elif "-48v" in nome and "20a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 2000W -48v/20a"})
                return
            elif "-48v" in nome and "10a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK 2000W -48v/10a"})
                return
        elif "mini max" in nome:
            if ("13,8v" in nome or "13.8v" in nome) and "2a" in nome and "p4" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MINI MAX P4 13,8v/2a"})
                return
            elif ("13,8v" in nome or "13.8v" in nome) and "2a" in nome and "poe" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MINI MAX P4 13,8v/2a"})
                return
            elif ("13,8v" in nome or "13.8v" in nome) and "2a" in nome and "p4" in nome and ("2 saidas" in nome or "2 cabos" in nome):
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MINI MAX 2 SAIDAS P4 13,8v/2a"})
                return
            elif ("13,8v" in nome or "13.8v" in nome) and "5a" in nome and "p4" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MINI MAX P4 13,8v/5a"})
                return
            elif ("13,8v" in nome or "13.8v" in nome) and "5a" in nome and "p4" in nome and ("3 saidas" in nome or "3 cabos" in nome):
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MINI MAX 3 SAIDAS P4 13,8v/5a"})
                return
            elif "27,5v" in nome and "3a" in nome and "p4" in nome and ("3 saidas" in nome or "3 cabos" in nome):
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MINI MAX 3 SAIDAS P4 27,5v/3a"})
                return
        elif "max energy" in nome:
            if "12v" in nome and "2a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MAX ENERGY 12v/2a"})
                return
            elif "24v" in nome and "1a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MAX ENERGY 24v/1a"})
                return
        elif "mini max duo" in nome and "gerenciavel" in nome:
            if ("13,8v" in nome or "13.8v" in nome) and "5a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MINI MAX DUO GERENCIAVEL 13,8v/5a"})
                return
            if "27,5v" in nome and "3a" in nome:
                items.append({"Vendedor": item["Vendedor"], "Produto": nome, "Marca": item["Marca"], "Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE NOBREAK MINI MAX DUO GERENCIAVEL 27,5v/3a"})
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
        # print("resposta ok")
        time.sleep(20)
        with open("produtos.xlsx", 'wb') as file:

            file.write(response.content)

    time.sleep(5)



    db = pd.read_excel("produtos.xlsx", engine='openpyxl')
                    
    for index, item in db.iterrows():
        SelecionarFonte(item)

        
df = pd.DataFrame(items)


df.to_excel("modelos_volt.xlsx", index=False)