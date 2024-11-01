import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter import messagebox
import subprocess
import argparse
from datetime import datetime, timedelta
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.common.exceptions import *
import time
import json


def juntar_planilhas():
    # Lista de arquivos a serem juntados
    arquivos = [
        "modelos_jfa.xlsx",
        "modelos_stetson.xlsx",
        "modelos_taramps.xlsx",
        "modelos_epever.xlsx",
        "modelos_hayonik.xlsx",
        "modelos_usina.xlsx",
        "modelos_volt.xlsx",
    ]
    
    # Data de hoje
    data_hoje = datetime.now().strftime('%Y-%m-%d')
    
    # Inicializar dataframe vazio
    df_juntado = pd.DataFrame()
    
    # Juntar planilhas
    for arquivo in arquivos:
        if os.path.exists(arquivo):
            df = pd.read_excel(arquivo)
            df['Data'] = data_hoje
            df_juntado = pd.concat([df_juntado, df])
    
    # Salvar planilha juntada
    df_juntado.to_excel('resultado_final.xlsx', index=False)

def chamar_script(dia_inicial, dia_final, cookie):
    arquivos = [
        "modelos_jfa.xlsx",
        "modelos_stetson.xlsx",
        "modelos_taramps.xlsx",
        "modelos_epever.xlsx",
        "modelos_hayonik.xlsx",
        "modelos_usina.xlsx",
        "modelos_volt.xlsx",
        "produtos.xlsx"
    ]
    
    for arquivo in arquivos:
        if os.path.exists(arquivo):
            os.remove(arquivo)
    
    scripts = ['jfa.py', 'usina.py', 'hayonik.py', 'epever.py', 'stetson.py', 'taramps.py']#'jfa.py', 'usina.py', 'hayonik.py', 'epever.py', 'stetson.py', 'taramps.py',
    
    dia_inicial = cal_inicial.get_date().strftime('%Y-%m-%d')
    dia_final = cal_final.get_date().strftime('%Y-%m-%d')
    janela.destroy()
    
    for script in scripts:
        comando = [
            'python',
            script,
            '--dia_inicial', dia_inicial,
            '--dia_final', dia_final,
            '--cookie', cookie
        ]
        subprocess.run(comando)
    messagebox.showinfo("Conclusão", "Todos os scripts foram executados com sucesso!")
    juntar_planilhas()

parser = argparse.ArgumentParser(description='Executar scripts com datas específicas.')
parser.add_argument('--dia_inicial', type=str, help='Data inicial no formato YYYY-MM-DD')
parser.add_argument('--dia_final', type=str, help='Data final no formato YYYY-MM-DD')
parser.add_argument('--cookie', type=str, help='Cookie')

args = parser.parse_args()

if not args.dia_inicial or not args.dia_final:
    data_anterior = datetime.now() - timedelta(days=1)
    dia_padrao = data_anterior.strftime('%Y-%m-%d')
    dia_inicial = args.dia_inicial or dia_padrao
    dia_final = args.dia_final or dia_padrao
    cookie = args.cookie

service = Service()
options = webdriver.ChromeOptions()
options.add_argument("--disable-gpu")
options.add_argument("--disable-extensions")
prefs = {"profile.managed_default_content_settings.images": 2}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=service, options=options)
driver.get("https://www.google.com.br/?hl=pt-BR")
time.sleep(3)
try:
    driver.get("https://corp.shoppingdeprecos.com.br/login")
    counter = 0
    while True:
        test = driver.find_elements(By.XPATH, '//*[@id="email"]')
        if test:
            break
        else:
            counter += 1
            if counter > 20:
                break
            time.sleep(0.5)
    driver.find_element(By.XPATH, '//*[@id="email"]').send_keys("loja@jfaeletronicos.com")
    driver.find_element(By.XPATH, '//*[@id="password"]').send_keys("922982PC")
    driver.find_element(By.XPATH, '//*[@id="btnLogin"]').click()
except TimeoutException as e:
    print(f"Timeout ao tentar carregar a página ou encontrar um elemento: {e}")
except NoSuchElementException as e:
    print(f"Elemento não encontrado na página: {e}")
except WebDriverException as e:
    print(f"Erro no WebDriver: {e}")

time.sleep(3)
driver.get("https://corp.shoppingdeprecos.com.br/vendedores/vendasMarca")

cookies_list = []

cookies = driver.get_cookies()
for cookie in cookies:
    objeto = cookie['name']
    value = cookie['value']
    cookies_list.append(f"{objeto}={value};")

cookie = "".join(cookies_list)

janela = tk.Tk()
janela.title('Market Share')
data_atual = datetime.now()
ttk.Label(janela, text='Data Inicial:').grid(column=0, row=0, padx=10, pady=10)
cal_inicial = DateEntry(janela, width=22, background='darkblue', foreground='white', borderwidth=2, locale='pt_BR', day=data_atual.day - 1)
cal_inicial.grid(column=1, row=0, padx=10, pady=10)

ttk.Label(janela, text='Data Final:').grid(column=0, row=1, padx=10, pady=10)
cal_final = DateEntry(janela, width=22, background='darkblue', foreground='white', borderwidth=2, locale='pt_BR', day=data_atual.day - 1)
cal_final.grid(column=1, row=1, padx=10, pady=10)

ttk.Button(janela, text='Executar', command=lambda: chamar_script("", "", cookie)).grid(column=0, row=2, columnspan=2, pady=10)

janela.protocol("WM_DELETE_WINDOW", lambda: janela.quit())
janela.mainloop()