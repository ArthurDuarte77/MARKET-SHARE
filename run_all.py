from tqdm import tqdm
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter import messagebox
import subprocess
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
from datetime import datetime, timedelta
import time

# Função para remover arquivo de resultado, se existir
def remover_arquivo_resultado():
    if os.path.exists("resultado_final.xlsx"):
        os.remove("resultado_final.xlsx")

remover_arquivo_resultado()

# Função para juntar planilhas
def juntar_planilhas():
    arquivos = [
        "modelos_jfa.xlsx",
        "modelos_stetson.xlsx",
        "modelos_taramps.xlsx",
        "modelos_epever.xlsx",
        "modelos_hayonik.xlsx",
        "modelos_usina.xlsx",
        "modelos_volt.xlsx",
        "modelos_tataliken.xlsx",
        "modelos_knup.xlsx",
        "modelos_amfer.xlsx",
    ]
    
    data_hoje = datetime.now().strftime('%Y-%m-%d')
    df_juntado = pd.DataFrame()
    
    for arquivo in arquivos:
        if os.path.exists(arquivo):
            df = pd.read_excel(arquivo)
            df['Data'] = data_hoje
            df_juntado = pd.concat([df_juntado, df])
    
    df_juntado.to_excel('resultado_final.xlsx', index=False)

# Função para executar o script principal
import json  # Add this import at the top of your script

def chamar_script(dia_inicial, dia_final, cookie):
    janela.destroy()
    dia_inicial = datetime.strptime(dia_inicial, '%Y-%m-%d')
    dia_final = datetime.strptime(dia_final, '%Y-%m-%d')
    
    # Prepare the selected brands
    selected = []
    if var_jfa.get():
        selected.append("JFA")
    if var_usina.get():
        selected.append("Usina")
    if var_taramps.get():
        selected.append("Taramps")
    if var_amfer.get():
        selected.append("Amfer")
    if var_hayonik.get():
        selected.append("Hayonik")
    if var_knup.get():
        selected.append("Knup")
    if var_stetson.get():
        selected.append("Stetson")
    if var_volt.get():
        selected.append("Volt")

    # Serialize the list to a JSON string
    selected_json = json.dumps(selected)
    
    # Check if "Por Dia" is selected
    if var_por_dia.get():
        total_days = (dia_final - dia_inicial).days + 1
        for _ in tqdm(range(total_days), desc="Processing dates:", unit="day"):
            formatted_date = dia_inicial.strftime('%Y-%m-%d')
            comando = [
                'python',
                'main.py',
                '--dia_inicial', formatted_date,
                '--dia_final', formatted_date,
                '--cookie', cookie,
                '--escolha', selected_json
            ]
            subprocess.run(comando)
            dia_inicial += timedelta(days=1)
    else:
        # Execute the script once with the initial and final dates
        comando = [
            'python',
            'main.py',
            '--dia_inicial', dia_inicial.strftime('%Y-%m-%d'),
            '--dia_final', dia_final.strftime('%Y-%m-%d'),
            '--cookie', cookie,
            '--escolha', selected_json
        ]
        subprocess.run(comando)

    messagebox.showinfo("Conclusão", "Todos os scripts foram executados com sucesso!")

# Função de login com Selenium
def configurar_driver():
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

    return "".join(cookies_list)

    time.sleep(2)
    
    driver.quit()


# Inicialização da GUI
janela = tk.Tk()
janela.title('Market Share')

# Variable for the "Por Dia" option
var_por_dia = tk.BooleanVar()

# Create and position the "Por Dia" checkbox
check_por_dia = tk.Checkbutton(janela, text="Por Dia", variable=var_por_dia)
check_por_dia.grid(column=0, row=0, padx=5, pady=5, sticky="w")

data_atual = datetime.now()
ttk.Label(janela, text='Data Inicial:').grid(column=0, row=1, padx=10, pady=10)
cal_inicial = DateEntry(janela, width=22, background='darkblue', foreground='white', borderwidth=2, locale='pt_BR', day=data_atual.day - 1)
cal_inicial.grid(column=1, row=1, padx=10, pady=10)

ttk.Label(janela, text='Data Final:').grid(column=0, row=2, padx=10, pady=10)
cal_final = DateEntry(janela, width=22, background='darkblue', foreground='white', borderwidth=2, locale='pt_BR', day=data_atual.day - 1)
cal_final.grid(column=1, row=2, padx=10, pady=10)

# Variáveis para checkboxes
var_jfa = tk.BooleanVar()
var_usina = tk.BooleanVar()
var_taramps = tk.BooleanVar()
var_amfer = tk.BooleanVar()
var_hayonik = tk.BooleanVar()
var_knup = tk.BooleanVar()
var_stetson = tk.BooleanVar()
var_volt = tk.BooleanVar()

# Criação dos checkboxes
check_jfa = tk.Checkbutton(janela, text="JFA", variable=var_jfa)
check_usina = tk.Checkbutton(janela, text="Usina", variable=var_usina)
check_taramps = tk.Checkbutton(janela, text="Taramps", variable=var_taramps)
check_amfer = tk.Checkbutton(janela, text="Amfer", variable=var_amfer)
check_hayonik = tk.Checkbutton(janela, text="Hayonik", variable=var_hayonik)
check_knup = tk.Checkbutton(janela, text="Epever", variable=var_knup)
check_stetson = tk.Checkbutton(janela, text="Stetson", variable=var_stetson)
check_volt = tk.Checkbutton(janela, text="Volt", variable=var_volt)

# Posicionamento dos checkboxes
check_jfa.grid(column=0, row=4, padx=5, pady=5, sticky="w")
check_usina.grid(column=0, row=5, padx=5, pady=5, sticky="w")
check_taramps.grid(column=0, row=6, padx=5, pady=5, sticky="w")
check_amfer.grid(column=0, row=7, padx=5, pady=5, sticky="w")
check_hayonik.grid(column=0, row=8, padx=5, pady=5, sticky="w")
check_knup.grid(column=0, row=9, padx=5, pady=5, sticky="w")
check_stetson.grid(column=0, row=10, padx=5, pady=5, sticky="w")
check_volt.grid(column=0, row=11, padx=5, pady=5, sticky="w")

# Botão para executar o script
ttk.Button(janela, text='Executar', command=lambda: chamar_script(
    cal_inicial.get_date().strftime('%Y-%m-%d'),
    cal_final.get_date().strftime('%Y-%m-%d'),
    configurar_driver()
)).grid(column=0, row=12, columnspan=2, pady=10)

janela.protocol("WM_DELETE_WINDOW", lambda: janela.quit())
janela.mainloop()