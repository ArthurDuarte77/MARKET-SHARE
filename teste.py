
import requests
import pandas as pd
import unidecode
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.common.exceptions import *
import subprocess
import argparse
import time
from datetime import datetime, timedelta



def processar_planilha(df, url, dia_inicial):
    headers = {
        'Content-Type': 'application/json'
    }

    for index, row in df.iterrows():
        dados = row.to_dict()
        frete_gratis = True if dados.get('Frete Grátis') == 'Sim' else False
        
        if dados.get('Produto2') not in ['POSSIVEL FONTE', 'POSSIVEL CONTROLE']:
            payload = json.dumps({
                'seller': dados.get('Vendedor'),
                'product': dados.get('Produto'),
                'brand': dados.get('Marca'),
                'freeShipping': frete_gratis,
                'quantity': dados.get('Qtde'),
                'unitPrice': dados.get('Preço Unitário'),
                'total': dados.get('Total'),
                'model': dados.get('Produto2'),
                'date': dia_inicial
            })

            try:
                resposta = requests.post(url, data=payload, headers=headers)
                print(f"Status code para linha {index + 1}: {resposta.status_code}")
            except requests.exceptions.RequestException as e:
                print(f"Erro ao fazer a requisição para a linha {index + 1}: {e}")

service = Service()
options = webdriver.ChromeOptions()
titulo_arquivo = ""
# options.add_argument("--headless=new")
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
                break;
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

def chamar_script(dia_inicial, dia_final, script_name):
    comando = [
        'python',
        script_name,
        '--dia_inicial', dia_inicial,
        '--dia_final', dia_final,
        '--cookie', cookie
    ]
    
    resultado = subprocess.run(comando, capture_output=True, text=True)
    return resultado

    
for i in range(1, 2):
    
    dia_inicial = f"2024-10-{str(i).zfill(2)}"
    dia_final = f"2024-10-{str(i).zfill(2)}"
    dia_inicial_script = f"2024-10-{str(i).zfill(2)}"
    dia_final_script = f"2024-10-{str(i).zfill(2)}"
    
    
    chamar_script(dia_inicial_script, dia_final_script, 'taramps.py')

    # Chamar o segundo script (Usina)
    chamar_script(dia_inicial_script, dia_final_script, 'usina.py')

    # Chamar o terceiro script (Stetsom)
    chamar_script(dia_inicial_script, dia_final_script, 'stetson.py')

    # Chamar o quarto script (JFA)
    chamar_script(dia_inicial_script, dia_final_script, 'jfa.py')

    # Processar a planilha gerada pelo script Taramps
    df_taramps = pd.read_excel('modelos_taramps.xlsx')
    processar_planilha(df_taramps, 'https://expertinvest.com.br/api/v1/taramps', dia_inicial[:-2] + str(int(dia_inicial[-2:]) + 1).zfill(2))

    # Processar a planilha gerada pelo script Usina
    df_usina = pd.read_excel('modelos_usina.xlsx')
    processar_planilha(df_usina, 'https://expertinvest.com.br/api/v1/usina', dia_inicial[:-2] + str(int(dia_inicial[-2:]) + 1).zfill(2))

    # Processar a planilha gerada pelo script Stetsom
    df_stetsom = pd.read_excel('modelos_stetson.xlsx')
    processar_planilha(df_stetsom, 'https://expertinvest.com.br/api/v1/stetsom', dia_inicial[:-2] + str(int(dia_inicial[-2:]) + 1).zfill(2))

    # Processar a planilha gerada pelo script JFA
    df_jfa = pd.read_excel('modelos_jfa.xlsx')
    processar_planilha(df_jfa, 'https://expertinvest.com.br/api/v1/jfa', dia_inicial[:-2] + str(int(dia_inicial[-2:]) + 1).zfill(2))
    
print("Processo concluído para todos os scripts.")