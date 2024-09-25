
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
                'date': dia_inicial.split('T')[0]
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

# Definir as datas de início e fim
data_inicial = datetime(2024, 9, 1)
data_final = datetime(2024, 9, 22)


# Iterar do dia inicial até o dia final
data_atual = data_inicial
while data_atual <= data_final:
    dia_inicial = data_atual.strftime('%Y-%m-%d')
    dia_final = dia_inicial  # Cada iteração processa um único dia

    # Chamar o primeiro script (Taramps)
    chamar_script(dia_inicial, dia_final, 'taramps.py')
    
    # Chamar o segundo script (Usina)
    chamar_script(dia_inicial, dia_final, 'usina.py')
    
    # Chamar o terceiro script (Stetsom)
    chamar_script(dia_inicial, dia_final, 'stetson.py')
    
    # Chamar o quarto script (JFA)
    chamar_script(dia_inicial, dia_final, 'jfa.py')
    
    # Processar a planilha gerada pelo script Taramps
    df_taramps = pd.read_excel('modelos_taramps.xlsx')
    processar_planilha(df_taramps, 'http://localhost:8090/api/v1/taramps', dia_inicial)
    
    # Processar a planilha gerada pelo script Usina
    df_usina = pd.read_excel('modelos_usina.xlsx')
    processar_planilha(df_usina, 'http://localhost:8090/api/v1/usina', dia_inicial)
    
    # Processar a planilha gerada pelo script Stetsom
    df_stetsom = pd.read_excel('modelos_stetson.xlsx')
    processar_planilha(df_stetsom, 'http://localhost:8090/api/v1/stetsom', dia_inicial)
    
    # Processar a planilha gerada pelo script JFA
    df_jfa = pd.read_excel('modelos_jfa.xlsx')
    processar_planilha(df_jfa, 'http://localhost:8090/api/v1/jfa', dia_inicial)
    
    # print("Processo concluído para todos os scripts.")
    data_atual += timedelta(days=1)
