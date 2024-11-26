import json
import argparse
import subprocess
import pandas as pd
import os
from datetime import datetime, timedelta
from tkinter import messagebox

def juntar_planilhas(data_hoje):
    #print(data_hoje)
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
    
    if os.path.exists("resultado_final.xlsx"):
        df_juntado = pd.read_excel("resultado_final.xlsx")
    else:
        df_juntado = pd.DataFrame()
    
    for arquivo in arquivos:
        if os.path.exists(arquivo):
            df = pd.read_excel(arquivo)
            df['Data'] = data_hoje
            df_juntado = pd.concat([df_juntado, df], ignore_index=True)
    
    df_juntado.to_excel('resultado_final.xlsx', index=False)

def chamar_script(dia_inicial, dia_final, cookie, escolha):
    arquivos = [
        "modelos_jfa.xlsx",
        "modelos_stetson.xlsx",
        "modelos_taramps.xlsx",
        "modelos_epever.xlsx",
        "modelos_hayonik.xlsx",
        "modelos_usina.xlsx",
        "modelos_volt.xlsx",
        "produtos.xlsx",
        "modelos_tataliken.xlsx",
        "modelos_knup.xlsx",
        "modelos_amfer.xlsx",
    ]
    
    for arquivo in arquivos:
        if os.path.exists(arquivo):
            os.remove(arquivo)
    scripts = []
    print(escolha)
    if "JFA" in escolha:
        scripts.append("jfa-ia.py")
    elif "Usina" in escolha:
        scripts.append("usina.py")
    elif "Taramps" in escolha:
        scripts.append("taramps.py")
    elif "Amfer" in escolha:
        scripts.append("amfer.py")
    elif "Hayonik" in escolha:
        scripts.append("hayonik.py")
    elif "Knup" in escolha:
        scripts.append("knup.py")
    elif "Stetson" in escolha:
        scripts.append("stetson.py")
    elif "Volt" in escolha:
        scripts.append("volt.py")
        
    #['amfer.py', 'hayonik.py', 'jfa-ia.py', 'knup.py', 'stetson.py', 'taramps.py', 'volt.py', 'usina.py']
    
    for script in scripts:
        comando = [
            'python',
            script,
            '--dia_inicial', dia_inicial,
            '--dia_final', dia_final,
            '--cookie', cookie
        ]
        try:
            subprocess.run(comando, check=True)
            ##print(f"Script {script} executado com sucesso.")
        except subprocess.CalledProcessError as e:
            #print(f"Erro ao executar o script {script}: {e}")
            messagebox.showerror("Erro", f"Erro ao executar o script {script}: {e}")

    juntar_planilhas(dia_inicial)

def main():
    parser = argparse.ArgumentParser(description='Executar scripts com datas específicas.')
    parser.add_argument('--dia_inicial', type=str, required=True, help='Data inicial no formato YYYY-MM-DD')
    parser.add_argument('--dia_final', type=str, required=True, help='Data final no formato YYYY-MM-DD')
    parser.add_argument('--cookie', type=str, required=True, help='Cookie')
    parser.add_argument('--escolha', required=True, help='Escolha')

    args = parser.parse_args()
    escolha_list = json.loads(args.escolha)

    chamar_script(args.dia_inicial, args.dia_final, args.cookie, escolha_list)

if __name__ == "__main__":
    main()