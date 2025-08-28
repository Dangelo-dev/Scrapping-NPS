import re
import os
import time
import tkinter as tk
import pandas as pd
from tkcalendar import DateEntry
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import dotenv_values

def selecionar_datas_visual():
    root = tk.Tk()
    root.title("Selecione o período")
    tk.Label(root, text="Data Inicial:").grid(row=0, column=0)
    tk.Label(root, text="Data Final:").grid(row=1, column=0)
    cal_inicio = DateEntry(root, date_pattern='dd/mm/yyyy')
    cal_fim = DateEntry(root, date_pattern='dd/mm/yyyy')
    cal_inicio.grid(row=0, column=1)
    cal_fim.grid(row=1, column=1)
    resultado = {}

    def confirmar():
        resultado['min_data'] = cal_inicio.get()
        resultado['max_data'] = cal_fim.get()
        root.destroy()

    tk.Button(root, text="OK", command=confirmar).grid(row=2, column=0, columnspan=2)
    root.mainloop()
    return resultado['min_data'], resultado['max_data']

def carregar_env():
    config = dotenv_values(os.path.join(os.path.dirname(__file__), ".env"))
    return config

def iniciar_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def login(driver, url, user, password):
    driver.get(url)
    time.sleep(3)
    driver.find_element(By.XPATH, '//*[@id="signInName"]').send_keys(user)
    driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(password)
    driver.find_element(By.XPATH, '//*[@id="next"]').click()
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[2]/div/div[4]/div/div[1]/div[2]/div[2]/div/section/div/div[1]/div/div[1]')))

def aplicar_filtro_periodo(driver, min_data, max_data, xpath_inicio, xpath_fim, xpath_abrir_calendario, xpath_abrir_filtro, xpath_programa_pesquisa):
    # Abrir filtro
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath_abrir_filtro)))
    driver.find_element(By.XPATH, xpath_abrir_filtro).click()
    time.sleep(2)
    # Abrir calendário
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath_abrir_calendario)))
    driver.find_element(By.XPATH, xpath_abrir_calendario).click()
    time.sleep(2)
    # Preencher datas
    driver.find_element(By.XPATH, xpath_inicio).send_keys(min_data)
    driver.find_element(By.XPATH, xpath_fim).send_keys(max_data)
    # Remover filtro programa de pesquisas
    driver.find_element(By.XPATH, xpath_programa_pesquisa).click()
    time.sleep(2)

def selecionar_loja(driver, codigo_loja, xpath_lista_lojas, xpath_checkbox_loja, xpath_aplicar, abrir_lista=True):
    time.sleep(2)
    if abrir_lista:
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath_lista_lojas)))
        driver.find_element(By.XPATH, xpath_lista_lojas).click()
        time.sleep(2)
    checkbox_xpath = xpath_checkbox_loja.format(codigo_loja)
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
    driver.find_element(By.XPATH, checkbox_xpath).click()
    time.sleep(2)
    driver.find_element(By.XPATH, xpath_aplicar).click()
    time.sleep(5)

def desmarcar_loja(driver, codigo_loja, xpath_abrir_filtro, xpath_lista_lojas, xpath_checkbox_loja):
    # Abrir filtro
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath_abrir_filtro)))
    driver.find_element(By.XPATH, xpath_abrir_filtro).click()
    time.sleep(2)
    # Abrir lista de lojas
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath_lista_lojas)))
    driver.find_element(By.XPATH, xpath_lista_lojas).click()
    time.sleep(2)
    # Desmarcar loja
    checkbox_xpath = xpath_checkbox_loja.format(codigo_loja)
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
    driver.find_element(By.XPATH, checkbox_xpath).click()
    time.sleep(2)

def extrair_dados_loja(driver, pdv, avaliacoes_xpath, min_data, max_data):
    dados = []
    for avaliacao, xpaths in avaliacoes_xpath.items():
        try:
            positivos_texto = driver.find_element(By.XPATH, xpaths['positivo']).text
            neutros_texto = driver.find_element(By.XPATH, xpaths['neutro']).text
            negativos_texto = driver.find_element(By.XPATH, xpaths['negativo']).text

            # Extrai o número entre parênteses
            positivos = int(re.search(r'\((\d+)\)', positivos_texto).group(1))
            neutros = int(re.search(r'\((\d+)\)', neutros_texto).group(1))
            negativos = int(re.search(r'\((\d+)\)', negativos_texto).group(1))

            total = positivos + neutros + negativos
            dados.append({
                "PDV": pdv,
                "Avaliação": avaliacao,
                "Analises Positivas": positivos,
                "Analises Neutras": neutros,
                "Analises Negativas": negativos,
                "Total Analises": total,
                "Min Data": min_data,
                "Max Data": max_data
            })
        except Exception as e:
            print(f"Loja {pdv} ({avaliacao}): nenhuma avaliação disponível.")
    return dados

def main():
    env = carregar_env()

    min_data, max_data = selecionar_datas_visual()

    # Lista de xpaths
    xpath_abrir_filtro = '//*[@id="main-content-for-skip-link"]/div[1]/div[1]/div/div/div/div/div/div/div/button'
    xpath_abrir_calendario = '//*[@id="main-content-for-skip-link"]/div[1]/div/form/div/div[1]/div/div/div[1]/div/div[1]/div[1]/div[2]/div/div[2]/button'
    xpath_inicio = '//input[@placeholder="dd/mm/yyyy" and @aria-label="Período de Tempo Início"]'
    xpath_fim = '//input[@placeholder="dd/mm/yyyy" and @aria-label="Período de Tempo Fim"]'
    xpath_aplicar = '//*[@id="main-content-for-skip-link"]/div[1]/div/form/div/div[2]/div/button[3]/canvas'
    xpath_lista_lojas = '//*[@id="pfe_unitid-control"]'
    xpath_checkbox_loja = '//button[@data-label="{}"]'
    xpath_programa_pesquisa = '//*[@id="main-content-for-skip-link"]/div[1]/div/form/div/div[1]/div/div/div[1]/div/div[2]/div[4]/div[1]/div/button'

    avaliacoes_xpath_default = {
        "NPS BOT": {
            "positivo": '//*[@id="main-content-for-skip-link"]/div[2]/div/div[6]/div/div[1]/div[2]/div[2]/div/section/div/div[1]/div[2]/div[1]/div',
            "neutro": '//*[@id="main-content-for-skip-link"]/div[2]/div/div[6]/div/div[1]/div[2]/div[2]/div/section/div/div[1]/div[2]/div[2]/div',
            "negativo": '//*[@id="main-content-for-skip-link"]/div[2]/div/div[6]/div/div[1]/div[2]/div[2]/div/section/div/div[1]/div[2]/div[3]/div'
        },
        "NPS Clique & Retire": {
            "positivo": '//*[@id="main-content-for-skip-link"]/div[2]/div/div[4]/div/div[3]/div[2]/div[2]/div/section/div/div/div[2]/div[1]/div',
            "neutro": '//*[@id="main-content-for-skip-link"]/div[2]/div/div[4]/div/div[3]/div[2]/div[2]/div/section/div/div/div[2]/div[2]/div',
            "negativo": '//*[@id="main-content-for-skip-link"]/div[2]/div/div[4]/div/div[3]/div[2]/div[2]/div/section/div/div/div[2]/div[3]/div'
        }
    }
    avaliacoes_xpath_911557 = {
        "NPS QDB": {
            "positivo": '//*[@id="main-content-for-skip-link"]/div[2]/div/div[6]/div/div[2]/div[2]/div[2]/div/section/div/div[1]/div[2]/div[1]/div',
            "neutro": '//*[@id="main-content-for-skip-link"]/div[2]/div/div[6]/div/div[2]/div[2]/div[2]/div/section/div/div[1]/div[2]/div[2]/div',
            "negativo": '//*[@id="main-content-for-skip-link"]/div[2]/div/div[6]/div/div[2]/div[2]/div[2]/div/section/div/div[1]/div[2]/div[3]/div'
        },
        "NPS Clique & Retire": {
            "positivo": '//*[@id="main-content-for-skip-link"]/div[2]/div/div[4]/div/div[3]/div[2]/div[2]/div/section/div/div/div[2]/div[1]/div',
            "neutro": '//*[@id="main-content-for-skip-link"]/div[2]/div/div[4]/div/div[3]/div[2]/div[2]/div/section/div/div/div[2]/div[2]/div',
            "negativo": '//*[@id="main-content-for-skip-link"]/div[2]/div/div[4]/div/div[3]/div[2]/div[2]/div/section/div/div/div[2]/div[3]/div'
        }
    }
    

    driver = iniciar_driver()
    login(driver, env["URL"], env["USER"], env["PASSWORD"])
    aplicar_filtro_periodo(driver, min_data, max_data, xpath_inicio, xpath_fim, xpath_abrir_calendario, xpath_abrir_filtro, xpath_programa_pesquisa)

    # Abrir lista de lojas para capturar os códigos disponíveis
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath_lista_lojas)))
    driver.find_element(By.XPATH, xpath_lista_lojas).click()
    time.sleep(3)

    # Capturar todos os códigos disponíveis
    botoes_lojas = driver.find_elements(By.XPATH, '//button[@data-label]')
    codigos_lojas = [btn.get_attribute('data-label') for btn in botoes_lojas]

    # Fechar o dropdown de lojas após capturar os códigos
    driver.find_element(By.XPATH, xpath_lista_lojas).click()
    time.sleep(2)

    resultados = []

    for idx, codigo_loja in enumerate(codigos_lojas):
        try:
            if idx > 0:
                desmarcar_loja(driver, codigos_lojas[idx-1], xpath_abrir_filtro, xpath_lista_lojas, xpath_checkbox_loja)
                selecionar_loja(driver, codigo_loja, xpath_lista_lojas, xpath_checkbox_loja, xpath_aplicar, abrir_lista=False)
            else:
                selecionar_loja(driver, codigo_loja, xpath_lista_lojas, xpath_checkbox_loja, xpath_aplicar)
            if codigo_loja == "911557":
                avaliacoes_xpath = avaliacoes_xpath_911557
            else:
                avaliacoes_xpath = avaliacoes_xpath_default
            dados_loja = extrair_dados_loja(driver, codigo_loja, avaliacoes_xpath, min_data, max_data)
            resultados.extend(dados_loja)
            print(f"Dados extraídos da loja {codigo_loja}")
        except Exception as e:
            print(f"Erro ao extrair dados da loja {codigo_loja}: {e}")

    driver.quit()
    
    # Extrai AnoMes de max_data (esperando formato dd/mm/yyyy)
    ano_mes = max_data[-4:] + max_data[3:5]  # Pega ano e mês
    nome_arquivo = f"NPS {ano_mes}.xlsx"

    df = pd.DataFrame(resultados)
    df.to_excel(nome_arquivo, index=False)

if __name__ == "__main__":
    main()
