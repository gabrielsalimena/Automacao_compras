from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import win32com.client as win32
from tabulate import tabulate

servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servico)
driver.maximize_window()

buscas = pd.read_excel(r'C:\Users\Pichau\Desktop\Projetos Python\2\buscas.xlsx')
valores_google = {}
valores_bp = {}


# google.shopping
for nome,termos,p_min,p_max in buscas.values:
    driver.get('https://shopping.google.com.br/')
    driver.find_element(By.XPATH,'//*[@id="REsRA"]').send_keys(nome,Keys.ENTER)
    anuncios = driver.find_elements(By.CLASS_NAME,'i0X6df')
    while len(anuncios) < 20:
        time.sleep(1)
    lista_final = []
    for anuncio in anuncios:
        nome_anuncio = anuncio.find_element(By.CLASS_NAME,'tAxDx').text
        lista_termo = (termos.split())
        for termo in lista_termo:
            if termo.casefold() in nome_anuncio.casefold():
                continue

        link = anuncio.find_element(By.TAG_NAME, 'a').get_attribute('href')

        valor = anuncio.find_element(By.CLASS_NAME,'a8Pemb').text
        if valor == ' ' or '':
            continue
        valor = valor.replace('R$','').replace('.','').replace(',','.').strip().replace(' ','_')

        try:
            if '_' in valor:
                index = valor.find('_')
                valor = valor[:index]
                valor = float(valor)
            else:
                valor = float(valor)
        except:
            continue

        if valor < p_max and valor > p_min:
            lista_final.append([nome_anuncio,valor,link])
    valores_google[nome] = lista_final
#google.buscape
for nome,termos,p_min,p_max in buscas.values:
    driver.get('https://www.buscape.com.br/')
    driver.find_element(By.XPATH,'//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(nome,Keys.ENTER)
    time.sleep(2)
    driver.find_element(By.XPATH,'//*[@id="__next"]/div/div[6]/div[2]/div[2]/div/select/option[3]').click()

    anuncios = driver.find_elements(By.CLASS_NAME,'Paper_Paper__HIHv0')
    lista_final = []
    for anuncio in anuncios:
        nome_anuncio = anuncio.find_element(By.TAG_NAME,'h2').text

        lista_termo = (termos.split())
        for termo in lista_termo:
            if termo.casefold() in nome_anuncio.casefold():
                continue

        link = anuncio.find_element(By.TAG_NAME, 'a').get_attribute('href')


        valor = anuncio.find_element(By.TAG_NAME,'p').text

        if valor == ' ' or '':
            continue
        valor = valor.replace('R$','').replace('.','').replace(',','.').strip().replace(' ','_')


        if '_' in valor:
            index = valor.find('_')
            valor = valor[:index]
            valor = float(valor)
        else:
            valor = float(valor)

        a = [nome_anuncio,valor,link]

        if valor < p_max and valor > p_min:
            lista_final.append(a)
        if len(lista_final) < 3:
            continue
    valores_bp[nome] = lista_final

for nome,termos,p_min,p_max in buscas.values:
    tabela_bp = tabulate(valores_bp[nome])
    tabela_google = tabulate(valores_google[nome])

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'gabriel.s@gmail.com'
    mail.Subject = f'Preços de: {nome} solicitados'

    mail.Body = f'''
    
    Bom dia, segue preços abaixo:
    Preços Buscapé:
    {tabela_bp}
    Preços Google.Shopping:
    {tabela_google}
    '''

    mail.Send()
