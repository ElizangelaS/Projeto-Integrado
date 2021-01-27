#!/usr/bin/env python
# coding: utf-8

# <h1><center>SSP - Segurança Pública</center></h1>
# <h2><center>Feminicídio</center></h2>
# 

# Esse código faz consulta no site SSP, e baixa todos os arquivos com dados sobre Feminicídio de 2016 até de 2020, de todos os meses. Depois envia para o Google Storage.
# 
# 
# <img src="https://selenium-python.readthedocs.io/_static/logo.png"/>

# # Web Scraping

# ### Importando Bibliotecas

# Responsavel pela automação de testes no navegador.<br> 
#     <font color='green'>from </font>selenium <font color='green'>import </font>webdriver </br> 
# 
# Responsavel pela "espera". Faz com que um programa seja obrigado a aguardar os segundos determinados pelo sleep, para poder executar o próximo código.<br>
#     <font color='green'>from </font>time <font color='green'>import </font>sleep </br>
# 
# Com essa biblioteca, permite que acessamos ferramentas Microsoft no Windows com Python.<br>
#     <font color='green'>import </font>win32com.client <font color='green'>as </font>win32 </br>
# 
# Ele permite que listemos os arquivos de um diretório, usando expressões semelhantes as que usamos no shell, 
# como por exemplo: *.<br>
#     <font color='green'>import </font>glob <br>
#     
# Essa é uma das biblitecas para realizar análise de dados.<br>
#     <font color='green'>import </font>pandas <font color='green'>as </font>pd </br>
#     
# Essa biblioteca é responsavel por fazer uma "ponte" com o GCP e enviar os dados para o Storage.<br>
#     <font color='green'>from </font>google.cloud <font color='green'>import </font>storage </br>

# In[5]:


# !pip install selenium
# !pip install google-cloud-storage


# In[6]:


from selenium import webdriver
from time import sleep
import win32com.client as win32
import glob
import pandas as pd
from google.cloud import storage


# ### Inicia o browser

# In[ ]:


driver = webdriver.Chrome('chromedriver')


# ### Entra no site desejado

# In[ ]:


driver.get('http://www.ssp.sp.gov.br/transparenciassp/Consulta.aspx')


# ### Seleciona categoria Feminicídio e depois seleciona o mês de janeiro

# In[ ]:


botao1 = driver.find_element_by_xpath('//*[@id="cphBody_btnFeminicidio"]').click()
sleep(5)
botao2 = driver.find_element_by_xpath('//*[@id="cphBody_listMes1"]').click()


# In[ ]:


def ano2020():
    """
    Botão 1: Seleciona o feminicidio na página.
    Botão 2: Seleciona janeiro de 2020.
    Botão 3: Pega todos os meses de 2020 conforme o range.
    Botão 4: Baixa o arquivo com os dados.
    Botão 5: Clica no botão fechar do mês de dezembro.
    """
    botao1 = driver.find_element_by_xpath('//*[@id="cphBody_btnFeminicidio"]').click()
    sleep(5)
    botao2 = driver.find_element_by_xpath('//*[@id="cphBody_listMes1"]').click()
    for c in range(1, 13):
        botao3 = driver.find_element_by_xpath(f'//*[@id="cphBody_listMes{c}"]').click()
        botao4 = driver.find_element_by_xpath('//*[@id="cphBody_ExportarBOLink"]').click()
        sleep(5)
        botao5 = driver.find_element_by_xpath('//*[@id="cphBody_divModal"]/div/div/div[3]/button')
        driver.execute_script("arguments[0].click()", botao5)
        sleep(5)


# In[ ]:


def ano2019():
    """
    Botão 6: Seleciona o ano 2019.
    Botão 7: Seleciona janeiro de 2019.
    Botão 8: Pega todos os meses de 2019 conforme o range.
    Botão 9: Baixa o arquivo com os dados.
    """
    botao6 = driver.find_element_by_xpath('//*[@id="cphBody_lkAno19"]').click()
    botao7 = driver.find_element_by_xpath('//*[@id="cphBody_lkMes1"]').click()
    for c in range(1, 13):
        botao8 = driver.find_element_by_xpath(f'//*[@id="cphBody_listMes{c}"]').click()
        sleep(5)
        botao9 = driver.find_element_by_xpath('//*[@id="cphBody_ExportarBOLink"]').click()
        sleep(5)


# In[ ]:


def ano2018():
    """
    Botão 10: Seleciona o ano 2018.
    Botão 11: Seleciona janeiro de 2018.
    Botão 12: Pega todos os meses de 2018 conforme o range.
    Botão 13: Baixa o arquivo com os dados.
    """
    botao10 = driver.find_element_by_xpath('//*[@id="cphBody_lkAno18"]').click()
    botao11 = driver.find_element_by_xpath('//*[@id="cphBody_lkMes1"]').click()
    for c in range(1, 13):
        botao12 = driver.find_element_by_xpath(f'//*[@id="cphBody_listMes{c}"]').click()
        sleep(5)
        botao13 = driver.find_element_by_xpath('//*[@id="cphBody_ExportarBOLink"]').click()
        sleep(5)


# In[ ]:


def ano2017():
    """
    Botão 14: Seleciona o ano 2017.
    Botão 15: Seleciona janeiro de 2017.
    Botão 16: Pega todos os meses de 2017 conforme o range.
    Botão 17: Baixa o arquivo com os dados.
    """
    botao14 = driver.find_element_by_xpath('//*[@id="cphBody_lkAno17"]').click()
    botao15 = driver.find_element_by_xpath('//*[@id="cphBody_lkMes1"]').click()
    for c in range(1, 13):
        botao16 = driver.find_element_by_xpath(f'//*[@id="cphBody_listMes{c}"]').click()
        sleep(5)
        botao17 = driver.find_element_by_xpath('//*[@id="cphBody_ExportarBOLink"]').click()
        sleep(5)


# In[ ]:


def ano2016():
    """
    Botão 18: Seleciona o ano 2016.
    Botão 19: Seleciona janeiro de 2016.
    Botão 20: Pega todos os meses de 2016 conforme o range.
    Botão 21: Baixa o arquivo com os dados.
    """
    botao18 = driver.find_element_by_xpath('//*[@id="cphBody_lkAno16"]').click()
    botao19 = driver.find_element_by_xpath('//*[@id="cphBody_lkMes1"]').click()
    for c in range(1, 13):
        botao20 = driver.find_element_by_xpath(f'//*[@id="cphBody_listMes{c}"]').click()
        sleep(5)
        botao21 = driver.find_element_by_xpath('//*[@id="cphBody_ExportarBOLink"]').click()
        sleep(5)


# In[ ]:


ano2020()
sleep(10)
ano2019()
sleep(10)
ano2018()
sleep(10)
ano2017()
sleep(10)
ano2016()


# ### Converte os arquivos XLS para XLSX do ano 2016 a 2020 de janeiro à dezembro

# In[ ]:


for ano in range (2016,2021):
    for i in range(1,13):
        try:
            fname = fr'C:\Users\elizangelasoaress\Downloads\DadosBO_{ano}_{i}(FEMINICÍDIO).xls' 
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(fname)
            wb.SaveAs(fname+"x", FileFormat = 51) # FileFormat = 51 is for .xlsx extension
            wb.Close() # FileFormat = 56 is for .xls extension
            excel.Application.Quit()
        except:
            pass


# ### Listando os arquivos

# In[9]:


d1 = glob.glob(r"C:\Users\elizangelasoaress\Downloads\*.xlsx")


# ### Unificação de arquivos e conversão para CSV

# In[ ]:


all_data = pd.DataFrame()
for f in glob.glob(r"C:\Users\elizangelasoaress\Downloads\*.xlsx"):
    df = pd.read_excel(f)
    all_data = all_data.append(df,ignore_index=True)


# In[ ]:


all_data.to_csv(r'C:\Users\elizangelasoaress\Downloads\Base-consolidada.csv')


# ### Conversão de tipo primitivo

# In[ ]:


df = pd.read_csv(r"C:\Users\elizangelasoaress\Downloads\Base-consolidada.csv")


# In[ ]:


df['DATAOCORRENCIA'] = pd.to_datetime(df['DATAOCORRENCIA'])


# In[ ]:


df.to_csv(r"C:\Users\elizangelasoaress\Downloads\Base-Fem.csv")


# In[ ]:


df = pd.read_csv(r"C:\Users\elizangelasoaress\Downloads\Base-Fem.csv")


# ### Migração da VM para Google Storage

# In[ ]:


storage_client = storage.Client.from_service_account_json('blueshift-academy.json') # Autenticação do GCP
bucket = storage_client.get_bucket('proj_integrado-stg') # Busca o bucket
blob = bucket.blob('NewFolder') # Cria uma nova pasta
blob.upload_from_filename(r'C:\Users\elizangelasoaress\Downloads\Base-Fem.csv') # Faz o upload do arquivo
bucket.rename_blob(blob, "BaseConsolidadaFem.csv") # Renomeia o arquivo


# ### Finalizando a aplicação

# In[ ]:


driver.close()

