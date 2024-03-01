import pandas as pd
import time
import os 
import selenium 
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from cx_Freeze import setup, executable
import tkinter as tk
from tkinter import filedialog

chrome_options = Options()
#chrome_options.add_argument("--headless")

ts =  time.sleep
nav = webdriver.Chrome(options= chrome_options)

input_chip = "/html/body/div[2]/div/div[2]/form/div/div/div[2]/div/span[1]/div/div[1]/div/input"
filtrar = "/html/body/div[2]/div/div[2]/form/div/div/div[2]/div/span[1]/div/div[5]/div/div/input[1]"
naoha = "/html/body/div[2]/div/div[2]/form/div/div/div[2]/div/span[2]/div/div[1]/table/tbody/tr/td"


file_path = filedialog.askopenfilename(filetypes=[('Arquivos Excel', '*.xlsx *.xls')])
tabela = pd.read_excel(file_path)
n_iteracoes = int(input("Digite o número de chips desejado: "))

def verif_arqia():

    print("Executando coluna: arqia")
    idx = 0
    
    nfe =  nav.find_element
   
    nav.get("http://genesis.sighra.com.br/restrito/listaSimCards.xhtml")
    ts(1)
    nfe("xpath" , input_chip).clear()
    
    for linha in tabela.index:
        idx +=1
        print(f'linha =  {idx} ARQIA')
        num = tabela.loc[linha, "ARQIA"]
        try:
         num = int(num)
        except ValueError:
            break


        nfe("xpath" , input_chip).send_keys(num)
        ts(.5)
        nfe('xpath' , filtrar).click()
        ts(.5)
        try:
            status = nfe("xpath" , "/html/body/div[2]/div/div[2]/form/div/div/div[2]/div/span[2]/div/div[1]/table/tbody/tr/td[8]").text
            
            if status  == "Ativo":
                tabela.at[linha, "STATUS ARQIA"] =  "ATIVO"
            elif status == "Novo e Inativo":
                tabela.at[linha, "STATUS ARQIA"] =  "NOVO E INATIVO"
            elif status ==  "Cancelado":
                tabela.at[linha, "STATUS ARQIA"] =  "CANCELADO"
        except:
            tabela.at[linha, "STATUS ARQIA"] =  "NÃO HÁ SIM CARD"
        

        

        nfe("xpath" , input_chip).clear()
        ts(.5)
        if idx >= n_iteracoes:
            tabela["ARQIA"] = pd.to_numeric(tabela["ARQIA"], errors="coerce")
            tabela.to_excel("chip_verificado.xlsx" , index= False)
            print("Finalizado ============================================== Finalizado ============================================== Finalizado ========================================== Finalizado ")
            break
    










def verif_vivo():
    print("Executando coluna: vivo")
    
    idx = 0
  
    nfe =  nav.find_element
    genesis = ("http://genesis.sighra.com.br/login.xhtml") 
    nav.get(genesis)
    nfe("xpath" , "/html/body/div/div/div/div/div[2]/div/form/div[2]/input").send_keys("gmoura")
    ts(.5)
    nfe("xpath","/html/body/div/div/div/div/div[2]/div/form/div[3]/input").send_keys("gmourar123")
    ts(.5)
    nfe("xpath" , "/html/body/div/div/div/div/div[2]/div/form/div[5]/div/input").click()
    ts(.7)
    nav.get("http://genesis.sighra.com.br/restrito/listaSimCards.xhtml")

    for linha in tabela.index:
        idx +=1
        print(f'linha =  {idx} VIVO')
        num = tabela.loc[linha, "VIVO"]
        try:
            num = int(num)
        except ValueError:
            break

        nfe("xpath" , input_chip).send_keys(num)
        ts(.5)
        nfe('xpath' , filtrar).click()
        ts(.5)
        try:
            status = nfe("xpath" , "/html/body/div[2]/div/div[2]/form/div/div/div[2]/div/span[2]/div/div[1]/table/tbody/tr/td[8]").text
            
            if status  == "Ativo":
                tabela.at[linha, "STATUS VIVO"] =  "ATIVO"
            elif status == "Novo e Inativo":
                tabela.at[linha, "STATUS VIVO"] =  "NOVO E INATIVO"
            elif status ==  "Cancelado":
                tabela.at[linha, "STATUS VIVO"] =  "CANCELADO"
        except:
            tabela.at[linha, "STATUS VIVO"] =  "NÃO HÁ SIM CARD"
        

        

        nfe("xpath" , input_chip).clear()
        ts(.5)
        if idx >= n_iteracoes:
            tabela["VIVO"] = pd.to_numeric(tabela["VIVO"], errors="coerce")
            tabela.to_excel("chip_verificado.xlsx" , index= False)  
            break


verif_vivo()
verif_arqia()











