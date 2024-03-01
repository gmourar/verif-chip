import smtplib
import schedule
import time
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE
from email import encoders
import datetime , timedelta
import selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from cx_Freeze import setup, Executable
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import tkinter as tk
from tkinter import filedialog
from tqdm import tqdm


janela = tk.Tk()

chrome_options = Options()
chrome_options.add_argument("--headless")
root = tk.Tk()
root.withdraw()

nav = webdriver.Chrome(options= chrome_options)


def get():
    nav.get(sighramanager)
    nfe('xpath' , '/html/body/form/table/tbody/tr[1]/td[2]/div/div[2]/table/tbody/tr[1]/td[2]/input').send_keys('gmoura')
    time.sleep(.5)
    nfe('xpath', '/html/body/form/table/tbody/tr[1]/td[2]/div/div[2]/table/tbody/tr[2]/td[2]/input').send_keys('gmourar123')
    
    time.sleep(.5)
    nfe('xpath', '/html/body/form/table/tbody/tr[1]/td[2]/div/div[2]/table/tbody/tr[4]/td/input').click()
    time.sleep(1)
    nav.find_element('xpath' , '/html/body/form/div/div/table/tbody/tr[2]/td/table/tbody/tr/td[5]/div/div[1]').click()
    time.sleep(1)
    nav.find_element('xpath' , '/html/body/form/div/div/table/tbody/tr[2]/td/table/tbody/tr/td[5]/div/div[2]/div/div/div[14]').click()
    

#get()

def tab():
    file_path = filedialog.askopenfilename(filetypes=[('Arquivos Excel', '*.xlsx *.xls')])
    tabela = pd.read_excel(file_path)
    total_iterations = len(tabela.index)
    print(tabela)
    nfe =  nav.find_element


    sighramanager = ('http://187.61.13.196/mn00000/login.jsf')
    print(f"====================================== /////   {file_path} ~~ Foi selecionado. ///// ======================================================")
    nav.get(sighramanager)
    nfe('xpath' , '/html/body/form/table/tbody/tr[1]/td[2]/div/div[2]/table/tbody/tr[1]/td[2]/input').send_keys('gmoura')
    time.sleep(.5)
    nfe('xpath', '/html/body/form/table/tbody/tr[1]/td[2]/div/div[2]/table/tbody/tr[2]/td[2]/input').send_keys('gmourar123')
    
    time.sleep(.5)
    nfe('xpath', '/html/body/form/table/tbody/tr[1]/td[2]/div/div[2]/table/tbody/tr[4]/td/input').click()
    time.sleep(1)
    nav.find_element('xpath' , '/html/body/form/div/div/table/tbody/tr[2]/td/table/tbody/tr/td[5]/div/div[1]').click()
    time.sleep(1)
    nav.find_element('xpath' , '/html/body/form/div/div/table/tbody/tr[2]/td/table/tbody/tr/td[5]/div/div[2]/div/div/div[14]').click()
    tabela = pd.read_excel(file_path)
    idx = 0
    for linha in tabela.index:
        progress_bar = tqdm(total=100, desc='Progresso')
        idx += 1
    
        progress_bar.set_description(f'Executando linha {idx}/{total_iterations}')
        num = tabela.loc[linha,"SERIAL"]
        num = int(num)
        
        progress_bar.update(10)
        nav.find_element("xpath", "/html/body/form/div/div/table/tbody/tr[4]/td/div/div[2]/div/table/tbody/tr/td/fieldset/table/tbody/tr[7]/td[2]/input").send_keys(num)
        progress_bar.update(10)
        nfe("xpath" , "/html/body/form/div/div/table/tbody/tr[5]/td/fieldset/table/tbody/tr/td/input[2]").click()
        progress_bar.update(10)
      
       
        time.sleep(15)
        progress_bar.update(10)
        try:
            table = nav.find_element('xpath', "/html/body/form/div/div/table/tbody/tr[7]/td/div/span/table/tbody")
            linhas = len(table.find_elements('xpath', "/html/body/form/div/div/table/tbody/tr[7]/td/div/span/table/tbody/tr"))
            
            progress_bar.update(10)
            
            
            if linhas == 1:
                
                nfe("id" , f"form1:tableFormServico:0:j_id_jsp_2119280986_191").click()
            else:
                linhas -= 1
                
            
                nflinha = nfe("id" , f"form1:tableFormServico:{linhas}:j_id_jsp_2119280986_191").click()
                
            progress_bar.update(10)
            time.sleep(5)
            progress_bar.update(10)
            obs = nfe("id" , "formDetailsServ:dtlLfseObs").text
            progress_bar.update(10)
            tabela.loc[linha, "Observação"] = obs
            time.sleep(.5)
            

            nfe("id" , "formDetailsServ:j_id_jsp_2119280986_375").click()
            progress_bar.update(10)
            time.sleep(1)
            
            progress_bar.update(10)
            nav.find_element("xpath", "/html/body/form/div/div/table/tbody/tr[4]/td/div/div[2]/div/table/tbody/tr/td/fieldset/table/tbody/tr[7]/td[2]/input").clear()
        
            progress_bar.update(10)
        except NoSuchElementException:
            print(f"Elemento não encontrado na linha {idx}. Pulando para a próxima.")
            nfe('id' , 'btnMsgsFechar').click() 

            nav.find_element("xpath", "/html/body/form/div/div/table/tbody/tr[4]/td/div/div[2]/div/table/tbody/tr/td/fieldset/table/tbody/tr[7]/td[2]/input").clear()

            continue
        progress_bar.close()
        print(f"linha {idx} executada!")
        
        
        
            
    tabela.to_excel("tabela_atualizada.xlsx", index=False)
    print("As observações foram adcionadas na planilha 'tabela_atualizada.xlsx")
    input("Pressione Enter para finalizar...")
 
        







       



def verificar ():
    chrome_options.add_argument("--headless")
    file_path = filedialog.askopenfilename(filetypes=[('Arquivos Excel', '*.xlsx *.xls')])
    tabela = pd.read_excel(file_path)
    print(tabela)
    nfe =  nav.find_element


    sighramanager = ('http://187.61.13.196/mn00000/login.jsf')
    print(f"====================================== /////   {file_path} ~~ Foi selecionado. ///// ======================================================")
    
    nav.get(sighramanager)
    nfe('xpath' , '/html/body/form/table/tbody/tr[1]/td[2]/div/div[2]/table/tbody/tr[1]/td[2]/input').send_keys('gmoura')
    time.sleep(.5)
    nfe('xpath', '/html/body/form/table/tbody/tr[1]/td[2]/div/div[2]/table/tbody/tr[2]/td[2]/input').send_keys('gmourar123')
    
    time.sleep(.5)
    nfe('xpath', '/html/body/form/table/tbody/tr[1]/td[2]/div/div[2]/table/tbody/tr[4]/td/input').click()
    tabela = pd.read_excel(file_path)
    nfe('xpath' , '/html/body/form/div/div/table/tbody/tr[2]/td/table/tbody/tr/td[19]/div/div[1]').click()
    time.sleep(.7)
    nfe('xpath' , '/html/body/form/div/div/table/tbody/tr[2]/td/table/tbody/tr/td[19]/div/div[2]/div/div/div[4]/span[2]').click()
    time.sleep(1)
    idx = 0
    primeira_linha = 1
    for linha in tabela.index:
        idx += 1
        num = tabela.loc[linha,"SERIAL"]
        num = int(num)
        print(f"Executando linha {idx}")

        nfe('xpath' , '/html/body/form/div/div/table/tbody/tr[4]/td/div/div[2]/table/tbody/tr[1]/td/fieldset/table/tbody/tr[1]/td[2]/input').send_keys(num)
        time.sleep(.7)
        nfe('xpath' , '/html/body/form/div/div/table/tbody/tr[4]/td/div/div[2]/table/tbody/tr[2]/td/fieldset/div[1]/input[1]').click()
        time.sleep(20)
        table = nav.find_element('xpath', "/html/body/form/div/div/table/tbody/tr[4]/td/div/div[2]/table/tbody/tr[4]/td/div/span/table/tbody")
        linhas = len(table.find_elements('xpath', "/html/body/form/div/div/table/tbody/tr[4]/td/div/div[2]/table/tbody/tr[4]/td/div/span/table/tbody/tr"))
        print(linhas)
        
       
        
        data_inicio = nfe('xpath' , f'/html/body/form/div/div/table/tbody/tr[4]/td/div/div[2]/table/tbody/tr[4]/td/div/span/table/tbody/tr[{linhas}]/td[14]').text
        data_inicio = data_inicio.split()[0]
        data_final = nfe("xpath" , '/html/body/form/div/div/table/tbody/tr[4]/td/div/div[2]/table/tbody/tr[4]/td/div/span/table/tbody/tr[1]/td[14]').text
        data_final = data_final.split()[0]
        
        dt_inicio = datetime.datetime.strptime(data_inicio, "%d/%m/%y")
        dt_final = datetime.datetime.strptime(data_final, "%d/%m/%y")

        diferenca = dt_final - dt_inicio

        if diferenca < datetime.timedelta(days=365):
           
             tabela.at[linha, "Garantia"] = "Está na Garantia"
             print(f"linha {idx} , está na Garantia")

        elif diferenca > datetime.timedelta(days=365):
            tabela.at[linha , "Garantia"] = "Sem Garantia"
            print(f"linha {idx} , está sem Garantia")
        

        time.sleep(1)
        nfe('xpath' , '/html/body/form/div/div/table/tbody/tr[4]/td/div/div[2]/table/tbody/tr[1]/td/fieldset/table/tbody/tr[1]/td[2]/input').clear()

    tabela.to_excel("tabela_garantia.xlsx", index=False)
    print("A Garantia dos Equipamentos foi verificada e adicionada na planilha 'tabela_garantia.xlsx")
    input("Pressione Enter para finalizar...")
 

largura = 200
altura = 100
dimensoes = f"{largura}x{altura}"
janela.geometry(dimensoes)

botao1 = tk.Button(janela, text="Observação", command=tab)
botao2 = tk.Button(janela, text="Verificar Garantia", command=verificar)
botao1.grid(row=0, column=0)
botao2.grid(row=0, column=1)
    

janela.mainloop()
        



        





        

    






    



    
        
    
    

 