import xlsxwriter  
from selenium import webdriver
import selenium.webdriver.support.expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from time import sleep
from datetime import date
from selenium.webdriver.common.keys import Keys

options = Options()
options.add_experimental_option("detach", True)
navegador = webdriver.Chrome(ChromeDriverManager().install(), options=options)
navegador2 = webdriver.Chrome(ChromeDriverManager().install(), options=options)

#Processos seletivos da UNIFAL
navegador.get("https://sistemas.unifal-mg.edu.br/app/graduacao/inscricaograduacao/relatorios/resultados.php?tipo=html&edital_id=366")

cont = 1


while(True):
    try:
        curso = navegador.find_element(By.XPATH, '//*[@id="conteudo"]/p['+str(cont)+']/a').get_attribute('href')
          
    except:
        break

    print(curso)
    navegador2.get(curso)
    
    sleep(5)

    try:
        print(navegador2.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[3]/th').text)

    except:
        print("Não há")




    cont = cont + 1