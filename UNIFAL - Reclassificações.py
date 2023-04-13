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

#Listas para alocar as variáveis
classificacao = []
alunos = []
cursos = []
nota_final = []
situacao = []
modalidade = []

#Buscar no site pelos links com as páginas de cada um dos cursos
while(True):
    try:
        paginas = navegador.find_element(By.XPATH, '//*[@id="conteudo"]/p['+str(cont)+']/a').get_attribute('href')
        
    except:
        #Caso não existir esse curso, chegou-se ao final da página
        break

    #Para cada um dos cursos, acessar a página
    navegador2.get(paginas)
    
    sleep(5)

    try:
        #Verificar a existência do cabeçalho - e portanto, de aprovados - com a informação de modalidade
        navegador2.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[3]/th').text
        cont_alunos = 5

        while(True):
            try:
                #Tentar extratir a informações de todos os alunos
                classificacao.append(navegador2.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[' + str(cont_alunos)+']/td[1]').text)
                alunos.append(navegador2.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[' + str(cont_alunos)+']/td[2]').text)
                nota_final.append(navegador2.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[' + str(cont_alunos)+']/td[4]').text)
                situacao.append(navegador2.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[' + str(cont_alunos)+']/td[5]').text)
                cursos.append(navegador2.find_element(By.XPATH, '//*[@id="conteudo"]/h2').text)
                modalidade.append(navegador2.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[3]/th').text)
            
            except:
                #Caso não existir esse aluno, chegou-se ao final da página
                break

            cont_alunos = cont_alunos + 1

        print(modalidade)

    except:
        print("Não há")

    cont = cont + 1

print(modalidade)

#Abrir um arquivo .xlsx a partir do caminho
book = xlsxwriter.Workbook("")     
sheet = book.add_worksheet()  

#Cabeçalho do arquivo
sheet.write(0, 0, 'Classificação')
sheet.write(0, 1, 'Aluno')
sheet.write(0, 2, 'Curso')
sheet.write(0, 3, 'Nota final')
sheet.write(0, 4, 'Situação')
sheet.write(0, 5, 'Modalidade')

#Todas as listas possuem o mesmo número de elementos
for i in range(len(alunos)):
    sheet.write(i+1, 0, classificacao[i])
    sheet.write(i+1, 1, alunos[i])
    sheet.write(i+1, 2, cursos[i])
    sheet.write(i+1, 3, nota_final[i])
    sheet.write(i+1, 4, situacao[i])
    sheet.write(i+1, 5, modalidade[i])

#Fechar o arquivo   
book.close()