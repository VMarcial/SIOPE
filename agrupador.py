from selenium import webdriver
import time
import os
from selenium.webdriver.firefox.options import Options
import pandas as pd
import random



# Arrumando o Firefox
Soptions = Options()
Soptions.set_preference("browser.download.folderList",2)
Soptions.set_preference("browser.download.manager.showWhenStarting", False)
Soptions.set_preference("browser.download.dir","/data")
Soptions.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream,application/vnd.ms-excel")


#Lista das cidades
df = pd.read_excel('C:/Users/vinic/Downloads/DTB_2019/RELATORIO_DTB_BRASIL_MUNICIPIO.xlsx') # can also index sheet by name or fetch all sheets
IDCidades = df['CódigoCompleto'].tolist()

tamanho = len(IDCidades)
i=0
while i < len(IDCidades):
    print(i)
    if int(str(IDCidades[i])[:2]) != 33:
        IDCidades.pop(i)
    else:
        i+=1
print(IDCidades)
print(len(IDCidades))

#x = int(input("Digite qual parte do programa você vai começar 0/5570: "))
#y = int(input("Digite qual parte do programa você vai terminar " + str(x) + "/" + str(len(IDCidades)) + ": "))

#IDCidades = IDCidades[x:y]

u = 0
nonconcluidos = []

for i in IDCidades:
    if os.path.isfile("C:/Users/vinic/Downloads/" + str(i)[:6] + ".xlsx"):
        IDCidades.remove(i)
        u += 1



for i in IDCidades:
    while os.path.isfile("C:/Users/vinic/Downloads/Planilha.xlsx"):
        source = "C:/Users/vinic/Downloads/Planilha.xlsx"
        renamer = random.randrange(0, 1000000)
        dest = "C:/Users/vinic/Downloads/Erro" + str(renamer) + ".xlsx"
        try:
            os.rename(source, dest)
        except:
            print("erro")
    os.system("cls")
    print("Já foram ", u, "/", len(IDCidades))
    driver = webdriver.Firefox(options = Soptions)
    print("Acessando o Site")
    siope = "https://www.fnde.gov.br/siope/consultarRemuneracaoMunicipal.do?acao=pesquisar&cod_uf=" + str(i)[:2] + "&municipios=" + str(i)[:6]+"&anos=2018&mes=0"
    driver.get(siope)
    print("Iniciando Download do Excel")
    driver.execute_script("exportarExcel()")
    espera = 0
    k = True
    k2 = True
    while not os.path.isfile("C:/Users/vinic/Downloads/Planilha.xlsx"):
        print("Esperando pelo Download")
        time.sleep(5)
        if not os.path.isfile("C:/Users/vinic/Downloads/Planilha.xlsx"):
            driver.execute_script("exportarExcel()")
        if espera > 36:
            print("Não concluido")
            k = False
            k2 = False
            nonconcluidos.append(i)
            break
        espera += 1
    #Verifica se o download já acabou de um jeito mais ou menos: se o tamnho do arquivo mudou
    while k:
        if os.stat("C:/Users/vinic/Downloads/Planilha.xlsx").st_size != 0:
            size1 = os.stat("C:/Users/vinic/Downloads/Planilha.xlsx").st_size
            time.sleep(10)
            size2 = os.stat("C:/Users/vinic/Downloads/Planilha.xlsx").st_size
            if size1 == size2:
                k = False
    #Muda o nome do arquivo pro padrão
    if k2:
        print("Renomeando Arquivo")
        source = "C:/Users/vinic/Downloads/Planilha.xlsx"
        dest =  "C:/Users/vinic/Downloads/" + str(i)[:6] + ".xlsx"
        try:
            os.rename(source, dest)
        except:
            dest =  "C:/Users/vinic/Downloads/" + str(i)[:6] + "(1).xlsx"
            os.rename (source, dest)
    driver.close()
    u +=1


os.system("cls")
#print("Já foram ", u, "/", str(y-x))
print("Não conseguimos o download de:")
print(nonconcluidos)
print("Concluido!!!")