from selenium import webdriver
import time
import os
from selenium.webdriver.firefox.options import Options
import pandas as pd
import random

#Opções do programa
IDEstado = 33 #Dois primeiros digitos de identificação do Estado
Usuario = "ThinkPad"


# AArrumando as preferências do Firefox
Soptions = Options()
Soptions.set_preference("browser.download.folderList",2)
Soptions.set_preference("browser.download.manager.showWhenStarting", False)
Soptions.set_preference("browser.download.dir","/data")
Soptions.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream,application/vnd.ms-excel")


#Obtendo a lista das cidades. A pasta deve estar em Downloads
df = pd.read_excel('C:/Users/' + str(Usuario) + '/Downloads/DTB_2019/RELATORIO_DTB_BRASIL_MUNICIPIO.xlsx')
IDCidades = df['CódigoCompleto'].tolist()

#Restringindo a apenas as cidades do Estado Selecionado
i=0
while i < len(IDCidades):
    print(i)
    if int(str(IDCidades[i])[:2]) != 35:
        IDCidades.pop(i)
    else:
        i+=1


u = 0 # Essa variável conta quantos arquivos já foram baixado
nonconcluidos = [] # Essa variavel conta os que não foi possível baixar
tamanho = len(IDCidades) # Número total de cidades do Estado


i=0
while i < len(IDCidades):
    print(i)
    if os.path.isfile("C:/Users/" + str(Usuario) + "/Downloads/" + str(IDCidades[i])[:6] + ".xlsx"):
        IDCidades.pop(i)
        u +=1
    else:
        i+=1


for i in IDCidades:
    
    # Verifica se não tem nenhum "Resto de Download"
    if os.path.isfile("C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx.part"):
        while os.path.isfile("C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx.part"):
            source = "C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx.part"
            renamer = random.randrange(0, 1000000)
            dest = "C:/Users/" + str(Usuario) + "/Downloads/Erro" + str(renamer) + ".xlsx.part"
            try:
                os.rename(source, dest)
            except:
                print("Erro")
    
    # Verifica se não tem nada com o nome Planilha.xlsx, se tiver renomeia pra outra coisa
    while os.path.isfile("C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx"):
        source = "C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx"
        renamer = random.randrange(0, 1000000)
        dest = "C:/Users/" + str(Usuario) + "/Downloads/Erro" + str(renamer) + ".xlsx"
        try:
            os.rename(source, dest)
        except:
            print("erro")
    
    #Display de Progresso 
    os.system("cls")
    print("Já foram ", u, "/", tamanho)

    # Abrindo o site
    driver = webdriver.Firefox(options = Soptions)
    print("Acessando o Site")
    siope = "https://www.fnde.gov.br/siope/consultarRemuneracaoMunicipal.do?acao=pesquisar&cod_uf=" + str(i)[:2] + "&municipios=" + str(i)[:6]+"&anos=2018&mes=0"
    driver.get(siope)

    # Iniciando o Download
    print("Iniciando Download do Excel")
    driver.execute_script("exportarExcel()")

    # Garantindo que o downlod ocorrá, se não ocorrer o ID será adiciona em nonconluido
    espera = 0 # Espera maxima pelo site, por padrão 3 minutos
    # Sinalizadores se outras partes do programa devem rodar caso não haja download
    k = True    # Sinalizador para o verificador de existência do arquivo
    k2 = True   # Sinalizador para o renomeador de arquivo 
    while not os.path.isfile("C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx"):
        print("Esperando pelo Download")
        time.sleep(5)
        if not os.path.isfile("C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx"):
            driver.execute_script("exportarExcel()")
        if espera > 36:
            print("Não concluido")
            if os.path.isfile("C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx.part"):
                while os.path.isfile("C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx.part"):
                    source = "C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx.part"
                    renamer = random.randrange(0, 1000000)
                    dest = "C:/Users/" + str(Usuario) + "/Downloads/Erro" + str(renamer) + ".xlsx.part"
                    try:
                        os.rename(source, dest)
                    except:
                        print("Erro")
            k = False
            k2 = False
            nonconcluidos.append(i)
            break
        espera += 1
    #Verifica se o download já acabou de um jeito mais ou menos: se o tamnho do arquivo mudou
    while k:
        if os.stat("C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx").st_size != 0:
            size1 = os.stat("C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx").st_size
            time.sleep(10)
            size2 = os.stat("C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx").st_size
            if size1 == size2:
                k = False
    #Muda o nome do arquivo pro padrão
    if k2:
        print("Renomeando Arquivo")
        source = "C:/Users/" + str(Usuario) + "/Downloads/Planilha.xlsx"
        dest =  "C:/Users/" + str(Usuario) + "/Downloads/" + str(i)[:6] + ".xlsx"
        try:
            os.rename(source, dest)
        except:
            dest =  "C:/Users/" + str(Usuario) + "/Downloads/" + str(i)[:6] + "(1).xlsx"
            os.rename (source, dest)
    driver.close()
    u +=1


os.system("cls")
#print("Já foram ", u, "/", str(y-x))
print("Não conseguimos o download de:")
print(nonconcluidos)
print("Concluido!!!")