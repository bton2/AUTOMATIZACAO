from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import openpyxl

driver = webdriver.Chrome()
# 1- Navegar até o site:
# "https://contabilidade-devaprender.netlify.app/"
driver.get("https://contabilidade-devaprender.netlify.app/")
sleep(5)
# Nota: para obter os valores dos campos nos navegadores, precisamos clicar com o botão direito 
# em cima do campo e pedir para inspecionar. Depois precisamos encontrar uma forma única de
#  identificar o campo, para isso precisamos usar uma forma chamada XPATH.
# Ex: tag[@atributo='valor]
# só lembrando pegar um atributo que não se repete, entre página que está em forma de inspect
# clicar ctcl+f e pesquisar: //input[@id='email], achando um valor único, então é só usá-lo
# //input[@id='email]
# 2 - Digitar o e-mail
email = driver.find_element(By.XPATH,"//input[@id='email']")
sleep(2)
email.send_keys('admin@contabilidade.com')
# 3 - Digitar a senha
senha = driver.find_element(By.XPATH,"//input[@id='senha']")
sleep(2)
senha.send_keys('contabilidade123456')
# 4 - Clicar em entrar
botao_entrar = driver.find_element(By.XPATH,"//button[@id='Entrar']")
sleep(2)
botao_entrar.click()
sleep(3)
# 5 - Extrair as informações da planilha
empresas = openpyxl.load_workbook('/home/ben/Documents/AUTOMATIZACAO/empresas.xlsx')
paginas_empresas = empresas['dados empresas']
for linha in paginas_empresas.iter_rows(min_row=2,values_only=True):
    nome_empresa, email, telefone, endereco, cnpj, area_atuacao, quantidade_de_funcionarios, data_fundacao = linha
# 5 - Clicar em cada campo e preencher com a informaçao extraída da planilha
    driver.find_element(By.ID,'nomeEmpresa').send_keys(nome_empresa)
    sleep(2)
    driver.find_element(By.ID,'emailEmpresa').send_keys(email)
    sleep(2)
    driver.find_element(By.ID,'telefoneEmpresa').send_keys(telefone)
    sleep(2)
    driver.find_element(By.ID,'enderecoEmpresa').send_keys(endereco)
    sleep(2)
    driver.find_element(By.ID,'cnpj').send_keys(cnpj)
    sleep(2)
    driver.find_element(By.ID,'areaAtuacao').send_keys(area_atuacao)
    sleep(2)
    driver.find_element(By.ID,'numeroFuncionarios').send_keys(quantidade_de_funcionarios)
    sleep(2)
    driver.find_element(By.ID,'dataFundacao').send_keys(data_fundacao)
    sleep(2)

    # 6 - Clicar em cadastrar
    driver.find_element(By.ID,'Cadastrar').click()
    sleep(3)
# 7 - Repito os passos 5 e 6  
