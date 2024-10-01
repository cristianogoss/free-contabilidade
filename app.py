from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from docx import Document
from selenium.webdriver.chrome.options import Options
import os

chrome_options = Options()
arguments = ['--lang=pt-BR', '--window-size=800,800','--incognito']
for argument in arguments:
    chrome_options.add_argument(argument)
driver = webdriver.Chrome(options=chrome_options)
driver.get("https://contabil-devaprender.netlify.app/")

def campo_dados():
    # Campo Email
    campo_email = driver.find_element(By.XPATH,'/html/body/div/div/form/div[1]/input')
    sleep(2)
    campo_email.click()
    campo_email.send_keys('cristiano@gmail.com')
    sleep(2)
        
    # Campo Senha
    campo_senha = driver.find_element(By.XPATH, '/html/body/div/div/form/div[2]/input')
    sleep(2)
    campo_senha.click()
    campo_senha.send_keys("123456")
    sleep(2)

    # clicar botão entrar
    botao_entrar = driver.find_element(By.XPATH, '/html/body/div/div/form/button')
    botao_entrar.click()
    sleep(3)

    # Clicar botao acessar sistema
    botoes_sistemas = driver.find_elements(By.XPATH, '/html/body/div/div/div/div[1]/div/div/a')
    sleep(2)
    botoes_sistemas[0].click()

campo_dados()

def processar_documento_word(caminho_arquivo_word):
# Extrair dados do arquivo word
    arquivo_word = Document(caminho_arquivo_word)

    ativo_circulante = ''
    caixa_equivalentes = ''
    contas_receber = ''
    estoques = ''
    ativo_nao_circulante = ''
    imobilizado = ''
    intangivel = ''
    total_ativo = ''

    # Acessar as tabelas
    for tabela in arquivo_word.tables:
        for linha in tabela.rows:
            if 'Ativo Circulante' in linha.cells[0].text.strip():
                ativo_circulante = linha.cells[1].text.strip()
            elif 'Caixa e Equivalentes' in linha.cells[0].text.strip():
                caixa_equivalentes = linha.cells[1].text.strip()
            elif 'Contas a Receber' in linha.cells[0].text.strip():
                contas_receber = linha.cells[1].text.strip()
            elif 'Estoques' in linha.cells[0].text.strip():
                estoques = linha.cells[1].text.strip()
            elif 'Ativo Não Circulante' in linha.cells[0].text.strip():
                ativo_nao_circulante = linha.cells[1].text.strip()
            elif 'Imobilizado' in linha.cells[0].text.strip():
                imobilizado = linha.cells[1].text.strip()
            elif 'Intangível' in linha.cells[0].text.strip():
                intangivel = linha.cells[1].text.strip()
            elif 'Total do Ativo' in linha.cells[0].text.strip():
                total_ativo = linha.cells[1].text.strip()

    # Acessar campo ativo circulante
    campo_ativo_circulante = driver.find_element(By.XPATH, '//*[@id="ativo_circulante"]')
    sleep(1)
    campo_ativo_circulante.click()
    campo_ativo_circulante.send_keys(ativo_circulante)

    # Acessar campo caixa equivalentes
    campo_caixa_equivalentes = driver.find_element(By.XPATH, '//*[@id="caixa_equivalentes"]')
    sleep(1)
    campo_caixa_equivalentes.click()
    campo_caixa_equivalentes.send_keys(caixa_equivalentes)

    # Acessar campo contas a receber
    campo_contas_receber = driver.find_element(By.XPATH, '//*[@id="contas_receber"]')
    sleep(1)
    campo_contas_receber.click()
    campo_contas_receber.send_keys(contas_receber)

    # Acessar campo estoques
    campo_estoques = driver.find_element(By.XPATH, '//*[@id="estoques"]')
    sleep(1)
    campo_estoques.click()
    campo_estoques.send_keys(estoques)

    # Acessar campo ativo nao circulante
    campo_ativo_nao_circulante = driver.find_element(By.XPATH, '//*[@id="ativo_nao_circulante"]')
    sleep(1)
    campo_ativo_nao_circulante.click()
    campo_ativo_nao_circulante.send_keys(ativo_nao_circulante)

    # Campo imobilizado
    campo_imobilizado = driver.find_element(By.XPATH, '//*[@id="imobilizado"]')
    sleep(1)
    campo_imobilizado.click()
    campo_imobilizado.send_keys(imobilizado)

    # Campo intangivel
    campo_intangivel = driver.find_element(By.XPATH, '//*[@id="intangivel"]')
    sleep(1)
    campo_intangivel.click()
    campo_intangivel.send_keys(intangivel)

    # Campo total ativo
    campo_total_ativo= driver.find_element(By.XPATH, '//*[@id="total_ativo"]')
    sleep(1)
    campo_total_ativo.click()
    campo_total_ativo.send_keys(total_ativo)

    # Botao cadastrar
    botao_cadastrar = driver.find_element(By.XPATH, '//*[@id="balanco-form"]/div[2]/button')
    botao_cadastrar.click()


pasta_relatorios = r'C:\Users\crigo\OneDrive\Área de Trabalho\free contabilidade\relatorios'
for nome_arquivo in os.listdir(pasta_relatorios):
    if nome_arquivo.endswith('.docx'):    
        caminho_arquivo = os.path.join(pasta_relatorios, nome_arquivo)
        processar_documento_word(caminho_arquivo)
    