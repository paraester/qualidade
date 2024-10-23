import logging
import time
import os
from datetime import datetime
import pyautogui  # Para controlar o sistema operacional
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys  # Importar para usar a tecla Enter
import credentials  # Importar o arquivo de credenciais (usuário e senha)

# Ativar logs para debugging
logging.basicConfig(level=logging.INFO)

# Função para configurar o Firefox para download com escolha da pasta
def configurar_firefox_para_download():
    options = webdriver.FirefoxOptions()

    # Desativar o download automático e mostrar o diálogo de download
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)  # Permitir escolher onde salvar o arquivo
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "")
    profile.set_preference("browser.download.useDownloadDir", False)  # Forçar abrir a caixa de diálogo de download
    options.profile = profile
    return options

# Função para interagir com a caixa de diálogo de download
def salvar_arquivo_como(caminho_arquivo):
    time.sleep(2)  # Aguardar a caixa de diálogo abrir
    logging.info(f"Digitando o caminho do arquivo: {caminho_arquivo}")
    pyautogui.write(caminho_arquivo)  # Digitar o caminho do arquivo sem a extensão ".csv"
    time.sleep(1)
    pyautogui.press('enter')  # Confirmar o salvamento
    logging.info(f"Arquivo salvo como: {caminho_arquivo}")

try:
    # Configurar o WebDriver para Firefox com as preferências de download
    options = configurar_firefox_para_download()
    logging.info("Configurações de download definidas para o Firefox.")

    driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()), options=options)
    logging.info("Navegador iniciado com sucesso!")

    # Acessar a página de login
    driver.get("https://www.ppm.celepar.pr.gov.br/pm/#/staffing/staffing")
    logging.info("Página acessada com sucesso!")

    # Verificar se estamos na página correta
    wait = WebDriverWait(driver, 30)
    wait.until(EC.url_contains("#/staffing/staffing"))
    logging.info("Estamos na página de staffing.")

    # Usar contains no XPath para IDs dinâmicos do campo de username
    username_field = wait.until(EC.element_to_be_clickable((By.XPATH, '//input[contains(@id, "username")]')))
    
    # Simular a digitação lenta no campo de usuário
    for char in credentials.username:
        username_field.send_keys(char)
        time.sleep(0.2)
    logging.info("Nome de usuário inserido lentamente com sucesso!")

    # Usar contains no XPath para IDs dinâmicos do campo de senha
    password_field = wait.until(EC.element_to_be_clickable((By.XPATH, '//input[contains(@id, "password")]')))
    
    # Simular a digitação lenta no campo de senha
    for char in credentials.password:
        password_field.send_keys(char)
        time.sleep(0.2)
    logging.info("Senha inserida lentamente com sucesso!")

    # Enviar a tecla Enter para submeter o formulário
    password_field.send_keys(Keys.ENTER)
    logging.info("Tecla Enter enviada para submeter o formulário")

    # Esperar a próxima página carregar
    wait.until(EC.url_contains("#/staffing/staffing"))
    logging.info("Login realizado com sucesso!")

    # Aguardar 15 segundos antes de procurar o botão "Exportar para CSV"
    time.sleep(15)
    logging.info("Aguardando 15 segundos antes de procurar o botão 'Exportar para CSV'.")

    # Esperar o botão "Exportar para CSV" aparecer
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Exportar para CSV"]')))
    logging.info("Botão 'Exportar para CSV' localizado.")

    # Aguardar mais 5 segundos antes de clicar no botão
    time.sleep(5)
    logging.info("Aguardando 5 segundos antes de clicar no botão 'Exportar para CSV'.")

    # Clicar no botão "Exportar para CSV"
    export_button.click()
    logging.info("Botão 'Exportar para CSV' clicado.")

    # Aguardar 30 segundos para a exportação ser concluída
    time.sleep(30)
    logging.info("Aguardando 30 segundos após exportação.")

    # Procurar e clicar no ícone de notificações
    notification_bell = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Painel de notificações"]')))
    logging.info("Ícone de notificações encontrado.")
    
    # Aguardar mais 5 segundos antes de clicar no ícone de notificações
    time.sleep(5)
    notification_bell.click()
    logging.info("Ícone de notificações clicado.")

    # Esperar mais 5 segundos para que as notificações sejam exibidas
    time.sleep(5)

    # Procurar a primeira notificação relacionada à exportação de CSV
    first_notification = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[contains(@class, "notification-list") and contains(.,"Exportação de CSV concluída")]')))
    first_notification.click()
    logging.info("Primeira notificação de 'Exportação de CSV' clicada.")

    # Criar a pasta 'Dados' para armazenar o arquivo exportado
    pasta_dados = os.path.join(os.getcwd(), "Dados")
    if not os.path.exists(pasta_dados):
        os.makedirs(pasta_dados)
        logging.info(f"Pasta 'Dados' criada em {pasta_dados}")

    # Definir o nome do arquivo sem a extensão ".csv"
    nome_arquivo = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")  # Sem ".csv" aqui
    caminho_arquivo = os.path.join(pasta_dados, nome_arquivo)

    # Simulação de escolha da pasta e salvamento do arquivo com PyAutoGUI
    salvar_arquivo_como(caminho_arquivo)

except Exception as e:
    logging.error(f"Ocorreu um erro: {e}")

finally:
    # Aguardar 5 segundos antes de fechar o navegador
    time.sleep(5)
    if 'driver' in locals():
        driver.quit()
        logging.info("Navegador fechado.")
