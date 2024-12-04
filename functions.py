# functions.py
import logging
import re
import numpy as np
from config import LOG_FILE, CHROMEDRIVER_PATH, DOWNLOADS_PATH, CNPJ_FOLDER, CPF_FOLDER, CNPJ_CERTIDAO_FEDERAL, CNPJ_CPF_CERTIDAO_JUSBRASIL, CNPJ_CERTIDAO_FGTS, CNPJ_CPF_CNEP, CNPJ_CPF_CEIS, INVESTIGACAO_CNPJ, INVESTIGACAO_CPF, CLIP_LAUNDERING, CNPJ_CPF_ADVICE, OFAC, CSNU, CPF_ADC_CLIENTE, PLANILHA_PATH
import os
import base64
from time import sleep
from random import randint
from dotenv import load_dotenv
from selenium.webdriver.common.by import By
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from config import iniciar_driver,criar_pastas_relatorios

# Configuração do logging
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Adicionar um manipulador de console para ver mensagens em tempo real
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
console_handler.setFormatter(formatter)
logging.getLogger().addHandler(console_handler)

def log_message(message, level='info'):
    """Função para registrar mensagens no log."""
    level = level.lower()
    if level == 'info':
        logging.info(message)
    elif level == 'error':
        logging.error(message)
    elif level == 'warning':
        logging.warning(message)
    else:
        logging.debug(message)  # Para níveis não reconhecidos, registra como debug.
    
    print(f"{level.upper()}: {message}")

def save_pdf(driver, filepath):
    """Salva uma página como PDF."""
    result = driver.execute_cdp_cmd("Page.printToPDF", {
        "format": "A4",
        "printBackground": True
    })
    pdf_data = base64.b64decode(result['data'])
    with open(filepath, "wb") as file:
        file.write(pdf_data)


def consulta_pessoa_juridica_Portal_Transparencia(CNPJ_formatado, CNPJ):
    driver = iniciar_driver()
    driver.get(f'https://portaldatransparencia.gov.br/busca?termo={CNPJ_formatado}&pessoaJuridica=true')
    log_message(f"Consultando o {CNPJ} no Portal da Transparencia")
    try:
        # Localizar o botão pelo ID e clicar nele
        botao_rejeitar_cookies = driver.find_element(By.ID, "accept-minimal-btn")
        botao_rejeitar_cookies.click()
        log_message("Botão de rejeição de cookies clicado com sucesso.")
    except NoSuchElementException:
        log_message("Botão de rejeição de cookies não encontrado.")
    except ElementNotInteractableException:
        log_message("Botão de rejeição de cookies não está interagível.")
    sleep(randint(5, 10))
    element_text = driver.find_element(By.ID, 'countResultados').text
    
    if element_text == '0':
        save_pdf(driver, f"{CNPJ_FOLDER}/Consulta_PessoaJuridica_PortalTransparencia_{CNPJ}.pdf")
    else:
        driver.find_element(By.XPATH, "//body/main[1]/section[2]/div[1]/div[1]/div[1]/div[2]/ul[1]/div[1]/h4[1]").click()
        sleep(10)
        save_pdf(driver, f"{CNPJ_FOLDER}/Consulta_PessoaJuridica_PortalTransparencia_{CNPJ}.pdf")
    driver.quit()

def consulta_representantes_Portal_Transparencia(CNPJ, CPFs):
    for CPF in CPFs:
        try:
            driver = iniciar_driver()
            driver.get(f'https://portaldatransparencia.gov.br/pessoa-fisica/busca/lista?termo={CPF}&pagina=1&tamanhoPagina=10')
            log_message(f"Consultando CPF: {CPF}")
            sleep(15)
            try:
                # Localizar o botão pelo ID e clicar nele
                botao_rejeitar_cookies = driver.find_element(By.ID, "accept-minimal-btn")
                botao_rejeitar_cookies.click()
                log_message("Botão de rejeição de cookies clicado com sucesso.")
            except NoSuchElementException:
                log_message("Botão de rejeição de cookies não encontrado.")
            except ElementNotInteractableException:
                log_message("Botão de rejeição de cookies não está interagível.")
                
            element_text = driver.find_element(By.XPATH, "//strong[@id='countResultados']").text
            if element_text == '0':
                save_pdf(driver, f"{CPF_FOLDER}/Consulta_Representantes_PortalTransparencia_{CPF}.pdf")
            else:
                result = driver.find_element(By.CLASS_NAME, "link-busca-nome")
                result.click()
                sleep(randint(3, 10))
                save_pdf(driver, f"{CPF_FOLDER}/Consulta_Representantes_PortalTransparencia_{CPF}.pdf")
            log_message(f"Consulta para CPF {CPF} finalizada.")
        except NoSuchElementException:
            log_message(f"Elemento não encontrado para CPF: {CPF}", level='warning')
        except Exception as e:
            log_message(f"Erro ao consultar CPF {CPF}: {e}", level='error')
        finally:
            driver.quit()


def consulta_CNEP_Portal_Transparencia(CNPJ_formatado, CNPJ):
    driver = iniciar_driver()
    driver.get(f'https://portaldatransparencia.gov.br/busca?termo={CNPJ_formatado}&cnep=true')
    log_message(f"Consultando o {CNPJ} no Portal da Transparencia cnep")
    sleep(randint(3, 10))
    element_text = driver.find_element(By.XPATH, "//strong[@id='countResultados']").text
    
    if element_text == '0':
        save_pdf(driver, f"{CNPJ_FOLDER}/Consulta_CNEP_PortalTransparencia_{CNPJ}.pdf")
    else:
        for cont in range(1, int(element_text) + 1):
            driver.get(f'https://portaldatransparencia.gov.br/busca?termo={CNPJ_formatado}&cnep=true')
            sleep(randint(3, 10))
            driver.find_element(By.XPATH, f'/html[1]/body[1]/main[1]/section[2]/div[1]/div[1]/div[1]/div[2]/ul[1]/div[{cont}]/h4[1]/a[1]').click()
            save_pdf(driver, f"{CNPJ_FOLDER}/Consulta_CNEP_PortalTransparencia_{CNPJ}.pdf")
        driver.quit()

def consulta_CEIS_Portal_Transparencia(CNPJ_formatado, CNPJ):
    driver = iniciar_driver()
    driver.get(f'https://portaldatransparencia.gov.br/busca?termo={CNPJ_formatado}&ceis=true')
    log_message(f"Consultando o {CNPJ} no Portal da Transparencia CEIS")
    sleep(randint(3, 10))
    element_text = driver.find_element(By.XPATH, "//strong[@id='countResultados']").text
    
    if element_text == '0':
        save_pdf(driver, f"{CNPJ_FOLDER}/Consulta_CEIS_PortalTransparencia_{CNPJ}.pdf")
    else:
        for cont in range(1, int(element_text) + 1):
            driver.get(f'https://portaldatransparencia.gov.br/busca?termo={CNPJ_formatado}&ceis=true')
            sleep(randint(3, 10))
            driver.find_element(By.XPATH, f'/html[1]/body[1]/main[1]/section[2]/div[1]/div[1]/div[1]/div[2]/ul[1]/div[{cont}]/h4[1]/a[1]').click()
            save_pdf(driver, f"{CNPJ_FOLDER}/Consulta_CEIS_PortalTransparencia_{CNPJ}.pdf")
        driver.quit()

def consulta_OFAC(CNPJ_formatado, CNPJ):
    driver = iniciar_driver()
    driver.get('https://sanctionssearch.ofac.treas.gov/')
    log_message(f"Consultando o {CNPJ} na OFAC")
    sleep(randint(3, 10))
    driver.find_element(By.XPATH, '/html[1]/body[1]/form[1]/div[3]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[2]/table[1]/tbody[1]/tr[3]/td[3]/input[1]').send_keys(CNPJ_formatado)
    driver.find_element(By.XPATH, '/html[1]/body[1]/form[1]/div[3]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[2]/table[1]/tbody[1]/tr[5]/td[5]/input[1]').click()
    sleep(randint(3, 10))
    total_height = driver.execute_script("return document.body.scrollHeight")
    driver.set_window_size(1920, total_height)
    driver.save_screenshot(f'{CNPJ_FOLDER}/Consulta_OFAC_PortalTransparencia_{CNPJ}.png')
    driver.quit()

def consulta_CSNU(CNPJ_formatado, CNPJ):
    driver = iniciar_driver()
    driver.get('https://scsanctions.un.org/SEARCH/')
    log_message(f"Consultando o {CNPJ} no CSNU")
    sleep(randint(3, 10))
    driver.find_element(By.XPATH, "//input[@id='include']").send_keys(CNPJ_formatado)
    driver.find_element(By.XPATH, '/html[1]/body[1]/center[1]/form[1]/table[1]/tbody[1]/tr[26]/td[3]/input[1]').click()
    sleep(randint(3, 10))
    total_height = driver.execute_script("return document.body.scrollHeight")
    driver.set_window_size(1920, total_height)
    driver.save_screenshot(f'{CNPJ_FOLDER}/Consulta_CSNU_{CNPJ}.png')
    driver.quit()

def consulta_PJ_Linkana(CNPJ_linkana, CNPJ):
    driver = iniciar_driver()
    driver.get('https://cnpj.linkana.com/')
    log_message(f"Consultando o {CNPJ} no Linkana")
    sleep(randint(3, 10))
    driver.find_element(By.XPATH, "//body/div[1]/div[1]/div[1]/form[1]/div[1]/input[1]").send_keys(CNPJ_linkana)
    driver.find_element(By.XPATH, "//button[contains(text(),'BUSCAR GRATUITAMENTE')]").click()
    sleep(randint(5, 10))
    driver.find_element(By.XPATH, "//body/div[1]/main[1]/div[1]/div[1]/a[1]").click()
    sleep(randint(5, 10))
    total_height = driver.execute_script("return document.body.scrollHeight")
    driver.set_window_size(1920, total_height)
    driver.save_screenshot(f'{CNPJ_FOLDER}/Consulta_PessoaJuridica_Linkana_{CNPJ}.png')
    driver.quit()

def consulta_Midia_Negativa(CNPJ_formatado, CNPJ, ADVICE_TECH_USERNAME, ADVICE_TECH_PASSWORD):
    dotenv_path = ".env"
    load_dotenv(dotenv_path)
    ADVICE_TECH_USERNAME = os.getenv('Advice_Tech_USERNAME')
    ADVICE_TECH_PASSWORD = os.getenv('Advice_Tech_PASSWORD')
    driver = iniciar_driver()
    driver.get('https://www.advicetech.com.br/ClipLaunderingWeb/')
    log_message(f"Consultando o {CNPJ} no ADVICETECH")
    sleep(10)
    driver.find_element(By.XPATH, '//body/app-root[1]/app-login[1]/div[1]/div[1]/form[1]/div[1]/div[2]/input[1]').send_keys(ADVICE_TECH_USERNAME)
    driver.find_element(By.XPATH, '//body/app-root[1]/app-login[1]/div[1]/div[1]/form[1]/div[1]/div[3]/input[1]').send_keys(ADVICE_TECH_PASSWORD)
    driver.find_element(By.XPATH, "//a[contains(text(),'Entrar')]").click()
    sleep(8)
    driver.find_element(By.XPATH, '//input[@id="txtPesquisaNome"]').send_keys(CNPJ_formatado)
    driver.find_element(By.XPATH, '//body/app-root[1]/div[1]/div[1]/div[1]/app-dashboard[1]/div[1]/div[1]/div[1]/app-pesquisar[1]/div[1]/div[1]/div[1]/form[1]/div[1]/div[1]/div[2]/div[2]/button[1]').click()
    sleep(8)
    total_height = driver.execute_script("return document.body.scrollHeight")
    driver.set_window_size(1920, total_height)
    driver.save_screenshot(f'{CNPJ_FOLDER}/Consulta_MidiaNegativa_{CNPJ}.png')
    driver.quit()

def consulta_processual_jusbrasil(CNPJ_formatado, CNPJ, Razao_Social):
    driver = iniciar_driver()
    driver.get(f'https://www.jusbrasil.com.br/consulta-processual/busca?q={CNPJ_formatado}')
    log_message(f"Consultando o {CNPJ} no JusBrasil")
    sleep(randint(4, 10))
    
    save_pdf(driver, f"{CNPJ_FOLDER}/Processos_Encontrados_Jusbrasil_{CNPJ}.pdf")
    
    try:
        num_links = driver.find_element(By.XPATH, '//*[@id="app-root"]/div/div/div[1]/main/div[1]/div[2]/div/span').text
        num_link = int(num_links.split()[0])
        num_link = min(num_link, 5)

        for cont in range(1, num_link + 1):
            driver.find_element(By.XPATH, f'/html/body/div[1]/div/div/div[1]/main/ul/li[{cont}]/div/div/div[2]/div[1]/a').click()
            sleep(2)
            cookie_popup = driver.find_elements(By.CSS_SELECTOR, '.icon.icon-close')
            if cookie_popup:
                cookie_popup[0].click()
            sleep(randint(3, 7))
            save_pdf(driver, f"{CNPJ_FOLDER}/Consulta_Processual_Jusbrasil_{CNPJ}_{cont}.pdf")
            driver.get(f'https://www.jusbrasil.com.br/consulta-processual/busca?q={CNPJ_formatado}')
            sleep(randint(4, 7))
    except Exception:
        log_message(f'Sem processos encontrados para: {Razao_Social}')
    
    driver.quit()

def is_cnpj_or_cpf(document):
    document = re.sub(r'[.\-\/]', '', str(document))
    if len(document) == 14:
        return 'CNPJ'
    elif len(document) == 11:
        return 'CPF'
    else:
        return 'INVALID'


def iniciar_processo(CNPJ_formatado, CNPJ, CPFs, CNPJ_linkana, ADVICE_TECH_USERNAME, ADVICE_TECH_PASSWORD, Razao_Social):
    """Inicia todas as consultas necessárias para o CNPJ especificado."""
    try:
        if not os.path.exists(f'{CNPJ_FOLDER}'):
            os.makedirs(f'{CNPJ_FOLDER}')

        criar_pastas_relatorios()
        consulta_pessoa_juridica_Portal_Transparencia(CNPJ_formatado, CNPJ)
        consulta_representantes_Portal_Transparencia(CNPJ, CPFs)
        consulta_CNEP_Portal_Transparencia(CNPJ_formatado, CNPJ)
        consulta_CEIS_Portal_Transparencia(CNPJ_formatado, CNPJ)
        consulta_PJ_Linkana(CNPJ_linkana, CNPJ)
        consulta_CSNU(CNPJ_formatado, CNPJ)
        consulta_OFAC(CNPJ_formatado, CNPJ)
        consulta_Midia_Negativa(CNPJ_formatado, CNPJ, ADVICE_TECH_USERNAME, ADVICE_TECH_PASSWORD)
        consulta_processual_jusbrasil(CNPJ_formatado, CNPJ, Razao_Social)

    except Exception as e:
        log_message(f"Erro no processamento de {CNPJ}: {e}", level='error')