# main.py
import os
import re
import dotenv
import pandas as pd
from time import sleep
import glob
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from functions import iniciar_processo, log_message
from config import PLANILHA_PATH, PLANILHA_ADIQ_PATH, PLANILHA_GLOBAL_PATH, PLANILHA_DUPLICIDADE_PATH, CONSULTA_CNPJ, CNPJ_REPORT, CONSULTA_CPF, CPF_REPORT, iniciar_driver

driver = iniciar_driver()

dotenv.load_dotenv()
user_name = os.getenv('USER_NAME')
password = os.getenv('PASSWORD')
adc_username = os.getenv('ADC_USERNAME')
adc_password = os.getenv('ADC_PASSWORD')

# Carregando as planilhas
df = pd.read_excel(PLANILHA_PATH)
planilha_adiq = pd.read_excel(PLANILHA_ADIQ_PATH)
planilha_global = pd.read_excel(PLANILHA_GLOBAL_PATH)
planilha_duplicidade = pd.read_excel(PLANILHA_DUPLICIDADE_PATH)
fim = len(df)

def conciliar_planilhas():
    # Adicionar colunas necessárias se não existirem
    colunas_para_atualizar = ['STATUS ADC LOJISTA', 'STATUS DO CBK ADIQ', 'STATUS DO CBK GLOBAL', 'MOBBUY', 'RECUPERADO_BNC', 'PENDÊNCIA PORTADOR']
    for col in colunas_para_atualizar:
        if col not in df.columns:
            df[col] = ''
        df[col] = df[col].astype(str)

    # Iteração pelos dados da planilha
    for index, cnpj in df['CNPJ/CPF'].items():
        try:
            cnpj_formatado = re.sub(r'\D', '', cnpj)
            cnpj_formatado = str(cnpj_formatado)
            log_message(f"Processando CNPJ/CPF: {cnpj_formatado}")

            # Consulta de Lojista
            driver.get(CONSULTA_CNPJ)
            username_field = driver.find_element(By.ID, 'txtUsuarioLogin')
            username_field.send_keys(adc_username)
            password_field = driver.find_element(By.ID, 'txtSenhaLogin')
            password_field.send_keys(adc_password)
            login_button = driver.find_element(By.XPATH, "//*[@id='btnOkLogin']")
            login_button.click()
            sleep(5)

            driver.get(CNPJ_REPORT)
            driver.find_element(By.ID, 'txtCpfCnpj').clear()
            driver.find_element(By.ID, 'txtCpfCnpj').send_keys(cnpj)
            driver.find_element(By.ID, 'btnPesquisar').click()
            sleep(8)

            try:
                tabela = driver.find_element(By.ID, 'multiPageParametros')
                df.at[index, 'STATUS ADC LOJISTA'] = 'CADASTRADO'
                log_message("CNPJ cadastrado!")

                # Simulação de download e leitura de planilha
                pasta_downloads = os.path.expanduser('~/Downloads')
                arquivos = glob.glob(os.path.join(pasta_downloads, '*.xls'))
                if arquivos:
                    arquivo_mais_recente = max(arquivos, key=os.path.getctime)
                    planilha_bnc_recuperado = pd.read_html(arquivo_mais_recente)
                    tabelas = planilha_bnc_recuperado[0]
                    tabelas.to_excel(f'planilha_formatada_{cnpj_formatado}.xlsx', index=False)
                    planilha_bnc = pd.read_excel(f'planilha_formatada_{cnpj_formatado}.xlsx')
                    log_message(f"Arquivo processado: {arquivo_mais_recente}")
                else:
                    log_message("Nenhum arquivo encontrado!")

                # Verificação na planilha CBK ADIQ
                resultado_adiq = planilha_adiq[planilha_adiq['CNPJ/CPF'] == cnpj]
                df.at[index, 'STATUS DO CBK ADIQ'] = 'SIM' if not resultado_adiq.empty else 'NÃO'
                log_message(f"Resultado Adiq: {str(resultado_adiq)}")

                # Verificação na planilha CBK GLOBAL
                resultado_global = planilha_global[planilha_global['CNPJ/CPF'] == cnpj]
                df.at[index, 'STATUS DO CBK GLOBAL'] = 'SIM' if not resultado_global.empty else 'NÃO'
                log_message(f"Resultado Global: {str(resultado_global)}")

                # Verificação na planilha de Duplicidade Mobbuy
                resultado_mobbuy = planilha_duplicidade[planilha_duplicidade['CNPJ/CPF'] == cnpj]
                df.at[index, 'MOBBUY'] = 'SIM' if not resultado_mobbuy.empty else 'NÃO'
                log_message(f"Resultado Mobbuy: {str(resultado_mobbuy)}")

                # Verificação de RECUPERADO_BNC
                resultado_gerencial = planilha_bnc[planilha_bnc['CPF/CNPJ'] == cnpj_formatado]
                df.at[index, 'RECUPERADO_BNC'] = 'SIM' if not resultado_gerencial.empty else 'NÃO'
                log_message(f"Resultado Gerencial: {str(resultado_gerencial)}")

            except NoSuchElementException:
                log_message("Sem cadastro encontrado.")
                df.at[index, 'STATUS ADC LOJISTA'] = 'NÃO CADASTRADO'
                df.at[index, 'STATUS DO CBK ADIQ'] = 'NÃO SE APLICA'
                df.at[index, 'STATUS DO CBK GLOBAL'] = 'NÃO SE APLICA'
                df.at[index, 'MOBBUY'] = 'NÃO SE APLICA'
                df.at[index, 'RECUPERADO_BNC'] = 'NÃO SE APLICA'

            # Verificação de PENDÊNCIA PORTADOR
            driver.get(CONSULTA_CPF)
            username_field = driver.find_element(By.ID, 'txtUsuarioLogin')
            username_field.send_keys(adc_username)
            password_field = driver.find_element(By.ID, 'txtSenhaLogin')
            password_field.send_keys(adc_password)
            login_button = driver.find_element(By.XPATH, "//*[@id='btnOkLogin']")
            login_button.click()
            sleep(5)

            driver.get(CPF_REPORT)
            driver.find_element(By.ID, 'txtValor').send_keys(cnpj_formatado)
            driver.find_element(By.ID, 'btnPesquisar').click()
            sleep(8)
            try:
                status = driver.find_element(By.ID, 'lblStatusAutomatico').text
                if re.match(r'^Bloqueado - Inadimplencia - \d+ Dias de Atraso $', status):
                    df.at[index, 'PENDÊNCIA PORTADOR'] = 'SIM'
                    log_message(f"Portador com pendência: {status}")
                elif re.match(r'^Bloqueado - Divida Renegociada em \d{2}/\d{2}/\d{4}$', status):
                    df.at[index, 'PENDÊNCIA PORTADOR'] = 'SIM'
                else:
                    df.at[index, 'PENDÊNCIA PORTADOR'] = 'NÃO'
            except NoSuchElementException:
                df.at[index, 'PENDÊNCIA PORTADOR'] = 'NÃO SE APLICA'

        except Exception as e:
            log_message(f"Erro ao processar {cnpj_formatado}: {e}", level='error')

    # Salvar a planilha final
    df.to_excel(PLANILHA_PATH, index=False)
    driver.quit()
    log_message("Processo concluído!")

# Iteração pelos dados da planilha
# Ajuste para os nomes corretos das colunas de CPF
cpf_columns = [
    'CPF REPRESENTANTE 1', 
    'CPF REPRESENTANTE 1.1', 
    'CPF REPRESENTANTE 1.2', 
    'CPF REPRESENTANTE 1.3', 
    'CPF REPRESENTANTE 1.4', 
    'CPF REPRESENTANTE 1.5'
]

log_message(f"Colunas disponíveis na planilha: {df.columns.tolist()}")

for id in range(len(df)):
    try:
        CNPJ_formatado = df.loc[id, 'CNPJ/CPF']

        if pd.isna(CNPJ_formatado):
            log_message(f"Skipping ID {id}: Missing CNPJ/CPF", level='warning')
            continue

        CNPJ = re.sub(r'[.\-\/]', '', str(CNPJ_formatado))
        Razao_Social = df.loc[id, 'FORNECEDOR'] if pd.notna(df.loc[id, 'FORNECEDOR']) else 'FORNECEDOR NÃO INFORMADO'

        # Coletar os CPFs dos representantes
        CPFs = []
        for coluna_cpf in cpf_columns:
            if coluna_cpf in df.columns:
                cpf_value = df.loc[id, coluna_cpf]
                if pd.notna(cpf_value):
                    cpf_cleaned = re.sub(r'[.\-]', '', str(cpf_value)).strip()
                    if cpf_cleaned.isdigit() and len(cpf_cleaned) == 11:
                        CPFs.append(cpf_cleaned)
                        log_message(f"CPF válido encontrado: {cpf_cleaned} (Coluna: {coluna_cpf})")
                    else:
                        log_message(f"CPF inválido ou mal formatado: {cpf_value} (Coluna: {coluna_cpf})", level='warning')

        if not CPFs:
            log_message(f"Nenhum CPF encontrado para o ID {id}. Continuando com CNPJ apenas.", level='warning')

        # Chamada do processo
        iniciar_processo(CNPJ_formatado, CNPJ, CPFs, CNPJ, user_name, password, Razao_Social)

    except Exception as e:
        log_message(f"Erro ao processar o ID {id}: {e}", level='error')

        continue

log_message("Processo finalizado com sucesso!")


conciliar_planilhas()



