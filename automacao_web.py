# -*- coding: utf-8 -*-
import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import config

# MODIFICADO: A função agora aceita um argumento para saber qual botão clicar
def login_e_download(period_button_text="Mês Corrente"):
    """Orquestra a automação web: login, navegação e download usando o filtro especificado."""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless") 
    options.add_argument("--start-maximized")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(options=options)
    
    try:
        wait = WebDriverWait(driver, 20)
        print("Acessando o site...")
        driver.get(config.SITE_URL)

        # Etapa 1: Lidar com o banner de cookies
        try:
            print("Procurando pelo banner de cookies...")
            cookie_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "ok_cookie")))
            cookie_button.click()
            print("Banner de cookies aceito.")
        except TimeoutException:
            print("Banner de cookies não encontrado, continuando...")

        # Etapa 2: Realizar o login
        print("Realizando login...")
        try:
            username_field = wait.until(EC.presence_of_element_located((By.NAME, "LoginName")))
            password_field = driver.find_element(By.NAME, "Password")
            username_field.send_keys(config.SITE_LOGIN)
            password_field.send_keys(config.SITE_SENHA)
            driver.find_element(By.ID, "button_processLogin").click()
            resumo_horas_link = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.LINK_TEXT, "Resumo de Horas por Profissional")))
            print("Login bem-sucedido.")
            resumo_horas_link.click()
        except TimeoutException:
            print("\nERRO CRÍTICO: Login falhou ou a página não carregou corretamente.")
            raise Exception("Não foi possível encontrar o link 'Resumo de Horas por Profissional' após o login.")
        
        # Etapa 3: Mudar para a nova janela de filtros
        print("Aguardando a página de filtros carregar...")
        wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[1])

        # Etapa 4: Aplicar filtro de data
        print(f"Aplicando filtro de data '{period_button_text}'...")
        date_input = wait.until(EC.visibility_of_element_located((By.ID, "P_DATA_show")))
        date_input.click()

        time.sleep(1) 
        
        try:
            print(f"Procurando pela opção '{period_button_text}'...")
            # MODIFICADO: O seletor agora usa a variável para encontrar o botão certo
            xpath_selector = f"//li[normalize-space()='{period_button_text}']"
            period_option = wait.until(EC.presence_of_element_located((By.XPATH, xpath_selector)))
            
            driver.execute_script("arguments[0].click();", period_option)
            print(f"Filtro '{period_button_text}' aplicado com sucesso.")
        except TimeoutException as e:
            print(f"\nERRO CRÍTICO: Não foi possível encontrar a opção '{period_button_text}' no menu de datas.")
            raise e

        # Etapa 5: Baixar o arquivo
        arquivos_antes = set(os.listdir(config.PASTA_DOWNLOADS))
        print("Exportando para Excel...")
        driver.find_element(By.ID, "button_ExecuteXSL").click()
        print("Aguardando download...")
        time.sleep(10)

        # Etapa 6: Monitorar o download
        caminho_arquivo_baixado = None
        start_time = time.time()
        while time.time() - start_time < 60:
            arquivos_depois = set(os.listdir(config.PASTA_DOWNLOADS))
            novos_arquivos = arquivos_depois - arquivos_antes
            if novos_arquivos and not any(f.endswith('.crdownload') for f in novos_arquivos):
                nome_do_arquivo = novos_arquivos.pop()
                caminho_arquivo_baixado = os.path.join(config.PASTA_DOWNLOADS, nome_do_arquivo)
                print(f"Novo arquivo detectado: {nome_do_arquivo}")
                break
            time.sleep(1)

        if not caminho_arquivo_baixado:
            raise TimeoutException("O download do arquivo demorou mais de 60 segundos.")
        
        time.sleep(2)
        print("Download concluído.")
        
        # Etapa 7: Mover o arquivo
        caminho_destino = os.path.join(os.getcwd(), os.path.basename(caminho_arquivo_baixado))
        shutil.move(caminho_arquivo_baixado, caminho_destino)
        print(f"Arquivo movido para a pasta do projeto: {caminho_destino}")
        
        return caminho_destino

    finally:
        if 'driver' in locals() and driver:
            print("Fechando navegador...")
            driver.quit()