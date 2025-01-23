from playwright.sync_api import sync_playwright
import time
import os
from datetime import datetime, timedelta

def on_download(download):
    """
    Callback para tratar o download do arquivo.
    Verifica se o arquivo tem extensão .xlsx e salva no local especificado.
    """
    try:
        # Calcula a data de ontem e formata como string
        yesterday = (datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")
        
        # Verifica se o arquivo baixado tem extensão .xlsx
        if download.suggested_filename.endswith(".xlsx"):
            # Define o caminho e o nome do arquivo a ser salvo com a data de ontem
            download_path = os.path.join(
                "C:/Users/FreddieCeribelli/LoyaltyCom/TI LOYALTYCOM - Documentos/6 - CHAT BOT/POWER BI ATENDIMENTO/DADOS", f"{yesterday}.xlsx"
            )
            print(f"Download iniciado: {download.url}")
            download.save_as(download_path)
            print(f"Arquivo salvo em: {download_path}")
        else:
            # Caso não seja um Excel, salva o arquivo com o nome sugerido pelo servidor
            download_path = os.path.join("C:/Users/FreddieCeribelli/LoyaltyCom/TI LOYALTYCOM - Documentos/6 - CHAT BOT/POWER BI ATENDIMENTO/DADOS", download.suggested_filename)
            print(f"Aviso: Arquivo baixado não parece ser Excel. Salvando como: {download.suggested_filename}")
            download.save_as(download_path)
    except Exception as e:
        print(f"Erro ao processar o download: {e}")

# Início do script Playwright
with sync_playwright() as p:
    # Lança o navegador
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(
        accept_downloads=True  # Habilita o download de arquivos
    )
    
    # Cria uma nova página
    page = context.new_page()

    # Registra o evento de download
    page.on("download", on_download)

    # Navega para a URL desejada
    page.goto("https://salesiq.zoho.com/loyaltycom/settings/home")

    # Aguarda a tela carregar
    page.wait_for_load_state("load")

    # Realiza o login
    page.click("xpath=//*[@id='zw-template-inner']/div[1]/div[2]/div/div[4]/a")
    page.wait_for_load_state("load")
    page.fill("xpath=//*[@id='login_id']", "contato@loyaltycom.com.br")
    page.click("xpath=//*[@id='nextbtn']")
    page.wait_for_load_state("load")
    page.fill("xpath=//*[@id='password']", "LoyaltyCom@24")
    page.click("xpath=//*[@id='nextbtn']")
    page.wait_for_load_state("load")

    # Navega para a seção de chats
    page.click("xpath=//*[@id='siq-left-nav']/nav/ul/li[4]")
    page.wait_for_load_state("load")

    # Clica no filtro de chats fechados
    time.sleep(5)
    page.click("xpath=//*[@id='content']/div/div/div[1]/div[1]/div[1]")
    time.sleep(5)
    page.click("xpath=//*[@id='Closed']/div")
    time.sleep(5)
    page.click("xpath=//*[@id='toggler']")
    time.sleep(5)
    page.click("xpath=//*[@id='content']/div/div/div[1]/div[1]/div[4]/div[1]/div/div[1]/div[1]/input")
    time.sleep(5)
    page.click("xpath=//*[@id='content']/div/div/div[1]/div[1]/div[4]/div[1]/div/div[2]/div/div[2]/div/div/div/div[2]/span")
    time.sleep(5)
    page.click("xpath=//*[@id='content']/div/div/div[1]/div/div[4]/div[1]/div/footer/div/div[2]")
    time.sleep(5)
    page.click("xpath=//*[@id='content']/div/div/div[1]/div/div[4]/div[2]")
    
    # Clica no botão para baixar o arquivo
    download_button = page.locator("xpath=//*[@id='xlsx']/div")
    download_button.click()

    # Espera um tempo para garantir que o arquivo seja baixado
    time.sleep(10)
    
    # Fecha o navegador
    browser.close()
