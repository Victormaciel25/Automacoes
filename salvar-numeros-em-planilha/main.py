import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, ElementClickInterceptedException

# Configuração do Google Sheets
def connect_google_sheets(teste01):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\vcr00\Documents\Projects GitHub\Automacoes\salvar-numeros-em-planilha\credentials.json", scope)
    client = gspread.authorize(creds)
    print("Conectando ao Google Sheets...")
    try:
        sheet = client.open(teste01).sheet1 # Abre a primeira aba da planilha
    except Exception as e:
        print(f"Erro ao conectar com a planilha: {e}")
    return sheet

time.sleep(1)

# Função para copiar nomes e números do WhatsApp Web
def get_whatsapp_contacts(driver):
    contacts = []
    
    # Espera até que os chats estejam presentes na página
    WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'span[title]')))
    
    # Obtém os contatos da barra lateral de conversas
    chat_elements = driver.find_elements(By.CSS_SELECTOR, 'span[title]')  # Atributo "title" contém o nome
    
    for chat in chat_elements:
        try:
            # Espera até que o chat esteja clicável
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(chat))
            
            # Tenta clicar no chat
            name = chat.get_attribute("title")
            try:
                chat.click()  # Abre a conversa para pegar o número
            except ElementClickInterceptedException:
                print(f"Elemento bloqueado, tentando fechar diálogos...")
                # Fecha qualquer possível pop-up
                close_dialogs(driver)
                chat.click()  # Tenta clicar novamente
            
            # Espera carregar o chat
            time.sleep(2)
            
            # Número de telefone do contato
            number_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'header span[title]'))
            )
            number = number_element.get_attribute("title")
            
            contacts.append((name, number))
        
        except StaleElementReferenceException:
            print(f"Elemento {chat} ficou 'stale', tentando novamente...")
            chat = driver.find_element(By.CSS_SELECTOR, f'span[title="{name}"]')
            chat.click()
            time.sleep(2)
            
            number_element = driver.find_element(By.CSS_SELECTOR, 'header span[title]')
            number = number_element.get_attribute("title")
            contacts.append((name, number))
        except Exception as e:
            print(f"Erro ao processar o contato: {e}")
            continue
    
    return contacts

# Função para fechar possíveis diálogos ou pop-ups
def close_dialogs(driver):
    try:
        close_buttons = driver.find_elements(By.CSS_SELECTOR, 'div[role="dialog"] .close-button')  # Ajuste o seletor conforme necessário
        for button in close_buttons:
            button.click()
    except Exception as e:
        print(f"Erro ao tentar fechar diálogos: {e}")

# Função para atualizar os dados na planilha
def update_google_sheet(sheet, contacts):
    for i, contact in enumerate(contacts, start=2):  # Começa na linha 2 (assumindo cabeçalho na linha 1)
        sheet.update_cell(i, 1, contact[0])  # Nome na coluna 1
        sheet.update_cell(i, 2, contact[1])  # Número na coluna 2

# Configuração do Selenium (WhatsApp Web)
def start_whatsapp_driver():
     # Definir o caminho do chromedriver usando Service
    chrome_service = Service(r'C:\Users\vcr00\Documents\Projects GitHub\Automacoes\chromedriver.exe')
    
    # Iniciar o WebDriver com o serviço configurado
    driver = webdriver.Chrome(service=chrome_service)
    
    driver.get('https://web.whatsapp.com')
    
    # Espera pelo usuário escanear o QR code
    input("Pressione Enter após escanear o QR code no WhatsApp Web")
    
    return driver

if __name__ == "__main__":
    # Conectar ao Google Sheets
    sheet = connect_google_sheets("teste01")
    
    # Iniciar o WhatsApp Web via Selenium
    print("Iniciando o WhatsApp Web via Selenium...")
    driver = start_whatsapp_driver()

    time.sleep(1)
    
    # Obter os contatos do WhatsApp
    print("Obtendo os contatos do WhatsApp...")
    contacts = get_whatsapp_contacts(driver)

    time.sleep(1)
    
    # Atualizar a planilha com os contatos
    print("Atualizando a planilha com os contatos...")
    update_google_sheet(sheet, contacts)

    time.sleep(1)
    
    # Fechar o navegador
    print("Fechando o navegador...")
    driver.quit()
