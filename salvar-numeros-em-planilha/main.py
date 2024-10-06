import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Configuração do Google Sheets
def connect_google_sheets(sheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\vcr00\Documents\Projects GitHub\Automacoes\salvar-numeros-em-planilhacredentials.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).sheet1  # Abre a primeira aba da planilha
    return sheet

# Função para copiar nomes e números do WhatsApp Web
def get_whatsapp_contacts(driver):
    contacts = []
    
    # Exemplo: Obtém os contatos da barra lateral de conversas
    chat_elements = driver.find_elements(By.CSS_SELECTOR, 'span[title]')  # Atributo "title" contém o nome
    
    for chat in chat_elements:
        name = chat.get_attribute("title")
        chat.click()  # Abre a conversa para pegar o número
        
        # Espera carregar o chat
        time.sleep(2)
        
        try:
            # Número de telefone do contato
            number_element = driver.find_element(By.CSS_SELECTOR, 'header span[title]')
            number = number_element.get_attribute("title")
        except:
            number = "Número não encontrado"
        
        contacts.append((name, number))
    
    return contacts

# Função para atualizar os dados na planilha
def update_google_sheet(sheet, contacts):
    for i, contact in enumerate(contacts, start=2):  # Começa na linha 2 (assumindo cabeçalho na linha 1)
        sheet.update_cell(i, 1, contact[0])  # Nome na coluna 1
        sheet.update_cell(i, 2, contact[1])  # Número na coluna 2

# Configuração do Selenium (WhatsApp Web)
def start_whatsapp_driver():
    driver = webdriver.Chrome(executable_path='caminho_para_o_chromedriver')
    driver.get('https://web.whatsapp.com')
    
    # Espera pelo usuário escanear o QR code
    input("Pressione Enter após escanear o QR code no WhatsApp Web")
    
    return driver

if __name__ == "__main__":
    # Conectar ao Google Sheets
    sheet = connect_google_sheets("teste")
    
    # Iniciar o WhatsApp Web via Selenium
    driver = start_whatsapp_driver()
    
    # Obter os contatos do WhatsApp
    contacts = get_whatsapp_contacts(driver)
    
    # Atualizar a planilha com os contatos
    update_google_sheet(sheet, contacts)
    
    # Fechar o navegador
    driver.quit()
