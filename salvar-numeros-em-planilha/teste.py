import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1BRHCRnA6g6zm4pOC2Aqb_QawYpBmBTBuqQfhTsI2e4s"
SAMPLE_RANGE_NAME = "Página1!A1:B12"


def main():
  """Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.
  """
  creds = None
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          r"C:\Users\vcr00\Documents\Projects GitHub\Automacoes\salvar-numeros-em-planilha\credentialsOAuth.json", SCOPES)
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)

    # Ler informações do Google sheets
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    valores = result['values']
    print(valores)
  
    valores_adicionar = [
      ['imposto'],
      ]
    for i, linha in enumerate(valores):
      if i > 0:
        vendas = linha[1]
        vendas = float(vendas.replace('R$', '').replace('.', '').replace(',','.'))
        imposto = vendas * 0.1
        valores_adicionar.append([imposto])
        print(vendas)

    result = (
        sheet.values()
        .update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='C1', valueInputOption='USER_ENTERED', body={'values':valores_adicionar})
        .execute()
    )

    # Adicionar/editar uma informação

  

  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()