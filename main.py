from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path
import pandas as pd

# Alcances para acceder al directorio global
SCOPES = ['https://www.googleapis.com/auth/directory.readonly']


def cargar_nombres_desde_excel(archivo_excel):
    # Leer el archivo Excel
    df = pd.read_excel(archivo_excel, skiprows=11) # Alumno empieza en la fila 12

    # Obtener los valores de la columna "alumno" y almacenarlos en un array
    nombres_a_buscar = df['ALUMNO'].dropna().tolist()

    return nombres_a_buscar

def combinar_excels_con_array(archivo_legajos, datos_contactos, archivo_salida):
    # Cargar los datos de nombres y legajos desde el archivo Excel
    df_nombres_legajos = cargar_nombres_desde_excel(archivo_legajos)

    # Convertir el array de datos de contactos a un DataFrame
    df_emails = pd.DataFrame(datos_contactos)

    # Renombrar la columna 'alumno' a 'Nombre' y asegurarnos de que coincidan con 'Name'
    df_nombres_legajos = df_nombres_legajos.rename(columns={'alumno': 'Nombre'})
    df_emails = df_emails.rename(columns={'Name': 'Nombre'})

    # Combinar los DataFrames en base a la columna 'Nombre'
    df_combinado = pd.merge(df_nombres_legajos, df_emails, on='Nombre', how='left')

    # Guardar el DataFrame combinado en un nuevo archivo Excel
    df_combinado.to_excel(archivo_salida, index=False)


archivo_excel = 'los-3-listados-siu.xls' # Aca va la ruta del archivo
nombres_a_buscar = cargar_nombres_desde_excel(archivo_excel)

def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # Construir el servicio para la API de People
    service = build('people', 'v1', credentials=creds)

    # Lista para almacenar los datos
    datos_contactos = []

    # Buscar contactos para cada nombre en el array
    for nombre in nombres_a_buscar:
        results = service.people().searchDirectoryPeople(
            query=nombre,
            sources=['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'],
            readMask='names,emailAddresses').execute()
        people = results.get('people', [])

        # Recopilar los datos en la lista
        for person in people:
            names = person.get('names', [])
            email_addresses = person.get('emailAddresses', [])
            if names:
                display_name = names[0].get('displayName')
                email = email_addresses[0].get('value') if email_addresses else 'No Email'
                datos_contactos.append({'Name': display_name, 'Email': email})


    # Crear un DataFrame con los datos recopilados
    df = pd.DataFrame(datos_contactos)

    # Guardar los datos en un archivo Excel
    df.to_excel('contactos_universidad.xlsx', index=False)

if __name__ == '__main__':
    main()
