from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path
import pandas as pd
import re
import unicodedata

# Alcances para acceder al directorio global
SCOPES = ['https://www.googleapis.com/auth/directory.readonly']
def main():
    contacts_data = []
    def connectToGoogleApi():
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
        return service

    def get_names_legajo(file):
        # Leer el archivo Excel, comenzando desde la fila 12
        df = pd.read_excel(file, skiprows=11)  # La fila 12 es el índice 11

        # Seleccionar las columnas relevantes
        df_relevante = df[['ALUMNO', 'LEGAJO']]

        # Convertir el DataFrame a una lista de diccionarios
        datos_array = df_relevante.dropna().to_dict(orient='records')

        return datos_array

    def normalize_name(name, normalization_level = 0):
        if(normalization_level==0):
            # Dividir el nombre en partes usando la coma
            if ',' in name:
                last_name, first_names = name.split(', ', 1)
                # Tomar las dos primeras palabras después de la coma
                first_names_words = first_names.split()
                first_two_names = ' '.join(first_names_words[:2])  # Combina las dos primeras palabras
                # Combinar las palabras antes de la coma y las dos primeras después de la coma
                normalized_name = f'{first_two_names} {last_name}'
            else:
                # Si no hay coma, tomar el nombre completo
                normalized_name = name
        elif normalization_level == 1:
            if ',' in name:
                last_name, first_names = name.split(', ', 1)
                # Dividir las palabras después de la coma
                first_names_words = first_names.split()
                # Tomar solo la primera palabra después de la coma
                first_name = first_names_words[0] if first_names_words else ''
                # Combinar todas las palabras antes de la coma y la primera después de la coma
                normalized_name = f'{first_name} {last_name}'
            else:
                # Si no hay coma, tomar el nombre completo
                normalized_name = name


        # Eliminar tildes y normalizar caracteres
        normalized_name = unicodedata.normalize('NFD', normalized_name)
        normalized_name = ''.join(c for c in normalized_name if unicodedata.category(c) != 'Mn')  # Elimina marcas diacríticas

        # Eliminar símbolos especiales
        normalized_name = re.sub(r'[^\w\s]', '', normalized_name)

        # Eliminar espacios extra
        normalized_name = re.sub(r'\s+', ' ', normalized_name).strip()
        return normalized_name

    def generate_contacts_data(student, normalization_level):

        results = service.people().searchDirectoryPeople(
            query=normalize_name(student['ALUMNO'], normalization_level),
            sources=['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'],
            readMask='names,emailAddresses').execute()
        people = results.get('people', [])

        # Recopilar los datos en la lista
        legajo = student['LEGAJO']
        for person in people:
            names = person.get('names', [])
            email_addresses = person.get('emailAddresses', [])

            if names:
                display_name = names[0].get('displayName')
                email = email_addresses[0].get('value') if email_addresses else 'No Email'
                contacts_data.append({'LEGAJO':legajo, 'NOMBRE': display_name, 'EMAIL': email})

    def get_contacts_data(names_to_search, service):
        # Buscar contactos para cada nombre en el array

        for student in names_to_search:
            generate_contacts_data(student, 0)
            if not any(contact['LEGAJO'] == student['LEGAJO'] for contact in contacts_data):
                generate_contacts_data(student, 1)

            if not any(contact['LEGAJO'] == student['LEGAJO'] for contact in contacts_data):
                contacts_data.append({'LEGAJO':student['LEGAJO'], 'NOMBRE': student['ALUMNO'], 'EMAIL': '-'})

    input_file = 'los-3-listados-siu.xls'
    output_file = 'listado_con_mails.xlsx'
    service = connectToGoogleApi()
    names = get_names_legajo(input_file)
    get_contacts_data(names,service)

    # Crear un DataFrame con los datos recopilados
    df = pd.DataFrame(contacts_data)

    # Guardar los datos en un archivo Excel
    df.to_excel(output_file, index=False)

if __name__ == '__main__':
    main()
