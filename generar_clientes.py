import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

# Configuración
SHEET_ID = "12D-Ru8GNm0EE0NFg4kpcygqcPdqKpor9U4NOA-1TQRs"
SHEET_NAME = "Clientes"
CREDENTIALS_FILE = "credentials.json"

# Conectar a Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
gc = gspread.authorize(creds)

# Obtener datos
sheet = gc.open_by_key(SHEET_ID)
worksheet = sheet.worksheet(SHEET_NAME)
records = worksheet.get_all_records()

# Crear diccionario
clientes = {}
for row in records:
    cliente_id = str(row.get("ClienteID", "")).strip()
    if cliente_id:
        clientes[cliente_id] = {
            "nombre": row.get("Nombre", ""),
            "direccion": row.get("Direccion", "")
        }

# Guardar JSON
with open("clientes.json", "w", encoding="utf-8") as f:
    json.dump(clientes, f, ensure_ascii=False, indent=2)

print(f"✅ Clientes guardados: {len(clientes)}")
print(f"📁 Archivo: clientes.json")