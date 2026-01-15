"""
SWIFT Financial Messaging Automation Suite
------------------------------------------
Description: Enterprise-grade pipeline for automating SWIFT (MT/MX) message processing.
Features:
  - Secure SFTP extraction using Paramiko.
  - AES-256 Encryption for credential management (Fernet).
  - Hybrid Parsing Engine: Regex (MT) & XML ElementTree (ISO 20022).
  - Automated Reporting: PDF generation via ReportLab.
  - Cloud Synchronization: Real-time integration with AppSheet API.
  - Financial Normalization: Intelligent amount formatting for global currencies.

Author: Johann Espinosa
Version: 3.1.2
Environment: Production
"""

import os
import shutil
import re
import time
import base64
import getpass
import io
import paramiko
import textwrap
import requests
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation

# Third-party libraries
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Preformatted
from reportlab.lib.styles import ParagraphStyle
from dotenv import load_dotenv
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from openpyxl import Workbook, load_workbook

# --- CONFIGURATION & CONSTANTS ---

# Base directories (Relative paths for portability)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
INPUT_DIR = os.path.join(DATA_DIR, "inputs")
OUTPUT_DIR = os.path.join(DATA_DIR, "outputs")
BACKUP_DIR = os.path.join(DATA_DIR, "backup")
LOG_DIR = os.path.join(DATA_DIR, "logs")

# Ensure directories exist
for path in [DATA_DIR, INPUT_DIR, OUTPUT_DIR, BACKUP_DIR, LOG_DIR]:
    os.makedirs(path, exist_ok=True)

# Security Files
ARCHIVO_ENV = ".env"
ARCHIVO_ENC = ".env.enc"
SEQUENCE_FILE = os.path.join(DATA_DIR, "sequence.txt")
RUTA_EXCEL_LOCAL = os.path.join(DATA_DIR, "SWIFT_Transaction_Log.xlsx")

# Reporting Styles
ESTILO_PDF = ParagraphStyle(name='PreformattedStyle', fontName='Courier', fontSize=10, leading=12)

# Business Logic Mapping (Anonymized Categories)
MAPEO_CATEGORIAS = {
    "RESERVICIOSINTERNAC": "IN_INTL_SERVICES",
    "TRSERVICIOSINTERNAC": "OUT_INTL_SERVICES",
    "REINVERSIONES": "IN_INVESTMENTS",
    "TRINVERSIONES": "OUT_INVESTMENTS",
    "REDEUDA": "IN_DEBT_SERVICE",
    "TRDEUDA": "OUT_DEBT_SERVICE",
    "RECARTASDECREDITO": "IN_LETTERS_OF_CREDIT",
    "TRCARTASDECREDITO": "OUT_LETTERS_OF_CREDIT"
}

# MT950 Pattern Matching (Business Rules)
PATRONES_MT950 = {
    "FIN 950_BISBCHBBXXX": "BIS_STATEMENTS",
    "FIN 950_FRNYUS33XXX": "FED_RESERVE_STATEMENTS",
    "FIN 950_FLARCOBBXXX": "FLAR_STATEMENTS",
    "FIN 950_COBADEFFXXX": "COMMERZBANK_STATEMENTS",
    "FIN 950_BLAEPAPAXXX": "BLADEX_STATEMENTS",
    "FIN 950_BKCHPAPAXXX": "BOC_STATEMENTS",
    "FIN 950_MGTCBEBEXXX": "EUROCLEAR_STATEMENTS"
}

# --- SECURITY MODULE ---

def obtener_llave_desde_password(password, salt):
    """Derives a secure cryptographic key using PBKDF2HMAC."""
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100_000,
    )
    return base64.urlsafe_b64encode(kdf.derive(password.encode()))

def inicializar_seguridad():
    """
    Handles credential decryption and enforces mandatory key rotation.
    Protects sensitive environment variables (.env).
    """
    datos_secretos = None
    if not os.path.exists(ARCHIVO_ENC):
        if not os.path.exists(ARCHIVO_ENV):
            print(f"❌ CRITICAL ERROR: Configuration file {ARCHIVO_ENV} not found.")
            exit(1)
        print(f"⚠️ PLAIN TEXT DETECTED. INITIALIZING ENCRYPTION PROTOCOL.")
        with open(ARCHIVO_ENV, "rb") as f:
            datos_secretos = f.read()
    else:
        print("\n🔐 SECURITY CHECK: Credential validation required.")
        password_ayer = getpass.getpass("👉 Enter active security passphrase: ")
        with open(ARCHIVO_ENC, "rb") as f:
            contenido = f.read()
        salt_leido = contenido[:16]
        data_encriptada = contenido[16:]
        try:
            key_ayer = obtener_llave_desde_password(password_ayer, salt_leido)
            fernet = Fernet(key_ayer)
            datos_secretos = fernet.decrypt(data_encriptada)
            print("✅ Access Granted. Credentials loaded.")
        except Exception:
            print("❌ ACCESS DENIED: Invalid passphrase.")
            exit(1)

    load_dotenv(stream=io.StringIO(datos_secretos.decode('utf-8')))
    
    # Mandatory Key Rotation
    print("\n🔄 SECURITY ROTATION (Required for next cycle)")
    while True:
        password_nueva = getpass.getpass("👉 Set new passphrase: ")
        password_confirm = getpass.getpass("👉 Confirm new passphrase: ")
        if password_nueva == password_confirm and len(password_nueva) > 0:
            break
        print("❌ Passphrases do not match.")

    nuevo_salt = os.urandom(16)
    nueva_key = obtener_llave_desde_password(password_nueva, nuevo_salt)
    nueva_data = Fernet(nueva_key).encrypt(datos_secretos)
    
    with open(ARCHIVO_ENC, "wb") as f:
        f.write(nuevo_salt + nueva_data)
    
    # Secure cleanup
    if os.path.exists(ARCHIVO_ENV):
        os.remove(ARCHIVO_ENV)
        print("🗑️ Plaintext credentials removed secure.")
    print("✅ Security protocols updated.\n")

# --- SFTP MODULE ---

def descargar_desde_sftp():
    """Retrieves financial files from secure SFTP server."""
    HOST = os.getenv('SFTP_HOST')
    PORT = int(os.getenv('SFTP_PORT', 22))
    USER = os.getenv('SFTP_USER')
    PASS = os.getenv('SFTP_PASS')

    if not all([HOST, USER, PASS]):
        print("❌ ERROR: Missing SFTP configuration.")
        return

    # Task definition: (Remote Path, Local Path, Extension)
    # Using environment variables for remote paths to abstract infrastructure
    REMOTE_BASE = os.getenv('SFTP_REMOTE_BASE', '/swift/data')
    
    tareas = [
        (f"{REMOTE_BASE}/PdfServInter/pdfs", INPUT_DIR, ".pdf"),
        (f"{REMOTE_BASE}/PdfDeuda/pdfs", INPUT_DIR, ".pdf"),
        (f"{REMOTE_BASE}/Pdf_SSFI", INPUT_DIR, ".prt")
    ]
    
    transport = None
    sftp = None
    try:
        transport = paramiko.Transport((HOST, PORT))
        transport.connect(username=USER, password=PASS)
        sftp = paramiko.SFTPClient.from_transport(transport)
        
        for ruta_remota, ruta_local, extension in tareas:
            try:
                try:
                    archivos_en_nube = sftp.listdir(ruta_remota)
                except FileNotFoundError:
                    continue

                archivos_filtrados = [f for f in archivos_en_nube if f.lower().endswith(extension)]
                
                if archivos_filtrados:
                    print(f"⬇️ Downloading {len(archivos_filtrados)} files from {ruta_remota}...")
                    for archivo in archivos_filtrados:
                        remote_file_path = f"{ruta_remota}/{archivo}"
                        local_file_path = os.path.join(ruta_local, archivo)
                        
                        if not os.path.exists(local_file_path):
                            sftp.get(remote_file_path, local_file_path)
                            try:
                                sftp.remove(remote_file_path) # Clean up remote after success
                            except Exception as e:
                                print(f"⚠️ Warning: Could not remove remote file {archivo}: {e}")
            except Exception as e:
                print(f"⚠️ Error processing remote folder {ruta_remota}: {e}")
    except Exception as e:
        print(f"❌ SFTP Connection Error: {e}")
    finally:
        if sftp: sftp.close()
        if transport: transport.close()

# --- CLOUD INTEGRATION (APPSHEET) ---

def enviar_a_appsheet_api(datos):
    """Push processed transaction data to AppSheet via REST API."""
    APP_ID = os.getenv('APPSHEET_APP_ID')
    ACCESS_KEY = os.getenv('APPSHEET_ACCESS_KEY')
    TABLE_NAME = os.getenv('APPSHEET_TABLE_NAME')

    if not all([APP_ID, ACCESS_KEY, TABLE_NAME]):
        # Fail silently if not configured (optional feature)
        return

    url = f"https://api.appsheet.com/api/v2/apps/{APP_ID}/tables/{TABLE_NAME}/Action"
    headers = {"ApplicationAccessKey": ACCESS_KEY, "Content-Type": "application/json"}

    def formatear_para_appsheet(valor):
        try:
            return str(float(valor)).replace('.', ',')
        except:
            return "0,0"

    fila = {
        "Fecha Carga": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Nombre Archivo": datos.get('archivo', ''),
        "Tipo Mensaje": datos.get('tipo', ''),
        "Referencia Principal": datos.get('ref', ''),
        "BIC Sender": datos.get('bic', ''),
        "Monto 1": formatear_para_appsheet(datos.get('monto_1', 0)),
        "Tercero": datos.get('tercero', '')
    }

    payload = {
        "Action": "Add",
        "Properties": {"Locale": "es-EC", "Timezone": "America/Guayaquil"},
        "Rows": [fila]
    }

    try:
        response = requests.post(url, headers=headers, json=payload)
        if response.status_code == 200:
            print(f"🚀 Synced with AppSheet: {datos.get('ref')}")
        else:
            print(f"⚠️ AppSheet API Error: {response.text}")
    except Exception as e:
        print(f"⚠️ API Connection failed: {e}")

# --- DATA PARSING & NORMALIZATION ---

def detectar_tipo_mensaje(input_file):
    """Identify SWIFT message standard (MT vs MX/ISO20022)."""
    try:
        with open(input_file, 'r', encoding='utf-8', errors='ignore') as f:
            contenido = f.read(1000).lower()
    except Exception:
        return "UNKNOWN"

    es_fin_explicito = re.search(r'swift (output|input)\s*:\s*fin \d{3}', contenido)
    tiene_formato_mx = "primary format : mx" in contenido
    tiene_iso = "businessservice: swift.cbprplus" in contenido

    if (tiene_formato_mx or tiene_iso) and not ("secondary format : mx" in contenido):
        return "MX"
    return "MT"

def limpiar_monto_inteligente(valor):
    """
    Intelligent normalization for financial amounts.
    Handles EU (1.000,00) vs US (1,000.00) formats automatically.
    """
    if not valor or str(valor) == 'UNKNOWN_AMOUNT': return 0.0
    s = str(valor).strip()
    try:
        # Mixed separators case (e.g. 1.234,56 or 1,234.56)
        if '.' in s and ',' in s:
            if s.rfind(',') > s.rfind('.'): # Coma at end -> European
                s = s.replace('.', '').replace(',', '.')
            else: # Dot at end -> American
                s = s.replace(',', '')
        # Only Coma (1234,56) -> European
        elif ',' in s:
            s = s.replace(',', '.')
        # Only Dot (1.234) -> Ambiguous, verify decimal places
        elif '.' in s and s.count('.') > 1:
            s = s.replace('.', '')
        return float(s)
    except:
        return 0.0

def guardar_en_excel_local(datos):
    """Appends transaction logs to a local Excel database."""
    ENCABEZADOS = ["Fecha Carga", "Nombre Archivo", "Formato", "Tipo", "Referencia", "BIC", "Monto 1", "Tercero"]
    
    fila = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        datos.get('archivo', ''), datos.get('formato', ''), datos.get('tipo', ''),
        datos.get('ref', ''), datos.get('bic', ''),
        datos.get('monto_1', 0.0), datos.get('tercero', '')
    ]

    try:
        if not os.path.exists(RUTA_EXCEL_LOCAL):
            wb = Workbook()
            ws = wb.active
            ws.title = "SWIFT_Logs"
            ws.append(ENCABEZADOS)
        else:
            wb = load_workbook(RUTA_EXCEL_LOCAL)
            ws = wb.active
        
        ws.append(fila)
        wb.save(RUTA_EXCEL_LOCAL)
        print(f"📗 Logged to Excel: {datos.get('ref')}")
    except PermissionError:
        print("❌ ERROR: Excel file is open. Close it to proceed.")
    except Exception as e:
        print(f"⚠️ Excel Error: {e}")

# --- PARSING ENGINES (MT & MX) ---

def extract_data_from_iso20022(input_file):
    """
    Parses ISO 20022 XML content to extract financial metadata.
    Uses Regex for high-performance extraction on large files.
    """
    with open(input_file, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    # Isolate XML payload if wrapped
    json_start = content.find('{')
    json_end = content.find('}', json_start)
    if json_start != -1 and json_end != -1:
        content = content[:json_start] + content[json_end + 1:]

    # Extraction Logic
    message_type = (re.search(r'MessageDefinitionIdentifier\s*[:>]\s*([\w.]+)', content) or 
                    type('obj', (object,), {'group': lambda x: "UNKNOWN"})).group(1)
    
    bic_sender = (re.search(r'From[\s\S]{0,2000}?BICFI\s*[:>\s]*([A-Z0-9]{8,11})', content, re.IGNORECASE) or 
                  type('obj', (object,), {'group': lambda x: "UNKNOWN"})).group(1)
    
    reference = (re.search(r'<MsgId>([^<]+)</MsgId>', content) or 
                 re.search(r'MessageIdentification\s*[:>]\s*([A-Z0-9\-]+)', content) or
                 type('obj', (object,), {'group': lambda x: "UNKNOWN"})).group(1)

    # Amount extraction strategies
    raw_amounts = re.findall(r'Amount\s*[:>]\s*([\d.,]+)', content)
    if not raw_amounts:
        raw_amounts = re.findall(r'<InstdAmt[^>]*>([\d.,]+)</InstdAmt>', content)
    amount = raw_amounts[-1] if raw_amounts else "0.00"

    return {
        "message_type": message_type,
        "bic_sender": bic_sender,
        "amount": amount,
        "reference": reference,
        "direction": "MX",
        "account": "N/A", # Simplified for generic version
        "debtor_name": "N/A"
    }

def extract_data_from_lines(lines):
    """Parses legacy MT message format (FIN)."""
    # Simplified regex extraction for demonstration
    type_message = 'UNKNOWN'
    bic = 'UNKNOWN'
    amount = '0.00'
    reference = 'UNKNOWN'
    
    for line in lines:
        if 'Swift Output' in line or 'Swift Input' in line:
            m = re.search(r'FIN (\d{3})', line)
            if m: type_message = f"FIN {m.group(1)}"
        if 'Sender :' in line:
            m = re.search(r'Sender\s*:\s+(\S+)', line)
            if m: bic = m.group(1)
        if '20:' in line and reference == 'UNKNOWN':
             # Logic to get next line for ref
             pass 
        if 'Amount :' in line:
             m = re.search(r'Amount\s*:\s*#([\d.,]+)#', line)
             if m: amount = m.group(1)
             
    return type_message, bic, amount, reference

# --- PDF GENERATION ---

def get_next_sequence_number(filepath):
    """Manages sequential file naming for daily reconciliations."""
    hoy_str = datetime.now().strftime("%Y-%m-%d")
    numero = 1
    if os.path.exists(filepath):
        try:
            with open(filepath, "r") as f:
                data = f.read().strip().split(",")
                if len(data) == 2 and data[0] == hoy_str:
                    numero = int(data[1]) + 1
        except: pass
        
    with open(filepath, "w") as f:
        f.write(f"{hoy_str},{numero}")
    return str(numero).zfill(3)

def txt_to_pdf(input_file, output_folder, backup_folder, sequence_file, data_dict=None):
    """Converts raw message text to a standardized PDF report."""
    try:
        with open(input_file, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
    except Exception as e:
        print(f"Error reading {input_file}: {e}")
        return

    # Clean lines for PDF
    wrapped_lines = []
    for line in lines:
        clean = line.replace('\t', '    ').rstrip()
        wrapped_lines.extend(textwrap.wrap(clean, width=80))

    seq = get_next_sequence_number(sequence_file)
    current_date = datetime.now().strftime("%Y-%m-%d")
    
    # Dynamic Naming
    ref = data_dict.get('ref', 'UNK') if data_dict else 'UNK'
    tipo = data_dict.get('tipo', 'MSG') if data_dict else 'MSG'
    pdf_name = f"{tipo}_{ref}_({seq})_({current_date}).pdf".replace(":", "-").replace("/", "-")
    
    output_path = os.path.join(output_folder, pdf_name)
    
    try:
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        story = [Preformatted(l, ESTILO_PDF) for l in wrapped_lines]
        doc.build(story)
        
        # Move original to backup
        shutil.move(input_file, os.path.join(backup_folder, os.path.basename(input_file)))
        return pdf_name
    except Exception as e:
        print(f"❌ Error generating PDF: {e}")

# --- MAIN CONTROLLER ---

def process_file(input_file):
    """Orchestrates the processing of a single file."""
    tipo = detectar_tipo_mensaje(input_file)
    nombre = os.path.basename(input_file)
    
    datos = {
        'archivo': nombre, 
        'formato': tipo,
        'monto_1': 0.0,
        'monto_2': 0.0
    }

    if tipo == "MX":
        extracted = extract_data_from_iso20022(input_file)
        datos.update({
            'tipo': extracted['message_type'],
            'ref': extracted['reference'],
            'bic': extracted['bic_sender'],
            'monto_1': limpiar_monto_inteligente(extracted['amount'])
        })
        # PDF Generation for MX
        txt_to_pdf(input_file, OUTPUT_DIR, BACKUP_DIR, SEQUENCE_FILE, datos)

    elif tipo == "MT":
        with open(input_file, 'r') as f: lines = f.readlines()
        t_msg, bic, amt, ref = extract_data_from_lines(lines)
        datos.update({
            'tipo': t_msg,
            'ref': ref,
            'bic': bic,
            'monto_1': limpiar_monto_inteligente(amt)
        })
        # PDF Generation for MT
        txt_to_pdf(input_file, OUTPUT_DIR, BACKUP_DIR, SEQUENCE_FILE, datos)

    # Persistence
    guardar_en_excel_local(datos)
    enviar_a_appsheet_api(datos)

if __name__ == "__main__":
    inicializar_seguridad()
    
    print("\n✅ SWIFT AUTOMATION ENGINE STARTED")
    print(f"📂 Watching directory: {INPUT_DIR}")
    print(f"🔄 Cycle Interval: 5 minutes\n")

    while True:
        print(f"▶️ CYCLE START: {datetime.now().strftime('%H:%M:%S')}")
        
        try:
            # 1. Download Phase
            descargar_desde_sftp()
            
            # 2. Processing Phase
            archivos = os.listdir(INPUT_DIR)
            if archivos:
                print(f"⚡ Processing {len(archivos)} files...")
                for f in archivos:
                    full_path = os.path.join(INPUT_DIR, f)
                    if os.path.isfile(full_path):
                        process_file(full_path)
            else:
                print("💤 No new files pending.")

            # 3. Cleanup/Rotation logic would go here
            
            print("✅ Cycle Complete. Sleeping...")
            time.sleep(300)

        except KeyboardInterrupt:
            print("\n🛑 Execution stopped by user.")
            break
        except Exception as e:
            print(f"⚠️ UNEXPECTED ERROR: {e}")
            time.sleep(60)