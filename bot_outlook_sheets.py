import base64
import io
import os
import re
import sqlite3
import sys
import threading
import time
from datetime import datetime

import customtkinter as ctk
import pdfplumber
import pypdf
import pytesseract
import pythoncom
import requests
import win32com.client
from PIL import Image
from pdf2image import convert_from_bytes
from tkinter import ttk

# --- CONFIGURAÇÃO VISUAL DO CUSTOMTKINTER ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_tesseract_path():
    """Retorna o caminho do tesseract.exe, compatível com PyInstaller."""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = BASE_DIR
    return os.path.join(base_path, "Tesseract-OCR", "tesseract.exe")

pytesseract.pytesseract.tesseract_cmd = get_tesseract_path()
tessdata_dir = os.path.join(os.path.dirname(get_tesseract_path()), "tessdata")
os.environ["TESSDATA_PREFIX"] = tessdata_dir

PARENT_FOLDER_ID = "1cWI8Wqg_njEgvJPuGRLX7ZA50Qkc0lHh"  
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxnx9jb8YAJMJr5VwvPsMlycWRPpCCbK5_ftV_l9nuhDzP1CMAYUQZLpIXsYagLlOz8_A/exec"
DB_PATH = os.path.join(BASE_DIR, "outlook_local_logs.db")
BOT_RODANDO = False 

DICIONARIO_TRIBUTOS = {
    "2172": "2172 - COFINS - CONTRIB P/ FIN. SEG. SOCIAL",
    "2372": "2372 - CSLL - CONTRIB. SOCIAL SOBRE LUCRO LIQUIDO",
    "2089": "2089 - IRPJ - LUCRO PRESUMIDO",
    "8109": "8109 - PIS - FATURAMENTO",
    "5952": "5952 - RETENÇÃO PIS/COFINS/CSLL",
    "1708": "1708 - IRRF - RENDIMENTO DO TRABALHO",
    "5928": "5928 - IRRF - RENDIMENTOS DE DECORRENTES DO TRABALHO",
    "8301": "8301 - PIS - FATURAMENTO",
    "6912": "6912 - PIS - REAL",
    "0561": "0561 - IRRF - TRABALHO ASSALARIADO",
    "3208": "3208 - IRRF - ALUGUÉIS E ROYALTIES",
}

REGRAS_DOCUMENTOS = {
    "DARF": [
        "documento de arrecadação", 
        "receitas federais", 
        "darf", 
        "arrecadação", 
        "receita federal", 
        "Composição do documento de arrecadação"
    ],
    "RECIBO_SPED": [
        "recibo de entrega",
        "escrituracao fiscal digital",
        "efd-contribuicoes",
        "sped",
        "efd"
    ],
}

MEMORIA_CNPJS = {}


def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            sender_name TEXT,
            folder_name TEXT,
            file_name TEXT,
            status TEXT,
            ocr_text TEXT,
            drive_url TEXT,
            duration_sec REAL,
            bot_instance TEXT DEFAULT 'Outlook Local Bot'
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS processed_messages (
            msg_id TEXT PRIMARY KEY
        )
    ''')
    conn.commit()
    conn.close()

def is_message_already_processed(msg_id):
    if not msg_id:
        return False
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM processed_messages WHERE msg_id = ?", (msg_id,))
        row = cursor.fetchone()
        conn.close()
        return row is not None
    except Exception:
        return False

def save_log(sender_name, folder_name, file_name, status, ocr_text="", drive_url="", duration_sec=0.0, msg_id=None):
    try:
        with sqlite3.connect(DB_PATH, timeout=30.0) as conn:
            cursor = conn.cursor()
            timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute(
                """
                INSERT INTO logs (timestamp, sender_name, folder_name, file_name, status, ocr_text, drive_url, duration_sec, bot_instance)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    timestamp_str,
                    sender_name,
                    folder_name,
                    file_name,
                    status,
                    ocr_text,
                    drive_url,
                    round(duration_sec, 2),
                    "Outlook Local Bot"
                ),
            )
            if msg_id and ("Success" in status or "Ignored" in status):
                cursor.execute(
                    "INSERT OR IGNORE INTO processed_messages (msg_id) VALUES (?)",
                    (msg_id,),
                )
            conn.commit()
    except Exception as e:
        print(f"❌ Erro ao salvar log no SQLite: {e}")



def extrair_dados_pdf_stream(stream_conteudo, nome_arquivo, memoria_cnpjs):
    """Extrai texto e campos estruturados de DARF diretamente da memória."""
    texto_raw = ""
    
    try:
        with pdfplumber.open(stream_conteudo) as pdf:
            for pagina in pdf.pages:
                texto_raw += (pagina.extract_text() or "") + "\n"
    except Exception as e:
        print(f"[pdfplumber error] {e}")

    if not texto_raw.strip():
        try:
            stream_conteudo.seek(0)
            file_bytes = stream_conteudo.read()
            images = convert_from_bytes(file_bytes)
            for img in images:
                texto_raw += pytesseract.image_to_string(img, lang="por") + "\n"
        except Exception as e:
            print(f"[OCR Fallback error] {e}")

    cnpj_limpo = None
    m_cnpj = re.search(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', texto_raw)
    if m_cnpj:
        cnpj_limpo = re.sub(r'\D', '', m_cnpj.group(1))

    empresa = "Não localizado"
    m_emp = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\s+([A-Za-z0-9\s\-\/\.\,\&]{5,})', texto_raw)
    if m_emp:
        cand = m_emp.group(1).strip()
        cand = re.split(r'\n|Período|Data|Número|Pagar|Observações|CNPJ', cand, flags=re.IGNORECASE)[0].strip()
        if len(cand) > 3 and not re.match(r'^\d{2}/', cand) and not re.search(r'Razão|CNPJ', cand, re.IGNORECASE):
            empresa = cand

    if empresa == "Não localizado":
        m_lbl = re.search(r'(?:Razão Social|Nome/Razão Social|Nome)\s*[:\n]?\s*(.+)', texto_raw, re.IGNORECASE)
        if m_lbl:
            cand = m_lbl.group(1).strip()
            cand = re.split(r'\n|Período|Data|Número|Pagar|CNPJ', cand, flags=re.IGNORECASE)[0].strip()
            if not re.match(r'^\d{2}/\d{2}/\d{4}', cand) and not re.match(r'^\d{2}\.', cand):
                empresa = cand

    if empresa != "Não localizado" and cnpj_limpo:
        memoria_cnpjs[cnpj_limpo] = empresa

    if empresa == "Não localizado" and cnpj_limpo and cnpj_limpo in memoria_cnpjs:
        empresa = memoria_cnpjs[cnpj_limpo]

    if empresa == "Não localizado":
        nome_limpo = re.sub(r'^\d+[\-_]?', '', nome_arquivo)
        nome_limpo = re.sub(r'\s*\d*\.pdf$', '', nome_limpo, flags=re.IGNORECASE).strip()
        empresa = nome_limpo

    empresa = re.sub(r'\s+(Período|Data|Número|CNPJ|Pagar).*$', '', empresa, flags=re.IGNORECASE).strip()

    competencia = "Não localizado"
    
    m_extenso = re.search(
        r'\b(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)[/\s](\d{4})\b', 
        texto_raw, 
        re.IGNORECASE
    )
    
    m_comp = re.search(r'PA\s*[:\.]?\s*(\d{2}/\d{2,4})', texto_raw, re.IGNORECASE)

    if m_extenso:
        meses = {
            "janeiro": "01", "fevereiro": "02", "março": "03", "marco": "03",
            "abril": "04", "maio": "05", "junho": "06", "julho": "07",
            "agosto": "08", "setembro": "09", "outubro": "10", "novembro": "11", "dezembro": "12"
        }
        mes_num = meses[m_extenso.group(1).lower()]
        ano = m_extenso.group(2)
        competencia = f"{mes_num}/{ano}"
    elif m_comp:
        competencia = m_comp.group(1).strip()
    else:
        m_pa_completo = re.search(r'PA\s*[:\.]?\s*\d{2}/(\d{2}/\d{4})', texto_raw, re.IGNORECASE)
        if m_pa_completo:
            competencia = m_pa_completo.group(1)

    tributo = "Não localizado"
    m_trib = re.search(r'^\s*(\d{4})\b\s*([A-Za-z0-9\s\-\/\.\(\)]+)', texto_raw, re.MULTILINE)
    if m_trib:
        cod, desc = m_trib.group(1).strip(), m_trib.group(2).strip()
        if cod not in ["0000", "2025", "2026", "2027"]:
            if any(k in desc.upper() for k in ["PIS", "COFINS", "CSLL", "IRPJ", "IRRF", "FATURAMENTO", "CONTRIB", "LUCRO"]):
                desc_limpa = re.sub(r'\s+\d+[\.,].*$', '', desc)
                tributo = f"{cod} - {desc_limpa}"
            elif cod in DICIONARIO_TRIBUTOS:
                tributo = DICIONARIO_TRIBUTOS[cod]

    if tributo == "Não localizado":
        codigos = re.findall(r'\b(\d{4})\b', texto_raw)
        for cod in codigos:
            if cod in DICIONARIO_TRIBUTOS:
                tributo = DICIONARIO_TRIBUTOS[cod]
                break

    valor = "Não localizado"
    m_val = re.search(r'Valor Total(?: do Documento)?\s*[:\n]?\s*R?\$\s*([\d\.,]+)', texto_raw, re.IGNORECASE)
    if not m_val:
        m_val = re.search(r'Valor\s*[:\n]\s*R?\$\s*([\d\.,]+)', texto_raw, re.IGNORECASE)

    if m_val:
        valor = m_val.group(1).strip()
    else:
        valores = re.findall(r'\b\d{1,3}(?:\.\d{3})*,\d{2}\b', texto_raw)
        if valores:
            valor = valores[-1]

    dados_extraidos = {
        "empresa": empresa,
        "competencia": competencia,
        "tributo": tributo,
        "valor": valor
    }

    return texto_raw, dados_extraidos


def classificar_documento(texto_ocr):
    if not texto_ocr:
        return None
    texto_lower = texto_ocr.lower()
    for tipo_doc, palavras_chave in REGRAS_DOCUMENTOS.items():
        coincidencias = sum(1 for kw in palavras_chave if kw in texto_lower)
        if coincidencias >= 2:
            return tipo_doc
    return None


def send_file_to_apps_script(file_name, mime_type, base64_data, folder_name, dados_extraidos=None):
    payload = {
        "fileName": file_name,
        "mimeType": mime_type,
        "base64Data": base64_data,
        "folderName": folder_name,
        "parentFolderId": PARENT_FOLDER_ID,
        "registros": [dados_extraidos] if dados_extraidos else []
    }
    headers = {"Content-Type": "application/json"}
    try:
        response = requests.post(WEB_APP_URL, json=payload, headers=headers, timeout=60)
        return response.json()
    except Exception as e:
        return {"status": "error", "message": str(e)}


def obter_titulo_email(message):
    try:
        subject = getattr(message, "Subject", None)
        if subject and str(subject).strip():
            titulo = str(subject).strip()
            titulo_limpo = re.sub(r'[\\/*?:"<>|]', '', titulo)
            return titulo_limpo if titulo_limpo else "Email_Sem_Assunto"
        return "Email_Sem_Assunto"
    except Exception:
        return "Email_Sem_Assunto"



def processar_emails_outlook_local(log_callback=None):
    pythoncom.CoInitialize()
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        unread_messages = messages.Restrict("[UnRead] = True")

        for message in list(unread_messages):
            if getattr(message, "Class", None) != 43:
                continue

            msg_id = getattr(message, "EntryID", None)
            if is_message_already_processed(msg_id):
                continue

            sender_name = obter_titulo_email(message)
            folder_name = sender_name
            subject = getattr(message, "Subject", "(Sem Assunto)")

            if log_callback:
                log_callback(f"📧 E-mail de: {sender_name} | Assunto: {subject}")

            attachments = getattr(message, "Attachments", None)
            if not attachments or attachments.Count == 0:
                message.UnRead = False
                save_log(sender_name, folder_name, "N/A", "Ignored_No_Attachments", msg_id=msg_id)
                continue

            for i in range(1, attachments.Count + 1):
                start_time = datetime.now()
                attachment = attachments.Item(i)
                file_name = attachment.FileName

                if not file_name.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg', '.tif', '.tiff')):
                    continue

                temp_path = os.path.join(BASE_DIR, f"temp_{file_name}")
                attachment.SaveAsFile(temp_path)

                with open(temp_path, "rb") as f:
                    file_bytes = f.read()

                if os.path.exists(temp_path):
                    os.remove(temp_path)

                file_b64 = base64.b64encode(file_bytes).decode('utf-8')
                mime_type = "application/pdf" if file_name.lower().endswith(".pdf") else "image/png"

                stream_conteudo = io.BytesIO(file_bytes)
                texto_extraido, dados_extraidos = extrair_dados_pdf_stream(
                    stream_conteudo, file_name, MEMORIA_CNPJS
                )

                tipo_documento = classificar_documento(texto_extraido)

                if not tipo_documento:
                    elapsed = (datetime.now() - start_time).total_seconds()
                    save_log(sender_name, folder_name, file_name, "Ignored_Unknown_Type", ocr_text=texto_extraido, duration_sec=elapsed, msg_id=msg_id)
                    continue

                if log_callback:
                    log_callback(f" └─ Lido: {file_name} | Empresa: {dados_extraidos['empresa']} | Comp: {dados_extraidos['competencia']} | Valor: {dados_extraidos['valor']}")

                apps_script_response = send_file_to_apps_script(
                    file_name, mime_type, file_b64, folder_name, dados_extraidos
                )
                elapsed = (datetime.now() - start_time).total_seconds()

                if apps_script_response and apps_script_response.get("status") == "success":
                    file_url = apps_script_response.get("fileUrl", "")
                    save_log(sender_name, folder_name, file_name, f"Success ({tipo_documento})", ocr_text=texto_extraido, drive_url=file_url, duration_sec=elapsed, msg_id=msg_id)
                else:
                    save_log(sender_name, folder_name, file_name, "Failed_Upload", ocr_text=texto_extraido, duration_sec=elapsed, msg_id=msg_id)

            message.UnRead = False

    except Exception as e:
        if log_callback:
            log_callback(f"❌ Erro Outlook: {e}")
            
    finally:
        pythoncom.CoUninitialize()



class OutlookBotApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Painel de Automação - Outlook OCR & DARFs Bot")
        self.geometry("1000x680")
        self.minsize(900, 600)

        init_db()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self.header_frame = ctk.CTkFrame(self, corner_radius=10)
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        self.title_label = ctk.CTkLabel(
            self.header_frame, 
            text=" Bot de Processamento de Documentos & DARFs", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.title_label.pack(side="left", padx=20, pady=15)

        self.status_badge = ctk.CTkLabel(
            self.header_frame,
            text="● INATIVO",
            text_color="#E74C3C",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.status_badge.pack(side="right", padx=20, pady=15)

        self.top_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.top_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.top_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.card_total = self.criar_card(self.top_frame, "Total Processados", "0", 0)
        self.card_sucesso = self.criar_card(self.top_frame, "Sucessos", "0", 1, color="#2ECC71")
        self.card_erros = self.criar_card(self.top_frame, "Erros / Ignorados", "0", 2, color="#E67E22")

        self.control_box = ctk.CTkFrame(self.top_frame)
        self.control_box.grid(row=0, column=3, padx=5, pady=5, sticky="nsew")

        self.btn_toggle = ctk.CTkButton(
            self.control_box, 
            text="Iniciar Bot", 
            fg_color="#27AE60", 
            hover_color="#219150",
            command=self.toggle_bot
        )
        self.btn_toggle.pack(expand=True, fill="both", padx=10, pady=10)

        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="nsew")
        
        self.tab_logs = self.tabview.add("Histórico de Logs")
        self.tab_console = self.tabview.add("Console ao Vivo")

        self.setup_table_ui()

        self.console_text = ctk.CTkTextbox(self.tab_console, font=ctk.CTkFont(family="Consolas", size=12))
        self.console_text.pack(fill="both", expand=True, padx=10, pady=10)

        self.atualizar_interface()

    def criar_card(self, parent, titulo, valor_inicial, col, color=None):
        frame = ctk.CTkFrame(parent)
        frame.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")
        
        lbl_title = ctk.CTkLabel(frame, text=titulo, font=ctk.CTkFont(size=12, weight="normal"))
        lbl_title.pack(anchor="w", padx=15, pady=(10, 0))

        lbl_val = ctk.CTkLabel(
            frame, 
            text=valor_inicial, 
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=color if color else ("black", "white")
        )
        lbl_val.pack(anchor="w", padx=15, pady=(0, 10))
        return lbl_val

    def setup_table_ui(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            background="#2B2B2B",
            foreground="white",
            fieldbackground="#2B2B2B",
            rowheight=28,
            font=("Segoe UI", 10)
        )
        style.configure("Treeview.Heading", background="#1F1F1F", foreground="white", font=("Segoe UI", 10, "bold"))
        style.map("Treeview", background=[("selected", "#1F538D")])

        cols = ("Data/Hora", "Remetente", "Arquivo", "Status", "Tempo (s)")
        self.tree = ttk.Treeview(self.tab_logs, columns=cols, show="headings")

        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=140)

        self.tree.column("Remetente", width=180)
        self.tree.column("Arquivo", width=220)

        scrollbar = ctk.CTkScrollbar(self.tab_logs, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

    def log_console(self, mensagem):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.console_text.insert("end", f"[{timestamp}] {mensagem}\n")
        self.console_text.see("end")

    def toggle_bot(self):
        global BOT_RODANDO
        if not BOT_RODANDO:
            BOT_RODANDO = True
            self.status_badge.configure(text="● EXECUTANDO", text_color="#2ECC71")
            self.btn_toggle.configure(text="Pausar Bot", fg_color="#E74C3C", hover_color="#C0392B")
            self.log_console("🚀 Bot iniciado...")
            threading.Thread(target=self.bot_worker, daemon=True).start()
        else:
            BOT_RODANDO = False
            self.status_badge.configure(text="● PAUSADO", text_color="#F1C40F")
            self.btn_toggle.configure(text="Retomar Bot", fg_color="#27AE60", hover_color="#219150")
            self.log_console("⏸️ Bot pausado pelo usuário.")

    def bot_worker(self):
        while BOT_RODANDO:
            self.log_console("🔍 Verificando caixa de entrada do Outlook...")
            processar_emails_outlook_local(log_callback=self.log_console)
            self.atualizar_interface()
            
            for _ in range(30):
                if not BOT_RODANDO:
                    break
                time.sleep(1)

    def atualizar_interface(self):
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()

            cursor.execute("SELECT COUNT(*) FROM logs")
            total = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM logs WHERE status LIKE 'Success%'")
            sucessos = cursor.fetchone()[0]

            erros = total - sucessos

            self.card_total.configure(text=str(total))
            self.card_sucesso.configure(text=str(sucessos))
            self.card_erros.configure(text=str(erros))

            for item in self.tree.get_children():
                self.tree.delete(item)

            cursor.execute("""
                SELECT timestamp, sender_name, file_name, status, duration_sec 
                FROM logs 
                ORDER BY id DESC LIMIT 50
            """)
            
            for row in cursor.fetchall():
                self.tree.insert("", "end", values=row)

            conn.close()
        except Exception as e:
            print(f"Erro ao atualizar interface: {e}")


if __name__ == "__main__":
    app = OutlookBotApp()
    app.mainloop()