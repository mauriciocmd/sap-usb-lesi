# file_reader.py

import os
import difflib
import re
import pyttsx3
import pythoncom
import msvcrt
import time
from docx import Document
from PyPDF2 import PdfReader

def _speak_interruptible(text):
    was_cancelled = False
    try:
        pythoncom.CoInitialize()
        engine = pyttsx3.init("sapi5")
        engine.setProperty('rate', 150)
        
        voices = engine.getProperty('voices')
        for v in voices:
            if "spanish" in v.name.lower() or "español" in v.name.lower():
                engine.setProperty('voice', v.id)
                break

        def onWord(name, location, length):
            nonlocal was_cancelled
            if msvcrt.kbhit():
                msvcrt.getch()
                was_cancelled = True
                engine.stop()

        engine.connect('started-word', onWord)

        clean_text = text.replace('\n', ' ').replace('\r', '')
        
        print("   [READER] Leyendo... (Presiona tecla para cancelar)")
        engine.say(clean_text)
        engine.runAndWait()
        
    except Exception as e:
        print(f"   [TTS ERROR] {e}")
    finally:
        try: pythoncom.CoUninitialize()
        except: pass
        
    return was_cancelled

def _extract_text(path):
    print(f"   [READER] Extrayendo: {os.path.basename(path)}")
    ext = os.path.splitext(path)[1].lower()
    text = ""
    try:
        if ext == '.pdf':
            reader = PdfReader(path)
            for page in reader.pages:
                t = page.extract_text()
                if t: text += t + " "
        elif ext == '.docx':
            doc = Document(path)
            text = "\n".join([p.text for p in doc.paragraphs])
        elif ext == '.txt':
            with open(path, 'r', encoding='utf-8') as f:
                text = f.read()
    except: return None
    return text

def _clean_query_name(raw_name):
    if not raw_name: return ""
    stopwords = ["documento", "archivo", "el", "la", "un", "leer", "abrir", "pdf", "word", "docx", "txt"]
    clean = raw_name.lower()
    for word in stopwords:
        clean = re.sub(r'\b' + word + r'\b', '', clean)
    return clean.strip()

def _get_target_path(folder_name):
    user_home = os.path.expanduser("~")
    if not folder_name: return os.path.join(user_home, "Downloads")
    folder = folder_name.lower().strip()
    mapping = { "descargas": "Downloads", "escritorio": "Desktop", "documentos": "Documents" }
    target = mapping.get(folder, "Downloads")
    return os.path.join(user_home, target)

def find_file_strict(filename_query, search_path):
    print(f"   [DEBUG] Buscando '{filename_query}' en '{search_path}'")
    if not filename_query or not os.path.exists(search_path): return None
    
    try:
        valid_exts = ('.pdf', '.docx', '.txt')
        files = [f for f in os.listdir(search_path) if f.lower().endswith(valid_exts)]
    except: return None

    query = filename_query.lower().strip()
    for f in files:
        if f.lower().startswith(query): return os.path.join(search_path, f)
    for f in files:
        if query in f.lower() and not f.startswith("~$"): return os.path.join(search_path, f)
    matches = difflib.get_close_matches(query, files, n=1, cutoff=0.4)
    if matches: return os.path.join(search_path, matches[0])
    return None

def execute_module(dto, dependencies):
    cmd = dto.get('comando')
    vars = dto.get('variables', {})
    system_speak = dependencies['tts']

    if cmd == "leer_documento":
        raw_fname = vars.get('file_name')
        fdir = vars.get('directory')
        fname = _clean_query_name(raw_fname)
        
        if len(fname) < 2:
            system_speak("Nombre de archivo no válido.")
            return

        target_path = _get_target_path(fdir)
        system_speak(f"Buscando {fname}...") 
        
        real_file = find_file_strict(fname, target_path)
        
        if not real_file:
            print("   [DEBUG] Buscando en Descargas...")
            target_path = _get_target_path("descargas")
            real_file = find_file_strict(fname, target_path)
        
        if real_file:
            nombre = os.path.basename(real_file)
            
            contenido = _extract_text(real_file)
            if not contenido or len(contenido.strip()) < 5:
                system_speak("El archivo está vacío.")
                return

            system_speak(f"Leyendo {nombre}. Presione cualquier tecla para cancelar la lectura.")
            
            time.sleep(3) 

            while msvcrt.kbhit(): msvcrt.getch()
            
            fue_cancelado = _speak_interruptible(contenido)

            if fue_cancelado:
                system_speak("Lectura cancelada.")
            else:
                system_speak("Fin del documento.")
        else:
            system_speak(f"No encontré el archivo {fname}.")
    else:
        system_speak("Comando no reconocido.")