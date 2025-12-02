# word_session.py

import os
import re
import threading
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

class WordSession:
    def __init__(self):
        self.doc = None
        self.is_active = False
        self.download_path = os.path.join(os.path.expanduser("~"), "Downloads")

    def start_session(self, initial_filename: str = None) -> str:
        self.doc = Document()
        self.is_active = True
        
        msg = "Editor de Word iniciado. Puedes dictar títulos, párrafos u oraciones."
        if initial_filename:
            msg += f" (Archivo pre-nombrado: {initial_filename})"
        return msg

    def _clean_filename(self, text: str) -> str:
        text = re.sub(r'\s+en\s+(pdf|word|docx)$', '', text, flags=re.IGNORECASE)
        clean = re.sub(r'[<>:"/\\|?*]', '', text)
        return clean.strip()

    def process_dictation(self, text: str) -> str:

        if not self.is_active or not self.doc:
            return "No hay sesión activa. Primero di 'Crear Word'."

        text_lower = text.lower()

        if "guardar documento" in text_lower:
            match = re.search(r'guardar documento con el nombre\s+(.+)', text_lower)
            if match:
                raw_name = match.group(1)
                filename = self._clean_filename(raw_name)
                # Ejecutamos el guardado
                return self._save_file(filename)
            else:
                return "Entendí guardar, pero no escuché el nombre. Di: 'Guardar documento con el nombre X'"

        if text_lower.startswith("título") or text_lower.startswith("titulo"):
            content = text[6:].strip()
            self._add_formatted_paragraph(content, size=16, bold=True, align="CENTER")
            return f"Título agregado."

        elif text_lower.startswith("subtítulo") or text_lower.startswith("subtitulo"):
            content = text[9:].strip()
            self._add_formatted_paragraph(content, size=14, bold=True, align="LEFT")
            return f"Subtítulo agregado."

        elif text_lower.startswith("párrafo") or text_lower.startswith("parrafo"):
            content = text[7:].strip()
            self._add_formatted_paragraph(content, size=12, bold=False, align="JUSTIFY")
            return "Párrafo agregado."

        elif text_lower.startswith("oración") or text_lower.startswith("oracion"):
            content = text[7:].strip()
            self._append_to_last_paragraph(content)
            return "Oración añadida."

        else:
            self._add_formatted_paragraph(text, size=12, bold=False, align="JUSTIFY")
            return "Texto agregado."

    def _add_formatted_paragraph(self, text, size, bold, align):
        if not text: return
        
        text = self._normalize_punctuation(text)
        
        p = self.doc.add_paragraph()
        run = p.add_run(text)
        
        run.font.size = Pt(size)
        run.bold = bold
        
        if align == "CENTER": p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "LEFT": p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif align == "JUSTIFY": p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def _append_to_last_paragraph(self, text):
        if not text: return
        text = self._normalize_punctuation(text)
        
        if len(self.doc.paragraphs) > 0:
            p = self.doc.paragraphs[-1]
            p.add_run(" " + text)
        else:
            self._add_formatted_paragraph(text, 12, False, "LEFT")

    def _normalize_punctuation(self, text):
        replacements = {
            " coma": ",", " punto": ".", " dos puntos": ":", 
            " abre interrogación": "¿", " cierra interrogación": "?"
        }
        for k, v in replacements.items():
            text = text.replace(k, v)
        return text[0].upper() + text[1:] if text else text

    def _save_file_thread(self, full_path, doc_ref):
        try:
            doc_ref.save(full_path)
            print(f">>> [THREAD] Archivo guardado: {full_path}")
        except Exception as e:
            print(f">>> [THREAD] Error al guardar: {e}")

    def _save_file(self, filename):
        full_path = os.path.join(self.download_path, f"{filename}.docx")
        
        doc_to_save = self.doc
        
        self.is_active = False 
        self.doc = None
        
        save_thread = threading.Thread(target=self._save_file_thread, args=(full_path, doc_to_save))
        save_thread.start()
        
        return f"Guardando documento {filename}.docx en Descargas..."

# Instancia Global
word_session = WordSession()

# ZONA DE PRUEBAS AUTOMÁTICAS
if __name__ == "__main__":
    print("\n--- PRUEBA INTERNA DEL MÓDULO DE OFIMÁTICA ---")
    
    print(f"1. Inicio: {word_session.start_session()}")
    
    comandos = [
        "Titulo la tumba de las luciernagas",
        "subtitulo capitulo 1",
        "Parrafo no estaba listo para lo que venía pues siempre quise estar cerca coma pero no pude punto",
        "Oracion No todo era lo que parece",
        "Esto es una frase suelta que se escribirá como párrafo normal",
        "guardar documento con el nombre Pesar del niño en pdf"
    ]
    
    for cmd in comandos:
        print(f"\n🎤 Input: '{cmd}'")
        respuesta = word_session.process_dictation(cmd)
        print(f"Output: {respuesta}")
    
    import time
    time.sleep(1) 
    
    ruta_final = os.path.join(os.path.expanduser("~"), "Downloads", "Pesar del niño.docx")
    if os.path.exists(ruta_final):
        print(f"\nÉXITO: El archivo fue creado en: {ruta_final}")
    else:
        print(f"\nERROR: No se encontró el archivo (Puede que el hilo siga guardando).")