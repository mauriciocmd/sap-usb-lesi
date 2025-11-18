import os
from PyPDF2 import PdfReader
from docx import Document

def _extract_text_from_pdf(file_path: str) -> str:
    try:
        reader = PdfReader(file_path)
        
        full_text = []
        for page in reader.pages:
            full_text.append(page.extract_text() or "")
            
        return " ".join(full_text)
    
    except Exception as e:
        print(f"ERROR: Fallo al leer el PDF {file_path}. Detalle: {e}")
        return "Error: No fue posible leer el archivo PDF solicitado."


def _extract_text_from_docx(file_path: str) -> str:
    try:
        document = Document(file_path)
        
        full_text = []
        for paragraph in document.paragraphs:
            full_text.append(paragraph.text)
            
        return "\n".join(full_text)
    
    except Exception as e:
        print(f"ERROR: Fallo al leer el DOCX {file_path}. Detalle: {e}")
        return "Error: No fue posible leer el archivo Word solicitado."
        

def read_document(file_path: str) -> str:
    if not os.path.exists(file_path):
        return f"Error: Archivo no encontrado en la ruta {file_path}."
        
    _, file_extension = os.path.splitext(file_path)
    file_extension = file_extension.lower()

    if file_extension == '.pdf':
        return _extract_text_from_pdf(file_path)
        
    elif file_extension == '.docx':
        return _extract_text_from_docx(file_path)
        
    elif file_extension == '.txt':
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as e:
            print(f"ERROR: Fallo al leer TXT. Detalle: {e}")
            return "Error: Fallo al leer el archivo de texto."
            
    else:
        return f"Error: Tipo de archivo '{file_extension}' no soportado por el asistente. Solo soporta PDF, DOCX y TXT."


# --- EJEMPLO DE USO ---
if __name__ == '__main__':
    
    # 1. Prueba de un archivo PDF
    pdf_path = r"C:\Users\mauricio.callisaya\Downloads\Archivo de prueba.pdf" 
    print("--- LEYENDO PDF ---")
    pdf_content = read_document(pdf_path)
    print(pdf_content)

    # 2. Prueba de un archivo DOCX
    docx_path = r"C:\Users\mauricio.callisaya\Downloads\Tsa Guia.docx"
    print("\n--- LEYENDO DOCX ---")
    docx_content = read_document(docx_path)
    print(docx_content)    

    # 3. Prueba de archivo inexistente
    inexistente_path = r"C:\archivo_que_no_existe.pdf"
    print("\n--- LEYENDO INEXISTENTE ---")
    print(read_document(inexistente_path))