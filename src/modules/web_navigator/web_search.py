# web_search.py

import os
import time
import requests
import re
import pyttsx3
import pythoncom
import keyboard
import urllib.parse
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

class WebSession:
    def __init__(self):
        self.current_results = [] 
        self.is_active = False
        self.download_path = os.path.join(os.path.expanduser("~"), "Downloads")

    def _speak_local_interruptible(self, text):
        was_cancelled = False
        try:
            pythoncom.CoInitialize()
            engine = pyttsx3.init("sapi5")
            engine.setProperty('rate', 155)
            
            voices = engine.getProperty('voices')
            for v in voices:
                if "spanish" in v.name.lower():
                    engine.setProperty('voice', v.id)
                    break

            def onWord(name, location, length):
                nonlocal was_cancelled
                if keyboard.is_pressed('ctrl'):
                    was_cancelled = True
                    engine.stop()

            engine.connect('started-word', onWord)

            clean = text.replace('\n', ' ').strip()
            if clean:
                print("   [WEB READER] Leyendo... (CTRL para parar)")
                
                time.sleep(1)
                
                engine.say(clean)
                engine.runAndWait()
        except Exception as e:
            print(f"   [TTS ERROR] {e}")
        finally:
            try: pythoncom.CoUninitialize()
            except: pass
        
        return was_cancelled

    def _get_driver(self):
        opts = Options()
        opts.add_argument("--headless") 
        opts.add_argument("--disable-gpu")
        opts.add_argument("--log-level=3")
        opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=opts)

    def search_duckduckgo(self, query):
        self.current_results = []
        driver = None
        
        search_query = query
        if "pdf" in query.lower() and "filetype" not in query.lower():
            search_query += " filetype:pdf"

        try:
            print(f"   [WEB] Buscando: {search_query}")
            driver = self._get_driver()
            driver.get(f"https://duckduckgo.com/html/?q={search_query}")
            time.sleep(3)
            
            elements = driver.find_elements(By.CSS_SELECTOR, ".result")
            
            count = 0
            res_text = "Encontré esto: "
            
            for el in elements:
                try:
                    link = el.find_element(By.CSS_SELECTOR, "a.result__a")
                    title = link.text
                    href = link.get_attribute("href")
                    
                    if title and href:
                        count += 1
                        self.current_results.append({"id": count, "title": title, "url": href})
                        clean_title = title[:80] 
                        res_text += f"Opción {count}: {clean_title}. "
                        
                        if count >= 3: break
                except: continue

            if count == 0:
                print("   [DEBUG] 0 resultados. HTML estructura cambiada.")
                return "No encontré resultados. Intenta otra búsqueda."

            return res_text + " Di el número."

        except Exception as e:
            print(f"Error Web: {e}")
            return "Error de conexión."
        finally:
            if driver: driver.quit()

    def _resolve_url(self, url):
        if "duckduckgo.com/l/" in url:
            try:
                parsed = urllib.parse.urlparse(url)
                params = urllib.parse.parse_qs(parsed.query)
                if 'uddg' in params: return params['uddg'][0]
            except: pass
        return url

    def _download_file(self, url, title):
        try:
            final_url = self._resolve_url(url)
            clean_title = re.sub(r'[<>:"/\\|?*]', '', title)[:30].strip()
            ext = final_url.split('.')[-1]
            if len(ext) > 4 or "?" in ext: ext = "pdf"
            
            filename = f"{clean_title}.{ext}"
            path = os.path.join(self.download_path, filename)
            
            print(f"   [WEB] Descargando {filename}...")
            headers = {'User-Agent': 'Mozilla/5.0'}
            response = requests.get(final_url, headers=headers, stream=True)
            
            with open(path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            return f"Archivo guardado en Descargas: {filename}"
        except: return "Error en la descarga."

    def _read_web_page(self, url, system_speak):
        try:
            final_url = self._resolve_url(url)
            headers = {'User-Agent': 'Mozilla/5.0'}
            
            print(f"   [WEB] Entrando a: {final_url[:40]}...")
            
            try:
                head = requests.head(final_url, headers=headers, timeout=5, allow_redirects=True)
                ctype = head.headers.get('content-type', '').lower()
                if 'pdf' in ctype or 'application' in ctype:
                    return self._download_file(final_url, "Documento Web")
            except: pass

            page = requests.get(final_url, headers=headers, timeout=10)
            page.encoding = page.apparent_encoding 
            soup = BeautifulSoup(page.content, 'html.parser')
            
            for tag in soup(["script", "style", "nav", "footer", "header", "aside", "form"]):
                tag.decompose()

            chunks = []
            if soup.title: chunks.append(f"Título: {soup.title.string}")
            
            content_area = soup.find('main') or soup.body

            if content_area:
                for element in content_area.find_all(['h1', 'h2', 'p']):
                    text = element.get_text().strip()
                    if len(text) > 40:
                        chunks.append(text)

            if not chunks:
                return "La página no tiene texto legible."

            system_speak("Iniciando lectura. Presione CONTROL o ESCAPE para detener.")
            time.sleep(2)
            
            full_text = " ".join(chunks)
            cancelled = self._speak_local_interruptible(full_text)

            return "Lectura cancelada." if cancelled else "Fin de la página."

        except Exception as e:
            return "No pude leer la página."

    def process_selection(self, selection_text, system_speak):
        if not self.current_results:
            return "Primero debes buscar algo."

        idx = -1
        match = re.search(r'\b(\d)\b', selection_text)
        if match: idx = int(match.group(1))
        else:
            nums = {"uno": 1, "dos": 2, "tres": 3}
            for w, n in nums.items():
                if w in selection_text: 
                    idx = n
                    break
        
        if idx > 0 and idx <= len(self.current_results):
            res = self.current_results[idx-1]
            url = res['url']
            title = res['title']
            
            is_file = False
            if url.lower().endswith(('.pdf', '.docx', '.doc')) or "pdf" in title.lower():
                is_file = True
                
            if is_file:
                system_speak(f"Es un archivo. Descargando {title[:20]}...")
                return self._download_file(url, title)
            else:
                system_speak(f"Entrando a {title[:20]}...")
                return self._read_web_page(url, system_speak)
        
        return "Opción no válida."

web_session = WebSession()

def execute_module(dto, dependencies):
    cmd = dto.get('comando')
    vars = dto.get('variables', {})
    main_tts = dependencies['tts']

    if cmd == "investigar_web":
        query = vars.get('query')
        if not query:
            main_tts("Dime qué buscar.")
            return
        
        main_tts(f"Investigando {query}...")
        res_text = web_session.search_duckduckgo(query)
        
        main_tts(res_text)
    
    elif cmd == "seleccionar_web":
        selection = vars.get('selection')
        if not selection:
            main_tts("Dime el número.")
            return
        msg = web_session.process_selection(selection, main_tts)
        if msg: main_tts(msg)

    elif cmd == "detener_lectura":
        main_tts("Presiona CONTROL para detener.")