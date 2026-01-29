"""
MÓDULO TEAMS MANAGER (V7 - Final con Maximize)
------------------------------------------------------
Mejoras:
  - Ventana maximizada para evitar menús ocultos.
  - Estrategia de búsqueda de equipos por encabezados (H2/H3).
  - Tiempos de espera aumentados para conexiones lentas.
  - Modo Exclusivo (Bloquea PLN general).
"""

import os
import time
import keyboard
import re
import pyttsx3
import pythoncom
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

class TeamsSession:
    def __init__(self):
        self.driver = None
        self.is_active = False
        self.download_path = os.path.join(os.path.expanduser("~"), "Downloads")
        self.profile_path = os.path.join(os.path.expanduser("~"), "ChromeProfileLesi")
        self.current_teams = {} 

    def _speak_interruptible(self, text):
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
                if keyboard.is_pressed('ctrl') or keyboard.is_pressed('esc'):
                    was_cancelled = True
                    engine.stop()

            engine.connect('started-word', onWord)
            
            clean = text.replace('\n', ' ').strip()
            if clean:
                time.sleep(0.5)
                engine.say(clean)
                engine.runAndWait()
        except: pass
        finally:
            try: pythoncom.CoUninitialize()
            except: pass
        return was_cancelled

    def _speak_local(self, text):
        self._speak_interruptible(text)

    def _init_driver(self):
        if self.driver: return self.driver
        opts = Options()
        opts.add_argument("--disable-gpu")
        opts.add_argument("--start-maximized") # <--- CORRECCIÓN CLAVE: Maximizar ventana
        opts.add_argument("--log-level=3")
        opts.add_argument(f"user-data-dir={self.profile_path}")
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=opts)
        return self.driver

    def check_login(self):
        self._speak_local("Abriendo Teams, espera un momento...")
        driver = self._init_driver()
        
        # URL directa a la Home
        driver.get("https://teams.microsoft.com/v2/")
        
        # Aumentamos espera a 8s porque Teams es pesado
        time.sleep(8) 
        
        url = driver.current_url.lower()
        if "login.microsoftonline" in url or "signin" in url:
            self._speak_local("Pantalla de login detectada. Solicita ayuda visual.")
            return False
        
        try:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
            self.is_active = True
            return True
        except Exception as e:
            print(f"   [TEAMS ERROR] Timeout: {e}")
            self._speak_local("Teams tarda demasiado en cargar.")
            return False

    def list_teams(self):
        if not self.check_login(): return

        self._speak_local("Analizando estructura de Teams...")
        
        try:
            # Espera de seguridad para que React renderice las tarjetas
            time.sleep(8)

            # --- DIAGNÓSTICO: ¿Qué ve el robot? ---
            # Esto imprimirá en tu consola el título de la página y la URL para confirmar dónde estamos
            print(f"   [DEBUG] Título Página: {self.driver.title}")
            print(f"   [DEBUG] URL Actual: {self.driver.current_url}")
            
            teams_elements = []
            
            # --- ESTRATEGIA MAESTRA: ACCESIBILIDAD (ARIA) ---
            # En el Nuevo Teams, las tarjetas de equipo suelen tener aria-label="Nombre del equipo"
            # O están dentro de un grid con role="gridcell"
            
            print("   [DEBUG] Probando selectores de accesibilidad...")
            
            # Intento 1: Buscar cualquier elemento que tenga un label que diga "equipo" o "team"
            # Esto suele atrapar las tarjetas del grid
            try:
                teams_elements = self.driver.find_elements(By.CSS_SELECTOR, "[data-tid='team-name-text']")
                if not teams_elements:
                     # Intento 2: Buscar tarjetas por clase genérica de tarjeta
                     teams_elements = self.driver.find_elements(By.CSS_SELECTOR, "[class*='fui-Card']")
                
                if not teams_elements:
                    # Intento 3 (El más agresivo): Buscar TODOS los textos grandes dentro de roles de botón o enlace
                    # Esto agarra el texto de las tarjetas en vista lista o cuadrícula
                    teams_elements = self.driver.find_elements(By.XPATH, "//div[@role='gridcell']//span | //div[@role='listitem']//span")
            except: pass

            self.current_teams = {}
            nombres = []
            
            # Si aún así no encuentra nada, vamos a imprimir TODO el texto visible para ver qué pasa
            if not teams_elements:
                print("   [DEBUG] No encontré selectores. Imprimiendo texto visible del Body:")
                try:
                    body_txt = self.driver.find_element(By.TAG_NAME, "body").text
                    print(f"   --- TEXTO VISIBLE: ---\n{body_txt[:300]}...\n----------------------")
                    
                    # Último recurso: Analizar el texto bruto del body línea por línea
                    lines = body_txt.split('\n')
                    for line in lines:
                        clean = line.strip()
                        if len(clean) > 3 and len(clean) < 40: # Un nombre de equipo suele ser corto
                            # Filtros de palabras prohibidas (menús, botones)
                            banned = ["general", "microsoft", "teams", "archivos", "chat", "actividad", "calendario", "unirse", "crear"]
                            if not any(b in clean.lower() for b in banned):
                                # Creamos un "elemento falso" para compatibilidad
                                if clean not in nombres:
                                    self.current_teams[clean.lower()] = clean
                                    nombres.append(clean)
                except: pass
            else:
                # Procesamiento normal si encontró elementos
                for el in teams_elements:
                    try:
                        name = el.text.strip()
                        if name and len(name) > 3:
                             # Filtro estricto para no leer la hora o palabras de sistema
                            banned = ["general", "oculto", "opciones", "más", "equipo", "modificar"]
                            if not any(b in name.lower() for b in banned):
                                if name not in nombres:
                                    self.current_teams[name.lower()] = name
                                    nombres.append(name)
                    except: continue

            # --- RESULTADO FINAL ---
            if not nombres:
                self._speak_local("Sigo sin detectar los nombres. Revisa la consola.")
                return

            # CORRECCIÓN: Limitar la lectura verbal a los primeros 5 para no aburrir
            primeros_equipos = nombres[:10] 
            texto_a_leer = ", ".join(primeros_equipos)
            
            print(f"   [TEAMS] Total encontrados: {len(nombres)}")
            
            # Avisamos cuántos hay y leemos los primeros
            self._speak_local(f"Encontré {len(nombres)} equipos. Los primeros son: {texto_a_leer} y otros más.")
            self._speak_local("Dime 'Entrar a' seguido del nombre del equipo.")

        except Exception as e:
            print(f"   [TEAMS CRASH] {e}")
            self._speak_local("Error crítico leyendo la pantalla.")

    def _normalize(self, text):
        # Función auxiliar para quitar tildes y hacer minusculas
        replacements = (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"))
        text = text.lower()
        for a, b in replacements:
            text = text.replace(a, b)
        return text

    def enter_team_files(self, team_name_query):
        if not self.driver: return
        
        target_name = None
        # Normalizamos entrada
        def normalize(t): return t.lower().replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
        query_norm = normalize(team_name_query)

        for name_key, real_name in self.current_teams.items():
            if query_norm in normalize(name_key):
                target_name = real_name
                break
        
        if not target_name:
            self._speak_local(f"No encontré el equipo {team_name_query}.")
            return

        try:
            self._speak_local(f"Entrando a {target_name[:25]}...")
            
            # 1. CLIC EN EL EQUIPO (Esto ya funcionaba)
            clicked = False
            try:
                xpath = f"//*[@aria-label='{target_name}'] | //*[@title='{target_name}']"
                elements = self.driver.find_elements(By.XPATH, xpath)
                if not elements:
                    xpath = f"//*[contains(text(), '{target_name}')]"
                    elements = self.driver.find_elements(By.XPATH, xpath)
                
                if elements:
                    self.driver.execute_script("arguments[0].click();", elements[0])
                    clicked = True
            except: pass

            if not clicked:
                self._speak_local("No pude hacer clic en la tarjeta del equipo.")
                return

            time.sleep(5) # Esperar a que cargue la vista del equipo

            # 2. ASEGURAR CANAL GENERAL (NUEVO PASO CRÍTICO)
            # A veces al entrar no estamos en "General". Buscamos el canal General explícitamente.
            try:
                print("   [DEBUG] Buscando canal 'General'...")
                # Selectores típicos del canal General en la barra lateral
                xpath_general = "//div[contains(@class, 'fui-TreeItem') and .//span[text()='General']]"
                general_channel = self.driver.find_elements(By.XPATH, xpath_general)
                
                if general_channel:
                    general_channel[0].click()
                    time.sleep(3)
                else:
                    print("   [DEBUG] No vi el canal General, asumo que ya estamos dentro.")
            except: pass

            # 3. BUSCAR PESTAÑA ARCHIVOS (SELECTORES AMPLIADOS)
            self._speak_local("Buscando pestaña Archivos...")
            try:
                # Selectores para New Teams (pueden ser botones, tabs o spans)
                # Buscamos por texto exacto "Archivos" o "Files" en cualquier lugar clicable
                xpath_files = "//*[text()='Archivos' or text()='Files']"
                
                # Esperamos hasta 10s a que aparezca
                files_tab = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, xpath_files))
                )
                # Clic JS es más seguro
                self.driver.execute_script("arguments[0].click();", files_tab)
                
                self._speak_local("Leyendo lista de documentos...")
                time.sleep(6) # Archivos tarda en cargar
                
                # 4. LEER DOCUMENTOS (NUEVO SELECTOR DE GRID)
                # En New Teams los archivos son filas en un grid (role="row")
                # El nombre suele estar en un div con data-tid="name-cell" o similar
                files = self.driver.find_elements(By.CSS_SELECTOR, "[data-tid*='file-name']")
                
                if not files:
                    # Fallback agresivo: Buscar cualquier texto que parezca un archivo (con extensión)
                    files = self.driver.find_elements(By.XPATH, "//div[@role='gridcell']//span[contains(text(), '.')]")

                nombres = []
                for f in files:
                    txt = f.text.strip()
                    if txt and "." in txt and txt not in nombres:
                        nombres.append(txt)
                
                # Limitar a 10 archivos
                nombres = nombres[:10]
                
                if nombres:
                    self._speak_local("Encontré estos archivos (CONTROL para parar):")
                    self._speak_interruptible(", ".join(nombres))
                    self._speak_local("Di 'Descargar [Nombre]' para bajar uno.")
                else:
                    self._speak_local("La carpeta parece vacía o no cargó la lista.")

            except Exception as e:
                print(f"   [TAB ERROR DETALLE] {e}")
                self._speak_local("No encontré la pestaña Archivos. Puede que esté oculta en el menú 'Más'.")

        except Exception as e:
            self._speak_local("Error de navegación desconocido.")

    def download_file(self, file_name_query):
        try:
            # Buscar en celdas o links
            files = self.driver.find_elements(By.CSS_SELECTOR, "[data-tid='file-name-cell']")
            if not files: files = self.driver.find_elements(By.CSS_SELECTOR, "[role='gridcell'] a")
            
            target_file = None
            for f in files:
                if file_name_query.lower() in f.text.lower():
                    target_file = f
                    break
            
            if target_file:
                self._speak_local(f"Descargando {target_file.text}...")
                
                actions = webdriver.ActionChains(self.driver)
                actions.context_click(target_file).perform()
                time.sleep(2) # Espera aumentada para menú contextual
                
                menu_items = self.driver.find_elements(By.CSS_SELECTOR, "[role='menuitem']")
                for item in menu_items:
                    if "descargar" in item.text.lower() or "download" in item.text.lower():
                        item.click()
                        self._speak_local("Iniciado. Revisa Descargas.")
                        return
                self._speak_local("No hallé la opción descargar.")
            else:
                self._speak_local("No veo ese archivo.")
        except: self._speak_local("Error técnico.")

    def upload_personal_file(self, file_name_query):
        try:
            local_file = None
            for f in os.listdir(self.download_path):
                if file_name_query.lower() in f.lower():
                    local_file = os.path.join(self.download_path, f)
                    break
            
            if not local_file:
                self._speak_local("Archivo no encontrado en Descargas.")
                return

            self._speak_local(f"Subiendo {os.path.basename(local_file)}...")
            # Truco: Navegar directo a URL de archivos si falla el botón
            # self.driver.get("https://teams.microsoft.com/v2/files") 
            # (Comentado porque depende del equipo, mejor usar la interfaz actual)
            
            try:
                inp = self.driver.find_element(By.CSS_SELECTOR, "input[type='file']")
                inp.send_keys(local_file)
                self._speak_local("Subiendo...")
            except: self._speak_local("Fallo en subida. No veo el botón Cargar.")
        except: self._speak_local("Error.")

    def close(self):
        if self.driver:
            self.driver.quit()
            self.driver = None
        self.is_active = False
        self._speak_local("Teams cerrado.")

    def process_dictation(self, text):
        text = text.lower().strip()

        if any(x in text for x in ["salir", "cerrar", "terminar", "adiós"]):
            self.close()
            return "Saliendo."

        if "listar" in text or "ver equipos" in text or "mis equipos" in text:
            self.list_teams()
            return None

        match_enter = re.search(r'(entrar|ir|abrir)\s+(al|a|en)?\s*(equipo|clase|grupo)?\s+(?P<name>.+)', text)
        if match_enter:
            name = match_enter.group('name')
            if len(name) > 2:
                self.enter_team_files(name)
                return None

        match_down = re.search(r'(descargar|bajar)\s+(el\s+)?(archivo|documento)?\s*(?P<name>.+)', text)
        if match_down:
            self.download_file(match_down.group('name'))
            return None

        match_up = re.search(r'(subir|cargar)\s+(el\s+)?(archivo|documento)?\s*(?P<name>.+)', text)
        if match_up:
            self.upload_personal_file(match_up.group('name'))
            return None

        print(f"   [TEAMS IGNORADO] {text}")
        return None

teams_manager = TeamsSession()

def execute_module(dto, dependencies):
    cmd = dto.get('comando')
    if cmd == "abrir_teams":
        teams_manager.list_teams()