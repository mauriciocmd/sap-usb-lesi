# main.py

import sys
import os
import json
import logging
import time
import pyttsx3
import pythoncom
from datetime import datetime
import modules.web_navigator.teams_manager as mod_teams

if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if BASE_DIR not in sys.path: sys.path.append(BASE_DIR)

user_docs = os.path.join(os.path.expanduser("~"), "Documents")
log_folder = os.path.join(user_docs, "Registros_Lesi")

if not os.path.exists(log_folder):
    try:
        os.makedirs(log_folder)
    except: pass

log_filename = f"Lesi_Log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
log_path = os.path.join(log_folder, log_filename)

handlers_list = [logging.StreamHandler(sys.stdout)]
if os.path.exists(log_folder):
    handlers_list.append(logging.FileHandler(log_path, encoding='utf-8'))

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    datefmt='%H:%M:%S',
    handlers=handlers_list
)
logger = logging.getLogger("Main")

try:
    from services.stt import ear_service
    from core.pln import process_command, initialize_pln_model
    
    import modules.os_control.file_reader as mod_file_reader
    import modules.office_auto.word_session as mod_office
    import modules.os_control.system_ops as mod_system
    import modules.web_navigator.web_search as mod_web
except ImportError as e:
    logger.critical(f"Error CRÍTICO de imports: {e}")
    time.sleep(5) 
    sys.exit(1)

def speak_main(text: str):
    if not text: return
    try:
        pythoncom.CoInitialize()
        engine = pyttsx3.init("sapi5")
        engine.setProperty('rate', 150)
        engine.setProperty('volume', 1.0)
        
        voices = engine.getProperty('voices')
        for v in voices:
            if "spanish" in v.name.lower() or "español" in v.name.lower():
                engine.setProperty('voice', v.id)
                break
        
        print(f"   [SISTEMA] {text}")
        engine.say(text)
        engine.runAndWait()
        
    except Exception as e:
        logger.error(f"Error TTS Main: {e}")
    finally:
        try: pythoncom.CoUninitialize()
        except: pass

def main():
    logger.info("Iniciando sistema...")
    if not initialize_pln_model(): 
        logger.error("Fallo al inicializar PLN")
        return

    deps = { "tts": speak_main, "stt": ear_service }

    while True:
        try:
            despierto, comando_inicial = ear_service.wait_for_wake_word()

            if despierto:
                logger.info("¡Wake Word detectada!")
                texto_a_procesar = comando_inicial if comando_inicial else None
                
                if not texto_a_procesar:
                    speak_main("Dime.")

                while True:
                    if not texto_a_procesar:
                        texto = ear_service.listen(timeout=100)
                        if not texto:
                            logger.info("Timeout. Volviendo a reposo.")
                            speak_main("Hasta luego.")
                            break 
                    else:
                        texto = texto_a_procesar
                        texto_a_procesar = None

                    print(f"🗣️: {texto}")
                    logger.info(f"Usuario dijo: {texto}")

                    if mod_office.word_session.is_active:
                        if "guardar" in texto or "cerrar" in texto:
                            mod_office.word_session.process_dictation(texto)
                        else:
                            mod_office.word_session.process_dictation(texto)
                        continue
                    
                    # --- B. MODO TEAMS (Prioridad 2 - NUEVO BLOQUEO) ---
                    if mod_teams.teams_manager.is_active:
                        # Enviamos todo el texto directamente al gestor de Teams
                        # Él tiene su propio mini-cerebro (Regex) para entender
                        mod_teams.teams_manager.process_dictation(texto)
                        continue # Salta el PLN (¡Esto evita que se abra Word!)

                    resultados = process_command(texto)
                    should_sleep = False 

                    for cmd in resultados:
                        modulo = cmd.get('modulo')
                        intencion = cmd.get('comando')
                        logger.info(f"Routing a: {modulo} -> {intencion}")

                        if modulo == "file_reader":
                            mod_file_reader.execute_module(cmd, deps)
                        
                        elif modulo == "office_auto":
                            if intencion == "crear_word":
                                mod_office.word_session.start_session()
                        
                        elif modulo == "interaction":
                            if intencion == "despedida":
                                speak_main("Adiós.")
                                should_sleep = True 
                                break 
                            else:
                                speak_main("Hola, estoy aquí.")
                        
                        elif modulo == "web_search":
                            mod_web.execute_module(cmd, deps)
                            
                        elif modulo == "teams_manager":
                            mod_teams.execute_module(cmd, deps)
                        
                        elif modulo == "os_control":
                            mod_system.execute_module(cmd, deps)

                        elif modulo == "unknown":
                            speak_main("No te entendí.")
                    
                    if should_sleep:
                        break

        except KeyboardInterrupt:
            logger.info("Apagado manual.")
            break
        except Exception as e:
            logger.error(f"Error Main Loop: {e}")
            time.sleep(1)

if __name__ == "__main__":
    main()