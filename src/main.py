# main.py

import sys
import os
import json
import logging
import time
import pyttsx3
import pythoncom

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s', datefmt='%H:%M:%S')
logger = logging.getLogger("Main")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path: sys.path.append(BASE_DIR)

try:
    from services.stt import ear_service
    from core.pln import process_command, initialize_pln_model
    
    import modules.os_control.file_reader as mod_file_reader
    import modules.office_auto.word_session as mod_office
    import modules.os_control.system_ops as mod_system
except ImportError as e:
    logger.critical(f"Error imports: {e}")
    sys.exit(1)

def speak_main(text: str):
    """
    Motor de voz exclusivo del Main.
    Crea una instancia, habla y se destruye para liberar el recurso.
    """
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
    if not initialize_pln_model(): return

    deps = {
        "tts": speak_main, 
        "stt": ear_service
    }

    #speak_main("Sistema listo y modularizado.")

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

                    if mod_office.word_session.is_active:
                        if "guardar" in texto or "cerrar" in texto:
                            mod_office.word_session.process_dictation(texto)
                        else:
                            mod_office.word_session.process_dictation(texto)
                        continue

                    resultados = process_command(texto)
                    should_sleep = False 

                    for cmd in resultados:
                        modulo = cmd.get('modulo')
                        intencion = cmd.get('comando')
                        logger.info(f"Routing a: {modulo}")

                        if modulo == "file_reader":
                            mod_file_reader.execute_module(cmd, deps)
                        
                        elif modulo == "office_auto":
                            if intencion == "crear_word":
                                mod_office.word_session.start_session()
                        
                        elif modulo == "os_control":
                            if intencion == "consultar_hora":
                                speak_main(mod_system.get_time())
                            elif intencion == "ajustar_volumen":
                                speak_main(mod_system.set_volume(cmd['variables'].get('level')))

                        elif modulo == "interaction":
                            if intencion == "despedida":
                                speak_main("Adiós.")
                                should_sleep = True 
                                break 
                            else:
                                speak_main("Hola, estoy aquí.")

                        elif modulo == "unknown":
                            speak_main("No te entendí.")
                    
                    if should_sleep:
                        break

        except KeyboardInterrupt:
            logger.info("Apagado manual.")
            break
        except Exception as e:
            logger.error(f"Error Main: {e}")
            time.sleep(1)

if __name__ == "__main__":
    main()