# main.py

import sys
import os
import json
import logging
import time

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s', datefmt='%H:%M:%S')
logger = logging.getLogger("Main")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path: sys.path.append(BASE_DIR)

try:
    from services.stt import ear_service
    from services.tts import speak, initialize_tts_engine
    from core.pln import process_command, initialize_pln_model
    
    import modules.os_control.file_reader as mod_file_reader
    import modules.office_auto.word_session as mod_office
    import modules.os_control.system_ops as mod_system
    
except ImportError as e:
    logger.critical(f"Error imports: {e}")
    sys.exit(1)

def main():
    initialize_tts_engine()
    if not initialize_pln_model(): return

    deps = {"tts": speak, "stt": ear_service}
    
    print("\n--- SISTEMA LISTO ---")

    while True:
        try:
            despierto, comando_inicial = ear_service.wait_for_wake_word()

            if despierto:
                logger.info("¡Wake Word detectada!")
                
                if comando_inicial:
                    texto_a_procesar = comando_inicial
                    logger.info(f"Comando rápido: {texto_a_procesar}")
                else:
                    speak("Dime.")
                    texto_a_procesar = None

                while True:
                    if not texto_a_procesar:
                        texto = ear_service.listen(timeout=100)
                        
                        if not texto:
                            logger.info("Tiempo de espera agotado. Volviendo a reposo.")
                            speak("Hasta luego.") 
                            break 
                    else:
                        texto = texto_a_procesar
                        texto_a_procesar = None 

                    print(f"🗣️: {texto}")
                    
                    if mod_office.word_session.is_active:
                        if "guardar" in texto:
                            speak(mod_office.word_session.process_dictation(texto))
                        else:
                            speak(mod_office.word_session.process_dictation(texto))
                        continue

                    resultados = process_command(texto)
                    
                    for cmd in resultados:
                        modulo = cmd.get('modulo')
                        logger.info(f"Routing a: {modulo}")

                        if modulo == "file_reader":
                            mod_file_reader.execute_module(cmd, deps)
                        
                        elif modulo == "office_auto":
                            if cmd['comando'] == "crear_word":
                                name = cmd['variables'].get('new_file_name')
                                speak(mod_office.word_session.start_session(name))
                        
                        elif modulo == "os_control":
                            if cmd['comando'] == "consultar_hora":
                                speak(mod_system.get_time())
                            elif cmd['comando'] == "ajustar_volumen":
                                speak(mod_system.set_volume(cmd['variables'].get('level')))

                        elif modulo == "interaction":
                            if cmd['comando'] == "despedida":
                                speak("Adiós.")
                                texto_a_procesar = None 
                                break 
                            else:
                                speak("Hola.")

                        elif modulo == "unknown":
                            speak("No entendí.")
                    
                    if "despedida" in [c.get('comando') for c in resultados]:
                        break

        except KeyboardInterrupt:
            logger.info("Apagado manual.")
            break
        except Exception as e:
            logger.error(f"Error Main: {e}")
            time.sleep(1)

if __name__ == "__main__":
    main()