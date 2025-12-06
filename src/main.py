# main.py

import sys
import os
import time
import json
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - [%(levelname)s] - %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger("AsistenteMain")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.append(BASE_DIR)

try:
    from services.stt import ear_service
    from services.tts import speak, initialize_tts_engine
    from core.pln import process_command, initialize_pln_model
except ImportError as e:
    logger.critical(f"Error importando módulos: {e}")
    sys.exit(1)

def inicializar_sistema():
    """Carga recursos pesados al inicio."""
    logger.info("Iniciando servicios del sistema...")
    
    try:
        initialize_tts_engine()
        logger.info("Motor TTS: OK")
    except Exception as e:
        logger.error(f"Fallo en TTS: {e}")

    try:
        initialize_pln_model()
        logger.info("Motor PLN: OK")
    except Exception as e:
        logger.critical(f"Fallo crítico en PLN: {e}")
        return False

    if ear_service.model is None:
        logger.critical("Motor STT: Fallo (Modelo no encontrado)")
        return False
    logger.info("Motor STT: OK")

    return True

def procesar_entrada_usuario(texto_usuario: str):
    """Envía el texto al cerebro y gestiona la respuesta."""
    if not texto_usuario:
        return

    logger.info(f"Procesando entrada: '{texto_usuario}'")

    try:
        resultados = process_command(texto_usuario)
        
        if not resultados:
            logger.warning("PLN no devolvió resultados.")
            return

        json_output = json.dumps(resultados, indent=4, ensure_ascii=False)
        print("\n--- RESULTADO DEL PLN ---")
        print(json_output)
        print("-------------------------\n")

        comando_principal = resultados[0]
        intent = comando_principal.get('comando')
        
        if intent == "error":
            logger.warning(f"Error interno en PLN: {comando_principal.get('variables')}")
        
        elif intent == "desconocido" or intent is None:
            logger.info("Comando no reconocido en el entrenamiento.")
            
        else:
            logger.info(f"Intención detectada: {intent}")

    except Exception as e:
        logger.error(f"Error procesando comando: {e}")

def main():
    if not inicializar_sistema():
        logger.critical("No se pudo iniciar el sistema. Cerrando.")
        return

    speak("Sistema en línea y listo.")
    logger.info("--- BUCLE PRINCIPAL INICIADO ---")

    while True:
        try:
            despierto, comando_inicial = ear_service.wait_for_wake_word()

            if despierto:
                logger.info("¡Wake Word detectada!")
                
                if comando_inicial:
                    logger.info(f"Comando rápido: {comando_inicial}")
                    procesar_entrada_usuario(comando_inicial)
                
                else:
                    speak("Dime.")
                    
                    while True:
                        texto = ear_service.listen_active(silence_limit=120)
                        
                        if texto:
                            procesar_entrada_usuario(texto)
                            
                            if "adiós" in texto or "suspender asistente" in texto:
                                speak("Hasta luego.")
                                break
                        else:
                            logger.info("Tiempo de espera agotado. Volviendo a reposo.")
                            break

        except KeyboardInterrupt:
            logger.info("Apagado manual por usuario.")
            break
        except Exception as e:
            logger.error(f"Error en bucle principal: {e}")
            time.sleep(1)

if __name__ == "__main__":
    main()