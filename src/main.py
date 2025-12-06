# main.py

import sys
import os
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.append(BASE_DIR)

try:
    from services.stt import ear_service
except ImportError as e:
    print(f"\nERROR CRÍTICO DE IMPORTACIÓN: {e}")
    print("Asegúrate de ejecutar desde la raíz del proyecto: python src/main.py")
    sys.exit(1)

def main():
    print("\n==============================================")
    print("   ASISTENTE VIRTUAL")
    print("==============================================\n")
    print(">>> Iniciando...")

    if ear_service.model is None:
        print("Error: El modelo de voz no se pudo cargar. Revisa la consola.")
        return

    print("Sistema listo. Entrando en bucle principal.")

    try:
        while True:
            despierto, comando_inicial = ear_service.wait_for_wake_word()

            if despierto:
                print("\n[MAIN] ¡DESPERTADO!")
                
                if not comando_inicial:
                    print("[SIMULACIÓN TTS] 'Hola, estoy escuchando.'")
                else:
                    print(f"[MAIN] Comando Rápido Detectado: '{comando_inicial}'")
                    print(f"   -> Enviando al Cerebro: {comando_inicial}")
                    
                while True:
                    texto_usuario = ear_service.listen_active(silence_limit=120)
                    
                    if texto_usuario:
                        print(f"[MAIN] Recibido del STT: '{texto_usuario}'")
                        
                        print(f"   -> Enviando al Cerebro (PLN)...")
                        
                        if "adiós" in texto_usuario or "apagar" in texto_usuario:
                            print("[SIMULACIÓN TTS] 'Hasta luego.'")
                            break
                            
                    else:
                        print("[MAIN] Timeout por inactividad.")
                        print("[SIMULACIÓN TTS] 'Hasta luego.'")
                        break

            else:
                print("[MAIN] Reiniciando ciclo de escucha...")
                time.sleep(1)

    except KeyboardInterrupt:
        print("\n\n[MAIN] Sistema apagado manualmente por el usuario.")
        sys.exit(0)
    except Exception as e:
        print(f"\n[MAIN] Error no controlado: {e}")
        
if __name__ == "__main__":
    main()