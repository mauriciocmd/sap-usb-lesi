#modulo de reconocimiento de voz stt.py

import os
import sys
import json
import pyaudio
import winsound
from vosk import Model, KaldiRecognizer, SetLogLevel
from typing import Optional

SetLogLevel(-1)


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(os.path.dirname(BASE_DIR), "core", "data")

MODEL_PATH = os.path.join(DATA_DIR, "model")
SOUND_ON = os.path.join(DATA_DIR, "sounds", "on.wav")
SOUND_OFF = os.path.join(DATA_DIR, "sounds", "off.wav")

class Ear:
    def __init__(self):
        print(">>> Inicializando sistema auditivo")

        # Carga modelo Vosk
        if not os.path.exists(MODEL_PATH):
            print(f"ERROR: No se encontró modelo Vosk en: {MODEL_PATH}")
            self.model = None
            return

        try:
            self.model = Model(MODEL_PATH)
            print(">>> Modelo Vosk cargado correctamente.")
        except Exception as e:
            print("Error cargando modelo:", e)
            self.model = None

        if not os.path.exists(SOUND_ON) or not os.path.exists(SOUND_OFF):
            print("⚠ ADVERTENCIA: No se encontraron sonidos ON/OFF.")
            print("   El sistema funcionará, pero sin feedback auditivo.")

    def _play(self, path: str):
        if os.path.exists(path):
            winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)

    #   Modo hibernacion “Lesi”
    def wait_for_wake_word(self, keyword: str = "lesi") -> bool:
        if not self.model: return False

        print(f"\nMODO REPOSO: Esperando '{keyword}'")

        p = pyaudio.PyAudio()
        stream = p.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=2048)
        recognizer = KaldiRecognizer(self.model, 16000)
        
        triggers = ["lesi", "le si", "lessi", "leci", "les y", "le sí", "lazy"]

        try:
            while True:
                data = stream.read(2048, exception_on_overflow=False)
                
                if recognizer.AcceptWaveform(data):
                    result = json.loads(recognizer.Result())
                    text = result.get("text", "").lower()

                    if not text: continue

                    for trg in triggers:
                        if trg in text:
                            print(f"Activado, detectado: '{text}'")
                            return True

        except KeyboardInterrupt:
            raise KeyboardInterrupt 
        finally:
            stream.stop_stream()
            stream.close()
            p.terminate()

    def listen(self, timeout: int = 15) -> Optional[str]:
        if not self.model: return None

        self._play(SOUND_ON)
        
        p = pyaudio.PyAudio()
        FRAME_RATE = 16000
        CHUNK = 1024
        max_cycles = int((FRAME_RATE / CHUNK) * timeout)
        cycle = 0

        stream = p.open(format=pyaudio.paInt16, channels=1, rate=FRAME_RATE, input=True, frames_per_buffer=CHUNK)
        recognizer = KaldiRecognizer(self.model, FRAME_RATE)
        
        detected_text = None
        print(f"\nESCUCHANDO COMANDO (máx {timeout}s)...")

        try:
            while cycle < max_cycles:
                cycle += 1
                data = stream.read(CHUNK, exception_on_overflow=False)

                if recognizer.AcceptWaveform(data):
                    result = json.loads(recognizer.Result())
                    text = result.get("text", "").strip()

                    if len(text) > 1:
                        detected_text = text.lower()
                        break
        
        except KeyboardInterrupt:
            print("\nInterrupción de usuario.")
            return None
        except Exception as e:
            print(f"Error micrófono: {e}")
        finally:
            self._play(SOUND_OFF)
            stream.stop_stream()
            stream.close()
            p.terminate()

        if detected_text:
            print(f"COMANDO DETECTADO: '{detected_text}'")
            return detected_text
        
        print("Tiempo agotado o silencio")
        return None

ear_service = Ear()

# PRUEBA
if __name__ == "__main__":
    print("\n--- SISTEMA DE VOZ INICIADO ---")
    
    try:
        while True:
            is_awake = ear_service.wait_for_wake_word()
            
            if is_awake:
                cmd = ear_service.listen(timeout=15)
                
                if cmd:
                    print(f"➡ Enviando al PLN: {cmd}")
                    if "salir" in cmd or "apagar" in cmd:
                        print("Apagando sistema.")
                        break
                else:
                    print("No se escuchó comando. Regresando a dormir.")
                    
    except KeyboardInterrupt:
        print("\n\nSISTEMA DETENIDO POR EL USUARIO")
        sys.exit(0)