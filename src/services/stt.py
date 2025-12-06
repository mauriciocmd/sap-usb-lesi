"""
MÓDULO STT: OÍDOS DEL SISTEMA (Vosk - Prioridad Funcionalidad)
--------------------------------------------------------------
Responsabilidad: 
  - Escucha continua con alta sensibilidad (Buffer 1024).
  - Detecta "Lesi" solo o "Lesi + Comando".
  - Mantiene la escucha activa por 2 minutos (120s) tras cada comando.
  - Retorna texto limpio al orquestador.
"""

import os
import sys
import json
import time
import pyaudio
import winsound
from vosk import Model, KaldiRecognizer, SetLogLevel
from typing import Optional, Tuple

SetLogLevel(-1)

def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_path()
if getattr(sys, 'frozen', False):
    DATA_DIR = os.path.join(BASE_DIR, "src", "core", "data")
else:
    DATA_DIR = os.path.join(os.path.dirname(BASE_DIR), "core", "data")

MODEL_PATH = os.path.join(DATA_DIR, "model")
SOUND_ON = os.path.join(DATA_DIR, "sounds", "on.wav")
SOUND_OFF = os.path.join(DATA_DIR, "sounds", "off.wav")

class Ear:
    def __init__(self):
        print(">>> [STT] Cargando modelo auditivo")
        if not os.path.exists(MODEL_PATH):
            print(f"ERROR: No existe modelo en {MODEL_PATH}")
            self.model = None
            return
        try:
            self.model = Model(MODEL_PATH)
            print(">>> [STT] Motor listo. Sensibilidad Alta.")
        except Exception as e:
            print(f"Error al cargar Vosk: {e}")
            self.model = None

    def _play(self, path):
        try:
            if os.path.exists(path):
                winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
        except: pass

    def wait_for_wake_word(self) -> Tuple[bool, str]:
        if not self.model: return False, ""

        p = pyaudio.PyAudio()
        stream = p.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=1024)
        rec = KaldiRecognizer(self.model, 16000)
        
        print("\n💤 MODO REPOSO: Esperando 'Lesi'...")
        
        triggers = [
            "lesi", "le si", "lazy", "lessie", "leci", "lissy", "lessy",
            "le sí", "les y", "decí", "deci", "desir",
            "lecil", "decile", "dile", "desi", "lacy",
            "bessie", "messi", "lezy"
        ]

        try:
            while True:
                data = stream.read(1024, exception_on_overflow=False)
                
                if rec.AcceptWaveform(data):
                    result = json.loads(rec.Result())
                    text = result.get("text", "").lower()

                    if not text: continue

                    for trg in triggers:
                        if trg in text:
                            self._play(SOUND_ON)
                            parts = text.split(trg, 1)
                            remainder = parts[1].strip() if len(parts) > 1 else ""
                            
                            if remainder:
                                print(f"ACTIVADO RÁPIDO. Comando: '{remainder}'")
                            else:
                                print(f"ACTIVADO.")
                                
                            return True, remainder

        except KeyboardInterrupt:
            raise KeyboardInterrupt
        except Exception as e:
            print(f"Error reposo: {e}")
            return False, ""
        finally:
            stream.stop_stream()
            stream.close()
            p.terminate()

    def listen_active(self, silence_limit=120) -> Optional[str]:
        if not self.model: return None

        p = pyaudio.PyAudio()
        stream = p.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=1024)
        rec = KaldiRecognizer(self.model, 16000)
        
        print(f"\nESCUCHANDO")
        
        start_time = time.time()

        try:
            while True:
                elapsed = time.time() - start_time
                if elapsed > silence_limit:
                    print("Tiempo de espera agotado.")
                    self._play(SOUND_OFF)
                    return None
                data = stream.read(1024, exception_on_overflow=False)

                if rec.AcceptWaveform(data):
                    result = json.loads(rec.Result())
                    text = result.get("text", "").strip()

                    if len(text) > 1:
                        if len(text) <= 2 and text not in ["si", "no", "ok", "ir"]:
                            continue
                            
                        print(f"🗣️  '{text}'")
                        return text.lower()

        except KeyboardInterrupt:
            raise KeyboardInterrupt
        except Exception as e:
            print(f"Error activo: {e}")
            return None
        finally:
            stream.stop_stream()
            stream.close()
            p.terminate()

ear_service = Ear()


# PRUEBA DEL FLUJO DE NEGOCIO (SIMULACIÓN MAIN)
if __name__ == "__main__":
    print("--- INICIANDO PRUEBA DE FLUJO REAL ---")
    try:
        while True:
            is_awake, initial_cmd = ear_service.wait_for_wake_word()
            if is_awake:
                if initial_cmd:
                    print(f"[MAIN] Procesando comando inmediato: '{initial_cmd}'")
                else:
                    print("[TTS] 'Hola, estoy aquí.'")
                while True:
                    comando = ear_service.listen_active(silence_limit=120)
                    if comando:
                        print(f"[MAIN] Ejecutando: '{comando}'") 
                        if "adiós" in comando or "apagar" in comando:
                            print("[TTS] 'Hasta luego.'")
                            break
                    else:
                        print("[TTS] 'Hasta luego.' (Por inactividad)")
                        break
    except KeyboardInterrupt:
        print("\nSistema detenido.")