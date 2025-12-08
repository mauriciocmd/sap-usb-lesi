# stt.py

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
        print(">>> [STT] Cargando modelo auditivo...")
        if not os.path.exists(MODEL_PATH):
            print(f"ERROR: No existe modelo en {MODEL_PATH}")
            self.model = None
            return
        try:
            self.model = Model(MODEL_PATH)
            print(">>> [STT] Motor listo.")
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
        
        print("\nMODO REPOSO: Esperando 'Lesi'...")
        
        triggers = ["lesi", "le si", "lazy", "lessie", "leci", "lissy", "lessy", 
                    "le sí", "les y", "decí", "deci", "desir"]

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

    def listen(self, timeout=120) -> Optional[str]:
        if not self.model: return None

        p = pyaudio.PyAudio()
        stream = p.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=1024)
        rec = KaldiRecognizer(self.model, 16000)
        
        print(f"\nESCUCHANDO... (Cierre tras {timeout}s de silencio)")
        
        start_time = time.time()

        try:
            while True:
                elapsed = time.time() - start_time
                if elapsed > timeout:
                    print("Tiempo agotado.")
                    self._play(SOUND_OFF)
                    return None 

                data = stream.read(1024, exception_on_overflow=False)

                if rec.AcceptWaveform(data):
                    result = json.loads(rec.Result())
                    text = result.get("text", "").strip()

                    if len(text) > 1:
                        if len(text) <= 2 and text not in ["si", "no", "ok"]:
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

if __name__ == "__main__":
    try:
        while True:
            despierto, cmd = ear_service.wait_for_wake_word()
            if despierto:
                print("Despierto!")
                ear_service.listen(timeout=5)
    except KeyboardInterrupt:
        pass