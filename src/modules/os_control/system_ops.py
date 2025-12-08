# system_ops.py

import datetime
import time
import re
import pyttsx3
import pythoncom
import keyboard 
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    HAS_AUDIO_LIB = True
except ImportError:
    HAS_AUDIO_LIB = False

class SystemManager:
    def _speak_local(self, text):
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
            print(f"   [SYSTEM] {text}")
            engine.say(text)
            engine.runAndWait()
        except: pass
        finally:
            try: pythoncom.CoUninitialize()
            except: pass

    def say_time(self):
        now = datetime.datetime.now()
        hora = now.strftime("%I").lstrip('0')
        minutos = now.strftime("%M")
        
        periodo = "de la mañana"
        if now.hour >= 12: periodo = "de la tarde"
        if now.hour >= 20: periodo = "de la noche"
        
        txt_min = "en punto" if minutos == "00" else f"y {minutos}"
        self._speak_local(f"Son las {hora} {txt_min} {periodo}.")

    def say_date(self):
        now = datetime.datetime.now()
        dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        
        texto = f"Hoy es {dias[now.weekday()]}, {now.day} de {meses[now.month-1]} de {now.year}."
        self._speak_local(texto)

    def _set_volume_keyboard(self, target_level):
        """
        Baja el volumen a 0 y lo sube pulsando la tecla virtualmente.
        Cada pulsación en Windows suele ser 2%.
        """
        print("   [SYSTEM] Usando método de teclado (Fallback)...")
        
        for _ in range(50):
            keyboard.send('volume down')
            time.sleep(0.01)
            
        steps = int(target_level / 2)
        for _ in range(steps):
            keyboard.send('volume up')
            time.sleep(0.02)
            
        return f"Volumen ajustado al {target_level} por ciento."

    def set_volume(self, level_str):
        if not level_str: 
            self._speak_local("Dime el número.")
            return

        level = None
        mapping = {
            "uno":1, "dos":2, "tres":3, "cuatro":4, "cinco":5, "diez":10, 
            "quince":15, "veinte":20, "treinta":30, "cuarenta":40, "cincuenta":50,
            "sesenta":60, "setenta":70, "ochenta":80, "noventa":90, "cien":100
        }
        
        for word, val in mapping.items():
            if word in str(level_str).lower():
                level = val
                break
        
        if level is None:
            nums = re.findall(r'\d+', str(level_str))
            if nums: level = int(nums[0])
        
        if level is None:
            self._speak_local("No entendí el nivel.")
            return

        level = max(0, min(100, level))

        try:
            if HAS_AUDIO_LIB:
                pythoncom.CoInitialize()
                
                devices = AudioUtilities.GetSpeakers()
                interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                volume = cast(interface, POINTER(IAudioEndpointVolume))
                
                scalar = level / 100.0
                volume.SetMasterVolumeLevelScalar(scalar, None)
                self._speak_local(f"Volumen al {level} por ciento.")
                return
        except Exception as e:
            print(f"   [WARN] Falló Pycaw ({e}). Usando teclado...")
        
        msg = self._set_volume_keyboard(level)
        self._speak_local(msg)

sys_manager = SystemManager()

def execute_module(dto, dependencies):
    cmd = dto.get('comando')
    vars = dto.get('variables', {})

    if cmd == "consultar_hora":
        sys_manager.say_time()
    elif cmd == "consultar_fecha":
        sys_manager.say_date()
    elif cmd == "ajustar_volumen":
        level = vars.get('level')
        sys_manager.set_volume(level)