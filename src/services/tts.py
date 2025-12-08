import pyttsx3
from typing import Optional

engine: Optional[pyttsx3.Engine] = None

def _setup_engine():
    global engine
    try:
        engine = pyttsx3.init()
        engine.setProperty('rate', 150)
        engine.setProperty('volume', 1.0)
        
        voices = engine.getProperty('voices')
        for v in voices:
            if "spanish" in v.name.lower() or "español" in v.name.lower():
                engine.setProperty('voice', v.id)
                break
    except Exception as e:
        print(f"Error setup TTS: {e}")

def initialize_tts_engine():
    global engine
    if engine is None:
        _setup_engine()
        print("INFO: Motor TTS inicializado.")

def speak(text: str) -> bool:
    global engine
    
    if not text: return False
    
    if engine is None:
        _setup_engine()

    try:
        if engine._inLoop:
            engine.endLoop()

        engine.say(text)
        engine.runAndWait()
        return True
        
    except Exception as e:
        print(f"ERROR TTS: {e}")
        try:
            _setup_engine()
            engine.say(text)
            engine.runAndWait()
            return True
        except:
            return False

if __name__ == '__main__':
    initialize_tts_engine()
    speak("Prueba de audio robusta.")