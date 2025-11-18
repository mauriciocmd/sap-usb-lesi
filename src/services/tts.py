import pyttsx3
from typing import Optional

# Variable global para almacenar el motor de voz (solo se inicializa una vez)
engine: Optional[pyttsx3.Engine] = None

def initialize_tts_engine(rate: int = 150, volume: float = 1.0) -> None:
    global engine
    
    if engine is not None:
        return

    try:
        engine = pyttsx3.init()
        
        engine.setProperty('rate', rate)
        engine.setProperty('volume', volume)

        voices = engine.getProperty('voices')
        spanish_voice_id = None
        
        for voice in voices:
            if "spanish" in voice.name.lower() or "español" in voice.name.lower() or "helena" in voice.name.lower():
                spanish_voice_id = voice.id
                break
        
        if spanish_voice_id:
            engine.setProperty('voice', spanish_voice_id)
            print(f"INFO: Motor TTS inicializado")
        else:
            print("ADVERTENCIA: No se encontró voz en español. Usando la voz predeterminada.")

    except Exception as e:
        engine = None
        print(f"ERROR FATAL: Fallo al inicializar el motor TTS. {e}")


def speak(text: str) -> bool:
    global engine
    
    if engine is None:
        print("ERROR: El motor TTS no fue inicializado correctamente. No se puede hablar.")
        return False
        
    try:
        engine.stop()
        
        engine.say(text)
        
        engine.runAndWait()
        
        return True
        
    except Exception as e:
        print(f"ERROR: Fallo durante la lectura por voz. Detalle: {e}")
        return False


if __name__ == '__main__':
    # Bloque de prueba
    print("--- PRUEBA DE MÓDULO TTS ---")
    initialize_tts_engine() 
    
    engine.setProperty('rate', 100)
    # Prueba: Lectura simple
    speak("Hola. Soy el Asistente Virtual SAP de la Universidad Salesiana.")