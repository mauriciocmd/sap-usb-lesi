# Módulo de Procesamiendo de Lenguaje Natural pln.py

import os
import sys
import json
import re
import joblib
import spacy
from typing import List, Dict, Any, Optional
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import Pipeline

current_dir = os.path.dirname(os.path.abspath(__file__))
root_dir = os.path.dirname(os.path.dirname(current_dir))
if root_dir not in sys.path: sys.path.append(root_dir)

try:
    from src.core.dtos import CommandDTO
except ImportError:
    sys.path.append(os.path.join(current_dir, '..', '..'))
    from src.core.dtos import CommandDTO

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
MODEL_PATH = os.path.join(DATA_DIR, 'intent_classifier.pkl')
TRAINING_DATA_PATH = os.path.join(DATA_DIR, 'training_data.json')

nlp: Optional[spacy.Language] = None
intent_classifier: Optional[Pipeline] = None
intent_config_cache: Dict[str, Dict] = {}


def _load_data_config():
    global intent_config_cache
    texts, labels = [], []
    intent_config_cache = {}

    if not os.path.exists(TRAINING_DATA_PATH): return [], []

    with open(TRAINING_DATA_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)

    for obj in data['intents']:
        name = obj['name']
        patterns = obj.get('patterns', [])
        patterns.sort(key=len, reverse=True)
        
        intent_config_cache[name] = {
            "patterns": patterns,
            "rules": obj.get('extraction_rules', [])
        }
        for ex in obj.get('examples', []):
            texts.append(ex)
            labels.append(name)
    return texts, labels

#Motor de extracción
def _generate_regex_from_pattern(pattern: str) -> str:

    regex = re.escape(pattern)

    regex = regex.replace(r'\ ', r'\s+')
    
    #variable al final
    if re.search(r'\\\{\w+\\\}$', regex):
        regex = re.sub(r'\\\{(\w+)\\\}$', r'(?P<\1>.+)', regex)
    
    #variables intermedias
    regex = re.sub(r'\\\{(\w+)\\\}', r'(?P<\1>.+?)', regex)
    
    return regex

def _extract_variables(text: str, intent_name: str) -> Dict[str, Any]:
    config = intent_config_cache.get(intent_name, {})
    patterns = config.get("patterns", [])
    rules = config.get("rules", [])
    extracted_data = {}

    for pattern in patterns:
        regex_str = _generate_regex_from_pattern(pattern)
        match = re.search(regex_str, text, re.IGNORECASE)
        if match:
            raw_data = match.groupdict()
            for k, v in raw_data.items():
                extracted_data[k] = v.strip()
            
            for rule in rules:
                key = rule['entity_key']
                if key not in extracted_data: extracted_data[key] = None
            return extracted_data

    for rule in rules:
        key = rule['entity_key']
        dtype = rule['type']
        
        if dtype == "number":
            nums = re.findall(r'\b(\d{1,3})\b', text)
            extracted_data[key] = nums[-1] if nums else None
            
        elif dtype == "text":
            if "file" in key or "document" in key:
                 match_ext = re.search(r'\b([\w\-\(\)\[\] ]+\.(pdf|docx|txt))\b', text, re.IGNORECASE)
                 extracted_data[key] = match_ext.group(1).strip() if match_ext else None
            else:
                extracted_data[key] = None

    return extracted_data

#Entrenamiento
def train_model():
    print(">>> [PLN] Entrenando...")
    texts, labels = _load_data_config()
    if not texts: return
    pipeline = Pipeline([
        ('tfidf', TfidfVectorizer(ngram_range=(1, 2))),
        ('clf', LogisticRegression(solver='liblinear', multi_class='auto'))
    ])
    pipeline.fit(texts, labels)
    if not os.path.exists(DATA_DIR): os.makedirs(DATA_DIR)
    joblib.dump(pipeline, MODEL_PATH)
    print(">>> [PLN] Guardado.")

def initialize_pln_model():
    global nlp, intent_classifier
    if nlp and intent_classifier: return
    try: nlp = spacy.load("es_core_web_sm")
    except: nlp = spacy.blank("es")
    
    if os.path.exists(MODEL_PATH):
        intent_classifier = joblib.load(MODEL_PATH)
        _load_data_config()
    else:
        train_model()
        intent_classifier = joblib.load(MODEL_PATH)


def process_command(raw_text: str) -> List[Dict]:
    global intent_classifier
    if not intent_classifier: initialize_pln_model()
    response_list = []
    
    segments = re.split(r'\b(y|e|además|luego|después)\b', raw_text, flags=re.IGNORECASE)
    clean_segments = [s.strip() for s in segments if len(s.strip()) > 2 and s.lower() not in ['y','e','además','luego','después']]
    if not clean_segments: clean_segments = [raw_text]

    for segment in clean_segments:
        try:
            pred_intent = intent_classifier.predict([segment])[0]
            variables = _extract_variables(segment, pred_intent)
            
            dto = CommandDTO(pred_intent, variables)
            response_list.append(dto.to_dict())
            
        except Exception as e:
            response_list.append(CommandDTO("error", {"detalle": str(e)}).to_dict())

    return response_list

# Pruebas
if __name__ == "__main__":
    print("--- INICIANDO SISTEMA ---")
    train_model()
    initialize_pln_model()
    
    input_usuario = "necesito un nuevo documento word, Crimenes de amor que es un resumen para mi amiga y empieza a leer el archivo Tarea 27 que lo tengo en descargas"
    print(f"\nUsuario: '{input_usuario}'\n")
    
    print(json.dumps(process_command(input_usuario), indent=4, ensure_ascii=False))