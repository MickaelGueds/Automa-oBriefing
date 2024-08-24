from docx import Document
import json
import re

def extract_n2_data(doc_path):
    doc = Document(doc_path)
    commitments_data = {}
    current_commitment = None
    extracting_observations = False
    current_observations = []

    for para in doc.paragraphs:
        text = para.text.strip()

        
        if any(keyword in text for keyword in ["Natureza", "Status do Compromisso atual", "Status do Compromisso pactuado", "REUNIÃO DE GOVERNANÇA N2"]):
            continue

        normalized_text = text.replace(" ", "").upper()  

        
        if "(CG-" in normalized_text or "(CG -" in normalized_text:
            if current_commitment and current_observations:
                
                commitments_data[current_commitment]["observations"] = current_observations

            current_commitment = text.strip()
            commitments_data[current_commitment] = {"observations": []}
            current_observations = []  
            extracting_observations = False  

        
        elif "OBSERVAÇÕES:" in text.upper():
            extracting_observations = True
            
            observations_part = re.split(r'\t+|\s{2,}', text.split("OBSERVAÇÕES:")[1].strip())
            current_observations.extend(observations_part)

        
        elif extracting_observations and current_commitment:
            # Detecta listas numeradas do tipo "1. ", "2. ", etc., mesmo que estejam mal formatadas
            if re.match(r'^\d+\.\s', text) or '\t' in text or re.search(r'\s{2,}', text):
                text = re.sub(r'^\d+\.\s+', '', text).strip()  # Remove o número e espaços extras
            if text:  # Adicionar apenas se o texto não estiver vazio
                current_observations.append(text)

    # Captura o último compromisso se as observações estiverem ativas
    if current_commitment and current_observations:
        commitments_data[current_commitment]["observations"] = current_observations

    return commitments_data

def save_data_to_json(data, file_name):
    with open(file_name, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# Exemplo de uso
doc_path = r'c:\Users\Mickael\Downloads\INVESTE - ATA N2 - 18-06-2024.docx'  
n2_data = extract_n2_data(doc_path)


save_data_to_json(n2_data, 'n2_data.json')


for commitment, data in n2_data.items():
    print(f"Compromisso: {commitment}")
    print("Observações:")
    for observation in data['observations']:
        print(f" - {observation}")
