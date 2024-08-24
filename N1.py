from docx import Document
import json
import re

def extract_n1_data(doc_path, target_orgao):
    doc = Document(doc_path)
    commitments_data = {}
    current_commitment = None
    current_orgao = None
    extracting_observations = False
    current_observations = []

    for para in doc.paragraphs:
        text = para.text.strip()

     
        if any(keyword in text for keyword in ["Natureza", "Status do Compromisso atual", "Status do Compromisso pactuado", "REUNIÃO DE GOVERNANÇA N1"]):
            continue

        normalized_text = text.replace(" ", "").upper()  
        if "órgão:" in text.lower():
            current_orgao = text.split(":")[1].strip()
            extracting_observations = False  

        if current_orgao == target_orgao:
           
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
                if re.match(r'^\d+\.\s', text) or '\t' in text or re.search(r'\s{2,}', text):
                    text = re.sub(r'^\d+\.\s+', '', text).strip()  # Remove o número e espaços extras
                if text:  # Adicionar apenas se o texto não estiver vazio
                    current_observations.append(text)

   
    if current_commitment and current_observations:
        commitments_data[current_commitment]["observations"] = current_observations

    return commitments_data

doc_path = r'c:\Users\Mickael\Downloads\_ATA N1 - DESENVOLVIMENTO ECONÔMICO I - 15-08-2024docx.docx'  # Substitua pelo caminho real do documento
target_orgao = 'INVESTE'  

n2tes = extract_n1_data(doc_path, target_orgao)

for commitment, data in n2tes.items():
    print(f"Compromisso: {commitment}")
    print("Observações:")
    for observation in data['observations']:
        print(f" - {observation}")
