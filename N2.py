from docx import Document
import json
import re

def extract_n2_data(doc_path):
    doc = Document(doc_path)
    commitments_data = {}
    current_commitment = None
    current_natureza = None
    extracting_observations = False
    current_observations = []

    paragraphs = doc.paragraphs  # Lista de todos os parágrafos

    for i, para in enumerate(paragraphs):
        text = para.text.strip()

        # Ignorar Statuses e palavras-chave não relevantes
        if any(keyword in text for keyword in ["Status do Compromisso atual"]):
            continue

        normalized_text = text.replace(" ", "").upper()  

        # Identificar o compromisso
        if "(CG-" in normalized_text or "(CG -" in normalized_text:
            if current_commitment and (current_natureza or current_observations):
                # Salva o compromisso anterior antes de começar um novo
                commitments_data[current_commitment] = {
                    "natureza": current_natureza,
                    "observations": current_observations
                }

            current_commitment = text.strip()
            commitments_data[current_commitment] = {"natureza": None, "observations": []}
            current_observations = []  # Resetar observações
            current_natureza = None  # Resetar natureza
            extracting_observations = False  # Resetar a flag de extração de observações

        # Identificar e armazenar a natureza do compromisso
        elif "NATUREZA:" in normalized_text:
            current_natureza = text.split(":", 1)[1].strip()
            # Verificar se a natureza está na linha seguinte
            if not current_natureza and i < len(paragraphs) - 1:
                next_text = paragraphs[i + 1].text.strip()
                if "Status" not in next_text:
                    current_natureza = next_text  # Captura a linha seguinte como natureza

            # Atualiza a natureza do compromisso atual
            if current_commitment:
                commitments_data[current_commitment]["natureza"] = current_natureza

        # Iniciar extração das observações quando encontrar a palavra-chave
        elif "OBSERVAÇÕES:" in text.upper():
            extracting_observations = True
            observations_part = re.split(r'\t+|\s{2,}', text.split("OBSERVAÇÕES:")[1].strip())
            current_observations.extend(observations_part)

        # Continuar coletando observações
        elif extracting_observations and current_commitment:
            # Detecta listas numeradas do tipo "1. ", "2. ", etc., mesmo que estejam mal formatadas
            if re.match(r'^\d+\.\s', text) or '\t' in text or re.search(r'\s{2,}', text):
                text = re.sub(r'^\d+\.\s+', '', text).strip()  # Remove o número e espaços extras
            if text:  # Adicionar apenas se o texto não estiver vazio
                current_observations.append(text)

    # Captura o último compromisso se as observações estiverem ativas
    if current_commitment and (current_natureza or current_observations):
        commitments_data[current_commitment] = {
            "natureza": current_natureza,
            "observations": current_observations
        }

    return commitments_data

def save_data_to_json(data, file_name):
    with open(file_name, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# Exemplo de uso
# doc_path = r'c:\Users\Mickael\Downloads\INVESTE - ATA N2 - 18-06-2024.docx'  
# n2_data = extract_n2_data(doc_path)

# save_data_to_json(n2_data, 'n2_data.json')

# for commitment, data in n2_data.items():
#     print(f"Compromisso: {commitment}")
#     print(f"Natureza: {data['natureza']}")
#     print("Observações:")
#     for observation in data['observations']:
#         print(f" - {observation}")
