import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from N1 import extract_n1_data
from N2 import extract_n2_data
from N25 import extract_n25_data
import google.generativeai as genai
import threading
from docx import Document


genai.configure(api_key="AIzaSyAckosGZaBTj1SxNti0OyllgV80yjXWJXc")

def select_file_n1():
    file_path = filedialog.askopenfilename(title="Selecione o arquivo N1", filetypes=[("Arquivos DOCX", "*.docx")])
    if file_path:
        entry_n1.delete(0, tk.END)
        entry_n1.insert(0, file_path)

def select_file_n2():
    file_path = filedialog.askopenfilename(title="Selecione o arquivo N2", filetypes=[("Arquivos DOCX", "*.docx")])
    if file_path:
        entry_n2.delete(0, tk.END)
        entry_n2.insert(0, file_path)

def select_file_n25():
    file_path = filedialog.askopenfilename(title="Selecione o arquivo N2.5", filetypes=[("Arquivos DOCX", "*.docx")])
    if file_path:
        entry_n25.delete(0, tk.END)
        entry_n25.insert(0, file_path)

def process_files():
    n1_path = r'' + entry_n1.get()  
    n2_path = r'' + entry_n2.get()
    n25_path = r'' + entry_n25.get()
    target_orgao = entry_orgao.get().strip()  

    if not n1_path or not n2_path or not n25_path or not target_orgao:
        messagebox.showerror("Erro", "Todos os arquivos e o órgão precisam ser selecionados.")
        return

    
    n1_data = extract_n1_data(n1_path, target_orgao)
    n2_data = extract_n2_data(n2_path)
    n25_data = extract_n25_data(n25_path)

    combined_data = {
        "n1_data": n1_data,
        "n2_data": n2_data,
        "n25_data": n25_data
    }
    
    print(combined_data)


    start_loading_screen()  
    threading.Thread(target=send_to_google_ai_and_save, args=(combined_data,)).start()  # Processa em uma thread separada

def send_to_google_ai_and_save(data):
    briefing = send_to_google_ai(data)
    stop_loading_screen()
    if briefing:
        save_briefing(briefing)

def send_to_google_ai(data):
    try:
        prompt = (
            "Você é um redator profissional, e eu preciso que você reescreva as informações abaixo em um formato de relatório claro e coeso. "
            "As informações estão organizadas por compromissos, e cada compromisso pode ter ações realizadas, deliberações e pendências a serem resolvidas. "
            "Antes de iniciar coloque APENAS COMPROMISSOS DE NATUREZA DE POLITICAS PUBLICAS, ESQUEÇA GESTÃO E SERVIÇOS E OBRAS POR FAVOR,"
            "se não tiver nos dados compromissos de natureza de politicas publicas escreva: NÃO TEM COMPROMISSOS DE POLITICA PUBLICA \n"
            "NÃO ESCREVA MAIS DE UMA REUNIÃO N1,N2,N2.5 QUERO AS INFORMAÇÕES EM SUAS RESPECTIVAS REUNIÕES\n"
            "Ou seja abaixo da reunião n1 deve ficar os dados da n1, e da n2 n2 e da n2.5 as da n2.5, por favor não deixe desorganizado, organização é a chave aqui \n"
            "Por favor, organize o texto da seguinte forma:\n\n"
            "1. Comece identificando o compromisso, mencionando o código e o que ele envolve.\n"
            "2. Trate o texto para uma lógica e entendimento humano claro.\n"
            "3. Em cada informação que você receber deve ser tratada de maneira que cada compromisso fique na respectiva reunião; se for da N1, deve ficar com as informações da N1.\n"
            "4. Por mais que exista mais de um mesmo compromisso, lembre-se que cada um deve ser separado por reunião.\n\n"
            "5. Por favor deixe todos da n1 no espaço requerido para n1 e todos de n2 para o espaço da n2, e do n2.5 da mesma maneira"
            "6. Para os da n2.5 faça apenas o dos compromissos que apareceram no n1 e no n2"
            "Aqui estão os dados dos compromissos extraídos das atas:\n\n"
            f"N1: {data['n1_data']}\n\n"
            f"N2: {data['n2_data']}\n\n"
            f"N2.5: {data['n25_data']}\n\n"
        )

        model = genai.GenerativeModel(model_name="gemini-1.5-flash")

        response = model.generate_content([prompt])

        return response.text

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        messagebox.showerror("Erro", f"Ocorreu um erro ao processar com a IA: {e}")
        return None

def save_briefing(briefing):
    file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Arquivo Word", "*.docx")])
    if file_path:
       
        doc = Document()

        doc.add_heading('Briefing Processado pela IA', 0)

        # Processa o texto do briefing
        lines = briefing.split("\n")
        last_meeting = None
        for line in lines:
            clean_line = line.strip().replace("**", "").replace("##", "").strip()  # Remove ## e **
            
           
            if "Reunião de Governança" in clean_line:
                current_meeting = clean_line
                if current_meeting != last_meeting:
                    doc.add_heading(clean_line, level=1)  
                    last_meeting = current_meeting  
            elif clean_line.startswith("CG-"): 
                doc.add_heading(clean_line, level=2)  
            elif clean_line:  
                doc.add_paragraph(clean_line)

        
        doc.save(file_path)
        messagebox.showinfo("Sucesso", "Briefing salvo com sucesso!")

def start_loading_screen():
    global loading_screen
    loading_screen = tk.Toplevel(root)
    loading_screen.title("Processando com IA")
    ttk.Label(loading_screen, text="Processando briefing com a IA...").pack(padx=10, pady=10)
    progress = ttk.Progressbar(loading_screen, mode="indeterminate")
    progress.pack(padx=10, pady=10)
    progress.start()

def stop_loading_screen():
    loading_screen.destroy()
    
def close_application():
    root.destroy()

root = tk.Tk()
root.title("Gerador de Briefing")

tk.Label(root, text="Arquivo N1:").grid(row=0, column=0, padx=10, pady=10)
entry_n1 = tk.Entry(root, width=50)
entry_n1.grid(row=0, column=1, padx=10, pady=10)
btn_n1 = tk.Button(root, text="Selecionar", command=select_file_n1)
btn_n1.grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Arquivo N2:").grid(row=1, column=0, padx=10, pady=10)
entry_n2 = tk.Entry(root, width=50)
entry_n2.grid(row=1, column=1, padx=10, pady=10)
btn_n2 = tk.Button(root, text="Selecionar", command=select_file_n2)
btn_n2.grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Arquivo N2.5:").grid(row=2, column=0, padx=10, pady=10)
entry_n25 = tk.Entry(root, width=50)
entry_n25.grid(row=2, column=1, padx=10, pady=10)
btn_n25 = tk.Button(root, text="Selecionar", command=select_file_n25)
btn_n25.grid(row=2, column=2, padx=10, pady=10)

tk.Label(root, text="Órgão para a N1:").grid(row=3, column=0, padx=10, pady=10)
entry_orgao = tk.Entry(root, width=50)
entry_orgao.grid(row=3, column=1, padx=10, pady=10)

btn_process = tk.Button(root, text="Gerar Briefing", command=process_files)
btn_process.grid(row=4, column=1, pady=20)

btn_close = tk.Button(root, text="Fechar Aplicação", command=close_application)
btn_close.grid(row=5, column=1, pady=10)

tk.Label(root, text="Para criar um novo briefing, feche esta janela e abra novamente.", fg="red").grid(row=6, column=1, padx=10, pady=20)


root.mainloop()

