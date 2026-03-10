# -*- coding: utf-8 -*-
"""
Configurador para o Robô de Controle de Combustível.

Autor: Vinicius Andrade Moreira Rissi  
Licença: Proprietária (Uso Comercial Proibido sem Autorização)
"""
import tkinter as tk
from tkinter import filedialog, messagebox
import json
import os
import sys
import datetime
import win32com.client
import win32timezone

# --- Configuração de Caminhos e Constantes ---
# Determina o diretório base (seja rodando como script ou como executável compilado)
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")
HISTORY_FILE = os.path.join(BASE_DIR, "history.json")

# --- Funções de Persistência (Carregar/Salvar) ---
def load_settings():
    """Carrega as configurações do arquivo JSON, retornando vazio se falhar."""
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_settings():
    """Coleta dados da interface, valida e salva no arquivo JSON."""
    data = {
        "email_account": entry_email.get().strip(),
        "outlook_folder": entry_outlook_folder.get().strip(),
        "save_path": entry_save_path.get().strip(),
        "manual_path": entry_manual_path.get().strip(),
    }
    
    if not all([data["email_account"], data["outlook_folder"], data["save_path"], data["manual_path"]]):
        messagebox.showwarning("Atenção", "Por favor, preencha todos os campos.")
        return

    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")
        root.destroy()
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar: {e}")

# --- Funções Auxiliares da Interface ---
def select_folder(entry_widget):
    """Abre o seletor de diretórios do Windows e preenche o campo de texto."""
    folder = filedialog.askdirectory()
    if folder:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, folder)

# --- Funções de Lógica do Outlook ---
def mark_existing_as_processed():
    """
    Conecta ao Outlook, varre a pasta selecionada e marca todos os e-mails
    atuais como 'processados' no histórico, para que sejam ignorados futuramente.
    """
    email_account = entry_email.get().strip()
    folder_path = entry_outlook_folder.get().strip()
    
    if not email_account or not folder_path:
        messagebox.showwarning("Atenção", "Preencha os campos de E-mail e Pasta do Outlook antes de executar.")
        return

    if not messagebox.askyesno("Confirmar", f"Deseja marcar todos os e-mails da pasta '{folder_path}' como processados?\n\nO robô irá IGNORAR estes e-mails e processar apenas os que chegarem a partir de agora."):
        return

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Tenta encontrar a conta correta
        account_folder = None
        for folder in outlook.Folders:
            if folder.Name.lower() == email_account.lower():
                account_folder = folder
                break
        
        if not account_folder:
            account_folder = outlook.Folders.Item(1)
        
        # Navega até a pasta alvo
        target_folder = account_folder
        try:
            for f_name in folder_path.split("\\"):
                target_folder = target_folder.Folders[f_name]
        except Exception:
            messagebox.showerror("Erro", f"Pasta '{folder_path}' não encontrada na conta '{account_folder.Name}'.")
            return

        items = target_folder.Items
        new_ids = set()
        count = 0
        
        for item in items:
            if item.Class == 43: # MailItem
                try:
                    subject = item.Subject if item.Subject else ""
                    email_id = f"{item.ReceivedTime.isoformat()}|{subject}"
                    new_ids.add(email_id)
                    count += 1
                except Exception as e:
                    print(f"Erro ao ler item: {e}") # Ajuda no debug se executado via console
                    continue
        
        # Atualiza histórico
        current_history = set()
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    current_history = set(json.load(f))
            except:
                pass
        
        updated_history = current_history.union(new_ids)
        
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(list(updated_history), f, indent=2, ensure_ascii=False)
            
        messagebox.showinfo("Sucesso", f"{count} e-mails identificados e salvos no histórico.\nO robô irá ignorar estes e-mails nas próximas execuções.")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao acessar o Outlook: {str(e)}")

def reset_history():
    """Apaga o arquivo de histórico, forçando o robô a reprocessar todos os e-mails."""
    if messagebox.askyesno(
        "Confirmar Reset",
        "Você tem certeza que deseja limpar o histórico de e-mails?\n\n"
        "Isso fará com que o robô analise TODOS os e-mails da pasta novamente na próxima execução. "
        "Esta ação não pode ser desfeita."
    ):
        if os.path.exists(HISTORY_FILE):
            try:
                os.remove(HISTORY_FILE)
                messagebox.showinfo("Sucesso", "Histórico de e-mails foi resetado com sucesso.")
            except OSError as e:
                messagebox.showerror("Erro", f"Falha ao deletar o arquivo de histórico: {e}")
        else:
            messagebox.showinfo("Informação", "Nenhum arquivo de histórico encontrado para resetar.")

VERSION = "1.0.3"

# --- Configuração da Interface Gráfica (GUI) ---
root = tk.Tk()
root.title(f"Configuração - Controle de Combustível v{VERSION}")
root.geometry("550x350")

icon_path = os.path.join(BASE_DIR, "doc", "icone.ico")
if os.path.exists(icon_path):
    try:
        root.iconbitmap(icon_path)
    except Exception:
        pass

current_settings = load_settings()

# Criação dos Campos de Entrada
tk.Label(root, text="E-mail da conta no Outlook (ex: seu.nome@empresa.com):").pack(pady=5)
entry_email = tk.Entry(root, width=60)
entry_email.pack()
entry_email.insert(0, current_settings.get("email_account", ""))

tk.Label(root, text=r"Nome da Pasta no Outlook a ser varrida (ex: Inbox ou Inbox\Combustivel):").pack(pady=5)
entry_outlook_folder = tk.Entry(root, width=60)
entry_outlook_folder.pack()
entry_outlook_folder.insert(0, current_settings.get("outlook_folder", "Caixa de Entrada"))

tk.Label(root, text="Pasta onde salvar as pastas dos Veículos:").pack(pady=5)
frame_save = tk.Frame(root)
frame_save.pack()
entry_save_path = tk.Entry(frame_save, width=45)
entry_save_path.pack(side=tk.LEFT)
entry_save_path.insert(0, current_settings.get("save_path", ""))
tk.Button(frame_save, text="...", command=lambda: select_folder(entry_save_path)).pack(side=tk.LEFT)

tk.Label(root, text="Pasta para revisão manual (sem placa identificada):").pack(pady=5)
frame_manual = tk.Frame(root)
frame_manual.pack()
entry_manual_path = tk.Entry(frame_manual, width=45)
entry_manual_path.pack(side=tk.LEFT)
entry_manual_path.insert(0, current_settings.get("manual_path", ""))
tk.Button(frame_manual, text="...", command=lambda: select_folder(entry_manual_path)).pack(side=tk.LEFT)

# Botões de Ação
frame_buttons = tk.Frame(root)
frame_buttons.pack(pady=10)

tk.Button(frame_buttons, text="Esquecer Histórico", command=reset_history, bg="#f44336", fg="white").pack(side=tk.LEFT, padx=5)
tk.Button(frame_buttons, text="Marcar Atuais como Lidos", command=mark_existing_as_processed, bg="#2196F3", fg="white").pack(side=tk.LEFT, padx=5)

tk.Button(root, text="Salvar e Fechar", command=save_settings, bg="#4CAF50", fg="white", font=("Arial", 12, "bold")).pack(pady=20)
root.mainloop()
