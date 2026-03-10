# -*- coding: utf-8 -*-
"""
Robô Automatizador para Controle de Combustível.

Autor: Vinicius Andrade Moreira Rissi
Licença: Proprietária (Uso Comercial Proibido sem Autorização)
"""
import win32com.client
import win32timezone
import os
import sys
import re
import json
import datetime
from plyer import notification
from pathlib import Path

# --- Configurações Globais e Caminhos ---
# Determina o diretório base corretamente (seja script ou executável)
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")
LOG_FILE = os.path.join(BASE_DIR, "relatorio_execucoes.txt")
HTML_FILE = os.path.join(BASE_DIR, "ultimo_relatorio.html")
HISTORY_FILE = os.path.join(BASE_DIR, "history.json")
VERSION = "1.0.3"
APP_TITLE = f"Controles de Combustivel v{VERSION}"

# --- Funções de Persistência e Log ---
def load_settings():
    """Carrega as configurações do arquivo JSON."""
    if not os.path.exists(SETTINGS_FILE):
        return None
    with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_settings(settings):
    """Salva as configurações no arquivo JSON."""
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=4, ensure_ascii=False)

def log_message(message):
    """Registra mensagens com data/hora no arquivo de log."""
    timestamp = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")

def load_history():
    """Lê o histórico de e-mails já processados para evitar duplicidade."""
    if not os.path.exists(HISTORY_FILE):
        return set()
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return set(json.load(f))
    except (json.JSONDecodeError, IOError):
        log_message("Aviso: Arquivo de histórico (history.json) não encontrado ou corrompido. Criando um novo.")
        return set()

def save_history(existing_history, newly_processed_ids):
    """Atualiza e salva o arquivo de histórico com os novos e-mails processados."""
    if not newly_processed_ids:
        return  # Nada de novo para salvar

    updated_history = existing_history.union(newly_processed_ids)
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(list(updated_history), f, indent=2, ensure_ascii=False)
    except Exception as e:
        log_message(f"Erro ao salvar o histórico de e-mails: {str(e)}")

# --- Funções Auxiliares ---
def send_notification(title, message):
    """Exibe uma notificação nativa do Windows na área de trabalho."""
    icon_path = os.path.join(BASE_DIR, "doc", "icone.ico")
    # Plyer requer que o arquivo exista para usar o ícone
    if not os.path.exists(icon_path):
        icon_path = None

    try:
        notification.notify(
            title=title,
            message=message,
            app_name=APP_TITLE,
            app_icon=icon_path,
            timeout=10
        )
    except Exception as e:
        log_message(f"Erro na notificação: {str(e)}")

def get_folder(base_folder, folder_path_str):
    r"""Navega recursivamente pelas pastas do Outlook (ex: Inbox\Combustivel)."""
    folders = folder_path_str.split("\\")
    current_folder = base_folder
    try:
        for f_name in folders:
            current_folder = current_folder.Folders[f_name]
        return current_folder
    except Exception:
        return None

# --- Lógica de Extração e Relatórios ---
def extract_plate(text):
    """
    Busca padrões de placa: ABC-1234, ABC - 1234, ABC 1234, ABC1234, ABC1D23, ABC-1D23.
    Retorna a placa normalizada (sem hífen/espaço, maiúscula) ou None se não encontrar.
    """
    # Regex explica:
    # (?<![a-zA-Z0-9]) : Garante que NÃO existe letra ou número ANTES (evita "OBRA" -> "BRA")
    # (?![a-zA-Z0-9])  : Garante que NÃO existe letra ou número DEPOIS (evita "1940503" -> "1940")
    pattern = r"(?<![a-zA-Z0-9])([a-zA-Z]{3}\s*-?\s*[0-9][a-zA-Z0-9][0-9]{2})(?![a-zA-Z0-9])"
    matches = re.findall(pattern, text)
    
    months = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
    
    for raw_match in matches:
        clean_plate = raw_match.upper().replace("-", "").replace(" ", "")
        
        # Verificação de falso positivo: Mês + Ano (ex: FEV 2026)
        if clean_plate[:3] in months:
            suffix = clean_plate[3:]
            # Verifica se é puramente numérico (exclui Mercosul tipo 1D23) e parece ano
            if suffix.isdigit() and suffix.startswith("20"):
                try:
                    year_val = int(suffix[2:])
                    if 25 < year_val < 99:
                        continue # É uma data, ignora e tenta o próximo match
                except ValueError:
                    pass
        
        return clean_plate

    return None

def generate_html_report(events, summary_stats, saved_details, manual_details):
    """Gera um relatório visual em HTML com estatísticas e tabelas detalhadas."""
    
    html_content = f"""
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Relatório de Execução - Combustível</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; margin: 0; padding: 20px; color: #333; }}
            .container {{ max-width: 900px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }}
            h1 {{ color: #2c3e50; border-bottom: 2px solid #eee; padding-bottom: 10px; }}
            .summary {{ display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }}
            .card {{ flex: 1; padding: 20px; border-radius: 8px; color: white; text-align: center; min-width: 150px; }}
            .bg-blue {{ background-color: #3498db; }}
            .bg-green {{ background-color: #27ae60; }}
            .bg-red {{ background-color: #e74c3c; }}
            .card h2 {{ margin: 0; font-size: 36px; }}
            .clickable {{ cursor: pointer; transition: transform 0.2s; }}
            .clickable:hover {{ transform: scale(1.05); box-shadow: 0 8px 20px rgba(0,0,0,0.2); }}
            .card p {{ margin: 5px 0 0; font-size: 14px; text-transform: uppercase; letter-spacing: 1px; opacity: 0.9; }}
            
            .timeline {{ list-style: none; padding: 0; }}
            .timeline-item {{ padding: 15px; border-left: 4px solid #ddd; margin-bottom: 10px; background: #fff; border: 1px solid #eee; border-left-width: 5px; border-radius: 4px; }}
            .type-success {{ border-left-color: #27ae60; }}
            .type-warning {{ border-left-color: #f39c12; background-color: #fffdf5; }}
            .type-info {{ border-left-color: #3498db; }}
            .type-error {{ border-left-color: #c0392b; background-color: #fff5f5; }}
            
            .time {{ font-size: 12px; color: #999; margin-bottom: 5px; display: block; }}
            .message {{ font-size: 15px; }}
            .highlight {{ font-weight: bold; }}
            
            footer {{ margin-top: 40px; text-align: center; font-size: 12px; color: #aaa; }}

            /* Modal Styles */
            .modal {{ display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; overflow: auto; background-color: rgba(0,0,0,0.5); }}
            .modal-content {{ background-color: #fefefe; margin: 5% auto; padding: 20px; border: 1px solid #888; width: 80%; max-width: 800px; border-radius: 8px; box-shadow: 0 5px 15px rgba(0,0,0,0.3); animation: fadeIn 0.3s; }}
            .close {{ color: #aaa; float: right; font-size: 28px; font-weight: bold; cursor: pointer; }}
            .close:hover, .close:focus {{ color: #000; text-decoration: none; cursor: pointer; }}
            
            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 14px; }}
            th, td {{ padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }}
            th {{ background-color: #f8f9fa; color: #333; }}
            tr:hover {{ background-color: #f1f1f1; }}
            
            @keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(-20px); }} to {{ opacity: 1; transform: translateY(0); }} }}
        </style>
        <script>
            function openModal(id) {{
                document.getElementById(id).style.display = "block";
            }}
            function closeModal(id) {{
                document.getElementById(id).style.display = "none";
            }}
            // Fecha o modal se clicar fora dele
            window.onclick = function(event) {{
                if (event.target.classList.contains('modal')) {{
                    event.target.style.display = "none";
                }}
            }}
        </script>
    </head>
    <body>
        <div class="container">
            <h1>Relatório de Execução</h1>
            <p>Data da Execução: <strong>{datetime.datetime.now().strftime("%d/%m/%Y às %H:%M:%S")}</strong></p>
            
            <div class="summary">
                <div class="card bg-blue">
                    <h2>{summary_stats['new_emails']}</h2>
                    <p>E-mails Verificados</p>
                </div>
                <div class="card bg-green clickable" onclick="openModal('modal-saved')">
                    <h2>{summary_stats['saved']}</h2>
                    <p>Arquivos Salvos (Ver Lista)</p>
                </div>
                <div class="card bg-red clickable" onclick="openModal('modal-manual')">
                    <h2>{summary_stats['manual']}</h2>
                    <p>Revisão Manual (Ver Lista)</p>
                </div>
            </div>

            <h3>Detalhamento</h3>
            <ul class="timeline">
    """
    
    for event in events:
        # Define classe CSS baseada no tipo
        css_class = "type-info"
        icon = "ℹ️"
        if event['type'] == 'success':
            css_class = "type-success"
            icon = "✅"
        elif event['type'] == 'warning':
            css_class = "type-warning"
            icon = "⚠️"
        elif event['type'] == 'error':
            css_class = "type-error"
            icon = "❌"
            
        html_content += f"""
                <li class="timeline-item {css_class}">
                    <span class="time">{event['time']}</span>
                    <span class="message">{icon} {event['message']}</span>
                </li>
        """

    html_content += f"""
            </ul>
            <footer>
                Gerado automaticamente por Controles de Combustível v{VERSION}
            </footer>
        </div>

        <!-- Modal Saved -->
        <div id="modal-saved" class="modal">
            <div class="modal-content">
                <span class="close" onclick="closeModal('modal-saved')">&times;</span>
                <h2 style="color: #27ae60;">Arquivos Salvos com Sucesso</h2>
                <table>
                    <thead><tr><th>Placa</th><th>Arquivo</th><th>Pasta Destino</th></tr></thead>
                    <tbody>
    """
    for item in saved_details:
        html_content += f"<tr><td><b>{item['plate']}</b></td><td>{item['filename']}</td><td>{item['folder']}</td></tr>"

    html_content += """
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Modal Manual -->
        <div id="modal-manual" class="modal">
            <div class="modal-content">
                <span class="close" onclick="closeModal('modal-manual')">&times;</span>
                <h2 style="color: #e74c3c;">Arquivos para Revisão Manual</h2>
                <p>Estes arquivos não tiveram a placa identificada automaticamente.</p>
                <table>
                    <thead><tr><th>Arquivo Original</th></tr></thead>
                    <tbody>
    """
    for item in manual_details:
        html_content += f"<tr><td>{item['filename']}</td></tr>"

    html_content += """
                    </tbody>
                </table>
            </div>
        </div>
    </body>
    </html>
    """
    
    try:
        with open(HTML_FILE, "w", encoding="utf-8") as f:
            f.write(html_content)
    except Exception as e:
        log_message(f"Erro ao gerar HTML: {e}")

# --- Bloco Principal de Execução ---
def main():
    # 1. Carregamento de Configurações
    settings = load_settings()
    if not settings:
        send_notification(APP_TITLE, "Erro: Arquivo settings.json não encontrado. Execute o configurador.")
        return

    email_target = settings.get("email_account")
    folder_name = settings.get("outlook_folder")
    # Remove aspas extras que podem vir do configurador para evitar erros de caminho
    save_path_root = settings.get("save_path", "").strip().strip('"')
    manual_path = settings.get("manual_path", "").strip().strip('"')
    # Lista para armazenar eventos para o HTML
    html_events = []
    
    # Listas detalhadas para os modais
    saved_details_list = []
    manual_details_list = []

    # 2. Inicialização
    send_notification(APP_TITLE, f"Iniciando varredura na pasta '{folder_name}' do seu email, baixando PDFs e separando por placa.")
    
    log_message("=== INÍCIO DA EXECUÇÃO ===")
    html_events.append({"type": "info", "message": "Início da varredura.", "time": datetime.datetime.now().strftime("%H:%M:%S")})
    
    try:
        # 3. Preparação do Histórico e Conexão Outlook
        # Carrega o histórico de e-mails que já foram processados
        processed_history = load_history()
        processed_in_this_run = set()
        log_message(f"{len(processed_history)} e-mails já constam no histórico.")

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Tenta encontrar a conta correta
        account_folder = None
        for folder in outlook.Folders:
            if folder.Name.lower() == email_target.lower():
                account_folder = folder
                break
        
        if not account_folder:
            # Se não achar pelo nome exato, tenta usar o padrão (primeira conta)
            account_folder = outlook.Folders.Item(1)
            log_message(f"Aviso: Conta '{email_target}' não encontrada explicitamente. Usando '{account_folder.Name}'.")

        target_folder = get_folder(account_folder, folder_name)
        
        if not target_folder:
            err_msg = f"Pasta do Outlook '{folder_name}' não encontrada."
            log_message(err_msg)
            send_notification(APP_TITLE, err_msg)
            html_events.append({"type": "error", "message": err_msg, "time": datetime.datetime.now().strftime("%H:%M:%S")})
            return

        # 4. Processamento dos Itens (E-mails)
        # A varredura agora é feita em todos os itens, sem filtro de data inicial
        items = target_folder.Items

        saved_count = 0
        manual_count = 0
        new_emails_processed = 0
        
        log_message(f"Verificando todos os {len(items)} e-mails na pasta '{folder_name}'...")

        for item in items:
            try:
                # Verifica se é um item de email (pode ser convite, etc)
                if item.Class != 43: # 43 = MailItem
                    continue
                
                # Cria um ID único para o e-mail baseado na data/hora de recebimento e assunto
                try:
                    # Garante que o assunto não quebre a lógica se for None
                    subject = item.Subject if item.Subject else ""
                    email_id = f"{item.ReceivedTime.isoformat()}|{subject}"
                except Exception:
                    # Em casos raros, acessar propriedades do item pode falhar. Melhor pular.
                    log_message("Aviso: Não foi possível gerar ID para um item. Pulando.")
                    continue

                # Pula o e-mail se seu ID já estiver no histórico
                if email_id in processed_history:
                    continue

                new_emails_processed += 1
                
                # 4.1 Verificação de Anexos
                if item.Attachments.Count > 0:
                    for attachment in item.Attachments:
                        filename = attachment.FileName
                        if filename.lower().endswith(".pdf"):
                            plate = extract_plate(filename)
                            
                            final_folder = ""
                            
                            # 4.2 Definição do Destino (Com Placa vs Sem Placa)
                            if plate:
                                # Cria nome da pasta: "Veiculo - ABC1234"
                                folder_plate_name = f"Veiculo - {plate}"
                                final_folder = os.path.join(save_path_root, folder_plate_name)
                                if not os.path.exists(final_folder):
                                    os.makedirs(final_folder)
                                    log_message(f"Criada nova pasta de veículo: {folder_plate_name}")
                                    html_events.append({"type": "info", "message": f"Criada nova pasta: <b>{folder_plate_name}</b>", "time": datetime.datetime.now().strftime("%H:%M:%S")})
                                saved_details_list.append({"plate": plate, "filename": filename, "folder": folder_plate_name})
                                saved_count += 1
                            else:
                                # Sem placa identificada
                                final_folder = os.path.join(manual_path, "PDF sem identificação de placa")
                                if not os.path.exists(final_folder):
                                    os.makedirs(final_folder)
                                manual_count += 1
                                log_message(f"PDF sem placa identificada: {filename} -> Enviado para revisão manual.")
                                html_events.append({"type": "warning", "message": f"Sem placa: {filename} -> <i>Revisão Manual</i>", "time": datetime.datetime.now().strftime("%H:%M:%S")})
                                manual_details_list.append({"filename": filename})

                            # 4.3 Salvamento do Arquivo
                            # Adiciona timestamp no nome para evitar sobrescrever arquivos iguais
                            timestamp_fname = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                            save_name = f"{timestamp_fname}_{filename}"
                            full_save_path = os.path.join(final_folder, save_name)
                            
                            attachment.SaveAsFile(full_save_path)
                            
                            if plate:
                                log_message(f"Arquivo salvo: {filename} em {folder_plate_name}")
                                html_events.append({"type": "success", "message": f"Salvo: {filename} <br>Destino: {folder_plate_name}", "time": datetime.datetime.now().strftime("%H:%M:%S")})
                
                # Adiciona o ID ao conjunto de e-mails processados nesta execução,
                # apenas se o processamento do item (e seus anexos) for bem-sucedido.
                processed_in_this_run.add(email_id)

            except Exception as e:
                log_message(f"Erro ao processar item: {str(e)}")
                html_events.append({"type": "error", "message": f"Erro no item: {str(e)}", "time": datetime.datetime.now().strftime("%H:%M:%S")})

        # 5. Finalização
        # Salva o histórico atualizado com os novos e-mails processados
        save_history(processed_history, processed_in_this_run)

        # 6. Geração de Relatórios e Notificação Final
        summary = f"Processo finalizado. {saved_count} arquivos salvos, {manual_count} para revisão."
        log_message(f"Resumo: {new_emails_processed} novos e-mails processados. {summary}")
        log_message("=== FIM DA EXECUÇÃO ===\n")
        
        send_notification(APP_TITLE, summary)
        
        # Gera o HTML
        stats = {
            "new_emails": new_emails_processed,
            "saved": saved_count,
            "manual": manual_count
        }
        generate_html_report(html_events, stats, saved_details_list, manual_details_list)

    except Exception as e:
        err = f"Erro fatal na execução: {str(e)}"
        log_message(err)
        html_events.append({"type": "error", "message": f"Erro Fatal: {str(e)}", "time": datetime.datetime.now().strftime("%H:%M:%S")})
        send_notification(APP_TITLE, "Erro durante a execução. Verifique o log.")

if __name__ == "__main__":
    main()
