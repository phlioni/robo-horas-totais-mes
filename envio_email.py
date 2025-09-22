# -*- coding: utf-8 -*-
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication # MODIFICADO: Importa a classe correta para anexos
from datetime import datetime
import os

import config

def enviar_email_resumo_mensal(corpo_html, report_title, caminho_anexo=None):
    """Envia o e-mail com o resumo de horas do período e um anexo opcional."""
    print("Preparando e-mail de resumo...")

    destinatario_principal = config.MENSAL_DESTINATARIO_PRINCIPAL
    destinatarios_copia = config.MENSAL_DESTINATARIOS_COPIA
    
    if not destinatario_principal['email']:
        print("AVISO: E-mail de resumo não configurado. Pulando envio.")
        return

    destinatarios_lista = [destinatario_principal['email']] + destinatarios_copia
    
    msg = MIMEMultipart('related')
    msg['Subject'] = f"[RESUMO][APONTAMENTO] - {report_title}"
    msg['From'] = f"{config.ASSINATURA_NOME} <{config.EMAIL_REMETENTE}>"
    msg['To'] = destinatario_principal['email']
    if destinatarios_copia:
        msg['Cc'] = ", ".join(destinatarios_copia)
    
    msg.attach(MIMEText(corpo_html, 'html', 'utf-8'))

    # MODIFICADO: Lógica de anexo corrigida para usar MIMEApplication
    if caminho_anexo and os.path.exists(caminho_anexo):
        print(f"Anexando o arquivo: {os.path.basename(caminho_anexo)}")
        try:
            with open(caminho_anexo, "rb") as fil:
                # Cria um anexo do tipo application, que é o correto para arquivos binários como .xlsx
                part = MIMEApplication(
                    fil.read(),
                    Name=os.path.basename(caminho_anexo)
                )
            
            # Adiciona o cabeçalho para que o cliente de e-mail saiba que é um anexo com um nome de arquivo
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho_anexo)}"'
            msg.attach(part)
            print("Anexo adicionado ao e-mail corretamente.")
        except Exception as e:
            print(f"ERRO ao tentar anexar o arquivo: {e}")

    try:
        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
        server.starttls()
        server.login(config.EMAIL_REMETENTE, config.EMAIL_SENHA)
        texto_email = msg.as_string()
        server.sendmail(config.EMAIL_REMETENTE, destinatarios_lista, texto_email)
        server.quit()
        print(f"E-mail de resumo enviado com sucesso para: {', '.join(destinatarios_lista)}")
    except Exception as e:
        print(f"ERRO CRÍTICO ao enviar o e-mail de resumo: {e}")

def enviar_email_status_execucao(status_final, erro_msg, timing_report, tempo_total):
    """Envia um e-mail de status (sucesso ou falha) da execução do robô."""
    print("Preparando e-mail de status da execução...")
    
    if not config.STATUS_EMAIL_DESTINATARIO:
        print("AVISO: E-mail de status não configurado. Pulando envio.")
        return

    msg = MIMEMultipart()
    msg['Subject'] = f"[STATUS][ROBÔ APONTAMENTO] - Execução {status_final}"
    msg['From'] = f"{config.ASSINATURA_NOME} <{config.EMAIL_REMETENTE}>"
    msg['To'] = config.STATUS_EMAIL_DESTINATARIO

    status_cor = "green" if status_final.startswith("SUCESSO") else "red"

    tabela_html_linhas = ""
    for etapa, duracao in timing_report.items():
        tabela_html_linhas += f"<tr><td style='padding: 8px; border: 1px solid #dddddd;'>{etapa}</td><td style='padding: 8px; border: 1px solid #dddddd; text-align: right;'>{duracao:.2f} s</td></tr>"
    
    tabela_html_linhas += f"<tr style='background-color: #f2f2f2; font-weight: bold;'><td style='padding: 8px; border: 1px solid #dddddd;'>Tempo Total</td><td style='padding: 8px; border: 1px solid #dddddd; text-align: right;'>{tempo_total:.2f} s</td></tr>"
    
    corpo_html = f"""
    <html>
    <body style="font-family: Calibri, sans-serif;">
        <h2>Relatório de Execução do Robô de Apontamento</h2>
        <p><b>Status:</b> <span style="color: {status_cor}; font-weight: bold;">{status_final}</span></p>
        <p><b>Data e Hora:</b> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
        <h3>Tempos de Execução por Etapa:</h3>
        <table style="width: 600px; border-collapse: collapse;">
            <thead style="background-color: #4472C4; color: white;">
                <tr><th style="padding: 8px; border: 1px solid #dddddd; text-align: left;">Etapa</th><th style="padding: 8px; border: 1px solid #dddddd; text-align: right;">Duração</th></tr>
            </thead>
            <tbody>{tabela_html_linhas}</tbody>
        </table>
    """

    if "FALHA" in status_final:
        corpo_html += f"""
        <h3 style="color: red;">Detalhes do Erro:</h3>
        <pre style="font-family: 'Courier New', monospace; background-color: #f5f5f5; padding: 10px; border: 1px solid #ccc; white-space: pre-wrap; word-wrap: break-word;">{erro_msg}</pre>
        """

    corpo_html += "</body></html>"
    msg.attach(MIMEText(corpo_html, 'html'))

    try:
        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
        server.starttls()
        server.login(config.EMAIL_REMETENTE, config.EMAIL_SENHA)
        server.sendmail(config.EMAIL_REMETENTE, config.STATUS_EMAIL_DESTINATARIO, msg.as_string())
        server.quit()
        print("E-mail de status enviado com sucesso.")
    except Exception as e:
        print(f"ERRO CRÍTICO ao enviar o e-mail de status: {e}")