# -*- coding: utf-8 -*-
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

import config

# MODIFICADO: A função agora recebe o título do relatório para usar no assunto.
def enviar_email_resumo_mensal(corpo_html, report_title):
    """Envia o e-mail com o resumo de horas do período."""
    print("Preparando e-mail de resumo...")

    destinatario_principal = config.MENSAL_DESTINATARIO_PRINCIPAL
    destinatarios_copia = config.MENSAL_DESTINATARIOS_COPIA
    
    if not destinatario_principal['email']:
        print("AVISO: E-mail de resumo não configurado. Pulando envio.")
        return

    destinatarios_lista = [destinatario_principal['email']] + destinatarios_copia
    
    msg = MIMEMultipart('related')
    # MODIFICADO: Assunto do e-mail agora é dinâmico.
    msg['Subject'] = f"[RESUMO][APONTAMENTO] - {report_title}"
    msg['From'] = f"{config.ASSINATURA_NOME} <{config.EMAIL_REMETENTE}>"
    msg['To'] = destinatario_principal['email']
    if destinatarios_copia:
        msg['Cc'] = ", ".join(destinatarios_copia)
    
    msg.attach(MIMEText(corpo_html, 'html', 'utf-8'))

    try:
        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
        server.starttls()
        server.login(config.EMAIL_REMETENTE, config.EMAIL_SENHA)
        server.sendmail(config.EMAIL_REMETENTE, destinatarios_lista, msg.as_string())
        server.quit()
        print("E-mail de resumo enviado com sucesso.")
    except Exception as e:
        print(f"ERRO ao enviar e-mail de resumo: {e}")
        raise e

def enviar_email_status(timing_report, status_final, erro_msg=""):
    """Envia um e-mail final com o status da execução e os tempos de cada etapa."""
    print("Preparando e-mail de status da execução...")
    
    if not config.STATUS_EMAIL_DESTINATARIO:
        print("Destinatário do e-mail de status não configurado. Pulando envio.")
        return

    msg = MIMEMultipart()
    status_cor = "green" if "SUCESSO" in status_final else "red"
    msg['Subject'] = f"Status do Robô de Apontamentos: {status_final}"
    msg['From'] = f"Robô de Apontamentos <{config.EMAIL_REMETENTE}>"
    msg['To'] = config.STATUS_EMAIL_DESTINATARIO

    tabela_html_linhas = "".join([f"<tr><td style='padding: 8px; border: 1px solid #dddddd;'>{etapa}</td><td style='padding: 8px; border: 1px solid #dddddd; text-align: right;'>{tempo:.2f} segundos</td></tr>" for etapa, tempo in timing_report.items()])

    corpo_html = f"""
    <html><body>
        <h2>Relatório de Execução do Robô de Apontamentos</h2>
        <p><b>Status Final:</b> <span style="color: {status_cor}; font-weight: bold;">{status_final}</span></p>
        <p><b>Data e Hora:</b> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
        <h3>Tempos de Execução por Etapa:</h3>
        <table style="width: 600px; border-collapse: collapse;">
            <thead style="background-color: #4472C4; color: white;">
                <tr><th style="padding: 8px; border: 1px solid #dddddd; text-align: left;">Etapa</th><th style="padding: 8px; border: 1px solid #dddddd; text-align: right;">Duração</th></tr>
            </thead>
            <tbody>{tabela_html_linhas}</tbody>
        </table>
    """

    if status_final == "FALHA":
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
        print(f"E-mail de status enviado com sucesso para {config.STATUS_EMAIL_DESTINATARIO}.")
    except Exception as e:
        print(f"ERRO ao enviar e-mail de status: {e}")