# -*- coding: utf-8 -*-
import pandas as pd
from datetime import datetime
import feriados

# MODIFICADO: A função agora também filtra os dados pelo período.
def processar_planilha(caminho_arquivo, start_date, end_date):
    """Lê, limpa e filtra o relatório de horas baixado para o período de análise."""
    print(f"Processando a planilha e filtrando dados entre {start_date.strftime('%d/%m/%Y')} e {end_date.strftime('%d/%m/%Y')}...")
    try:
        df = pd.read_excel(caminho_arquivo, skiprows=18, engine='openpyxl')
        df.columns = df.columns.str.strip()
        
        colunas_necessarias = ['Data', 'Profissional', 'Situação', 'Horas']
        for col in colunas_necessarias:
            if col not in df.columns:
                raise ValueError(f"Coluna obrigatória '{col}' não encontrada no relatório.")

        df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
        df.dropna(subset=['Data', 'Profissional'], inplace=True)
        
        # NOVO: Filtra o DataFrame para incluir apenas as datas no período de análise
        df = df[(df['Data'] >= pd.to_datetime(start_date)) & (df['Data'] <= pd.to_datetime(end_date))]
        
        df['Profissional'] = df['Profissional'].astype(str).str.strip()
        df = df[df['Profissional'] != '']
        
        print("Planilha processada, filtrada e limpa com sucesso.")
        return df
    except Exception as e:
        print(f"ERRO CRÍTICO ao processar a planilha: {e}")
        raise e

# MODIFICADO: Função reescrita para analisar o período específico.
def analisar_horas_periodo(df_filtrado, start_date, end_date):
    """
    Analisa os dados do período e cria um DataFrame consolidado.
    """
    print("Iniciando análise de horas para o período...")
    
    # Cálculo de horas úteis para o período de análise
    horas_uteis_periodo = feriados.get_horas_uteis_no_periodo(start_date, end_date)

    # DataFrame base com horas lançadas
    df_consolidado = df_filtrado.groupby('Profissional').agg(
        Horas_Lancadas=('Horas', 'sum')
    ).reset_index()

    # Adiciona colunas relevantes para o período
    df_consolidado['Horas_Esperadas_Periodo'] = horas_uteis_periodo
    df_consolidado['Saldo_Periodo'] = df_consolidado['Horas_Lancadas'] - df_consolidado['Horas_Esperadas_Periodo']
    
    # Organiza a ordem das colunas para melhor visualização
    ordem_colunas = [
        'Profissional',
        'Horas_Lancadas',
        'Horas_Esperadas_Periodo',
        'Saldo_Periodo'
    ]
    df_consolidado = df_consolidado[ordem_colunas]
    
    print(f"Análise do período concluída para {len(df_consolidado)} profissionais.")
    return df_consolidado, horas_uteis_periodo

# MODIFICADO: Função reescrita para criar o novo corpo de e-mail.
def criar_corpo_email_resumo(df_consolidado, horas_uteis_periodo, report_title):
    """Cria o corpo do e-mail com a tabela consolidada do período."""
    print("Gerando corpo do e-mail consolidado...")
    
    data_extracao_str = datetime.now().strftime('%d/%m/%Y às %H:%M')
    titulo_email = f"Resumo de Apontamento de Horas - {report_title}"
    
    texto_principal = f"""
    <p>Prezados,</p>
    <p>Segue abaixo o resumo consolidado do status de apontamento de horas para o período de <b>{report_title}</b>.</p>
    <p>Este relatório foi gerado em <b>{data_extracao_str}</b> e considera um total de <b>{horas_uteis_periodo} horas úteis</b> no período.</p>
    """
    
    tabela_principal_html = _gerar_tabela_html("Resumo de Apontamentos", df_consolidado)

    total_lancado = df_consolidado['Horas_Lancadas'].sum()
    total_esperado = df_consolidado['Horas_Esperadas_Periodo'].sum()
    total_saldo = df_consolidado['Saldo_Periodo'].sum()
    df_total = pd.DataFrame([{'Total_Horas_Lancadas': total_lancado, 'Total_Horas_Esperadas': total_esperado, 'Total_Saldo': total_saldo}])
    tabela_totais_html = _gerar_tabela_html("Totalizador Geral do Período", df_total)

    return f"""
    <html>
    <body style="font-family: 'Calibri', sans-serif; font-size: 11pt; color: #333;">
        <h2>{titulo_email}</h2>
        {texto_principal}
        <hr style="border: 0; border-top: 1px solid #ccc; margin: 20px 0;">
        {tabela_principal_html}
        {tabela_totais_html}
        <p><i>Este e-mail foi gerado automaticamente pelo robô.</i></p>
    </body>
    </html>
    """

def _gerar_tabela_html(titulo, df):
    """Função auxiliar para criar uma tabela HTML com estilos e nomes de colunas em pt-BR."""
    if df.empty:
        return f'<h3 style="font-family: Calibri, sans-serif;">{titulo}</h3><p>Nenhum registro encontrado.</p>'

    df = df.copy()

    # MODIFICADO: Novo mapa de colunas, removendo as colunas de Mês.
    mapa_colunas = {
        'Profissional': 'Profissional',
        'Horas_Lancadas': 'Horas Lançadas',
        'Horas_Esperadas_Periodo': 'Horas Esperadas (Período)',
        'Saldo_Periodo': 'Saldo (Período)',
        'Total_Horas_Lancadas': 'Total Horas Lançadas',
        'Total_Horas_Esperadas': 'Total Horas Esperadas',
        'Total_Saldo': 'Total Saldo'
    }
    df.rename(columns=mapa_colunas, inplace=True)

    style_table = 'width: auto; max-width: 800px; border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 11pt; margin-bottom: 25px;'
    style_th = 'background-color: #4472C4; color: #ffffff; padding: 12px 15px; text-align: left; font-weight: bold; border: 1px solid #dddddd;'
    style_td = 'padding: 12px 15px; text-align: left; border: 1px solid #dddddd;'
    style_tr_even = 'background-color: #f8f8f8;'
    
    headers = "".join([f'<th style="{style_th}">{col}</th>' for col in df.columns])
    html_header = f'<thead><tr>{headers}</tr></thead>'
    
    html_body_rows = []
    for index, row in df.iterrows():
        tr_style = style_tr_even if index % 2 == 0 else ''
        row_html = f'<tr style="{tr_style}">'
        for cell_value in row:
            if isinstance(cell_value, (int, float)):
                 cell_value = f"{cell_value:,.2f}".replace(',', 'v').replace('.', ',').replace('v', '.')
            row_html += f'<td style="{style_td}">{cell_value}</td>'
        row_html += '</tr>'
        html_body_rows.append(row_html)
    html_body = f'<tbody>{"".join(html_body_rows)}</tbody>'

    return f"""
    <h3 style="font-family: Calibri, sans-serif;">{titulo}</h3>
    <table style="{style_table}">
        {html_header}
        {html_body}
    </table>
    """