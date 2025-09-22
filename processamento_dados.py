# -*- coding: utf-8 -*-
import pandas as pd
from datetime import datetime
import feriados
import config

def get_pessoas_ativas():
    """
    Lê a lista de pessoas ativas a partir do arquivo CSV/Excel definido no config.
    """
    try:
        print(f"Lendo arquivo de pessoas ativas de: {config.CAMINHO_PESSOAS_ATIVAS}")
        
        if config.CAMINHO_PESSOAS_ATIVAS.endswith('.xlsx'):
             df_ativas = pd.read_excel(config.CAMINHO_PESSOAS_ATIVAS, engine='openpyxl')
        else:
             df_ativas = pd.read_csv(config.CAMINHO_PESSOAS_ATIVAS, encoding='latin-1')

        df_ativas.columns = df_ativas.columns.str.strip()
        
        if 'Nome' not in df_ativas.columns:
            raise ValueError("A coluna 'Nome' não foi encontrada no arquivo de pessoas ativas.")
        
        df_ativas['Profissional'] = df_ativas['Nome'].str.strip()
        
        print(f"Encontradas {len(df_ativas)} pessoas ativas.")
        return df_ativas[['Profissional']]
    except FileNotFoundError:
        print(f"ERRO CRÍTICO: O arquivo de pessoas ativas não foi encontrado em '{config.CAMINHO_PESSOAS_ATIVAS}'. Verifique o nome e o local do arquivo.")
        raise
    except Exception as e:
        print(f"ERRO ao ler o arquivo de pessoas ativas: {e}")
        raise

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
        
        print("Filtrando o relatório para manter apenas as horas com situação 'Aprovado'.")
        df_aprovado = df[df['Situação'].str.strip() == 'Aprovado'].copy()
        
        df_filtrado = df_aprovado[(df_aprovado['Data'] >= pd.to_datetime(start_date)) & (df_aprovado['Data'] <= pd.to_datetime(end_date))].copy()
        
        df_filtrado['Profissional'] = df_filtrado['Profissional'].str.strip()
        df_filtrado['Horas'] = pd.to_numeric(df_filtrado['Horas'], errors='coerce').fillna(0)
        
        print("Processamento da planilha concluído.")
        return df_filtrado

    except Exception as e:
        print(f"ERRO ao processar a planilha: {e}")
        raise

def gerar_resumo_e_html(df_filtrado, start_date, end_date):
    """
    Cria o DataFrame de RESUMO e o corpo HTML do e-mail, incluindo a tabela de totais.
    """
    print("Gerando resumo de horas por profissional...")
    
    horas_esperadas_periodo = feriados.get_horas_uteis_no_periodo(start_date, end_date)
    
    if not df_filtrado.empty:
        resumo_profissionais = df_filtrado.groupby('Profissional')['Horas'].sum().reset_index()
        resumo_profissionais.rename(columns={'Horas': 'Horas Aprovadas'}, inplace=True)
    else:
        resumo_profissionais = pd.DataFrame(columns=['Profissional', 'Horas Aprovadas'])

    df_pessoas_ativas = get_pessoas_ativas()
    df_resumo_completo = pd.merge(df_pessoas_ativas, resumo_profissionais, on='Profissional', how='left')
    df_resumo_completo['Horas Aprovadas'] = df_resumo_completo['Horas Aprovadas'].fillna(0)
    df_resumo_completo['Total Horas Esperadas'] = horas_esperadas_periodo
    df_resumo_completo['Total Saldo'] = df_resumo_completo['Horas Aprovadas'] - df_resumo_completo['Total Horas Esperadas']
    df_resumo_completo.sort_values(by='Profissional', inplace=True)
    
    print("Resumo final gerado. Convertendo para HTML...")
    html_table_individual = dataframe_to_html(df_resumo_completo)

    total_aprovadas = df_resumo_completo['Horas Aprovadas'].sum()
    total_esperadas = df_resumo_completo['Total Horas Esperadas'].sum()
    saldo_geral = df_resumo_completo['Total Saldo'].sum()
    html_table_geral = criar_html_resumo_geral(total_aprovadas, total_esperadas, saldo_geral)

    corpo_html = f"""
    <html>
        <body style="font-family: Calibri, sans-serif;">
            <h2>Resumo de Apontamento de Horas</h2>
            <p>Olá,</p>
            <p>Segue abaixo o resumo de horas aprovadas para o período de <b>{start_date.strftime('%d/%m/%Y')}</b> a <b>{end_date.strftime('%d/%m/%Y')}</b>.</p>
            <p>Para análise e filtros, utilize o relatório completo em anexo.</p>
            {html_table_individual}
            {html_table_geral}
            <br>
            <p>Este é um e-mail automático enviado pelo Robô de Apontamento de Horas.</p>
        </body>
    </html>
    """
    
    return df_resumo_completo, corpo_html

def dataframe_to_html(df):
    """Converte um DataFrame pandas para uma string de tabela HTML com estilos."""
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
        for col_name, cell_value in row.items():
            align_style = 'text-align: right;' if isinstance(cell_value, (int, float)) else 'text-align: left;'
            color_style = ''
            if col_name == 'Total Saldo':
                if cell_value < 0:
                    color_style = 'color: red; font-weight: bold;'
                elif cell_value > 0:
                    color_style = 'color: green; font-weight: bold;'
            final_style = f"{style_td} {align_style} {color_style}"
            cell_display = f'{cell_value:.2f}' if isinstance(cell_value, (int, float)) else cell_value
            row_html += f'<td style="{final_style}">{cell_display}</td>'
        row_html += '</tr>'
        html_body_rows.append(row_html)
        
    html_body = f'<tbody>{"".join(html_body_rows)}</tbody>'
    return f'<table style="{style_table}">{html_header}{html_body}</table>'

def criar_html_resumo_geral(total_aprovadas, total_esperadas, saldo_geral):
    """Cria uma tabela HTML formatada para o resumo geral da equipe."""
    style_table_geral = 'width: auto; max-width: 450px; border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 11pt; margin-top: 25px;'
    style_td_label = 'padding: 10px 15px; text-align: left; border: 1px solid #dddddd; font-weight: bold;'
    style_td_valor = 'padding: 10px 15px; text-align: right; border: 1px solid #dddddd;'
    
    saldo_color_style = ''
    if saldo_geral < 0:
        saldo_color_style = 'color: red; font-weight: bold;'
    elif saldo_geral > 0:
        saldo_color_style = 'color: green; font-weight: bold;'

    return f"""
    <h3 style="font-family: Calibri, sans-serif; margin-top: 30px;">Resumo Geral da Equipe</h3>
    <table style="{style_table_geral}">
        <tbody>
            <tr>
                <td style="{style_td_label}">Total de Horas Aprovadas</td>
                <td style="{style_td_valor}">{total_aprovadas:.2f}</td>
            </tr>
            <tr style="background-color: #f8f8f8;">
                <td style="{style_td_label}">Total de Horas Esperadas</td>
                <td style="{style_td_valor}">{total_esperadas:.2f}</td>
            </tr>
            <tr>
                <td style="{style_td_label}">Saldo Geral</td>
                <td style="{style_td_valor} {saldo_color_style}">{saldo_geral:.2f}</td>
            </tr>
        </tbody>
    </table>
    """

def criar_excel_completo(df_detalhado, df_resumo, caminho_arquivo):
    """
    (MODIFICADO) Salva um arquivo Excel com duas abas:
    1. Relatorio Detalhado: Com todos os lançamentos de horas.
    2. Resumo de Horas: Com o resumo por profissional.
    """
    print(f"Criando arquivo Excel completo (Detalhado e Resumo) em: {caminho_arquivo}")

    colunas_desejadas = [
        'Data', 'Projeto', 'Profissional', 'Horas', 
        'Situação', 'Atividade', 'Descrição'
    ]
    colunas_existentes = [col for col in colunas_desejadas if col in df_detalhado.columns]
    df_para_excel_detalhado = df_detalhado[colunas_existentes]

    try:
        with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
            # --- Aba 1: Relatório Detalhado ---
            df_para_excel_detalhado.to_excel(writer, sheet_name='Relatorio Detalhado', index=False)
            worksheet1 = writer.sheets['Relatorio Detalhado']
            
            # --- Aba 2: Resumo de Horas ---
            df_resumo.to_excel(writer, sheet_name='Resumo de Horas', index=False)
            worksheet2 = writer.sheets['Resumo de Horas']
            
            # --- Formatações ---
            workbook = writer.book
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top',
                'fg_color': '#4472C4', 'font_color': 'white', 'border': 1
            })

            # Formatação Aba 1
            for col_num, value in enumerate(df_para_excel_detalhado.columns.values):
                worksheet1.write(0, col_num, value, header_format)
            for i, col in enumerate(df_para_excel_detalhado.columns):
                column_len = max(df_para_excel_detalhado[col].astype(str).map(len).max(), len(col)) + 3
                worksheet1.set_column(i, i, min(column_len, 50))
            
            # Formatação Aba 2
            for col_num, value in enumerate(df_resumo.columns.values):
                worksheet2.write(0, col_num, value, header_format)
            
            red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            saldo_col_index = df_resumo.columns.get_loc('Total Saldo')
            worksheet2.conditional_format(1, saldo_col_index, len(df_resumo), saldo_col_index,
                                         {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red_format})
            worksheet2.conditional_format(1, saldo_col_index, len(df_resumo), saldo_col_index,
                                         {'type': 'cell', 'criteria': '>', 'value': 0, 'format': green_format})

            for i, col in enumerate(df_resumo.columns):
                column_len = max(df_resumo[col].astype(str).map(len).max(), len(col)) + 3
                worksheet2.set_column(i, i, column_len)
        
        print("Arquivo Excel completo criado com sucesso.")
        return caminho_arquivo
    except Exception as e:
        print(f"ERRO ao criar o arquivo Excel: {e}")
        raise