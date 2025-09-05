# -*- coding: utf-8 -*-
from datetime import date, timedelta
import calendar

def get_analysis_period():
    """
    Determina o período de análise com base na regra de negócio.
    A análise sempre cobre do início do mês de referência até a última sexta-feira da semana anterior à data de hoje.
    
    Se a execução ocorrer no início de um mês, antes da primeira sexta-feira, 
    a análise será sobre o mês anterior completo.

    Retorna:
        dict: Um dicionário contendo 'button_key', 'start_date', 'end_date' e 'report_title'.
    """
    today = date.today()
    
    # Encontra o início da semana corrente (Segunda-feira = 0)
    start_of_current_week = today - timedelta(days=today.weekday())
    
    # A data final da análise é sempre a sexta-feira da semana anterior.
    # Se hoje é segunda, a semana passada terminou 3 dias atrás (sex). Se for domingo, 2 dias atrás.
    end_date = start_of_current_week - timedelta(days=3)

    # Verifica se a data de término pertence ao mês anterior
    if end_date.month < today.month or end_date.year < today.year:
        # --- Caso Especial: Analisar o MÊS ANTERIOR COMPLETO ---
        button_to_click = "Mês passado"
        
        # O mês de referência é o mês da data de término
        report_month = end_date.month
        report_year = end_date.year
        
        # Datas para o mês anterior completo
        start_date = date(report_year, report_month, 1)
        _, last_day = calendar.monthrange(report_year, report_month)
        final_end_date = date(report_year, report_month, last_day)

    else:
        # --- Caso Normal: Analisar o MÊS CORRENTE até a última sexta-feira ---
        button_to_click = "Mês Corrente"
        
        # O mês de referência é o mês atual
        report_month = today.month
        report_year = today.year

        # A data de início é o primeiro dia do mês corrente
        start_date = date(report_year, report_month, 1)
        final_end_date = end_date # A data final que já calculamos

    # Formata o título do relatório
    meses_pt = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    nome_mes_pt = meses_pt.get(report_month, "")
    report_title = f"{nome_mes_pt} de {report_year} (período de {start_date.strftime('%d/%m')} a {final_end_date.strftime('%d/%m')})"

    return {
        "button_key": button_to_click,
        "start_date": start_date,
        "end_date": final_end_date,
        "report_title": report_title
    }