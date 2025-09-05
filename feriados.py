# -*- coding: utf-8 -*-
import holidays
from datetime import date, timedelta
import calendar

def _get_feriados(ano):
    """Função auxiliar para centralizar a criação do calendário de feriados."""
    feriados_br = holidays.country_holidays('BR', state='SP')
    feriados_br.update({
        date(ano, 1, 26): "Aniversário de Santos",
        date(ano, 9, 8): "Nossa Senhora do Monte Serrat"
    })
    return feriados_br

def get_horas_uteis_no_periodo(start_date, end_date):
    """
    (NOVO) Calcula o total de horas úteis em um intervalo de datas específico.
    - Considera uma jornada de 8 horas por dia.
    - Desconta sábados, domingos e feriados.
    """
    # Coleta feriados para todos os anos que o período possa abranger
    anos = range(start_date.year, end_date.year + 1)
    feriados_todos = {}
    for ano in anos:
        feriados_todos.update(_get_feriados(ano))

    dias_uteis = 0
    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() < 5 and current_date not in feriados_todos:
            dias_uteis += 1
        current_date += timedelta(days=1)
    
    print(f"O período de {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')} tem {dias_uteis} dias úteis.")
    return dias_uteis * 8

# As funções antigas abaixo podem ser mantidas ou removidas se não forem mais usadas em outro lugar.
def get_total_horas_uteis_mes(ano, mes):
    feriados_br = _get_feriados(ano)
    dias_no_mes = calendar.monthrange(ano, mes)[1]
    dias_uteis = 0
    for dia in range(1, dias_no_mes + 1):
        data_atual = date(ano, mes, dia)
        if data_atual.weekday() < 5 and data_atual not in feriados_br:
            dias_uteis += 1
    return dias_uteis * 8

def get_horas_uteis_ate_hoje(ano, mes):
    feriados_br = _get_feriados(ano)
    hoje = date.today()
    if ano == hoje.year and mes == hoje.month:
        ultimo_dia_a_contar = hoje.day
    else:
        ultimo_dia_a_contar = calendar.monthrange(ano, mes)[1]
    dias_uteis_passados = 0
    for dia in range(1, ultimo_dia_a_contar + 1):
        data_atual = date(ano, mes, dia)
        if data_atual.weekday() < 5 and data_atual not in feriados_br:
            dias_uteis_passados += 1
    return dias_uteis_passados * 8