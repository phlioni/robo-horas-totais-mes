# -*- coding: utf-8 -*-
import os
import time
import traceback

import config
import automacao_web
import excel_handler
import processamento_dados
import envio_email
import date_logic # NOVO: Importa a lógica de data

def run():
    """Função principal que orquestra o relatório de resumo mensal de horas."""
    timing_report = {}
    status = "SUCESSO"
    error_message = ""
    caminho_arquivo_baixado = None
    
    start_total_time = time.time()

    try:
        # NOVO: Determina o período de análise antes de começar
        periodo_analise = date_logic.get_analysis_period()
        print(f"Período de análise determinado: {periodo_analise['report_title']}")
        print(f"Botão a ser clicado no site: '{periodo_analise['button_key']}'")

        # Etapa 1: Login e download do relatório com base no período
        start_step_time = time.time()
        # MODIFICADO: Passa o botão a ser clicado como argumento
        caminho_arquivo_baixado = automacao_web.login_e_download(periodo_analise['button_key'])
        timing_report["1. Download do Relatório"] = time.time() - start_step_time
        
        # Etapa 2: Desbloquear o arquivo Excel
        start_step_time = time.time()
        excel_handler.unprotect_and_save(caminho_arquivo_baixado)
        timing_report["2. Desbloqueio do Arquivo Excel"] = time.time() - start_step_time
        
        # Etapa 3: Processar a planilha, agora filtrando pelas datas corretas
        start_step_time = time.time()
        # MODIFICADO: Passa as datas de início e fim para o processamento
        df_filtrado = processamento_dados.processar_planilha(caminho_arquivo_baixado, periodo_analise['start_date'], periodo_analise['end_date'])
        timing_report["3. Processamento e Filtro da Planilha"] = time.time() - start_step_time
        
        if df_filtrado is not None and not df_filtrado.empty:
            # Etapa 4: Analisar horas e criar DataFrame consolidado
            start_step_time = time.time()
            # MODIFICADO: Usa as datas para calcular as horas esperadas
            df_consolidado, horas_uteis_periodo = processamento_dados.analisar_horas_periodo(df_filtrado, periodo_analise['start_date'], periodo_analise['end_date'])
            timing_report["4. Análise e Consolidação dos Dados"] = time.time() - start_step_time

            # Etapa 5: Gerar corpo do e-mail
            start_step_time = time.time()
            # MODIFICADO: Passa o título e as horas do período
            corpo_email_html = processamento_dados.criar_corpo_email_resumo(df_consolidado, horas_uteis_periodo, periodo_analise['report_title'])
            timing_report["5. Geração do Corpo de E-mail"] = time.time() - start_step_time

            # Etapa 6: Enviar e-mail com o resumo
            start_step_time = time.time()
            # MODIFICADO: Passa o título do relatório para o assunto do e-mail
            envio_email.enviar_email_resumo_mensal(corpo_email_html, periodo_analise['report_title'])
            timing_report["6. Envio do E-mail de Resumo"] = time.time() - start_step_time

        else:
            print(f"Nenhum dado de apontamento encontrado no período de {periodo_analise['start_date'].strftime('%d/%m/%Y')} a {periodo_analise['end_date'].strftime('%d/%m/%Y')}.")
            status = "SUCESSO (SEM DADOS)"

    except Exception as e:
        status = "FALHA"
        error_message = traceback.format_exc()
        print(f"\nOcorreu um erro crítico durante a execução:\n{error_message}")

    finally:
        # Limpa o arquivo baixado
        if caminho_arquivo_baixado and os.path.exists(caminho_arquivo_baixado):
            try:
                os.remove(caminho_arquivo_baixado)
                print(f"Arquivo temporário '{os.path.basename(caminho_arquivo_baixado)}' removido.")
            except OSError as e:
                print(f"Erro ao remover arquivo temporário: {e}")

        timing_report["Tempo Total de Execução"] = time.time() - start_total_time
        print("\nProcesso finalizado.")
        # Envia o e-mail de status da execução
        envio_email.enviar_email_status(timing_report, status, error_message)

if __name__ == "__main__":
    run()