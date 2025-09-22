# -*- coding: utf-8 -*-
import os
import time
import traceback

import config
import automacao_web
import excel_handler
import processamento_dados
import envio_email
import date_logic

def run():
    """Função principal que orquestra o relatório de resumo mensal de horas."""
    timing_report = {}
    status = "SUCESSO"
    error_message = ""
    caminho_arquivo_baixado = None
    caminho_relatorio_excel = None

    start_total_time = time.time()

    try:
        periodo_analise = date_logic.get_analysis_period()
        print(f"Período de análise determinado: {periodo_analise['report_title']}")
        print(f"Botão a ser clicado no site: '{periodo_analise['button_key']}'")

        # Etapa 1: Login e download do relatório
        start_step_time = time.time()
        caminho_arquivo_baixado = automacao_web.login_e_download(periodo_analise['button_key'])
        timing_report["1. Download do Relatório"] = time.time() - start_step_time
        
        # Etapa 2: Desbloquear o arquivo Excel baixado
        start_step_time = time.time()
        excel_handler.unprotect_and_save(caminho_arquivo_baixado)
        timing_report["2. Desbloqueio do Excel"] = time.time() - start_step_time
        
        # Etapa 3: Processar a planilha (filtrar por data e status 'Aprovado')
        start_step_time = time.time()
        df_filtrado = processamento_dados.processar_planilha(
            caminho_arquivo_baixado, 
            periodo_analise['start_date'], 
            periodo_analise['end_date']
        )
        timing_report["3. Processamento da Planilha"] = time.time() - start_step_time
        
        # Etapa 4: Gerar resumo e corpo do e-mail
        start_step_time = time.time()
        df_resumo, corpo_email_html = processamento_dados.gerar_resumo_e_html(
            df_filtrado, 
            periodo_analise['start_date'], 
            periodo_analise['end_date']
        )
        timing_report["4. Geração do Resumo e HTML"] = time.time() - start_step_time
        
        # MODIFICADO: Etapa 5: Criar o arquivo Excel COMPLETO (com 2 abas) para o anexo
        start_step_time = time.time()
        nome_arquivo_excel = f"Relatorio_Horas_{periodo_analise['report_title'].replace(' ', '_').replace('/', '-')}.xlsx"
        # AQUI ESTÁ A MUDANÇA: Passamos ambos os dataframes para a nova função
        caminho_relatorio_excel = processamento_dados.criar_excel_completo(df_filtrado, df_resumo, nome_arquivo_excel)
        timing_report["5. Criação do Anexo Excel Completo"] = time.time() - start_step_time

        # Etapa 6: Enviar e-mail com o resumo e anexo completo
        start_step_time = time.time()
        envio_email.enviar_email_resumo_mensal(
            corpo_email_html, 
            periodo_analise['report_title'], 
            caminho_relatorio_excel
        )
        timing_report["6. Envio do E-mail de Resumo"] = time.time() - start_step_time

    except Exception as e:
        status = "FALHA"
        error_message = traceback.format_exc()
        print(f"\nOcorreu um erro crítico durante a execução:\n{error_message}")

    finally:
        if caminho_arquivo_baixado and os.path.exists(caminho_arquivo_baixado):
            try:
                os.remove(caminho_arquivo_baixado)
                print(f"Arquivo temporário '{os.path.basename(caminho_arquivo_baixado)}' removido.")
            except OSError as e:
                print(f"Erro ao remover arquivo temporário: {e}")

        if caminho_relatorio_excel and os.path.exists(caminho_relatorio_excel):
            try:
                os.remove(caminho_relatorio_excel)
                print(f"Arquivo de anexo '{os.path.basename(caminho_relatorio_excel)}' removido.")
            except OSError as e:
                print(f"Erro ao remover arquivo de anexo: {e}")
        
        total_time = time.time() - start_total_time
        
        envio_email.enviar_email_status_execucao(status, error_message, timing_report, total_time)

if __name__ == '__main__':
    run()