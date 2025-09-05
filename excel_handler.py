# -*- coding: utf-8 -*-
import win32com.client as win32

def unprotect_and_save(file_path):
    """Abre um arquivo Excel, remove a proteção de 'Modo de Exibição Protegido' e salva."""
    print("Tentando desbloquear o arquivo Excel...")
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        workbook = excel.Workbooks.Open(file_path, UpdateLinks=False, ReadOnly=False)
        
        # A simples ação de abrir e salvar com o win32com remove a proteção.
        workbook.Save()
        workbook.Close(SaveChanges=True)
        excel.Quit()
        print("Arquivo salvo e desbloqueado com sucesso.")
    except Exception as e:
        print(f"ERRO ao tentar desbloquear o arquivo Excel: {e}")
        if 'excel' in locals():
            excel.Quit()
        raise e