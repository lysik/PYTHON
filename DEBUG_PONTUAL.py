import pythoncom
import win32com.client as win32

arquivo = r'C:\Users\Bruno\OneDrive - AGROCONSULT PARTICIPACOES LTDA\VINCULOS\COMUNICA - VBA\APURAÇÃO IRPJ E CSLL TRIMESTRAIS\2020\IRPJ CSLL PREVISÃO 4º TRIMESTRE 2020.xlsx'

pythoncom.CoInitialize()
excel = win32.DispatchEx('Excel.Application')
excel.Visible = True  # Deixa visível para debug

wb = excel.Workbooks.Open(arquivo)
links = wb.LinkSources(1)  # 1 = ExcelLinks

if links:
    print("Links encontrados:")
    for link in links:
        print(link)
else:
    print("Nenhum link encontrado.")

wb.Close(SaveChanges=False)
excel.Quit()
pythoncom.CoUninitialize()
