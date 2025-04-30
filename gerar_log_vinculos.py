import os
import csv
import pythoncom
import win32com.client as win32
from tqdm import tqdm

# Caminho base onde estão os arquivos
RAIZ_CAMINHO = r'C:\BASE-TESTE'
# Caminho para salvar o log CSV
CAMINHO_CSV_LOG = r'C:\logderede\log_vinculos_detectados.csv'

# Garante que a pasta de log existe
os.makedirs(os.path.dirname(CAMINHO_CSV_LOG), exist_ok=True)

def listar_arquivos_excel(caminho_inicial):
    arquivos_excel = []
    for root, dirs, files in os.walk(caminho_inicial):
        for file in files:
            if file.lower().endswith(('.xls', '.xlsx', '.xlsm')):
                arquivos_excel.append(os.path.join(root, file))
    return arquivos_excel

def extrair_vinculos_excel(caminho_arquivo):
    pythoncom.CoInitialize()
    excel = win32.DispatchEx('Excel.Application')
    excel.DisplayAlerts = False

    vinculos = []
    try:
        wb = excel.Workbooks.Open(caminho_arquivo, UpdateLinks=0)
        links = wb.LinkSources(1)  # 1 = ExcelLinks
        if links:
            vinculos.extend(links)
        wb.Close(False)
    except Exception as e:
        vinculos.append(f"[ERRO] {str(e)}")
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()
    return vinculos

def gerar_log_de_vinculos():
    arquivos = listar_arquivos_excel(RAIZ_CAMINHO)

    with open(CAMINHO_CSV_LOG, mode='w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Arquivo Excel', 'Vínculo Detectado'])

        for caminho_arquivo in tqdm(arquivos, desc="Analisando arquivos", unit="arquivo"):
            vinculos = extrair_vinculos_excel(caminho_arquivo)
            if vinculos:
                for link in vinculos:
                    writer.writerow([caminho_arquivo, link])
            else:
                writer.writerow([caminho_arquivo, 'Nenhum vínculo encontrado'])

    print(f"\n✅ Log salvo em: {CAMINHO_CSV_LOG}")

if __name__ == "__main__":
    gerar_log_de_vinculos()
