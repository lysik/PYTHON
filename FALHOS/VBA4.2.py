import os
import logging
import pythoncom
import win32com.client.gencache  # atualizado aqui
from tqdm import tqdm

# Configurações principais
RAIZ_CAMINHO = r'C:\BASE-TESTE'
LOG_CAMINHO = os.path.join(RAIZ_CAMINHO, 'logderede')
CAMINHO_ANTIGO_PARCIAL = r"C:\Users\brunolysik\AppData\Roaming\Microsoft\Excel\XLSTART\ADM\FINANCEIRO\CONTROLES INTERNOS"
CAMINHO_NOVO_PARCIAL = r"C:\AGRO-DADOS\OneDrive - AGROCONSULT PARTICIPACOES LTDA\COMUNICA - VBA"

# Cria a pasta de log, se não existir
if not os.path.exists(LOG_CAMINHO):
    os.makedirs(LOG_CAMINHO)

# Configura o log
logging.basicConfig(
    filename=os.path.join(LOG_CAMINHO, 'log_vinculos.txt'),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def atualizar_vinculos_excel(caminho_arquivo):
    excel = None
    try:
        pythoncom.CoInitialize()  # Inicializa comunicação COM

        # Usa EnsureDispatch para garantir compatibilidade com propriedades como .Visible
        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(caminho_arquivo)

        links = workbook.LinkSources(1)  # 1 = ExcelLinks
        if links:
            contador = 0
            for link in links:
                if CAMINHO_ANTIGO_PARCIAL in link:
                    novo_link = link.replace(CAMINHO_ANTIGO_PARCIAL, CAMINHO_NOVO_PARCIAL)
                    workbook.ChangeLink(link, novo_link, 1)
                    contador += 1
            if contador > 0:
                logging.info(f"[OK] {caminho_arquivo} - {contador} vínculo(s) atualizado(s).")
            else:
                logging.info(f"[INFO] {caminho_arquivo} - Nenhum vínculo atualizado.")
        else:
            logging.info(f"[INFO] {caminho_arquivo} - Nenhum vínculo encontrado.")

        workbook.Save()
        workbook.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()

    except Exception as e:
        logging.error(f"[ERRO] {caminho_arquivo} - {str(e)}")
        if excel:
            try:
                excel.Quit()
            except:
                pass
        pythoncom.CoUninitialize()

def listar_arquivos_excel(caminho_inicial):
    arquivos_excel = []
    for root, dirs, files in os.walk(caminho_inicial):
        for file in files:
            if file.lower().endswith(('.xls', '.xlsx', '.xlsm')):
                arquivos_excel.append(os.path.join(root, file))
    return arquivos_excel

if __name__ == "__main__":
    arquivos = listar_arquivos_excel(RAIZ_CAMINHO)

    for caminho_arquivo in tqdm(arquivos, desc="Processando arquivos", unit="arquivo"):
        atualizar_vinculos_excel(caminho_arquivo)

    print("\nProcesso concluído.")
