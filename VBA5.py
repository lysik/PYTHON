import os
import logging
import pythoncom
import win32com.client as win32
from tqdm import tqdm

# Configurações principais
RAIZ_CAMINHO = r'C:\BASE-TESTE'
LOG_CAMINHO = os.path.join(RAIZ_CAMINHO, 'logderede')

#####################################################################
#####################################################################
########### TROQUE O CAMINHO DO VÍNCULO NOVO AQUI: ##################
CAMINHO_NOVO_PARCIAL = r"https://agroconsultsc.sharepoint.com/sites/SHARE-000001/VINCULOS/COMUNICA - VBA"
#####################################################################
#####################################################################
#####################################################################

# Lista de possíveis caminhos antigos a substituir (caminhos locais e SharePoint)
CAMINHOS_ANTIGOS = [
    r"C:\Users\brunolysik\AppData\Roaming\Microsoft\Excel\XLSTART\ADM\FINANCEIRO\CONTROLES INTERNOS",
    r"https://agroconsultsc.sharepoint.com/sites/COMUNICA/DADOS/VBA",
    r"https://agroconsultsc.sharepoint.com/AGRO-DADOS/OneDrive - AGROCONSULT PARTICIPACOES LTDA",
    r"C:\AGRO-DADOS\OneDrive - AGROCONSULT PARTICIPACOES LTDA\COMUNICA - VBA"
    r"C:\Users\Bruno\AppData\Roaming\Microsoft\Excel\XLSTART\ADM\FINANCEIRO\CONTROLES INTERNOS"
    r"C:\ADM\FINANCEIRO\CONTROLES INTERNOS\CONTROLE FATURAMENTO"
    r"Z:\ADM\FINANCEIRO\CONTROLES INTERNOS\CONTROLE FATURAMENTO"
]

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
        excel = win32.DispatchEx('Excel.Application')  # Cria nova instância limpa do Excel
        excel.DisplayAlerts = False  # Evita caixas de mensagens

        workbook = excel.Workbooks.Open(caminho_arquivo)

        links = workbook.LinkSources(1)  # 1 = ExcelLinks
        if links:
            contador = 0
            for link in links:
                for caminho_antigo in CAMINHOS_ANTIGOS:
                    if caminho_antigo in link:
                        novo_link = link.replace(caminho_antigo, CAMINHO_NOVO_PARCIAL)
                        workbook.ChangeLink(link, novo_link, 1)
                        contador += 1
                        break
            if contador > 0:
                logging.info(f"[OK] {caminho_arquivo} - {contador} vínculo(s) atualizado(s).")
                print(f"[OK] {os.path.basename(caminho_arquivo)} - {contador} vínculo(s) atualizado(s).")
            else:
                logging.info(f"[INFO] {caminho_arquivo} - Nenhum vínculo atualizado.")
                print(f"[INFO] {os.path.basename(caminho_arquivo)} - Nenhum vínculo atualizado.")
        else:
            logging.info(f"[INFO] {caminho_arquivo} - Nenhum vínculo encontrado.")
            print(f"[INFO] {os.path.basename(caminho_arquivo)} - Nenhum vínculo encontrado.")

        workbook.Save()
        workbook.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()

    except Exception as e:
        logging.error(f"[ERRO] {caminho_arquivo} - {str(e)}")
        print(f"[ERRO] {os.path.basename(caminho_arquivo)} - {str(e)}")
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
