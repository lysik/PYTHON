import os
import shutil
import time

# Caminho base
base_path = r"C:\Users\Bruno\OneDrive - AGROCONSULT PARTICIPACOES LTDA\DESCARTE"

# Função para excluir arquivos e pastas com delay de 1 segundo
def delete_files_and_folders(path):
    # Excluir arquivos
    for item in os.listdir(path):
        item_path = os.path.join(path, item)
        
        if os.path.isfile(item_path):
            os.remove(item_path)
            print(f"Arquivo excluído: {item_path}")
        
        elif os.path.isdir(item_path):
            # Recursão para pastas dentro da pasta atual
            delete_files_and_folders(item_path)
            os.rmdir(item_path)
            print(f"Pasta excluída: {item_path}")
    
    # Aguardar 3 segundo após processar todos os arquivos e subpastas na pasta
    time.sleep(3)
    print(f"Aguardando 3 segundo antes de excluir o conteúdo de: {path}")

# Iniciar exclusão a partir da pasta base
delete_files_and_folders(base_path)
