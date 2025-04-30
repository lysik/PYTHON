import csv

# Caminho do arquivo CSV gerado
LOG_ORIGINAL_CSV = r"C:\logderede\log_vinculos_detectados.csv"
LOG_RESUMIDO_CSV = r"C:\logderede\log_vinculos_resumidos.csv"

# Dicionário para armazenar vínculos e suas contagens
vinculos_dict = {}

# Lê o arquivo CSV de log original
with open(LOG_ORIGINAL_CSV, mode='r', encoding='utf-8') as arquivo:
    leitor = csv.reader(arquivo)
    for linha in leitor:
        if len(linha) > 1:
            vinculo = linha[1].strip()  # Considerando que o vínculo está na segunda coluna
            if vinculo:
                vinculos_dict[vinculo] = vinculos_dict.get(vinculo, 0) + 1

# Escreve o log resumido com vínculos únicos e suas contagens
with open(LOG_RESUMIDO_CSV, mode='w', newline='', encoding='utf-8') as arquivo_resumido:
    escritor = csv.writer(arquivo_resumido)
    escritor.writerow(['Vínculo', 'Contagem'])  # Cabeçalho
    for vinculo, contagem in vinculos_dict.items():
        escritor.writerow([vinculo, contagem])

print(f"Log resumido gerado em: {LOG_RESUMIDO_CSV}")
