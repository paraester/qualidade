import os
import pandas as pd
import logging

# Ativar logs para debugging
logging.basicConfig(level=logging.INFO)

# Caminho da pasta 'Dados'
pasta_dados = os.path.join(os.getcwd(), "Dados")

# Função para encontrar o arquivo CSV dentro da pasta "Dados"
def encontrar_arquivo_csv(pasta_dados):
    logging.info(f"Procurando arquivos CSV na pasta: {pasta_dados}")
    for arquivo in os.listdir(pasta_dados):
        if arquivo.endswith('.csv'):
            return os.path.join(pasta_dados, arquivo)
    return None

# Substituições específicas
substituicoes = {
    "GMM - Gerência de Manutenção e Modernização de Soluções/": "",
     


}

# Função para manipular o CSV
def manipular_csv(caminho_arquivo):
    logging.info(f"Manipulando o arquivo CSV: {caminho_arquivo}")

    # Ler o CSV, pulando a primeira linha
    df = pd.read_csv(caminho_arquivo, skiprows=1)
    logging.info("Primeira linha removida com sucesso.")

    # Substituir strings específicas conforme o dicionário 'substituicoes'
    for chave, valor in substituicoes.items():
        df = df.replace(chave, valor, regex=True)
        logging.info(f"Substituição '{chave}' por '{valor}' realizada com sucesso.")

    # Salvar o arquivo novamente
    df.to_csv(caminho_arquivo, index=False)
    logging.info(f"Arquivo salvo novamente em: {caminho_arquivo}")

def main():
    # Procurar o arquivo CSV dentro da pasta 'Dados'
    caminho_arquivo = encontrar_arquivo_csv(pasta_dados)

    if caminho_arquivo:
        logging.info(f"Arquivo CSV encontrado: {caminho_arquivo}")

        # Manipular o arquivo CSV: remover a primeira linha e fazer as substituições
        manipular_csv(caminho_arquivo)
    else:
        logging.error("Nenhum arquivo CSV encontrado na pasta 'Dados'.")

if __name__ == "__main__":
    main()
