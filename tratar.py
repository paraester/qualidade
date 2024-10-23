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
    "GMM - Gerência de Manutenção e Modernização de Soluções": "GMM",
    "/Governo do Estado do Paraná/CELEPAR - Companhia de Tecnologia da Informação e Comunicação do Paraná/DP - Presidência/DDSI - Diretoria de Desenvolvimento, Serviços e Inovação/": "",
    "/Governo do Estado do Paraná/CELEPAR - Companhia de Tecnologia da Informação e Comunicação do Paraná/DP - Presidência/DDSI - Diretoria de Desenvolvimento, Serviços e Inovação": "DDSI",
    "GSI-A - Gerência de Sistemas de Informação - A/": "",
    "GSI-A - Gerência de Sistemas de Informação - A": "GSI-A",
    "GSI-B - Gerência de Sistemas de Informação - B/": "",
    "GSI-B - Gerência de Sistemas de Informação - B": "GSI-B",
    "GSI-C - Gerência de Sistemas de Informação - C/": "",
    "GSI-C - Gerência de Sistemas de Informação - C": "GSI-C",
    "GSI-D - Gerência de Sistemas de Informação - D/": "",
    "GSI-D - Gerência de Sistemas de Informação - D": "GSI-D",
    "GPS - Gerência de Produtos e Serviços/": "",
    "GPS - Gerência de Produtos e Serviços": "GPS",
    "GIA - Gerência de Inovação e Arquitetura/": "",
    "GIA - Gerência de Inovação e Arquitetura": "GIA",    
    "GPG - Gerência de Planejamento e Governança/COPPS - Coordenação de Portfólio, Projetos, Produtos e Serviços": "COPPS",  
    "GPG - Gerência de Planejamento e Governança/CONEC - Coordenação de Operação de Negócio e Capacidade": "CONEC",  
    "GPG - Gerência de Planejamento e Governança/COMEP - Coordenação de Metodologia e Processos": "COMEP", 
    "GPG - Gerência de Planejamento e Governança/COGEF - Coordenação de Gestão de Fornecimento": "COGEF",     
	"COAUT - Coordenação de Automação de Processos": "COAUT",     
	"COBRM - Coordenação de Relacionamento": "COBRM",     
	"CODEP - Coordenação de Design e Prototipação": "CODEP",     
	"CODEV - Coordenação de DevSecOps": "CODEV",     
	"COINA - Coordenação de Inovação Aplicada": "COINA",     
	"COPRE - Coordenação de Projetos Especiais": "COPRE",     
	"COPRO - Coordenação de Projetos Novos": "COPRO",     
	"COS-ADM - Coordenação de Soluções de Administração Pública": "COS-ADM",     
	"COS-EDU - Coordenação de Soluções de Educação": "COS-EDU",     
	"COS-GOV - Coordenação de Soluções de Governo": "COS-GOV",     
	"COS-GPE - Coordenação de Soluções de Gestão de Pessoas": "COS-GPE",     
	"COS-JUS - Coordenação de Soluções de Justiça e Fiscalização": "COS-JUS",     
	"COS-LOG - Coordenação de Soluções de Logística e Cidades": "COS-LOG",     
	"COS-MAG - Coordenação de Soluções de Meio Ambiente e Agronegócio": "COS-MAG",     
	"COS-PGO - Coordenação de Soluções de Produtos de Governo": "COS-PGO",     
	"COS-SAU - Coordenação de Soluções de Saúde": "COS-SAU",     
	"COSIN-A1 - Coordenação de Sistemas de Informação - A1": "COSIN-A1",     
	"COSIN-A2 - Coordenação de Sistemas de Informação - A2": "COSIN-A2",     
	"COSIN-A3 - Coordenação de Sistemas de Informação - A3": "COSIN-A3",     
	"COSIN-A4 - Coordenação de Sistemas de Informação - A4": "COSIN-A4",     
    "COSIN-A5 - Coordenação de Sistemas de Informação - A5": "COSIN-A5",     
    "COSIN-C1 - Coordenação de Sistemas de Informação - C1": "COSIN-C1",     
    "COSIN-C2 - Coordenação de Sistemas de Informação - C2": "COSIN-C2",     
    "COSIN-C3 - Coordenação de Sistemas de Informação - C3": "COSIN-C3",     
    "COSIN-D1 - Coordenação de Sistemas de Informação - D1": "COSIN-D1",     
    "COSIN-D2 - Coordenação de Sistemas de Informação - D2": "COSIN-D2",     
    "COSIN-D3 - Coordenação de Sistemas de Informação - D3": "COSIN-D3",     
    "COTAP - Coordenação de Tecnologia Aplicada": "COTAP",     
	"[INATIVO]COSIN-A6 - Coordenação de Sistemas de Informação A6": "COSIN-A6",     
	"[INATIVO]COSIN-A7 - Coordenação de Sistemas de Informação - A7": "COSIN-A7",     
	"[INATIVO]GSCM - Gerência de Serviços de Comunicação Multimídia": "GSCM",     
	"[INATIVO]GSI-G - Gerência de Sistemas de Informação G": "GSI-G",     
	"[INATIVO]GSI-G - Gerência de Sistemas de Informação G/[INATIVO]COSIN-G1 - Coordenação de Sistemas de Informação - G1": "[INATIVO]COSIN-G1",     
	"[INATIVO]GSI-G - Gerência de Sistemas de Informação G/[INATIVO]COSIN-G2 - Coordenação de Sistemas de Informação - G2": "[INATIVO]COSIN-G2",     


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
