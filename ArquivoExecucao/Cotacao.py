import requests 
import pandas as pd
from datetime import datetime

# Caminho do arquivo de log
log_file = "C:\Users\Marcola\Documents\BI---Cota-o\ArquivoExecucao\execution_log.txt"

# Função para adicionar mensagens ao log
def log_message(message):
    with open(log_file, "a") as log:
        log.write(f"{datetime.now()} - {message}\n")

# Registrando o início da execução
log_message("Início do script de cotações.")

# Fazendo a requisição da API
try:
    log_message("Iniciando requisição da API...")
    requisicao = requests.get("https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL,GBP-BRL")
    requisicao.raise_for_status()  # Verifica se ocorreu algum erro
    requisicao_dic = requisicao.json()
    log_message("Requisição bem-sucedida!")
except requests.exceptions.RequestException as e:
    log_message(f"Erro na requisição da API: {e}")
    exit(1)

# Pegando as cotações
try:
    cotacao_dolar = float(requisicao_dic['USDBRL']['bid'])
    cotacao_euro = float(requisicao_dic['EURBRL']['bid'])
    cotacao_btc = float(requisicao_dic['BTCBRL']['bid']) 
    cotacao_libra= float(requisicao_dic['GBPBRL']['bid']) 
    log_message(f"Cotações obtidas: Dólar: {cotacao_dolar}, Euro: {cotacao_euro}, Bitcoin: {cotacao_btc}, libra: {cotacao_libra}")
except KeyError as e:
    log_message(f"Erro ao acessar as cotações no JSON: {e}")
    exit(1)

# Criando novas linhas com os dados atualizados
data_atual = datetime.now()
novas_linhas = pd.DataFrame([{
    "Moeda": "Dólar", 
    "Cotação": cotacao_dolar, 
    "Data da última atualização": data_atual
}, {
    "Moeda": "Euro", 
    "Cotação": cotacao_euro, 
    "Data da última atualização": data_atual
}, {
    "Moeda": "Bitcoin", 
    "Cotação": cotacao_btc, 
    "Data da última atualização": data_atual
},{
    "Moeda": "Libra",
    "Cotação": cotacao_libra,
    "Data da última atualização": data_atual      
}])

# Caminho para o arquivo Excel onde as cotações serão salvas
arquivo = "C:\Users\Marcola\Documents\BI---Cota-o\ArquivoExceCotações.xlsx"

# Tentando carregar o arquivo Excel e adicionar os dados
try:
    log_message(f"Tentando carregar o arquivo Excel: {arquivo}")
    tabela = pd.read_excel(arquivo)
    log_message("Arquivo Excel carregado com sucesso.")
    
    # Se a tabela estiver vazia ou contiver apenas valores NaN, reinicia a tabela com as novas linhas
    if tabela.empty or tabela.isna().all().all():
        tabela = novas_linhas
    else:
        tabela = pd.concat([tabela, novas_linhas], ignore_index=True)
        log_message("Novas linhas adicionadas à tabela existente.")
except FileNotFoundError:
    # Se o arquivo não for encontrado, cria uma nova planilha com as novas linhas
    tabela = novas_linhas
    log_message(f"Arquivo não encontrado. Criando um novo arquivo com as cotações.")

# Salvando a planilha com as novas informações
try:
    tabela.to_excel(arquivo, index=False)
    log_message(f"Arquivo Excel salvo com sucesso em {arquivo}")
except Exception as e:
    log_message(f"Erro ao salvar o arquivo Excel: {e}")

# Registrando o final da execução
log_message("Execução do script concluída.")

#VINICIUS VIEIRA DA COSTA




