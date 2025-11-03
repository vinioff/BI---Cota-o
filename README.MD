# 📈 Atualizador de Cotações – Dólar, Euro, Bitcoin e Libra

Este projeto realiza a coleta automática de cotações atualizadas para **Dólar (USD)**, **Euro (EUR)**,**Bitcoin (BTC)** e **Libra (GBP)** em relação ao **Real Brasileiro (BRL)**, e salva essas informações em um arquivo Excel, mantendo um histórico de atualizações.

---

## 🚀 Funcionalidades

✅ Consulta das cotações via API  
✅ Atualização da planilha `Cotações.xlsx`  
✅ Execução do script Python via botão no Excel  

---

## 🧰 Requisitos

- ✅ Python 3.11+
- ✅ [Poetry](https://python-poetry.org/) instalado
- ✅ Pacotes:
  - `requests`
  - `pandas`
  - `openpyxl`
  ⚠️ IMPORTANTE: Como a macro VBA do Excel chama o Python nativo do seu computador (e não o ambiente do Poetry), você precisa garantir que os pacotes estejam instalados no Python que o sistema reconhece.

📦 Instalação
    🔁 Clone o repositório
        git clone https://github.com/Romulomrs/CotacaoDolar.git
        cd CotacaoDolar
    🧰 Instale as dependências com o Poetry (opcional, recomendado para testes manuais) Ideal para rodar o script manualmente em um ambiente isolado.
        poetry install
        poetry shell
    ⚠️ Instale os pacotes no Python do sistema
            Essencial para que a macro VBA consiga executar o script corretamente, pois ela utiliza o Python nativo do seu sistema operacional (e não o ambiente virtual do Poetry).
            Abra o terminal (CMD ou PowerShell) e execute:
                pip install requests pandas openpyxl
            ❗ Se estiver usando mais de um Python instalado, use where python e where pip para confirmar qual versão o sistema está usando por padrão.


⚙️ Como Usar

    🔒 Feche o arquivo Cotações.xlsx antes de rodar a macro.

    🔘 Clique no botão "Atualizar a cotação" no arquivo excel nomeado "Macro" para executar a macro.

    🧠 O script Python buscará as cotações e atualizará a planilha.
📎 Exemplo de Macro VBA
💡  💡 Lembre-se de substituir o caminho CAMINHO\DO\SEU\PROJETO pelo local correto onde você salvou o projeto!
        Sub ExecutarScriptPython()
            Dim comando As String
            comando = "cmd /c cd ""CAMINHO\DO\SEU\PROJETO\CotacaoDolar"" && poetry run python ArquivoExecucao\Cotacao.py"
            Shell comando, vbNormalFocus
        End Sub
📌 Exemplo real:
cd "C:\Users\SeuUsuario\Documents\Projetos\CotacaoDolar"

📊 Visualização das Cotações no Power BI
Caso você deseje visualizar os dados de cotação em um relatório do Power BI, é possível:

Conectar a planilha Cotações.xlsx ao Power BI.

Inserir filtros com segmentações de dados (Slicers) para selecionar a moeda desejada.
🖼️ Gráfico
    ![Image](https://github.com/user-attachments/assets/21ef3c36-e34b-4c3c-b072-7ad3f69da42c)
    ![Image](https://github.com/user-attachments/assets/1f790517-3dde-44e8-9713-95a3b1c99212)



