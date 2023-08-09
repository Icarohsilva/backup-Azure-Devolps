# Backup de Repositórios (Master e Main) do Azure DevOps

Este script Python permite fazer o backup dos repositórios do Azure DevOps, baixando os arquivos dos branches "master" e "main". Ele também gera um relatório em Excel sobre o status dos downloads e envia o relatório por e-mail.

## Pré-requisitos

Antes de executar este script, verifique se você possui as seguintes dependências instaladas:

    - Python (versão X.Y.Z)
    - Bibliotecas Python: requests, openpyxl
    - Conta no Azure DevOps com token de acesso da API

## Configuração

1. Clone este repositório para sua máquina local:

   ```bash
   git clone https://github.com/Icarohsilva/backup-Azure-Devolps.git
   cd backup-Azure-Devolps

2. Crie um ambiente virtual (recomendado) e ative-o:
   
       python -m venv venv
       source venv/bin/activate  # No Windows: venv\Scripts\activate
   
4. Instale as dependências do projeto: 

       pip install -r requirements.txt

5. Abra o arquivo  BackupAzure.py e edite as seguintes variáveis: 

        organization: Substitua pelo nome da sua organização no Azure DevOps.
        access_token: Substitua pelo token de acesso da API do Azure DevOps. 
        to_emails: Lista de endereços de e-mail dos destinatários do relatório. 
        Configurações do servidor SMTP:  host,  port,  username,  password,  enable_ssl. 

## Execução 

Execute o script com o seguinte comando: 

    python BackupAzure.py

## Resultados 

     Após a execução, o script fará o backup dos repositórios "master" e "main" do Azure DevOps.
     Salvará os arquivos ZIP na pasta  C:/BackupAzureDevolps/NOME_DA_PASTA.
     E gerará um relatório em Excel. O relatório será enviado por e-mail aos destinatários especificados. 

Autor:  Icaro Henrique Nunes Viana Silva
