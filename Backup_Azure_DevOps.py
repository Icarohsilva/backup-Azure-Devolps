import requests
import base64
import os
import openpyxl
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


# Criar um nome para pasta do projeto no C:
folder_pasta = "BackupAzureDevolps" #Inclua o o nome da pasta que deseja

# Caminho completo para a pasta do projeto
output_folder = os.path.join("C:/", folder_pasta)

# Verificar se a pasta já existe ou criar a nova pasta para salvar os backup
if not os.path.exists(output_folder):
    os.makedirs(output_folder)


# Criar um nome para incluir os arquivos de backup
current_date = datetime.now().strftime("%Y-%m-%d")
folder_name = f"{current_date}" #Inclua o o nome da pasta casa deseja

# Caminho completo para a pasta dos arquivos de backup daquele dia
output_folder = os.path.join("C:/BackupAzureDevolps", folder_name)

# Verificar se a pasta já existe ou criar a nova pasta para salvar os backup
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Crie um novo arquivo Excel  para salvar relatorio de execucao
report_filename = f"relatorio_logs_{current_date}.xlsx"
report_filepath = os.path.join(output_folder, report_filename)
workbook = openpyxl.Workbook()
sheet = workbook.active

# cabeçalhos ao relatório
sheet["A1"] = "Projeto"
sheet["B1"] = "Repositório"
sheet["C1"] = "Status do Download"
sheet["D1"] = "Arquivo"

# Inicialize uma variável para controlar a linha onde os dados serão inseridos no relatório
row = 2

# Substitua com suas informações de creendciamento do AZURE
organization = "INCLUA AQUI O NOME DA SUA ORGANIZACAO"
access_token = "TOKEN DE ACESSO DA API DO AZURE"

# Construa a URL da API para listar projetos
url = f"https://dev.azure.com/{organization}/_apis/projects?api-version=5.0-preview.1"

# Construa o cabeçalho com o token de acesso pessoal
headers = {
    "Authorization": f"Basic {base64.b64encode(f':{access_token}'.encode()).decode()}",
    "Content-Type": "application/json"
}

# Faça a solicitação GET para listar projetos
response = requests.get(url, headers=headers)
data = response.json()


# Loop através dos projetos
for project in data['value']:
    project_Id = project['id']
    project_name = project['name']
    print(f"Projeto: {project_name} Id: {project_Id}")

    # Construa a URL da API para listar repositórios
    repo_url = f"https://dev.azure.com/{organization}/{project_Id}/_apis/git/repositories?api-version=5.0"
    # Faça a solicitação GET para listar repositórios
    repo_response = requests.get(repo_url, headers=headers)
    repo_data = repo_response.json()
    # Loop através dos repositórios
    for repo in repo_data['value']:
        repo_name = repo['name']
        print(f"  Repositório: {repo_name} Id: {repo['id']}")

        # Construa a URL da API para listar e baixar arquivos do branch "master"
        branch = "master"  # Use "master" se o nome do branch for "main"
        file_url = f"https://dev.azure.com/{organization}/{project_Id}/_apis/git/repositories/{repo['id']}/items?path=/&versionDescriptor[versionOptions]=0&versionDescriptor[versionType]=0&versionDescriptor[version]={branch}&resolveLfs=true&%24format=zip&api-version=5.0&download=true"
        file_response = requests.get(file_url, headers=headers)


        # Verifique se a solicitação foi bem-sucedida
        if file_response.status_code == 200:
            # Salve o conteúdo do arquivo ZIP em um arquivo local
            zip_filename = f"{project_name}_{repo_name}_{branch}.zip"
            zip_filepath = os.path.join(output_folder, zip_filename)
            with open(zip_filepath, "wb") as zip_file:
                zip_file.write(file_response.content)
            print(f"    Arquivo ZIP baixado e salvo em: {zip_filepath}")
            sheet[f"A{row}"] = project_name
            sheet[f"B{row}"] = repo_name
            sheet[f"C{row}"] = "Baixado e Salvo" if file_response.status_code == 200 else f"Falha ({file_response.status_code})"
            sheet[f"D{row}"] = branch


        # Caso o master apresente erro, baixe outro branch
        elif file_response.status_code != 200:
            branch = "main"  # Use "main" se o nome do branch for "main"
            file_url = f"https://dev.azure.com/{organization}/{project_Id}/_apis/git/repositories/{repo['id']}/items?path=/&versionDescriptor[versionOptions]=0&versionDescriptor[versionType]=0&versionDescriptor[version]={branch}&resolveLfs=true&%24format=zip&api-version=5.0&download=true"
            file_response = requests.get(file_url, headers=headers)
            if file_response.status_code == 200:
                # Salve o conteúdo do arquivo ZIP em um arquivo local
                zip_filename = f"{project_name}_{repo_name}_{branch}.zip"
                zip_filepath = os.path.join(output_folder, zip_filename)
                with open(zip_filepath, "wb") as zip_file:
                    zip_file.write(file_response.content)
                print(f"    Arquivo ZIP baixado e salvo em: {zip_filepath}")
                sheet[f"A{row}"] = project_name
                sheet[f"B{row}"] = repo_name
                sheet[f"C{row}"] = "Baixado e Salvo" if file_response.status_code == 200 else f"Falha ({file_response.status_code})"
                sheet[f"D{row}"] = branch

        else:
            print(f"    Falha ao baixar o arquivo ZIP ({file_response.status_code})")
            sheet[f"A{row}"] = project_name
            sheet[f"B{row}"] = repo_name
            sheet[f"C{row}"] = f"Falha ({file_response.status_code})"
            sheet[f"D{row}"] = branch

    # Agora você pode salvar ou processar o conteúdo do arquivo conforme necessário
    print("=" * 40)
    row += 1

# Salve o arquivo Excel
workbook.save(report_filepath)
print(f"Relatório gerado e salvo em: {report_filepath}")

# Lista de endereços de e-mail dos destinatários
# Inclua os emails seprados por virgulas
to_emails = []


# Configurações do servidor SMTP
host = ""
port = 0 #Informar a porta de acesso
username = ""
password = ""
enable_ssl = True
subject = "Relatório de Logs"

# Crie o corpo do e-mail
msg = MIMEMultipart()
msg["From"] = username
msg["To"] = ", ".join(to_emails)  # Combine os endereços de e-mail separados por vírgula
msg["Subject"] = subject

# Adicione um texto ao corpo do e-mail
body = f"Segue em anexo o relatório de logs gerado dos dowloads de arquivos Master e Main dos Projetos do Azure baixado na data: {current_date}."
msg.attach(MIMEText(body, "plain"))

# Anexe o arquivo de relatório
with open(report_filepath, "rb") as attachment:
    part = MIMEApplication(attachment.read(), Name=os.path.basename(report_filepath))
    part["Content-Disposition"] = f'attachment; filename="{os.path.basename(report_filepath)}"'
    msg.attach(part)

try:
    # Inicialize o servidor SMTP
    server = smtplib.SMTP(host, port)
    if enable_ssl:
        server.starttls()

    # Faça login no servidor SMTP
    server.login(username, password)

    # Envie o e-mail
    server.sendmail(username, to_emails, msg.as_string())

    # Encerre a conexão com o servidor SMTP
    server.quit()

    print("E-mail enviado com sucesso!")

    # Salvar um arquivo de log
    log_content = f"E-mail enviado com sucesso para: {', '.join(to_emails)}\n"
    log_filename = os.path.join(output_folder, "email_log.txt")
    with open(log_filename, "w") as log_file:
        log_file.write(log_content)

except Exception as e:
    print(f"Erro ao enviar o e-mail: {e}")
    # Salvar um arquivo de log de erro
    log_content = f"Erro ao enviar o e-mail: {e}\n"
    log_filename = os.path.join(output_folder, "email_error_log.txt")
    with open(log_filename, "w") as log_file:
        log_file.write(log_content)

print("FIM")