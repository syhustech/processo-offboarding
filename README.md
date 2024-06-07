# Projeto Offboarding Mensal e Anual #

- Projeto desenvolvido utilizando BotCity como framework para facilitar o desenvolvimento do código fonte
Link: https://documentation.botcity.dev/frameworks/desktop/python/

- Existe um projeto criado no ambiente da Google Cloud Platform que armazena as credenciais para acesso ao Drive e ao Google Sheets para todos os scripts de automação de processos.
Link: https://console.cloud.google.com/welcome

- A arquitetura do script está em forma de monolito (único arquivo = bot.py)
- O arquivo google_credentials.py é usado para extrair parte da lógica de autenticação usada para drive e sheets
- A pasta resources armazena as imagens usadas durante a execução do script para encontrar os pontos clicáveis na tela pelo bot

- Os arquivos credentials.json, token_drive_v3.pickle e service_account.json são utilizados para manipular as planilhas no sheets e armazenar os arquivos no drive
    - service_account.json: credenciais da conta de serviço para acessar a planilha de controle de execução (Se for alterado no Google Cloud Platform será necessário gerar novamente)
    - credentials.json: Usado para dar acesso ao drive para a conta de serviço
    - token_drive_v3.pickle: arquivo gerado automaticamente para armazenar as credenciais de login após a primeira execução