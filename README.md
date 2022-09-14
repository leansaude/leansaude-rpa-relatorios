# RPA de obtenção de relatórios pendentes do Amplimed, produto Lean Stay
Identifica relatórios de visitas pendentes, obtém os PDFs do Amplimed, move os arquivos para o Google Drive e atualiza a planilha de gerenciamento no Google Sheet.

## Pré-requisitos
1. Possuir o python instalado localmente. Testado com python versão 3.10.5 em ambiente Windows 11.
2. Possuir o pip instalado localmente. (Ferramenta de gerenciamento de dependências python.)

## Observações
1. Ao clonar o repositório, não esqueça de criar um arquivo .env baseado no exemplo contido em .env.exemplo. 
Ajuste as configurações como desejar.

2. Também é necessário criar uma pasta "credentials" na raiz do repositório, contendo um arquivo .json referenciado na variável GOOGLE_APPLICATION_CREDENTIALS (em .env). Este .json são as credenciais para autenticação no Google Cloud. Obtenha o arquivo com o gerente do projeto.

3. Crie também uma pasta "pdfs" na raiz do repositório.

4. Este repositório faz uso de virtual envs do python. Para usá-lo, você precisa ter a seguinte dependência instalada globalmente em seu computador:
```
pip install virtualenv
```
Depois, inicie o virtual env na pasta do repo local:
```
cd <caminho-do-seu-repo>
virtualenv env
```

5. As dependências específicas deste projeto estão descritas em requirements.txt. 
Instale-as rodando:
```
cd <caminho-do-seu-repo>
env\Scripts\pip install -r requirements.txt
```

6. Execute o script usando o python contido no virtual env, e não o python global.
```
cd <caminho-do-seu-repo>
env\Scripts\python relatorios.py
```