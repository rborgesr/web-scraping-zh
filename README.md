# Web Scraping Automático de Dados Acadêmicos com Selenium
Este projeto foi aplicado e desenvolvido para a extração de dados com o objetivo de migração de sistema durante meu estágio. Utilizando Selenium para a navegação e interação com o site, e OpenPyXL para salvar os dados coletados em uma planilha Excel (.xlsx), o processo de coleta de informações financeiras e pessoais de alunos foi automatizado para facilitar a transição entre sistemas.

## Funcionalidades
Automatiza o login em um sistema de gestão de alunos.

Navega pelas páginas, coleta dados financeiros e pessoais de alunos.

Armazena os dados extraídos em uma planilha Excel.

Exibe progresso durante a execução, com tempo estimado para finalização.

## Pré-requisitos

Antes de executar o script, certifique-se de que os seguintes pacotes estão instalados:

- Python 3.x
- Selenium
- WebDriver Manager
- OpenPyXL
- Google Chrome (ou outro navegador compatível com Selenium)
- Chromedriver (instalado automaticamente pelo webdriver-manager)
  
#### Você pode instalar as dependências necessárias com o seguinte comando:


- pip install selenium webdriver-manager openpyxl


#### Durante a execução, o script solicitará as seguintes informações:

- Nome do arquivo Excel de saída: O nome do arquivo onde os dados serão salvos.
- Caminho da planilha existente: O caminho de uma planilha Excel contendo a lista de nomes dos alunos.
- Nome da aba: O nome da aba dentro da planilha de entrada.
- O script então realizará o login no sistema e começará a processar os nomes, coletando os dados financeiros e pessoais, e salvando tudo em uma nova planilha Excel.

# Estrutura do Código

## Bibliotecas Utilizadas:

- selenium: Para automação de navegação e interação com o site.
- webdriver-manager: Para gerenciar a instalação do ChromeDriver automaticamente.
- openpyxl: Para manipulação e geração de planilhas Excel.
- time e math: Para controle de tempo e cálculos de estimativa de execução.

## Fluxo do Código:

- Login Automático: O script acessa o site de login, preenche as credenciais e faz o login.
- Navegação e Coleta de Dados: Para cada aluno na lista, o script pesquisa no sistema, acessa os dados financeiros e coleta informações como nome, CPF, endereço, e-mail, etc.
- Armazenamento de Dados: Os dados coletados são salvos em um arquivo .xlsx nomeado pelo usuário.
### Exemplo de Saída:
#### A saída gerada será uma planilha Excel com as seguintes colunas:

|Nome|	CPF	|E-mail	|RG	|Telefone	|Data de Nascimento	|Sexo	|Logradouro	|Complemento	|CEP	|Bairro	|Cidade	|UF |

## Observações

- Certifique-se de que o sistema que está sendo acessado permite a extração de dados de forma automatizada, e que você está seguindo todas as normas de uso do site.
- É importante respeitar a privacidade dos dados coletados e utilizá-los de forma ética e legal.
