# Certificados Automatizados para Flávia Mega Hair :haircut_woman:

Este é um projeto desenvolvido para automatizar a geração de certificados para a empresa Flávia Mega Hair, permitindo que você crie facilmente certificados personalizados para seus participantes de cursos, workshops, etc.

## Funcionalidades

- Gera certificados personalizados a partir de uma planilha de dados.
- Suporta adição de nome do participante, nome do curso, datas e carga horária.
- Usa modelos de certificados em formato de imagem (jpg, png, etc.) para personalização.

## Como usar

### 1. Configuração inicial

- Clone este repositório:
    git clone https://github.com/ealboliveira/Certificado_Automatizado

- Instale as dependências necessárias:
    pip install openpyxl Pillow
    pip install -r requirements.txt

### 2. Preparação dos dados

- Certifique-se de ter uma planilha no formato correto com os dados dos participantes. Você pode usar o modelo fornecido em 'PlanilhaCertificados.xlsx' como referência.
- Certifique-se de ter um modelo de certificado em formato de imagem (jpg, png, etc.) para personalização. Você pode substituir 'Certificado.jpg' pelo seu próprio modelo.

### 3. Execução do script

- Abra um terminal e navegue até o diretório onde o repositório foi clonado.
- Execute o script Python:
    python app.py


### 4. Resultado

- Os certificados gerados serão salvos como arquivos PNG com o nome do participante.

## Suporte

Se precisar de ajuda ou tiver alguma dúvida, entre em contato com Eduardo Alberto de Oliveira pelo e-mail ealb.oliveira@gmail.com



