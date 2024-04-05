import openpyxl
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime

# Carregar a planilha Excel
workbook_alunos = openpyxl.load_workbook('PlanilhaCertificados.xlsx')
sheet_alunos = workbook_alunos['Planilha1']

# Iterar sobre as linhas da planilha
for linha in sheet_alunos.iter_rows(min_row=2):
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    data_inicio = linha[2].value.strftime('%d/%m/%Y') if linha[2].value else None
    data_termino = linha[3].value.strftime('%d/%m/%Y') if linha[3].value else None
    carga_horario = linha[4].value

    if nome_participante:
        # Definir as fontes para o texto no certificado
        fonte_geral = ImageFont.truetype('tahoma.ttf', 60)
        fonte_nome = ImageFont.truetype('tahomabd.ttf', 70)
        fonte_data =  ImageFont.truetype('tahoma.ttf', 40)

        # Abrir a imagem do modelo de certificado
        imagem = Image.open('Certificado.jpg')
        desenhar = ImageDraw.Draw(imagem)

        # Adicionar o nome do participante no certificado
        desenhar.text((712,558), nome_participante, fill='black', font=fonte_nome)
        # Adicionar o nome do curso
        if nome_curso:
            desenhar.text((712,658), nome_curso, fill='black', font=fonte_geral)
        # Adicionar a carga horária
        if carga_horario:
            desenhar.text((962,760), str(carga_horario), fill='black', font=fonte_geral)
        # Adicionar a data de início
        if data_inicio:
            desenhar.text((580,1120), data_inicio, fill='black', font=fonte_data)
        # Adicionar a data de término
        if data_termino:
            desenhar.text((580,1240), data_termino, fill='black', font=fonte_data)
        
        # Salvar o certificado com o nome do participante
        imagem.save(f'{nome_participante}_certificado.png')

print("Processamento concluído.")
