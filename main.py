#-----------------------------------------------------#

# Importar as bibliotecas necessárias

import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Carregar a planilha de alunos
planilha_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
# Acessar a primeira pagina da planilha
sheet_alunos = planilha_alunos['Sheet1']

# Para cada linha sobre as linhas da planilha a partir da linha 2
for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # Cada célula que contém a info que precisamos
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_de_participacao = linha[2].value
    carga_horaria = linha[5].value

    data_inicio = linha[3].value
    data_fim = linha[4].value

    data_emissao = linha[6].value

    # Transferir os dados para o certificado:

    # Definindo as fontes 
    font_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    font_geral = ImageFont.truetype('./tahoma.ttf', 75)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

    # Abrir a imagem do certificado e salvar em uma variável
    image = Image.open('certificado_padrao.jpg')
    # Definir uma variavel para desenhar na imagem quando chamada
    desenhar = ImageDraw.Draw(image)

    # Definir as coordenadas, oque será escrito, a cor e a fonte para cada campo
    desenhar.text((1020, 825), nome_participante, fill='black', font=font_nome)
    desenhar.text((1075, 955), nome_curso, fill='black', font=font_geral)
    desenhar.text((1440, 1070), tipo_de_participacao, fill='black', font=font_geral)
    desenhar.text((1485, 1192), f'{carga_horaria} horas', fill='black', font=font_geral)

    desenhar.text((750, 1780), f'{data_inicio}', fill='black', font=fonte_data)
    desenhar.text((750, 1930), f'{data_fim}', fill='black', font=fonte_data)

    desenhar.text((2220, 1930), f'{data_emissao}', fill='black', font=fonte_data)

    # Salvar a imagem com o nome do participante
    image.save(f'./certificados_feitos/{indice + 1}º Certificado - {nome_participante}.png')