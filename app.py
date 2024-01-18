import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Definição de fontes a serem utilizadas
fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

# Abrindo planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

# Pega os dados na planilha
for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)): # Caso queira fazer um teste rápido você pode inserir 'max_row=2' para limitar a impressão a somente um certificado
    # Acessar cada célula
    nome_do_participante = linha[1].value # Nome do participante
    nome_do_curso = linha[0].value # Nome do curso
    tipo_de_participacao = linha[2].value # Tipo de participação
    carga_horaria = linha[5].value # Carga horario do participante

    data_de_inicio = linha[3].value # Data de inicio do curso
    data_de_termino = linha[4].value # Data de termino do curso
    
    data_de_emissao = linha[6].value # Data de emissão do certificado

    # Transferir dados da planilha para a imagem do certificado
    imagem = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(imagem)
    desenhar.text((1020, 827), nome_do_participante, fill='black', font=fonte_nome)
    desenhar.text((1060, 954), nome_do_curso, fill='black', font=fonte_geral)
    desenhar.text((1435, 1065), tipo_de_participacao, fill='black', font=fonte_geral)
    desenhar.text((1490, 1188),str(carga_horaria), fill='black', font=fonte_geral)

    desenhar.text((750, 1770), data_de_inicio, fill='blue', font=fonte_data)
    desenhar.text((750, 1930), data_de_termino, fill='blue', font=fonte_data)

    desenhar.text((2220, 1930), data_de_emissao, fill='blue', font=fonte_data)

    imagem.save(f'./certificados/{indice}{nome_do_participante} certificado.png')