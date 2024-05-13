import openpyxl
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    nome_curso = linha[0].value #nome do curso
    nome_participante = linha[1].value 
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value

    fonte_nome = ImageFont.truetype('./LiberationSansNarrow-Bold.ttf', 90)
    fonte_geral = ImageFont.truetype('./LiberationSansNarrow-Regular.ttf', 80)

    image = Image.open('./certificado_padrao.jpg')
    sobrepor_imagem = ImageDraw.Draw(image)

    sobrepor_imagem.text((1020,833),nome_participante,fill='black',font=fonte_nome)
    sobrepor_imagem.text((1075,957),nome_curso,fill='black',font=fonte_geral)
    sobrepor_imagem.text((1440,1073),tipo_participacao,fill='black',font=fonte_geral)
    sobrepor_imagem.text((1500,1190),str(carga_horaria),fill='black',font=fonte_geral)
    sobrepor_imagem.text((735,1770),str(data_inicio),fill='black',font=fonte_geral)
    sobrepor_imagem.text((735,1920),str(data_final),fill='black',font=fonte_geral)
    sobrepor_imagem.text((2210,1920),str(data_emissao),fill='black',font=fonte_geral)

    image.save(f'./{indice}{nome_participante} certificado.png')