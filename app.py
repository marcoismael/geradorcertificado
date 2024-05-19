import openpyxl
from PIL import Image, ImageDraw, ImageFont
import datetime
from pathlib import Path

#Importando dados do arquivo xlsx
arquivo = openpyxl.load_workbook('./util/planilha_alunos.xlsx')
pagina_alunos = arquivo['Sheet1']

for indice,linha in enumerate(pagina_alunos.iter_rows(min_row=2)):
    
    nome_curso = linha[0].value
    nome_aluno = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    
    fonte_nome = ImageFont.truetype('./fonts/tahomabd.ttf',90)
    fonte_geral = ImageFont.truetype('./fonts/tahoma.ttf',80)
    fonte_data = ImageFont.truetype('./fonts/tahoma.ttf',55)
    data = datetime.datetime.now()

    #Criando Diretorio
    Path(f'./{nome_curso}').mkdir(exist_ok=True)

    #Setando Certificado
    imagem = Image.open('./util/certificado_padrao.jpg')
    preencher = ImageDraw.Draw(imagem)

    #Preenchimento do certificado.
    preencher.text((1020,827), nome_aluno,fill='black',font=fonte_nome)
    preencher.text((1070,950), nome_curso,fill='black',font=fonte_geral)
    preencher.text((1435,1065), tipo_participacao,fill='black',font=fonte_geral)
    preencher.text((1483,1188),str(carga_horaria),fill='black',font=fonte_geral)
    preencher.text((750,1770),data_inicio,fill='blue',font=fonte_data)
    preencher.text((750,1930),data_final,fill='blue',font=fonte_data)
    preencher.text((2238,1930),f'{data.day}-{data.month}-{data.year}',fill='black',font=fonte_data)
    
    #Salvando Certificado
    imagem.save(f'./{nome_curso}/{indice} - {nome_aluno} Certificado.png')

    
