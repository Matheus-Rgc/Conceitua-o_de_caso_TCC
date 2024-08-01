from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.style import WD_STYLE_TYPE

def Titulo_Terminal(a):
    print('=' * 36)
    print(a)
    print('=' * 36)
    return a

def linha():
    print('-' * 30)

def mostrar_exp(mostrar):
    if len(mostrar) == 0:
        return ' '
    else:
        return '\n'.join(mostrar)
    
def mostrar(mostrar):
    for c in mostrar:
        print(c)

def adicionar_paragrafo(doc, texto, estilo=None):
    paragrafo = doc.add_paragraph(texto, style=estilo)
    return paragrafo

Titulo_Terminal('  Conceitualização de Caso pro Word')
doc = Document()

# Criar estilos personalizados
styles = doc.styles

# Estilo de parágrafo
p_style = styles.add_style('Paragraph', WD_STYLE_TYPE.PARAGRAPH)
p_style.font.name = 'Arial'
p_style.font.size = Pt(11)
p_style.font.bold = False

# Estilo do título principal
head_style = styles.add_style('Head', WD_STYLE_TYPE.PARAGRAPH)
head_style.font.name = 'Arial'
head_style.font.size = Pt(22)
head_style.font.color.rgb = RGBColor(0, 0, 0)
head_style.font.bold = True

# Estilo dos subtítulos
subhead_style = styles.add_style('SubHead', WD_STYLE_TYPE.PARAGRAPH)
subhead_style.font.name = 'Arial'
subhead_style.font.size = Pt(14)
subhead_style.font.bold = True
subhead_style.font.color.rgb = RGBColor(0, 0, 255)

# Obter dados do usuário
name = str(input('Nome: '))
idade = str(input('Idade: '))

# Corrigir a entrada do gênero para aceitar apenas 'm' ou 'f'
genero = str(input('Genero: M/F: ')).lower()
while genero not in ['m', 'f']:
    genero = str(input('Genero: M/F: ')).lower()

tel = str(input('Telefone/Celular: '))
linha()
bio = str(input('Fatores biológicos/genéticos do paciente: '))
dev = str(input('Influências do Desenvolvimento: '))
situ = str(input('Questões situacionais: '))
pontos = str(input('Pontos fortes e recursos: '))
sintomas = str(input('Sintomas aparentes: '))

Titulo_Terminal('   Modelo Cognitivo')
a1 = int(input('Quantos modelos cognitivos gostaria de inserir: '))
linha()
lista = []
for c in range(0, a1):
    evento = str(input(f'Evento {c+1}: '))
    p_automatico = str(input('Pensamento Automatico: '))
    emotion = str(input('Emoções: '))
    comportamento = str(input('Comportamentos: '))
    linha()
    # Adicionar dados ao dicionário de lista
    lista.append({
        'Evento': evento,
        'Pensamento Automatico': p_automatico,
        'Emoções': emotion,
        'Comportamentos': comportamento
    })
hipotese = str(input('Hipotese de trabalho: '))
plano = str(input('Planos/Objetivos de Tratamento: '))


# Adicionar os dados ao documento Word
adicionar_paragrafo(doc, 'Conceituação de Caso', estilo='Head')
adicionar_paragrafo(doc, 'Dados básicos do paciente', estilo='SubHead')
adicionar_paragrafo(doc, f'Nome: {name}', estilo='Paragraph')
adicionar_paragrafo(doc, f'Idade: {idade}', estilo='Paragraph')
adicionar_paragrafo(doc, f'Gênero: {genero.upper()}', estilo='Paragraph')
adicionar_paragrafo(doc, f'Telefone: {tel}', estilo='Paragraph')

adicionar_paragrafo(doc, 'Fatores Biológicos/Genéticos:', estilo='SubHead')
adicionar_paragrafo(doc, bio, estilo='Paragraph')

adicionar_paragrafo(doc, 'Influências do Desenvolvimento:', estilo='SubHead')
adicionar_paragrafo(doc, dev, estilo='Paragraph')

adicionar_paragrafo(doc, 'Questões Situacionais:', estilo='SubHead')
adicionar_paragrafo(doc, situ, estilo='Paragraph')

adicionar_paragrafo(doc, 'Pontos fortes/recursos:', estilo='SubHead')
adicionar_paragrafo(doc, pontos, estilo='Paragraph')

adicionar_paragrafo(doc, 'Sinais e Sintomas:', estilo='SubHead')
adicionar_paragrafo(doc, sintomas, estilo='Paragraph')

adicionar_paragrafo(doc, 'Modelos Cognitivos', estilo='Head')
for i, item in enumerate(lista, start=1):
    adicionar_paragrafo(doc, f'Modelo Cognitivo {i}', estilo='SubHead')
    adicionar_paragrafo(doc, f"Evento: {item['Evento']}", estilo='Paragraph')
    adicionar_paragrafo(doc, f"Pensamento Automatico: {item['Pensamento Automatico']}", estilo='Paragraph')
    adicionar_paragrafo(doc, f"Emoções: {item['Emoções']}", estilo='Paragraph')
    adicionar_paragrafo(doc, f"Comportamentos: {item['Comportamentos']}", estilo='Paragraph')
    linha()

adicionar_paragrafo(doc, 'Hipotese de Trabalho', estilo='SubHead')
adicionar_paragrafo(doc, hipotese, estilo='Paragraph')

adicionar_paragrafo(doc, 'Plano de Tratamento', estilo='SubHead')
adicionar_paragrafo(doc, plano, estilo='Paragraph')

# Salvar o documento
doc.save(f'conceituacao_de_caso_{name}.docx')
print(f"Documento salvo como 'conceituacao_de_caso_(nome do paciente).docx'")

