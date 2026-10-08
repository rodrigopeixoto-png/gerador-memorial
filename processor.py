import ifcopenshell
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO

def adicionar_tabela_formatada(doc, cabecalhos, dados):
    """Função auxiliar para criar tabelas com estilo profissional"""
    tabela = doc.add_table(rows=1, cols=len(cabecalhos))
    tabela.style = 'Table Grid'
    
    # Formata cabeçalho
    hdr_cells = tabela.rows[0].cells
    for i, nome in enumerate(cabecalhos):
        hdr_cells[i].text = nome
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(10)
    
    # Preenche dados
    for linha in dados:
        row_cells = tabela.add_row().cells
        for i, valor in enumerate(linha):
            row_cells[i].text = str(valor)
            for paragraph in row_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
    return tabela

def gerar_memorial_docx(ifc_file_path):
    # 1. LEITURA BÁSICA DO IFC (Simulação de extração)
    # modelo = ifcopenshell.open(ifc_file_path)
    # Em um cenário real, você extrairia a lista de pavimentos dinamicamente aqui.
    
    doc = Document()
    
    # --- CAPA ---
    p_capa = doc.add_paragraph()
    p_capa.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_capa.add_run('PARANÁ PROJETOS\nEngenharia & Consultoria\n\n')
    run.bold = True
    run.font.size = Pt(16)
    
    run2 = p_capa.add_run('MEMORIAL DESCRITIVO E DE CÁLCULO\n\n')
    run2.bold = True
    run2.font.size = Pt(20)
    
    run3 = p_capa.add_run('PROJETO: IMPLANTAÇÃO DE SANTA FELICIDADE\nORDEM DE SERVIÇO N°. 024/2024 - LOTE 11\nESTADO DO PARANÁ\n')
    run3.font.size = Pt(14)
    doc.add_page_break()

    # --- HISTÓRICO DO DOCUMENTO ---
    doc.add_heading('Histórico do Documento', level=1)
    cabecalhos_hist = ['Revisão', 'Descrição', 'Editado', 'Verificado', 'Autorizado', 'Data']
    dados_hist = [
        ['00', 'Emissão Inicial', 'Eng. Responsável', 'Revisor', 'Diretoria', '19-11-2025']
    ]
    adicionar_tabela_formatada(doc, cabecalhos_hist, dados_hist)
    doc.add_page_break()

    # --- APRESENTAÇÃO E DESCRIÇÃO ---
    doc.add_heading('1. APRESENTAÇÃO', level=1)
    doc.add_paragraph('Trata-se do memorial descritivo e de cálculo da estrutura adotada para a presente implantação.')
    
    doc.add_heading('2. DESCRIÇÃO DO EDIFÍCIO', level=1)
    doc.add_paragraph('O edifício é constituído pelos seguintes pavimentos estruturais extraídos do modelo BIM:')
    
    cabecalhos_pav = ['Pavimentos', 'Piso a Piso (m)', 'Cota (m)', 'Área (m²)']
    dados_pav = [
        ['Topo Reservatório', '2,80', '10,35', '40,39'],
        ['Fundo Reservatório', '0,70', '7,55', '11,74'],
        ['Cobertura', '3,40', '6,85', '713,58'],
        ['Térreo', '0,40', '-0,13', '48,98'],
        ['Fundação', '0,00', '-0,53', '0,24']
    ]
    adicionar_tabela_formatada(doc, cabecalhos_pav, dados_pav)

    # --- NORMAS E SOFTWARES ---
    doc.add_heading('3. NORMA EM USO', level=1)
    doc.add_paragraph('Na análise, dimensionamento e detalhamento dos elementos estruturais deste edifício foram utilizadas as prescrições indicadas pelas seguintes normas:')
    normas = [
        'ABNT NBR 6118:2023 - Projeto de estruturas de concreto - Procedimento;',
        'ABNT NBR 6120:2019 - Cargas para o cálculo de estruturas de edificações;',
        'ABNT NBR 6123:2023 - Forças devido ao vento em edificações;',
        'ABNT NBR 8681:2003 - Ações e segurança nas estruturas - Procedimento;',
        'ABNT NBR 15200:2024 - Projeto de estruturas de concreto em situação de incêndio.'
    ]
    for norma in normas:
        doc.add_paragraph(norma, style='List Bullet')

    doc.add_heading('4. SOFTWARE UTILIZADO', level=1)
    doc.add_paragraph('Para a análise estrutural, dimensionamento e detalhamento estrutural foi utilizado o sistema TQS na versão V24.5.12 e modelagem parametrizada via rotinas automatizadas.')

    # --- MATERIAIS ---
    doc.add_heading('5. MATERIAIS', level=1)
    doc.add_heading('Concreto', level=2)
    doc.add_paragraph('A seguir são apresentados os valores de fck utilizados para cada um dos elementos estruturais:')
    
    cabecalhos_mat = ['Pavimento', 'Lajes (MPa)', 'Vigas (MPa)', 'Fundações (MPa)']
    dados_mat = [
        ['Cobertura', '30', '30', '30'],
        ['Térreo', '30', '30', '30'],
        ['Fundação', '30', '30', '30']
    ]
    adicionar_tabela_formatada(doc, cabecalhos_mat, dados_mat)
    
    doc.add_heading('Aço de armadura passiva', level=2)
    cabecalhos_aco = ['Tipo de barra', 'Es (MPa)', 'fyk (MPa)', 'Massa específica (kgf/m³)']
    dados_aco = [
        ['CA-50', '210000', '500', '7850'],
        ['CA-60', '210000', '600', '7850']
    ]
    adicionar_tabela_formatada(doc, cabecalhos_aco, dados_aco)

    doc.add_heading('Parâmetros de Durabilidade', level=2)
    doc.add_paragraph('Classe de agressividade: II - Moderada.')
    cabecalhos_cob = ['Elemento Estrutural', 'Cobrimento (cm)']
    dados_cob = [
        ['Lajes convencionais', '2.0'],
        ['Vigas', '2.5'],
        ['Pilares', '2.5'],
        ['Fundações', '2.5']
    ]
    adicionar_tabela_formatada(doc, cabecalhos_cob, dados_cob)

    # --- SALVAR EM MEMÓRIA ---
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer
