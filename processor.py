import ifcopenshell
import ifcopenshell.util.element as util
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- FUNÇÕES DE EXTRAÇÃO GENÉRICA DO IFC ---

def extrair_volume(elemento):
    """Busca qualquer propriedade que contenha 'Volume' de forma genérica."""
    psets = util.get_psets(elemento)
    for pset_nome, propriedades in psets.items():
        if isinstance(propriedades, dict):
            for prop_nome, valor in propriedades.items():
                if 'volume' in prop_nome.lower() and isinstance(valor, (int, float)):
                    return valor
    return 0.0

def extrair_material(elemento):
    """Busca a associação de material padrão do elemento."""
    if elemento.HasAssociations:
        for rel in elemento.HasAssociations:
            if rel.is_a('IfcRelAssociatesMaterial'):
                material = rel.RelatingMaterial
                if material.is_a('IfcMaterial'):
                    return material.Name
                elif material.is_a('IfcMaterialList') and len(material.Materials) > 0:
                    return material.Materials[0].Name
                elif material.is_a('IfcMaterialProfileSet'):
                    return material.MaterialProfiles[0].Material.Name
    return "Concreto Específico do Projeto"

def extrair_dados_elementos(modelo, ifc_class):
    """Gera uma lista de dados para a tabela do Word."""
    elementos = modelo.by_type(ifc_class)
    dados = []
    volume_total = 0.0
    
    for el in elementos:
        nome = el.Name if el.Name else "N/A"
        material = extrair_material(el)
        volume = extrair_volume(el)
        volume_total += volume
        
        # Formata o volume para 2 casas decimais, ou traço se for 0
        vol_str = f"{volume:.2f}" if volume > 0 else "-"
        
        dados.append([nome, material, vol_str])
        
    return dados, volume_total

# --- FUNÇÕES DE FORMATAÇÃO DO WORD ---

def adicionar_tabela_formatada(doc, cabecalhos, dados):
    """Cria uma tabela com bordas e cabeçalhos em negrito."""
    if not dados:
        doc.add_paragraph("Nenhum elemento encontrado nesta categoria.", style='Italic')
        return
        
    tabela = doc.add_table(rows=1, cols=len(cabecalhos))
    tabela.style = 'Table Grid'
    
    hdr_cells = tabela.rows[0].cells
    for i, nome in enumerate(cabecalhos):
        hdr_cells[i].text = nome
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(10)
    
    for linha in dados:
        row_cells = tabela.add_row().cells
        for i, valor in enumerate(linha):
            row_cells[i].text = str(valor)
            for paragraph in row_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)

# --- MOTOR PRINCIPAL ---

def gerar_memorial_docx(ifc_file_path):
    # 1. Carrega o modelo IFC
    modelo = ifcopenshell.open(ifc_file_path)
    
    # 2. Inicia o Documento
    doc = Document()
    
    # --- CAPA ---
    p_capa = doc.add_paragraph()
    p_capa.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_capa.add_run('MEMORIAL DESCRITIVO DE ESTRUTURAS\n\n')
    run.bold = True
    run.font.size = Pt(20)
    
    nome_projeto = modelo.by_type("IfcProject")[0].Name if modelo.by_type("IfcProject") else "Projeto Estrutural"
    p_capa.add_run(f'Modelo BIM: {nome_projeto}\nGerado Automaticamente via IFC')
    doc.add_page_break()

    # --- TEXTOS FIXOS BÁSICOS ---
    doc.add_heading('1. APRESENTAÇÃO', level=1)
    doc.add_paragraph('Este documento apresenta o memorial descritivo dos elementos estruturais de concreto armado, contendo as identificações, materiais especificados e quantitativos geométricos extraídos diretamente do modelo BIM (arquivo IFC) gerado pelo software de cálculo estrutural.')

    doc.add_heading('2. NORMAS UTILIZADAS', level=1)
    normas = [
        'ABNT NBR 6118 - Projeto de estruturas de concreto - Procedimento;',
        'ABNT NBR 6120 - Cargas para o cálculo de estruturas de edificações;'
    ]
    for norma in normas:
        doc.add_paragraph(norma, style='List Bullet')

    # --- EXTRAÇÃO DINÂMICA DE DADOS ---
    cabecalhos = ['Identificação', 'Material (Classe)', 'Volume Líquido (m³)']
    
    # Pilares
    doc.add_heading('3. PILARES', level=1)
    dados_pilares, vol_pilares = extrair_dados_elementos(modelo, "IfcColumn")
    adicionar_tabela_formatada(doc, cabecalhos, dados_pilares)
    doc.add_paragraph(f'\nVolume total de concreto em Pilares: {vol_pilares:.2f} m³').bold = True

    # Vigas
    doc.add_heading('4. VIGAS', level=1)
    dados_vigas, vol_vigas = extrair_dados_elementos(modelo, "IfcBeam")
    adicionar_tabela_formatada(doc, cabecalhos, dados_vigas)
    doc.add_paragraph(f'\nVolume total de concreto em Vigas: {vol_vigas:.2f} m³').bold = True

    # Lajes
    doc.add_heading('5. LAJES', level=1)
    dados_lajes, vol_lajes = extrair_dados_elementos(modelo, "IfcSlab")
    adicionar_tabela_formatada(doc, cabecalhos, dados_lajes)
    doc.add_paragraph(f'\nVolume total de concreto em Lajes: {vol_lajes:.2f} m³').bold = True

    # Fundações
    doc.add_heading('6. FUNDAÇÕES', level=1)
    dados_fund, vol_fund = extrair_dados_elementos(modelo, "IfcFooting")
    adicionar_tabela_formatada(doc, cabecalhos, dados_fund)
    doc.add_paragraph(f'\nVolume total de concreto em Fundações: {vol_fund:.2f} m³').bold = True
    
    # Resumo Global
    doc.add_heading('7. RESUMO DE MATERIAIS', level=1)
    vol_global = vol_pilares + vol_vigas + vol_lajes + vol_fund
    doc.add_paragraph(f'Volume global de concreto da estrutura: {vol_global:.2f} m³').bold = True

    # 3. Salva em buffer para o Streamlit
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer
