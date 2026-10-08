import os
import re
import ifcopenshell
import ifcopenshell.util.element as util
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- FUNÇÃO DE EXTRAÇÃO AVANÇADA (EBERICK + GENÉRICO) ---

def extrair_dados_eberick(elemento):
    """Lê o padrão específico do AltoQi Eberick e extrai Volume, Material e Armadura."""
    volume = 0.0
    material = "Concreto Especificado"
    armadura = "-"
    
    try:
        psets = util.get_psets(elemento)
        for pset_nome, propriedades in psets.items():
            if isinstance(propriedades, dict):
                for prop_nome, valor in propriedades.items():
                    nome_prop = str(prop_nome).lower()
                    valor_str = str(valor)
                    
                    # 1. BUSCA DE MATERIAL
                    # Padrão 1: Nome da propriedade contém o material (ex: Pilares/Vigas)
                    if 'concreto' in nome_prop and 'abatimento' in nome_prop:
                        partes = str(prop_nome).split(' - ')
                        if len(partes) >= 2:
                            material = partes[1].strip()
                        # Extrai o volume desta mesma linha
                        if isinstance(valor, (int, float)):
                            volume = float(valor)
                        elif isinstance(valor, str):
                            nums = re.findall(r"[-+]?\d*\.\d+|\d+", valor.replace(',', '.'))
                            if nums: volume = float(nums[0])
                            
                    # Padrão 2: Lajes e Fundações ("Classe de concreto")
                    elif 'classe de concreto' in nome_prop:
                        material = valor_str

                    # 2. BUSCA DE VOLUME (Se não achou no Padrão 1)
                    elif 'volume' in nome_prop and volume == 0.0:
                        if isinstance(valor, (int, float)):
                            volume = float(valor)
                        elif isinstance(valor, str):
                            nums = re.findall(r"[-+]?\d*\.\d+|\d+", valor.replace(',', '.'))
                            if nums: volume = float(nums[0])

                    # 3. BUSCA DE ARMADURA / TAXA DE ARMADURA
                    if 'taxa de armadura' in nome_prop:
                        armadura = valor_str
                        if '%' not in armadura: 
                            armadura += ' %'
                    elif 'armadura' in nome_prop and 'aço' in nome_prop and armadura == "-":
                        # Lê valores em kg (ex: "Armadura - Aço CA50...")
                        nums = re.findall(r"[-+]?\d*\.\d+|\d+", valor_str.replace(',', '.'))
                        if nums:
                            armadura = f"{float(nums[0]):.2f} kg"

        # Fallback de Material (Padrão IFC)
        if material == "Concreto Especificado" and hasattr(elemento, 'HasAssociations') and elemento.HasAssociations:
            for rel in elemento.HasAssociations:
                if rel.is_a('IfcRelAssociatesMaterial'):
                    mat = rel.RelatingMaterial
                    if mat.is_a('IfcMaterial'): material = mat.Name
                    elif mat.is_a('IfcMaterialList') and len(mat.Materials) > 0: material = mat.Materials[0].Name

    except Exception:
        pass

    return volume, material, armadura

def extrair_dados_elementos(elementos):
    """Gera a lista formatada para a tabela do Word."""
    dados = []
    volume_total = 0.0
    
    for el in elementos:
        nome = el.Name if el.Name else "N/A"
        volume, material, armadura = extrair_dados_eberick(el)
        volume_total += volume
        
        vol_str = f"{volume:.2f}" if volume > 0 else "0.00"
        dados.append([nome, material, armadura, vol_str])
        
    # Ordena alfabeticamente (Ex: L205, L206, P1, P2...)
    dados.sort(key=lambda x: x[0])
    return dados, volume_total

# --- FUNÇÕES DE FORMATAÇÃO DO WORD ---

def adicionar_tabela_formatada(doc, cabecalhos, dados):
    if not dados:
        doc.add_paragraph("Nenhum elemento encontrado neste pavimento.", style='Italic')
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
    modelo = ifcopenshell.open(ifc_file_path)
    doc = Document()
    
    # --- CAPA ---
    p_capa = doc.add_paragraph()
    p_capa.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    caminho_logo = 'fundo_transparente.png'
    if os.path.exists(caminho_logo):
        p_capa.add_run().add_picture(caminho_logo, width=Cm(6.0))
        p_capa.add_run('\n\n')
        
    run = p_capa.add_run('SECRETARIA DA SEGURANÇA PÚBLICA\nUnidade Técnica de Engenharia e Arquitetura - UTEA\n\n')
    run.bold = True
    run.font.size = Pt(14)
    
    run2 = p_capa.add_run('MEMORIAL DESCRITIVO DE ESTRUTURAS\n\n')
    run2.bold = True
    run2.font.size = Pt(20)
    
    projetos = modelo.by_type("IfcProject")
    nome_projeto = projetos[0].Name if projetos else "Projeto Estrutural"
    p_capa.add_run(f'Modelo BIM: {nome_projeto}\nGerado Automaticamente via IFC')
    doc.add_page_break()

    # --- AGRUPAMENTO POR PAVIMENTO ---
    pavimentos = {}
    classes_estruturais = ["IfcColumn", "IfcBeam", "IfcSlab", "IfcFooting"]
    
    for classe in classes_estruturais:
        for el in modelo.by_type(classe):
            container = util.get_container(el)
            nome_pav = container.Name if container else "Pavimento Indefinido"
            
            if nome_pav not in pavimentos:
                pavimentos[nome_pav] = {"IfcColumn": [], "IfcBeam": [], "IfcSlab": [], "IfcFooting": []}
            pavimentos[nome_pav][classe].append(el)

    # --- GERAÇÃO DAS SECÇÕES NO DOCUMENTO ---
    cabecalhos = ['Identificação', 'Material', 'Armadura (Taxa/Peso)', 'Volume (m³)']
    titulos_classes = {"IfcColumn": "PILARES", "IfcBeam": "VIGAS", "IfcSlab": "LAJES", "IfcFooting": "FUNDAÇÕES"}

    vol_global = 0.0
    contador_capitulo = 4

    for classe in classes_estruturais:
        doc.add_heading(f'{contador_capitulo}. {titulos_classes[classe]}', level=1)
        vol_total_classe = 0.0
        
        for nome_pav, elementos_do_pavimento in pavimentos.items():
            elementos = elementos_do_pavimento[classe]
            if elementos:
                doc.add_heading(f'Pavimento: {nome_pav}', level=2)
                dados, vol_parcial = extrair_dados_elementos(elementos)
                adicionar_tabela_formatada(doc, cabecalhos, dados)
                vol_total_classe += vol_parcial
                doc.add_paragraph()

        doc.add_paragraph(f'Volume total estimado em {titulos_classes[classe]}: {vol_total_classe:.2f} m³').bold = True
        vol_global += vol_total_classe
        contador_capitulo += 1
        doc.add_page_break()

    # --- Resumo Global ---
    doc.add_heading(f'{contador_capitulo}. RESUMO DE MATERIAIS', level=1)
    doc.add_paragraph(f'Volume global de betão/concreto da estrutura: {vol_global:.2f} m³').bold = True

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer
