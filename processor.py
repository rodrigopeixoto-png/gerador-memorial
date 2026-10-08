import os
import re
import ifcopenshell
import ifcopenshell.util.element as util
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- FUNÇÕES DE EXTRAÇÃO AVANÇADA ---

def extrair_volume(elemento):
    """Busca o volume nas BaseQuantities ou Psets de forma exaustiva."""
    try:
        # Busca nas propriedades e quantidades (Qto_BaseQuantities)
        psets = util.get_psets(elemento)
        for pset_nome, propriedades in psets.items():
            if isinstance(propriedades, dict):
                for prop_nome, valor in propriedades.items():
                    nome_prop = str(prop_nome).lower()
                    # Ignora propriedades de 'profile' (área da secção) e busca 'volume'
                    if 'volume' in nome_prop:
                        if isinstance(valor, (int, float)):
                            return float(valor)
                        elif isinstance(valor, str):
                            numeros = re.findall(r"[-+]?\d*\.\d+|\d+", valor.replace(',', '.'))
                            if numeros:
                                return float(numeros[0])
    except Exception:
        pass
    return 0.0

def extrair_material(elemento):
    """Busca a associação de material padrão do elemento."""
    if hasattr(elemento, 'HasAssociations') and elemento.HasAssociations:
        for rel in elemento.HasAssociations:
            if rel.is_a('IfcRelAssociatesMaterial'):
                mat = rel.RelatingMaterial
                if mat.is_a('IfcMaterial'):
                    return mat.Name
                elif mat.is_a('IfcMaterialList') and len(mat.Materials) > 0:
                    return mat.Materials[0].Name
                elif mat.is_a('IfcMaterialProfileSet'):
                    return mat.MaterialProfiles[0].Material.Name
    return "Betão Especificado"

def extrair_dados_elementos(elementos):
    """Gera uma lista de dados para a tabela do Word a partir de uma lista de elementos."""
    dados = []
    volume_total = 0.0
    
    for el in elementos:
        nome = el.Name if el.Name else "N/A"
        material = extrair_material(el)
        volume = extrair_volume(el)
        volume_total += volume
        
        # Formata o volume, se for 0, coloca 0.00 para evidenciar que foi lido
        vol_str = f"{volume:.2f}" if volume > 0 else "0.00"
        dados.append([nome, material, vol_str])
        
    # Ordena alfabeticamente pelo nome do elemento (Ex: P1, P2, P3)
    dados.sort(key=lambda x: x[0])
    return dados, volume_total

# --- FUNÇÕES DE FORMATAÇÃO DO WORD ---

def adicionar_tabela_formatada(doc, cabecalhos, dados):
    """Cria uma tabela com bordas e cabeçalhos em negrito."""
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
    # Cria um dicionário para organizar os elementos por Piso (IfcBuildingStorey)
    pavimentos = {}
    classes_estruturais = ["IfcColumn", "IfcBeam", "IfcSlab", "IfcFooting"]
    
    for classe in classes_estruturais:
        for el in modelo.by_type(classe):
            # Tenta descobrir em que pavimento o elemento está
            container = util.get_container(el)
            nome_pav = container.Name if container else "Pavimento Indefinido"
            
            if nome_pav not in pavimentos:
                pavimentos[nome_pav] = {"IfcColumn": [], "IfcBeam": [], "IfcSlab": [], "IfcFooting": []}
            
            pavimentos[nome_pav][classe].append(el)

    # --- GERAÇÃO DAS SECÇÕES NO DOCUMENTO ---
    cabecalhos = ['Identificação', 'Material (Classe)', 'Volume Líquido (m³)']
    
    # Mapeamento para títulos no Word
    titulos_classes = {
        "IfcColumn": "PILARES",
        "IfcBeam": "VIGAS",
        "IfcSlab": "LAJES",
        "IfcFooting": "FUNDAÇÕES"
    }

    vol_global = 0.0

    # Iterar sobre as classes estruturais para criar os capítulos (ex: 4. PILARES)
    contador_capitulo = 4
    for classe in classes_estruturais:
        doc.add_heading(f'{contador_capitulo}. {titulos_classes[classe]}', level=1)
        
        vol_total_classe = 0.0
        
        # Iterar sobre os pavimentos dentro de cada classe
        for nome_pav, elementos_do_pavimento in pavimentos.items():
            elementos = elementos_do_pavimento[classe]
            if elementos:
                # Título do Pavimento (ex: 4.1 Térreo)
                doc.add_heading(f'Pavimento: {nome_pav}', level=2)
                
                dados, vol_parcial = extrair_dados_elementos(elementos)
                adicionar_tabela_formatada(doc, cabecalhos, dados)
                
                vol_total_classe += vol_parcial
                doc.add_paragraph() # Espaço

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
