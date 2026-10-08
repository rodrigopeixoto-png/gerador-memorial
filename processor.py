import ifcopenshell
from docx import Document
from io import BytesIO

def gerar_memorial_docx(ifc_file_path):
    # 1. CARREGAR O MODELO IFC
    modelo = ifcopenshell.open(ifc_file_path)
    pilares = modelo.by_type("IfcColumn")
    
    # Lista para armazenar dados reais (simplificado para o exemplo)
    dados_pilares = []
    volume_total = 0.0
    
    for pilar in pilares:
        # Extração básica do nome (A lógica de Psets vai aqui)
        nome = pilar.Name if pilar.Name else "Pilar"
        
        # Simulação de dados que você extrairia das Properties/Quantities
        volume_simulado = 0.25 
        dados_pilares.append({
            "id": nome,
            "secao": "20x40", # Substituir pela lógica de extração de perfil
            "volume": volume_simulado,
            "fck": "C30",     # Substituir pela lógica de material
            "cobrimento": "3 cm"
        })
        volume_total += volume_simulado

    # 2. GERAR O DOCUMENTO WORD
    # Se tiver um template, use: doc = Document('template.docx')
    doc = Document()
    doc.add_heading('Memorial Descritivo Estrutural', level=0)
    doc.add_paragraph(f'Foram identificados {len(pilares)} pilares no modelo.')
    
    # Criar a Tabela
    tabela = doc.add_table(rows=1, cols=5)
    tabela.style = 'Table Grid'
    
    # Cabeçalho
    cabecalhos = ['Identific.', 'Seção (cm)', 'Volume (m³)', 'fck', 'Cobrim.']
    for i, nome_coluna in enumerate(cabecalhos):
        tabela.rows[0].cells[i].text = nome_coluna
        
    # Preencher dados
    for pilar in dados_pilares:
        row = tabela.add_row().cells
        row[0].text = pilar["id"]
        row[1].text = pilar["secao"]
        row[2].text = f"{pilar['volume']:.2f}"
        row[3].text = pilar["fck"]
        row[4].text = pilar["cobrimento"]
        
    doc.add_paragraph()
    doc.add_paragraph(f'Volume Total Estimado: {volume_total:.2f} m³').bold = True

    # 3. SALVAR EM MEMÓRIA (Essencial para o Streamlit)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0) # Retorna o cursor para o início do arquivo
    
    return buffer
