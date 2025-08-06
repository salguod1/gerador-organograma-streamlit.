import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
import io
from collections import defaultdict

# --- Configuração da Página ---
st.set_page_config(layout="wide", page_title="Gerador de Organograma Editável")
st.title("Gerador de Organograma Editável para PowerPoint 🏢")
st.info("Este aplicativo gera um organograma com formas e conectores que podem ser editados diretamente no PowerPoint.")

# --- Inicialização do Estado da Aplicação ---
if 'relationships' not in st.session_state:
    st.session_state.relationships = []

# --- Lógica da Aplicação (Interface do Streamlit - não muda) ---
col1, col2 = st.columns([1, 1])
with col1:
    st.header("1. Adicione as Relações Societárias")
    # ... (O resto da interface do Streamlit é idêntico ao anterior)
    with st.form("relationship_form", clear_on_submit=True):
        controladora = st.text_input("Nome da Empresa Controladora/Holding")
        subsidiaria = st.text_input("Nome da Empresa Subsidiária/Afiliada")
        percentual = st.number_input("Percentual de Posse (%)", min_value=0.0, max_value=100.0, step=0.1, format="%.1f")
        
        submitted = st.form_submit_button("➕ Adicionar Relação")
        if submitted:
            if not controladora or not subsidiaria:
                st.warning("Por favor, preencha o nome da controladora e da subsidiária.")
            else:
                st.session_state.relationships.append({
                    "Controladora": controladora.strip(),
                    "Subsidiária": subsidiaria.strip(),
                    "Percentual": percentual
                })
                st.success(f"Adicionado: {controladora} -> {percentual}% de {subsidiaria}")
with col2:
    st.header("2. Estrutura Atual")
    if st.session_state.relationships:
        df_display = pd.DataFrame(st.session_state.relationships)
        st.dataframe(df_display, use_container_width=True)
        if st.button("🗑️ Limpar Tudo"):
            st.session_state.relationships = []
            st.experimental_rerun()
    else:
        st.info("Nenhuma relação adicionada ainda.")

st.markdown("---")
st.header("3. Gerar a Apresentação Editável")

# --- LÓGICA DE GERAÇÃO DO PPT EDITÁVEL ---

def build_tree(relationships):
    """Constrói uma árvore hierárquica a partir da lista de relações."""
    tree = defaultdict(lambda: {'children': [], 'data': None})
    all_nodes = set()
    children_nodes = set()

    for rel in relationships:
        parent, child, percent = rel['Controladora'], rel['Subsidiária'], rel['Percentual']
        tree[parent]['children'].append({'name': child, 'percent': percent})
        all_nodes.add(parent)
        all_nodes.add(child)
        children_nodes.add(child)

    # Identifica os nós raiz (aqueles que nunca são filhos)
    root_nodes = list(all_nodes - children_nodes)
    if not root_nodes and all_nodes: # Caso de ciclo ou uma única empresa
        root_nodes = [list(all_nodes)[0]]
        
    return tree, root_nodes

def calculate_positions_recursive(node_name, tree, level, sibling_counts, positions, x_offset, level_widths):
    """Calcula recursivamente as posições de cada nó."""
    # Define as dimensões e espaçamentos
    BOX_WIDTH = Inches(2.0)
    BOX_HEIGHT = Inches(1.0)
    H_SPACING = Inches(0.5)
    V_SPACING = Inches(1.5)

    # Calcula a posição Y baseada no nível
    y = level * (BOX_HEIGHT + V_SPACING)

    # Calcula a posição X
    # O x_offset é o ponto de partida para este galho da árvore
    x = x_offset + sibling_counts[level] * (BOX_WIDTH + H_SPACING)
    
    positions[node_name] = {'x': x, 'y': y, 'width': BOX_WIDTH, 'height': BOX_HEIGHT}
    sibling_counts[level] += 1
    level_widths[level] = max(level_widths.get(level, 0), x + BOX_WIDTH)

    # Recursão para os filhos
    child_x_offset = x # O primeiro filho começa alinhado com o pai
    for i, child_info in enumerate(tree[node_name]['children']):
        child_name = child_info['name']
        # Ajusta o offset para os filhos subsequentes para que não se sobreponham
        if i > 0:
           child_x_offset = level_widths.get(level + 1, child_x_offset) + H_SPACING
        calculate_positions_recursive(child_name, tree, level + 1, sibling_counts, positions, child_x_offset, level_widths)

def draw_organogram(slide, relationships, positions, tree):
    """Desenha as formas e conectores no slide do PowerPoint."""
    shapes = {} # Armazena as formas criadas para poder conectar

    # 1. Desenha todas as caixas (formas)
    for name, pos in positions.items():
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, pos['x'], pos['y'], pos['width'], pos['height']
        )
        shape.text = name
        # Customização da aparência da caixa
        text_frame = shape.text_frame
        text_frame.word_wrap = True
        p = text_frame.paragraphs[0]
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
        
        shapes[name] = shape

    # 2. Desenha todos os conectores
    for rel in relationships:
        parent_name, child_name, percent = rel['Controladora'], rel['Subsidiária'], rel['Percentual']
        
        if parent_name in shapes and child_name in shapes:
            from_shape = shapes[parent_name]
            to_shape = shapes[child_name]

            # Adiciona o conector
            connector = slide.shapes.add_connector(
                MSO_CONNECTOR.ELBOW, 
                from_shape.left, from_shape.top, # Posições iniciais (serão ajustadas)
                to_shape.left, to_shape.top
            )
            
            # Conecta o início do conector à parte inferior da forma pai
            connector.begin_connect(from_shape, 3) # 3 = ponto de conexão central inferior
            # Conecta o fim do conector à parte superior da forma filha
            connector.end_connect(to_shape, 1) # 1 = ponto de conexão central superior
            
            # Adiciona o percentual como texto no meio do conector
            # Isso é um pouco mais complexo, então adicionamos uma caixa de texto perto do conector
            line_mid_x = connector.left + connector.width / 2
            line_mid_y = connector.top + connector.height / 2
            
            # Adiciona uma pequena caixa de texto para o percentual
            textbox = slide.shapes.add_textbox(
                line_mid_x - Inches(0.2), line_mid_y - Inches(0.1), 
                Inches(0.4), Inches(0.2)
            )
            textbox.text = f"{percent}%"
            p = textbox.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            # Remove o fundo e a borda da caixa de texto
            textbox.fill.background()
            textbox.line.fill.background()

# Botão para gerar a apresentação
if st.session_state.relationships:
    if st.button("🚀 Gerar Apresentação Editável", type="primary"):
        with st.spinner("Construindo organograma editável... Isso pode levar um momento."):
            
            # 1. Construir a árvore hierárquica
            tree, root_nodes = build_tree(st.session_state.relationships)

            # 2. Calcular as posições de cada nó
            positions = {}
            sibling_counts = defaultdict(int)
            level_widths = {}
            current_x_offset = Inches(0.5)

            for root_name in root_nodes:
                calculate_positions_recursive(root_name, tree, 0, sibling_counts, positions, current_x_offset, level_widths)
                # Atualiza o offset para o próximo organograma (se houver mais de uma raiz)
                current_x_offset = max(level_widths.values() or [0]) + Inches(1.0)
            
            # 3. Criar a apresentação e desenhar o organograma
            prs = Presentation()
            
            # --- CORREÇÃO APLICADA AQUI ---
            # Usar o layout "Título Apenas" ou "Título e Conteúdo" que TEM um placeholder de título.
            # O layout [5] é geralmente "Título Apenas" (Title Only).
            title_only_layout = prs.slide_layouts[5]
            slide = prs.slides.add_slide(title_only_layout)
            
            # Agora esta linha funcionará sem erro
            shapes = slide.shapes
            shapes.title.text = "Estrutura Societária Editável"
            # --- FIM DA CORREÇÃO ---

            # A função draw_organogram não tem um 'shapes' no slide.shapes.title.text
            # então ela não será afetada, mas passamos o slide como antes.
            draw_organogram(slide, st.session_state.relationships, positions, tree)

            # 4. Salvar e disponibilizar para download
            ppt_stream = io.BytesIO()
            prs.save(ppt_stream)
            ppt_stream.seek(0)
            
            st.success("Apresentação editável gerada com sucesso!")
            
            st.download_button(
                label="📥 Baixar Apresentação Editável (.pptx)",
                data=ppt_stream,
                file_name="organograma_editavel.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
else:
    st.warning("Adicione pelo menos uma relação para gerar o organograma.")
