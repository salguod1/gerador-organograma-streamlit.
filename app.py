import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
import io
from collections import defaultdict

# --- Configuração da Página ---
st.set_page_config(layout="wide", page_title="Gerador de Organograma Neumórfico")
st.title("Gerador de Organograma Neumórfico para PowerPoint 🎨")
st.info("Crie um organograma com um design moderno e sutil, totalmente editável no PowerPoint.")

# --- Interface do Streamlit (sem alterações) ---
if 'relationships' not in st.session_state:
    st.session_state.relationships = []
col1, col2 = st.columns([1, 1])
with col1:
    st.header("1. Adicione as Relações")
    with st.form("relationship_form", clear_on_submit=True):
        controladora = st.text_input("Nome da Empresa Controladora/Holding")
        subsidiaria = st.text_input("Nome da Empresa Subsidiária/Afiliada")
        percentual = st.number_input("Percentual de Posse (%)", min_value=0.0, max_value=100.0, step=0.1, format="%.1f")
        submitted = st.form_submit_button("➕ Adicionar Relação")
        if submitted:
            if not controladora or not subsidiaria:
                st.warning("Por favor, preencha todos os campos.")
            else:
                st.session_state.relationships.append({"Controladora": controladora.strip(), "Subsidiária": subsidiaria.strip(), "Percentual": percentual})
                st.success(f"Adicionado: {controladora} -> {percentual}% de {subsidiaria}")
with col2:
    st.header("2. Estrutura Atual")
    if st.session_state.relationships:
        st.dataframe(pd.DataFrame(st.session_state.relationships), use_container_width=True)
        if st.button("🗑️ Limpar Tudo"):
            st.session_state.relationships = []
            st.experimental_rerun()
    else:
        st.info("Nenhuma relação adicionada.")
st.markdown("---")
st.header("3. Gerar a Apresentação Neumórfica")

# --- LÓGICA DE GERAÇÃO DO PPT NEUMÓRFICO ---

def build_tree(relationships):
    tree = defaultdict(lambda: {'children': []})
    all_nodes, children_nodes = set(), set()
    for rel in relationships:
        parent, child, percent = rel['Controladora'], rel['Subsidiária'], rel['Percentual']
        tree[parent]['children'].append({'name': child, 'percent': percent})
        all_nodes.update([parent, child])
        children_nodes.add(child)
    root_nodes = list(all_nodes - children_nodes)
    if not root_nodes and all_nodes:
        root_nodes = [list(all_nodes)[0]]
    return tree, root_nodes

def calculate_positions_recursive(node_name, tree, level, sibling_counts, positions, x_offset, level_widths):
    BOX_WIDTH, BOX_HEIGHT = Inches(2.2), Inches(1.1)
    H_SPACING, V_SPACING = Inches(0.8), Inches(1.8)
    y = level * (BOX_HEIGHT + V_SPACING)
    x = x_offset + sibling_counts[level] * (BOX_WIDTH + H_SPACING)
    positions[node_name] = {'x': x, 'y': y, 'width': BOX_WIDTH, 'height': BOX_HEIGHT}
    sibling_counts[level] += 1
    level_widths[level] = max(level_widths.get(level, 0), x + BOX_WIDTH)
    child_x_offset = x
    for i, child_info in enumerate(tree[node_name]['children']):
        if i > 0:
            child_x_offset = level_widths.get(level + 1, child_x_offset) + H_SPACING
        calculate_positions_recursive(child_info['name'], tree, level + 1, sibling_counts, positions, child_x_offset, level_widths)

def draw_organogram_neumorphic(slide, relationships, positions, tree):
    """Desenha o organograma com estilo neumórfico."""
    shapes = {}
    
    # Paleta de cores Neumórfica
    BG_COLOR = RGBColor(224, 229, 236)  # Um cinza-azulado claro
    MAIN_COLOR = RGBColor(236, 240, 245) # Cor principal das caixas, quase igual ao fundo
    DARK_SHADOW = RGBColor(211, 217, 224)
    LIGHT_SHADOW = RGBColor(255, 255, 255)
    TEXT_COLOR = RGBColor(94, 108, 132)
    CONNECTOR_COLOR = RGBColor(188, 196, 208)

    # Define o fundo do slide
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = BG_COLOR

    # Parâmetros do efeito
    shadow_offset = Pt(4)

    # 1. Desenha as caixas (formas) com o efeito
    for name, pos in positions.items():
        # Desenha a sombra escura (primeiro, para ficar no fundo)
        dark_shadow_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, 
            pos['x'] + shadow_offset, pos['y'] + shadow_offset, 
            pos['width'], pos['height']
        )
        dark_shadow_shape.fill.solid()
        dark_shadow_shape.fill.fore_color.rgb = DARK_SHADOW
        dark_shadow_shape.line.fill.background() # Sem borda

        # Desenha a luz clara
        light_shadow_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, 
            pos['x'] - shadow_offset, pos['y'] - shadow_offset, 
            pos['width'], pos['height']
        )
        light_shadow_shape.fill.solid()
        light_shadow_shape.fill.fore_color.rgb = LIGHT_SHADOW
        light_shadow_shape.line.fill.background() # Sem borda
        
        # Desenha a forma principal por cima de tudo
        main_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, pos['x'], pos['y'], pos['width'], pos['height']
        )
        main_shape.fill.solid()
        main_shape.fill.fore_color.rgb = MAIN_COLOR
        main_shape.line.fill.background() # Sem borda

        # Adiciona o texto à forma principal
        main_shape.text = name
        text_frame = main_shape.text_frame
        text_frame.margin_bottom = Inches(0.05)
        text_frame.margin_top = Inches(0.05)
        text_frame.word_wrap = True
        p = text_frame.paragraphs[0]
        p.font.name = 'Calibri'
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = TEXT_COLOR
        
        shapes[name] = main_shape

    # 2. Desenha todos os conectores
    for rel in relationships:
        parent_name, child_name, percent = rel['Controladora'], rel['Subsidiária'], rel['Percentual']
        if parent_name in shapes and child_name in shapes:
            from_shape, to_shape = shapes[parent_name], shapes[child_name]
            connector = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW, 0, 0, 0, 0)
            connector.begin_connect(from_shape, 3)
            connector.end_connect(to_shape, 1)
            
            # Estiliza o conector
            line = connector.line
            line.color.rgb = CONNECTOR_COLOR
            line.width = Pt(1.5)

            # Adiciona a caixa de texto para o percentual
            textbox = slide.shapes.add_textbox(
                connector.left + connector.width/2 - Inches(0.3), 
                connector.top + connector.height/2 - Inches(0.15),
                Inches(0.6), Inches(0.3)
            )
            textbox.text = f"{percent}%"
            p = textbox.text_frame.paragraphs[0]
            p.font.size = Pt(9)
            p.font.color.rgb = TEXT_COLOR
            textbox.fill.background()
            textbox.line.fill.background()

# Botão para gerar a apresentação
if st.session_state.relationships:
    if st.button("🚀 Gerar Apresentação Neumórfica", type="primary"):
        with st.spinner("Criando um design neumórfico..."):
            tree, root_nodes = build_tree(st.session_state.relationships)
            positions, sibling_counts, level_widths = {}, defaultdict(int), {}
            current_x_offset = Inches(0.5)
            for root_name in root_nodes:
                calculate_positions_recursive(root_name, tree, 0, sibling_counts, positions, current_x_offset, level_widths)
                current_x_offset = max(level_widths.values() or [0]) + Inches(1.5)
            
            prs = Presentation()
            prs.slide_width = Inches(16)
            prs.slide_height = Inches(9)
            blank_slide_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(blank_slide_layout)
            
            draw_organogram_neumorphic(slide, st.session_state.relationships, positions, tree)

            ppt_stream = io.BytesIO()
            prs.save(ppt_stream)
            ppt_stream.seek(0)
            
            st.success("Apresentação neumórfica gerada com sucesso!")
            st.download_button(
                label="📥 Baixar Apresentação Neumórfica (.pptx)",
                data=ppt_stream,
                file_name="organograma_neumorfico.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
else:
    st.warning("Adicione pelo menos uma relação para começar.")
