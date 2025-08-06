import streamlit as st
import pandas as pd
from graphviz import Digraph
from pptx import Presentation
from pptx.util import Inches
import io

# --- Configuração da Página ---
st.set_page_config(layout="wide", page_title="Gerador de Organograma")
st.title("Gerador de Organograma Societário Interativo 🏢")

# (O resto do código é exatamente o mesmo...)

# --- Inicialização do Estado da Aplicação ---
if 'relationships' not in st.session_state:
    st.session_state.relationships = []

# --- Lógica da Aplicação ---
col1, col2 = st.columns([1, 1])

with col1:
    st.header("1. Adicione as Relações Societárias")
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
st.header("3. Gerar o Organograma e a Apresentação")

if st.session_state.relationships:
    if st.button("🚀 Gerar Arquivo PowerPoint", type="primary"):
        with st.spinner("Processando... Por favor, aguarde."):
            df = pd.DataFrame(st.session_state.relationships)
            
            # --- Lógica de Geração do Gráfico (Graphviz) ---
            dot = Digraph(comment='Estrutura Societária', format='png')
            dot.attr(rankdir='TB', splines='ortho', nodesep='0.6')
            dot.attr('node', shape='box', style='rounded,filled', fillcolor='lightblue')
            todas_empresas = set(df['Controladora']).union(set(df['Subsidiária']))
            for empresa in todas_empresas:
                dot.node(empresa, empresa)
            for _, row in df.iterrows():
                dot.edge(row['Controladora'], row['Subsidiária'], label=f"{row['Percentual']}%")

            img_bytes = dot.pipe()

            # --- Lógica de Geração do PPT (python-pptx) ---
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "Estrutura Societária"
            img_stream = io.BytesIO(img_bytes)
            slide.shapes.add_picture(img_stream, Inches(0.5), Inches(1.5), width=Inches(9))
            
            ppt_stream = io.BytesIO()
            prs.save(ppt_stream)
            ppt_stream.seek(0)
            
            st.success("Apresentação gerada com sucesso!")
            
            st.download_button(
                label="📥 Baixar Apresentação (.pptx)",
                data=ppt_stream,
                file_name="organograma_societario.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
else:
    st.warning("Adicione pelo menos uma relação para gerar o organograma.")
