import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from openai import OpenAI
import re

# Configuração da página
st.set_page_config(
    page_title="Dashboard de Orçamento",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dashboard Completo de Orçamento de Obra")
st.markdown("Faça upload da planilha de orçamento e explore análises detalhadas.")

# Sidebar para configurações
with st.sidebar:
    st.header("🔧 Configurações")
    usar_ia = st.checkbox("Usar IA local (LM Studio)", value=True)
    if usar_ia:
        lm_studio_url = st.text_input("URL do LM Studio", value="http://localhost:1234/v1")
        modelo_ia = st.text_input("Modelo", value="hermes-3-llama-3.2-3b")
        st.caption("Certifique-se de que o LM Studio está rodando com o modelo selecionado.")

# Upload do arquivo
uploaded_file = st.file_uploader("Carregue a planilha de orçamento (formato .xls ou .xlsx)", type=["xls", "xlsx"])

if uploaded_file is not None:
    # Ler o arquivo Excel
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name="Orçamento Global", header=None)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        st.stop()

    # --- Extração e limpeza dos dados ---
    # Encontrar a linha do cabeçalho (contém "Código")
    cabecalho_idx = None
    for i, row in df_raw.iterrows():
        if row.astype(str).str.contains("Código").any():
            cabecalho_idx = i
            break

    if cabecalho_idx is None:
        st.error("Não foi possível encontrar o cabeçalho 'Código' na planilha.")
        st.stop()

    # Extrair cabeçalho e dados
    header = df_raw.iloc[cabecalho_idx].fillna('').astype(str).tolist()
    header = [h.strip() for h in header]

    # Mapear colunas essenciais (pode haver variações nos nomes)
    col_map = {}
    for nome_esperado in ['Código', 'Descrição dos Serviços', 'Unit', 'Quant.', 'Preço Serv.', 'Preço Total']:
        for idx, h in enumerate(header):
            if nome_esperado in h:
                col_map[nome_esperado] = idx
                break

    if len(col_map) < 6:
        st.error("Colunas necessárias não encontradas. Verifique a planilha.")
        st.stop()

    # Extrair dados e criar DataFrame
    dados = df_raw.iloc[cabecalho_idx+1:].copy()
    dados = dados.iloc[:, [col_map['Código'], col_map['Descrição dos Serviços'], col_map['Unit'],
                           col_map['Quant.'], col_map['Preço Serv.'], col_map['Preço Total']]]
    dados.columns = ['Código', 'Descrição', 'Unidade', 'Quantidade', 'Preço Unitário', 'Preço Total']

    # Limpeza: converter para números, tratar vírgulas e pontos
    for col in ['Quantidade', 'Preço Unitário', 'Preço Total']:
        dados[col] = dados[col].astype(str).str.replace(',', '.').str.replace('[^0-9.]', '', regex=True)
        dados[col] = pd.to_numeric(dados[col], errors='coerce').fillna(0)

    # Remover linhas sem código (linhas de totais, vazias)
    dados = dados[dados['Código'].astype(str).str.match(r'^\d')].reset_index(drop=True)

    # Extrair código principal (primeiro nível, ex: 01.0)
    dados['Código Principal'] = dados['Código'].astype(str).apply(lambda x: x.split('.')[0] + '.0' if '.' in x else x)

    # Agrupar por código principal para resumo de categorias
    categorias = dados.groupby(['Código Principal', 'Descrição']).agg({'Preço Total': 'sum'}).reset_index()

    # Totais gerais
    total_geral = dados['Preço Total'].sum()
    n_itens = len(dados)
    n_categorias = categorias.shape[0]

    # --- Criação das abas ---
    abas = st.tabs([
        "📋 Visão Geral",
        "📋 Lista Completa",
        "💰 Análise de Preços",
        "📦 Análise de Quantidades",
        "🏷️ Distribuição por Categoria",
        "📊 Análises Estatísticas"
    ])

    # Aba 1: Visão Geral
    with abas[0]:
        st.subheader("📊 Resumo Executivo")

        col1, col2, col3 = st.columns(3)
        col1.metric("Total da Obra", f"R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        col2.metric("Número de Itens", n_itens)
        col3.metric("Categorias", n_categorias)

        # Top 10 itens mais caros
        st.subheader("🔝 Top 10 Itens Mais Caros")
        top10 = dados.nlargest(10, 'Preço Total')[['Descrição', 'Preço Total']]
        fig_top = px.bar(top10, x='Preço Total', y='Descrição', orientation='h',
                         title='Top 10 Itens por Valor Total',
                         labels={'Preço Total': 'R$', 'Descrição': ''})
        st.plotly_chart(fig_top, use_container_width=True)

        # Gráfico de pizza dos 5 principais itens
        top5 = dados.nlargest(5, 'Preço Total')
        fig_pizza = px.pie(top5, values='Preço Total', names='Descrição',
                           title='Distribuição dos 5 Itens Mais Caros')
        st.plotly_chart(fig_pizza, use_container_width=True)

        # Resumo com IA (opcional)
        if usar_ia:
            st.subheader("🤖 Resumo Gerado por IA")
            top5_text = "\n".join([f"- {row['Descrição']}: R$ {row['Preço Total']:,.2f}" for _, row in top5.iterrows()])
            categorias_text = "\n".join([f"- {row['Descrição']}: R$ {row['Preço Total']:,.2f}" for _, row in categorias.head(5).iterrows()])

            prompt = f"""
            Você é um engenheiro especialista em orçamentos de obras. Analise os dados abaixo e gere um resumo executivo profissional, destacando os principais custos, itens mais relevantes e observações pertinentes.

            Total da obra: R$ {total_geral:,.2f}
            Número de itens: {n_itens}
            Categorias: {n_categorias}

            Principais itens (top 5):
            {top5_text}

            Principais categorias (top 5 por custo):
            {categorias_text}

            Resumo:
            """

            if st.button("Gerar resumo com IA"):
                with st.spinner("Consultando IA local..."):
                    try:
                        client = OpenAI(base_url=lm_studio_url, api_key="not-needed")
                        response = client.chat.completions.create(
                            model=modelo_ia,
                            messages=[
                                {"role": "system", "content": "Você é um assistente especializado em orçamentos de engenharia."},
                                {"role": "user", "content": prompt}
                            ],
                            temperature=0.3,
                            max_tokens=600
                        )
                        st.write(response.choices[0].message.content)
                    except Exception as e:
                        st.error(f"Erro ao consultar IA: {e}")

    # Aba 2: Lista Completa
    with abas[1]:
        st.subheader("📋 Lista Completa de Itens")
        st.dataframe(dados, use_container_width=True, height=600)

        # Download da lista
        csv = dados.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download da lista (CSV)", data=csv, file_name="lista_itens.csv", mime="text/csv")

    # Aba 3: Análise de Preços
    with abas[2]:
        st.subheader("💰 Análise de Preços Unitários e Totais")

        col1, col2 = st.columns(2)
        with col1:
            # Distribuição dos preços unitários
            fig_hist_unit = px.histogram(dados, x='Preço Unitário', nbins=50,
                                          title='Distribuição dos Preços Unitários',
                                          labels={'Preço Unitário': 'R$'})
            st.plotly_chart(fig_hist_unit, use_container_width=True)

        with col2:
            # Boxplot dos preços unitários por unidade (para as principais unidades)
            top_unidades = dados['Unidade'].value_counts().nlargest(10).index
            df_filtrado = dados[dados['Unidade'].isin(top_unidades)]
            fig_box = px.box(df_filtrado, x='Unidade', y='Preço Unitário',
                              title='Boxplot dos Preços Unitários por Unidade')
            st.plotly_chart(fig_box, use_container_width=True)

        # Scatter plot: quantidade x preço unitário (com tamanho proporcional ao total)
        fig_scatter = px.scatter(dados, x='Quantidade', y='Preço Unitário', size='Preço Total',
                                 hover_name='Descrição', log_x=True, log_y=True,
                                 title='Relação Quantidade vs Preço Unitário (tamanho = valor total)')
        st.plotly_chart(fig_scatter, use_container_width=True)

    # Aba 4: Análise de Quantidades
    with abas[3]:
        st.subheader("📦 Análise de Quantidades por Unidade")

        # Agrupar quantidade total por unidade
        qtd_por_unidade = dados.groupby('Unidade')['Quantidade'].sum().reset_index().sort_values('Quantidade', ascending=False)
        fig_qtd_unidade = px.bar(qtd_por_unidade.head(15), x='Unidade', y='Quantidade',
                                  title='Quantidade Total por Unidade (top 15)')
        st.plotly_chart(fig_qtd_unidade, use_container_width=True)

        # Top itens por quantidade
        top_qtd = dados.nlargest(10, 'Quantidade')[['Descrição', 'Quantidade', 'Unidade']]
        fig_top_qtd = px.bar(top_qtd, x='Quantidade', y='Descrição', orientation='h',
                              title='Top 10 Itens por Quantidade')
        st.plotly_chart(fig_top_qtd, use_container_width=True)

        # Distribuição das unidades (contagem de itens)
        contagem_unidade = dados['Unidade'].value_counts().reset_index()
        contagem_unidade.columns = ['Unidade', 'Contagem']
        fig_cont_unidade = px.pie(contagem_unidade.head(10), values='Contagem', names='Unidade',
                                   title='Distribuição dos Tipos de Unidade (top 10)')
        st.plotly_chart(fig_cont_unidade, use_container_width=True)

    # Aba 5: Distribuição por Categoria
    with abas[4]:
        st.subheader("🏷️ Distribuição por Categoria (Código Principal)")

        # Gráfico de barras das categorias
        fig_cat = px.bar(categorias, x='Código Principal', y='Preço Total',
                         hover_data=['Descrição'],
                         title='Custo por Categoria',
                         labels={'Preço Total': 'R$', 'Código Principal': 'Categoria'})
        st.plotly_chart(fig_cat, use_container_width=True)

        # Tabela com detalhes das categorias
        st.dataframe(categorias, use_container_width=True)

    # Aba 6: Análises Estatísticas
    with abas[5]:
        st.subheader("📊 Análises Estatísticas dos Preços")

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Estatísticas dos Preços Totais**")
            st.write(dados['Preço Total'].describe())

        with col2:
            st.markdown("**Estatísticas dos Preços Unitários**")
            st.write(dados['Preço Unitário'].describe())

        # Outliers: itens com preço total muito acima da média
        media = dados['Preço Total'].mean()
        std = dados['Preço Total'].std()
        outliers = dados[dados['Preço Total'] > media + 3*std]
        if not outliers.empty:
            st.subheader("🚨 Possíveis Outliers (acima de 3 desvios padrão)")
            st.dataframe(outliers[['Descrição', 'Preço Total']])

        # Matriz de correlação (apenas numéricas)
        corr = dados[['Quantidade', 'Preço Unitário', 'Preço Total']].corr()
        fig_corr = px.imshow(corr, text_auto=True, title='Matriz de Correlação')
        st.plotly_chart(fig_corr, use_container_width=True)

else:
    st.info("Faça o upload da planilha para começar.")

