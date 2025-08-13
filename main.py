import streamlit as st
import pandas as pd
import tempfile
from pathlib import Path
from openpyxl import load_workbook

TEMPLATE_PATH = Path("Entradas/TEMPLATE-SITUAÇÃO DE PROJETOS COMPLEMENTARES.xlsx")

## Funções
def ajusta_lista(arquivo_excel, fases_selecionadas):
    # Abrindo o arquivo Excel
    tabela = pd.read_excel(arquivo_excel)

    #Renomeando colunas
    tabela = tabela[['Revisão de Projeto',
                    'Unnamed: 1',
                    'Unnamed: 2',
                    'Unnamed: 3',
                    'Unnamed: 4',
                    'Unnamed: 5',
                    'Unnamed: 8',
                    'Unnamed: 9']]
    index_tabela = {'Revisão de Projeto':'Documento',
                    'Unnamed: 1':'Código',
                    'Unnamed: 2':'Revisão',
                    'Unnamed: 3':'Extensão',
                    'Unnamed: 4':'Situação',
                    'Unnamed: 5':'Título',
                    'Unnamed: 8':'Disciplina',
                    'Unnamed: 9':'Fase'}
    tabela.rename(columns=index_tabela, inplace=True)

    #Removendo linhas desnecessárias
    tabela = tabela.drop(tabela.index[0])
    tabela = tabela.drop(tabela.index[0])

    # Selecionando colunas
    tabela = tabela[['Disciplina',
                    'Fase',
                    'Código',
                    'Revisão',
                    'Documento',
                    'Extensão',
                    'Situação',
                    'Título']]

    # Filtrando por fase
    tabela = tabela[tabela['Fase'].isin(fases_selecionadas)]    
    
    return tabela

def ajusta_planilha_template (arquivo_origem, arquivo_destino):
    
    # Caminhos dos arquivos
    novo_arquivo = Path(r"Saidas/[EMPREENDIMENTO]_Lista_Mestra.xlsx")

    # Intervalo de células que será copiado
    intervalo_inicial = "A2"
    intervalo_final = "H10000"

    # Abrir planilha de origem
    wb_origem = load_workbook(arquivo_origem)
    ws_origem = wb_origem["Sheet1"]

    # Obter linhas e colunas do intervalo
    celulas_origem = ws_origem[intervalo_inicial:intervalo_final]

    # Abrir planilha de destino
    wb_destino = load_workbook(arquivo_destino)
    ws_destino = wb_destino["0-DADOS"]

    # Ponto inicial para colar (A2 → linha 2, coluna 1)
    linha_inicio = 2
    coluna_inicio = 1

    # Copiar valores para o destino
    for i, linha in enumerate(celulas_origem):
        for j, celula in enumerate(linha):
            ws_destino.cell(row=linha_inicio + i, column=coluna_inicio + j, value=celula.value)

    # --- Salvar novo arquivo ---
    wb_destino.save(novo_arquivo)

    return novo_arquivo

########
if __name__ == '__main__':
    # Configurando a página inicial
    st.set_page_config(
        layout= 'wide', #formato da página,
        page_title='Gerar lista mestra de arquivos' #título da página
    )

    with st.sidebar:
        st.divider()
        st.write('Criado por Gilmar Ceregato @2025')
    
    st.title('Gerar lista mestra de arquivos')
    st.divider()

    # Cria botão para upload de arquivo excel?
    uploaded_file = st.file_uploader("Escolha o arquivo Excel", type=["xlsx"])

    if uploaded_file:
        st.write(f"Processando: {uploaded_file.name}")

        # 1. Criar arquivo temporário para salvar a tabela tratada
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
            temp_tratada_path = Path(temp_file.name)
        
        # 2. Ler o arquivo excel enviado e listar as fases de projeto dele
        tabela_tratada = pd.read_excel(uploaded_file)
        if "Fase" in tabela_tratada.columns:
            fases_unicas = tabela_tratada["Fase"].dropna().unique()
        elif "Unnamed: 9" in tabela_tratada.columns:
            fases_unicas = tabela_tratada["Unnamed: 9"].dropna().unique()
        else:
            fases_unicas = []
        
        # Remove "Fase" da lista, se existir
        fases_unicas = [f for f in fases_unicas if f != "Fase"]

        st.info(f"Valores únicos encontrados na coluna 'Fase': {list(fases_unicas)}")

        fases_selecionadas = st.multiselect(
            "Selecione as fases que deseja incluir no arquivo final:",
            options=sorted(fases_unicas),
            default=None
        )
        if not fases_selecionadas:
            st.warning("Selecione pelo menos uma fase para continuar.")
        else:

            # Botão para iniciar o processamento
            gerar = st.button("Gerar lista mestra")
            if gerar:
                # 2. Ajustar a lista e salvar arquivo tratado temporário
                tabela_tratada = pd.read_excel(uploaded_file)
                tabela_tratada = ajusta_lista(uploaded_file, fases_selecionadas)
                tabela_tratada.to_excel(temp_tratada_path, index=False)

                # 3. Aplicar o template na planilha tratada
                arquivo_final = ajusta_planilha_template(temp_tratada_path, TEMPLATE_PATH)

                # 4. Excluir o arquivo temporário
                if temp_tratada_path.exists():
                    temp_tratada_path.unlink()
                
                # 5. Botão para download do arquivo final
                with open(arquivo_final, "rb") as f:
                    st.download_button(
                        label="Baixar Planilha Ajustada",
                        data=f,
                        file_name=arquivo_final.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                st.success("Processamento concluído com sucesso!Faça o download da planilha ajustada.")