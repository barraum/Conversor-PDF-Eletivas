import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import os
import io # Necessário para o download do arquivo Excel

# --- Suas Constantes Globais e Função de Extração (copiadas da v6) ---
HEADER_TABLE_PARTS_LOWER = [
    "procedimento", "solic. vaga", "cartão sus", "data nasc."
]
SKIP_KEYWORDS_STARTSWITH = [
    "R E L A T Ó R I O  D E  I N T E R N A Ç Ã O  H O S P I T A L A R  -  A I H",
    "d e  2 0 1 7 ) ,  m u d a n ç a  d o  A r t  2 º"
]

def extract_data_from_pdf_multiline(pdf_bytes_io): # Modificado para aceitar bytes
    all_final_rows = []
    headers = ["Categoria de Procedimento", "Procedimento", "Solic. Vaga", "Cartão SUS", "Data Nasc."]
    procedimento_pattern = re.compile(r"(\d{8,10}\s*-\s*.*)")
    data_pattern = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    potential_cartao_sus_pattern = re.compile(r"^\d{10,}$")
    skip_keywords_contain = [
        "Emissão:", "PREFEITURA MUNICIPAL DE FRANCA", "Secretaria Municipal de Saúde",
        "RELATÓRIO DE INTERNAÇÃO HOSPITALAR - AIH", "Total procedimento:",
        "Categoria de Procedimento Procedimento Solic. Vaga Cartão SUS Data Nasc.",
        "Procedimento Solic. Vaga Cartão SUS Data Nasc.", "Pág."
    ]
    total_procedimentos_pdf_sum = 0
    regex_total_procedimento = re.compile(r"Total procedimento:\s*(\d{1,4})")
    
    try:
        # Modificado: Abrir PDF a partir de bytes
        doc = fitz.open(stream=pdf_bytes_io, filetype="pdf") 
    except Exception as e:
        st.error(f"Erro ao abrir o arquivo PDF: {e}")
        return None, 0

    current_record = {}
    categoria_buffer = []
    procedimento_buffer = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        lines = text.split('\n')
        for line_num, raw_line in enumerate(lines):
            line = raw_line.strip()
            if not line: continue

            found_skip_keyword = False
            triggered_skip_keyword = None
            for skc in skip_keywords_contain:
                if skc.lower() in line.lower():
                    found_skip_keyword = True; triggered_skip_keyword = skc; break
            if not found_skip_keyword:
                for sks in SKIP_KEYWORDS_STARTSWITH:
                    if line.lower().startswith(sks.lower()):
                        found_skip_keyword = True; triggered_skip_keyword = sks; break
            
            if found_skip_keyword:
                if triggered_skip_keyword and "total procedimento:" in triggered_skip_keyword.lower():
                    match_total = regex_total_procedimento.search(line)
                    if match_total:
                        try:
                            total_procedimentos_pdf_sum += int(match_total.group(1))
                        except ValueError:
                            st.warning(f"AVISO: Não foi possível converter o total da linha '{line}' para número.")
                
                if current_record.get("Solic. Vaga"): 
                    if categoria_buffer:
                        cat_text = " ".join(categoria_buffer).replace(" ,", ",").strip()
                        if not any(ht in cat_text.lower() for ht in HEADER_TABLE_PARTS_LOWER):
                            current_record["Categoria de Procedimento"] = cat_text
                    if procedimento_buffer:
                        current_record["Procedimento"] = "".join(procedimento_buffer).strip()
                    if (current_record.get("Categoria de Procedimento", "").strip() and \
                        current_record.get("Procedimento", "").strip()): 
                        all_final_rows.append([
                            current_record.get("Categoria de Procedimento", ""),
                            current_record.get("Procedimento", ""),
                            current_record.get("Solic. Vaga", ""),
                            current_record.get("Cartão SUS", ""), 
                            current_record.get("Data Nasc.", "") 
                        ])
                current_record = {}; categoria_buffer = []; procedimento_buffer = []
                continue

            if data_pattern.match(line):
                if not current_record.get("Solic. Vaga"):
                    if categoria_buffer:
                        cat_text = " ".join(categoria_buffer).replace(" ,",",").strip()
                        if not any(ht in cat_text.lower() for ht in HEADER_TABLE_PARTS_LOWER):
                            current_record["Categoria de Procedimento"] = cat_text
                        else: current_record = {}; categoria_buffer = []; procedimento_buffer = []; continue
                        categoria_buffer = []
                    if procedimento_buffer:
                        current_record["Procedimento"] = "".join(procedimento_buffer).strip(); procedimento_buffer = []
                    if current_record.get("Categoria de Procedimento") or current_record.get("Procedimento"):
                        current_record["Solic. Vaga"] = line
                    else: current_record = {}; categoria_buffer = []; procedimento_buffer = []; continue
                elif current_record.get("Solic. Vaga") and not current_record.get("Data Nasc."): 
                    current_record["Data Nasc."] = line
                    if (current_record.get("Categoria de Procedimento","").strip() and \
                        current_record.get("Procedimento","").strip() and \
                        current_record.get("Solic. Vaga","").strip()):
                        all_final_rows.append([
                            current_record.get("Categoria de Procedimento", ""), current_record.get("Procedimento", ""),
                            current_record.get("Solic. Vaga", ""), current_record.get("Cartão SUS", ""),
                            current_record.get("Data Nasc.", "")
                        ])
                    current_record = {}; categoria_buffer = []; procedimento_buffer = []
                continue 

            if potential_cartao_sus_pattern.match(line) and \
               current_record.get("Solic. Vaga") and not current_record.get("Cartão SUS"):
                current_record["Cartão SUS"] = line
                if categoria_buffer:
                    cat_text = " ".join(categoria_buffer).replace(" ,",",").strip()
                    if not any(ht in cat_text.lower() for ht in HEADER_TABLE_PARTS_LOWER):
                        current_record["Categoria de Procedimento"] = cat_text
                    else: current_record = {}; categoria_buffer = []; procedimento_buffer = []; continue
                    categoria_buffer = []
                if procedimento_buffer:
                    current_record["Procedimento"] = "".join(procedimento_buffer).strip(); procedimento_buffer = []
                continue 

            proc_match = procedimento_pattern.match(line)
            if proc_match:
                if categoria_buffer: 
                    cat_text = " ".join(categoria_buffer).replace(" ,",",").strip()
                    if not any(ht in cat_text.lower() for ht in HEADER_TABLE_PARTS_LOWER):
                        current_record["Categoria de Procedimento"] = cat_text
                    else: current_record = {}; categoria_buffer = []; procedimento_buffer = []; continue
                    categoria_buffer = []
                if current_record.get("Procedimento"): current_record = {} 
                procedimento_buffer.append(line.strip()) 
                continue

            if not current_record.get("Solic. Vaga"):
                if procedimento_buffer: procedimento_buffer.append(line.strip())
                elif not current_record.get("Procedimento"): 
                    if not any(ht in line.lower() for ht in HEADER_TABLE_PARTS_LOWER):
                        categoria_buffer.append(line.strip())
    
    if current_record.get("Solic. Vaga"): 
        if categoria_buffer: 
            cat_text_final = " ".join(categoria_buffer).replace(" ,",",").strip()
            if not any(ht in cat_text_final.lower() for ht in HEADER_TABLE_PARTS_LOWER):
                current_record["Categoria de Procedimento"] = cat_text_final
        if procedimento_buffer: 
            current_record["Procedimento"] = "".join(procedimento_buffer).strip()
        if (current_record.get("Categoria de Procedimento","").strip() and \
            current_record.get("Procedimento","").strip()):
            all_final_rows.append([
                current_record.get("Categoria de Procedimento",""), current_record.get("Procedimento",""),
                current_record.get("Solic. Vaga",""), current_record.get("Cartão SUS",""),
                current_record.get("Data Nasc.","") 
            ])
    doc.close()
    if not all_final_rows: 
        st.info("Nenhum dado tabular estruturado foi extraído do PDF.")
        return None, total_procedimentos_pdf_sum
    df = pd.DataFrame(all_final_rows, columns=headers)
    return df, total_procedimentos_pdf_sum
# --- Fim da Função de Extração ---

# --- Interface Streamlit ---
st.set_page_config(page_title="Extrator PDF para Excel", layout="wide")
st.title("📄 Extrator de Dados de PDF para Excel")
st.markdown("""
Carregue seu arquivo PDF (Relatório de Internação Hospitalar) para extrair os dados
e gerar uma planilha Excel.
""")

uploaded_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")

if uploaded_file is not None:
    st.info(f"Arquivo carregado: {uploaded_file.name}")
    
    # Para PyMuPDF abrir a partir de um arquivo carregado pelo Streamlit,
    # precisamos passar o conteúdo em bytes.
    pdf_bytes_io = io.BytesIO(uploaded_file.getvalue())

    with st.spinner("Processando PDF... Por favor, aguarde."):
        extracted_df, soma_total_pdf = extract_data_from_pdf_multiline(pdf_bytes_io)

    if extracted_df is not None and not extracted_df.empty:
        st.success("PDF processado com sucesso!")

        # Validação de Totais
        num_linhas_brutas_extraidas = len(extracted_df)
        st.subheader("📊 Validação de Totais")
        st.write(f"- Total de linhas de dados extraídas (bruto): `{num_linhas_brutas_extraidas}`")
        st.write(f"- Soma dos 'Total procedimento:' do PDF: `{soma_total_pdf}`")

        if soma_total_pdf > 0:
            if num_linhas_brutas_extraidas == soma_total_pdf:
                st.success("✅ VALIDAÇÃO OK: O número de linhas extraídas bate com a soma dos totais do PDF.")
            else:
                diferenca = num_linhas_brutas_extraidas - soma_total_pdf
                st.warning(f"⚠️ AVISO DE VALIDAÇÃO: Linhas extraídas ({num_linhas_brutas_extraidas}) "
                           f"NÃO BATEM com a soma dos totais do PDF ({soma_total_pdf}). "
                           f"Diferença: {diferenca}")
        else:
            st.info("ℹ️ Nenhum 'Total procedimento:' foi encontrado/somado do PDF para validação, ou a soma foi zero.")
        
        # Filtragem Final (como na função main anterior)
        df_para_salvar = extracted_df.copy()
        
        df_para_salvar_f1 = df_para_salvar[
            (df_para_salvar["Categoria de Procedimento"].str.strip().fillna('') != "") &
            (df_para_salvar["Categoria de Procedimento"].str.strip().fillna('') != "N/A") &
            (df_para_salvar["Procedimento"].str.strip().fillna('') != "") & 
            (df_para_salvar["Procedimento"].str.strip().fillna('') != "N/A") &
            (df_para_salvar["Solic. Vaga"].str.strip().fillna('') != "") &
            (df_para_salvar["Solic. Vaga"].str.strip().fillna('') != "N/A")
        ].copy()
        
        df_para_salvar_f2 = df_para_salvar_f1.copy()
        for ht_part in HEADER_TABLE_PARTS_LOWER: 
             df_para_salvar_f2 = df_para_salvar_f2[~df_para_salvar_f2["Categoria de Procedimento"].str.lower().str.contains(ht_part, na=False)]

        df_para_salvar_f3 = df_para_salvar_f2.copy()
        for sks in SKIP_KEYWORDS_STARTSWITH: 
            df_para_salvar_f3 = df_para_salvar_f3[~df_para_salvar_f3["Categoria de Procedimento"].str.strip().fillna('').str.lower().eq(sks.lower())]
        
        df_final_filtrado = df_para_salvar_f3
        
        st.subheader("📋 Pré-visualização dos Dados Extraídos (Primeiras 10 linhas)")
        st.dataframe(df_final_filtrado.head(10))
        st.info(f"Total de linhas a serem salvas no Excel (após filtros): {len(df_final_filtrado)}")

        # Preparar arquivo Excel para download
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            df_final_filtrado.to_excel(writer, index=False, sheet_name='DadosExtraidos')
        # O nome do arquivo Excel será baseado no nome do PDF original
        excel_file_name = os.path.splitext(uploaded_file.name)[0] + "_extraido.xlsx"
        
        st.download_button(
            label="📥 Baixar Arquivo Excel",
            data=output_excel.getvalue(), # .getvalue() é importante aqui
            file_name=excel_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    elif extracted_df is None and soma_total_pdf is not None: # Se a extração falhou mas temos totais
        st.subheader("📊 Validação de Totais")
        st.write(f"- Total de linhas de dados extraídas (bruto): `0`")
        st.write(f"- Soma dos 'Total procedimento:' do PDF: `{soma_total_pdf}`")
        if soma_total_pdf > 0:
             st.warning(f"⚠️ AVISO DE VALIDAÇÃO: Nenhum dado foi extraído, mas o PDF indicava {soma_total_pdf} procedimentos nas linhas 'Total procedimento:'.")
        else:
            st.info("ℹ️ Nenhum 'Total procedimento:' foi encontrado/somado do PDF.")
    else:
        # Este caso é se extracted_df é None e soma_total_pdf também é (ou 0 se houve erro ao abrir PDF)
        # A mensagem de "Nenhum dado tabular estruturado..." já foi mostrada dentro da função de extração
        pass