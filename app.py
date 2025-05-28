
import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
import xml.etree.ElementTree as ET

# Configuração da página
st.set_page_config(page_title="Leitor de XML | Regime Tributário", layout="centered")

st.title("🔍 Leitor de XML - Regime Tributário da Empresa")
st.write(
    "Envie um arquivo ZIP contendo arquivos XML de NF-e. "
    "O sistema irá extrair CNPJ, Nome da Empresa e o Regime Tributário."
)

# Função para mapear o CRT
def map_crt(crt):
    return {
        '1': 'Simples Nacional',
        '2': 'Simples Nacional, excesso sublimite de receita bruta',
        '3': 'Regime Normal',
        '4': 'Microempreendedor Individual'
    }.get(crt, 'Não identificado')

# Upload do arquivo ZIP
uploaded_file = st.file_uploader("📥 Envie o arquivo ZIP contendo XMLs", type="zip")

if uploaded_file is not None:
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "arquivo.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_file.getvalue())

        # Extrair o ZIP
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        # Procurar todos os arquivos XML dentro das pastas
        xml_files = []
        for root_dir, dirs, files in os.walk(tmpdir):
            for file in files:
                if file.lower().endswith('.xml'):
                    xml_files.append(os.path.join(root_dir, file))

        # Processar os XMLs
        dados = []
        ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}

        for xml_file in xml_files:
            try:
                tree = ET.parse(xml_file)
                root = tree.getroot()

                cnpj = root.find('.//ns:emit/ns:CNPJ', ns)
                nome = root.find('.//ns:emit/ns:xNome', ns)
                crt = root.find('.//ns:emit/ns:CRT', ns)

                if cnpj is not None and nome is not None and crt is not None:
                    dados.append({
                        'CNPJ': cnpj.text,
                        'Nome da Empresa': nome.text,
                        'Regime Tributário': map_crt(crt.text)
                    })
            except Exception as e:
                st.warning(f"Erro ao processar o arquivo: {xml_file}\nErro: {e}")

        if dados:
            df = pd.DataFrame(dados)
            st.success(f"✅ {len(df)} arquivos processados com sucesso!")
            st.dataframe(df)

            # Gerar planilha Excel
            output = pd.ExcelWriter(os.path.join(tmpdir, 'Regime_Tributario.xlsx'), engine='xlsxwriter')
            df.to_excel(output, index=False, sheet_name='Regime')
            output.close()

            with open(os.path.join(tmpdir, 'Regime_Tributario.xlsx'), 'rb') as f:
                st.download_button(
                    label="📥 Baixar Planilha Excel",
                    data=f,
                    file_name="Regime_Tributario.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ Nenhum dado foi extraído. Verifique se os arquivos XML estão no padrão correto.")
