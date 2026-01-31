import pandas as pd
import streamlit as st
import os

def processar_vendas_ml(file_vendas, mes):
    if not file_vendas:
        return "Erro: Nenhuma planilha de Vendas Mercado Livre foi carregada." 

    try:
        df_vendas = pd.read_excel(file_vendas, header=5)  
        df_vendas.drop(columns=[col for col in df_vendas.columns if "Unnamed" in col], inplace=True, errors='ignore')
        df_vendas.dropna(axis=1, how='all', inplace=True)
        df_vendas.columns = df_vendas.columns.str.strip()

        # Colunas obrigatórias
        required_columns = [
            'Estado',
            'Tarifa de venda e impostos (BRL)',
            'Receita por produtos (BRL)',
            'Forma de entrega',
            'Receita por envio (BRL)',
            'Tarifas de envio (BRL)'
        ]
        missing_columns = [col for col in required_columns if col not in df_vendas.columns]
        if missing_columns:
            return f"Erro: As seguintes colunas não foram encontradas na planilha: {', '.join(missing_columns)}"

        df_vendas['Estado'] = df_vendas['Estado'].astype(str).str.strip().str.lower()
        df_vendas['Forma de entrega'] = df_vendas['Forma de entrega'].astype(str).str.strip().str.lower()

        # ---------------- CANCELADOS ----------------
        df_cancelados = df_vendas[df_vendas['Estado'].str.contains(r'cancelad[oa]|cancelou', regex=True, na=False)]
        df_cancelados['Tarifa de venda e impostos (BRL)'] = pd.to_numeric(
            df_cancelados['Tarifa de venda e impostos (BRL)'], errors='coerce'
        )
        total_comissao_cancelados = df_cancelados['Tarifa de venda e impostos (BRL)'].sum()

        # ---------------- DEVOLUÇÕES (OPERACIONAL - ORIGINAL) ----------------
        df_devolucoes = df_vendas[df_vendas['Estado'].str.contains(r'devoluç|devolvid[oa]', regex=True, na=False)]
        df_devolucoes['Receita por produtos (BRL)'] = pd.to_numeric(
            df_devolucoes['Receita por produtos (BRL)'], errors='coerce'
        )
        total_devolucoes = df_devolucoes['Receita por produtos (BRL)'].sum()

        # ---------------- REEMBOLSOS DE CAIXA (NOVO - SEM MEXER NO ORIGINAL) ----------------
        df_reembolsos_caixa = df_vendas[
            df_vendas['Estado'].str.contains('te demos o dinheiro', na=False)
        ]

        df_reembolsos_caixa['Receita por produtos (BRL)'] = pd.to_numeric(
            df_reembolsos_caixa['Receita por produtos (BRL)'], errors='coerce'
        )

        total_reembolsos_caixa = df_reembolsos_caixa['Receita por produtos (BRL)'].sum()
        qtd_reembolsos_caixa = len(df_reembolsos_caixa)

        # ---------------- FATURAMENTO (SEM CANCELADOS) ----------------
        df_faturamento = df_vendas[
            ~df_vendas['Estado'].str.contains(r'cancelad[oa]|cancelou', regex=True, na=False)
        ]
        df_faturamento['Receita por produtos (BRL)'] = pd.to_numeric(
            df_faturamento['Receita por produtos (BRL)'], errors='coerce'
        )
        total_faturamento = df_faturamento['Receita por produtos (BRL)'].sum()

        # ---------------- FLEX ----------------
        df_flex = df_vendas[
            (df_vendas['Forma de entrega'].str.contains(r'mercado envios flex', regex=True, na=False)) & 
            ~df_vendas['Estado'].str.contains(r'cancelad[oa]|cancelou', regex=True, na=False)
        ]
        df_flex['Receita por envio (BRL)'] = pd.to_numeric(
            df_flex['Receita por envio (BRL)'], errors='coerce'
        )
        total_flex = df_flex['Receita por envio (BRL)'].sum()

        # ---------------- FRETE ----------------
        df_frete = df_vendas[
            df_vendas['Forma de entrega'].str.contains(
                r'mercado envios full|correios|pontos de envio', regex=True, na=False
            )
        ]
        df_frete['Receita por envio (BRL)'] = pd.to_numeric(
            df_frete['Receita por envio (BRL)'], errors='coerce'
        )
        df_frete['Tarifas de envio (BRL)'] = pd.to_numeric(
            df_frete['Tarifas de envio (BRL)'], errors='coerce'
        )
        total_receita_envio = df_frete['Receita por envio (BRL)'].sum()
        total_tarifas_envio = df_frete['Tarifas de envio (BRL)'].sum()

        # ---------------- EXPORTAÇÃO ----------------
        output_dir = 'uploads'
        os.makedirs(output_dir, exist_ok=True)
        output_filepath = os.path.join(output_dir, f"Relatorio_MercadoLivre_{mes}.xlsx")

        with pd.ExcelWriter(output_filepath, engine='openpyxl') as writer:
            pd.DataFrame({
                'Mês': [mes],
                'Comissão Cancelados (BRL)': [total_comissao_cancelados]
            }).to_excel(writer, sheet_name='Comissão Cancelados', index=False)

            pd.DataFrame({
                'Mês': [mes],
                'Faturamento Total (BRL)': [total_faturamento]
            }).to_excel(writer, sheet_name='Faturamento Total', index=False)

            pd.DataFrame({
                'Mês': [mes],
                'Total Devoluções (BRL)': [total_devolucoes]
            }).to_excel(writer, sheet_name='Devoluções', index=False)

            # NOVA ABA (SEM IMPACTAR AS OUTRAS)
            pd.DataFrame({
                'Mês': [mes],
                'Valor reembolsado (BRL)': [total_reembolsos_caixa],
                'Quantidade de registros': [qtd_reembolsos_caixa]
            }).to_excel(writer, sheet_name='Reembolsos Caixa ML', index=False)

            pd.DataFrame({
                'Mês': [mes],
                'Total Flex (BRL)': [total_flex]
            }).to_excel(writer, sheet_name='Flex', index=False)

            pd.DataFrame({
                'Mês': [mes],
                'Receita por envio (BRL)': [total_receita_envio],
                'Tarifas de envio (BRL)': [total_tarifas_envio]
            }).to_excel(writer, sheet_name='Frete', index=False)

        return output_filepath

    except Exception as e:
        return f"Erro ao processar as vendas do Mercado Livre: {e}"


# =========================================================
# MERCADO PAGO (INALTERADO)
# =========================================================
def processar_planilhas_pago(files_pago, mes):
    output_dir = 'uploads'
    os.makedirs(output_dir, exist_ok=True)
    output_filepath = os.path.join(output_dir, f"Relatorio_MercadoPago_{mes}.xlsx")

    all_data = []

    for file in files_pago:
        df = pd.read_excel(file, header=7)
        if not df.empty:
            all_data.append(df)

    if not all_data:
        return "Erro: Nenhuma planilha válida foi carregada."

    df_merged = pd.concat(all_data, ignore_index=True)
    df_merged.columns = df_merged.columns.str.strip().str.lower()
    df_merged.columns = df_merged.columns.astype(str)

    possible_date_columns = [col for col in df_merged.columns if "data da tarifa" in col]
    if not possible_date_columns:
        return "Erro: A coluna 'Data da tarifa' não foi encontrada nas planilhas."

    date_column = possible_date_columns[0]
    df_merged[date_column] = pd.to_datetime(df_merged[date_column], errors='coerce')
    df_filtrado = df_merged[df_merged[date_column].dt.month == mes]

    if df_filtrado.empty:
        return f"Erro: Nenhum dado encontrado para o mês {mes}."

    possible_detalhes_col = [col for col in df_filtrado.columns if "detalhe" in col]
    possible_valor_col = [col for col in df_filtrado.columns if "valor da tarifa" in col]

    if not possible_detalhes_col or not possible_valor_col:
        return "Erro: As colunas 'Detalhes' e 'Valor da tarifa' não foram encontradas."

    detalhes_col = possible_detalhes_col[0]
    valor_col = possible_valor_col[0]

    df_filtrado[valor_col] = pd.to_numeric(df_filtrado[valor_col], errors='coerce')
    df_pivot = df_filtrado.pivot_table(index=detalhes_col, values=valor_col, aggfunc='sum')

    with pd.ExcelWriter(output_filepath, engine='openpyxl') as writer:
        df_filtrado.to_excel(writer, sheet_name="Dados Filtrados", index=False)
        df_pivot.to_excel(writer, sheet_name="Tabela Dinâmica", index=True)

    return output_filepath


def main():
    st.title("📊 Relatórios de Tarifas e Vendas - Marketplace")

    meses_dict = {m: i+1 for i, m in enumerate([
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ])}
    mes_numero = meses_dict[st.selectbox("📅 Selecione o mês:", list(meses_dict.keys()))]

    st.header("📌 Planilhas de Mercado Pago")
    files_pago = st.file_uploader(
        "🔽 Envie as Planilhas de Mercado Pago",
        type=["xls", "xlsx"],
        accept_multiple_files=True
    )

    if st.button("📊 Gerar Relatório Mercado Pago"):
        if files_pago:
            st.info("🔄 Processando Mercado Pago...")
            output_pago = processar_planilhas_pago(files_pago, mes_numero)
            if os.path.exists(output_pago):
                with open(output_pago, "rb") as f:
                    st.download_button(
                        "📥 Baixar Relatório MercadoPago", 
                        data=f, 
                        file_name=f"Relatorio_MercadoPago_{mes_numero}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error(output_pago)

    st.header("📌 Planilha de Vendas Mercado Livre")
    file_vendas = st.file_uploader(
        "🔽 Envie a Planilha de Vendas Mercado Livre",
        type=["xls", "xlsx"]
    )

    if st.button("📊 Gerar Relatório Mercado Livre"):
        if file_vendas:
            st.info("🔄 Processando Mercado Livre...")
            output_vendas = processar_vendas_ml(file_vendas, mes_numero)
            if os.path.exists(output_vendas):
                with open(output_vendas, "rb") as f:
                    st.download_button(
                        "📥 Baixar Relatório MercadoLivre", 
                        data=f, 
                        file_name=f"Relatorio_MercadoLivre_{mes_numero}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error(output_vendas)

if __name__ == '__main__':
    main()
