import streamlit as st
import pandas as pd
import io

st.title("Atualizador de Arquivo MSC")

# Upload dos arquivos
uploaded_csv = st.file_uploader("Faça upload do arquivo da MSC no formato .CSV", type=["csv"])
uploaded_xlsx = st.file_uploader("Faça upload do arquivo de distribuição por fontes no formato .XLSX", type=["xlsx"])

# Só mostra o botão se os dois arquivos forem enviados
if uploaded_csv and uploaded_xlsx:
    # Lista todas as abas do XLSX
    xls = pd.ExcelFile(uploaded_xlsx)
    sheet_names = xls.sheet_names
    uploaded_xlsx.seek(0)  # reseta o ponteiro do arquivo
    
    st.success(f"Arquivos carregados com sucesso! ✅\n\nForam detectadas as abas: {', '.join(sheet_names)}")
    
    if st.button("Confirmar e processar"):
        
        # Lê o CSV em memória
        msc_lista = uploaded_csv.read().decode("utf-8").splitlines()

        msc_nova = msc_lista.copy()
        itens_processados = []

        # Itera sobre cada aba (conta)
        for conta in sheet_names:
            df = pd.read_excel(uploaded_xlsx, sheet_name=conta, dtype=str)

            # POs únicos nessa aba
            pos_unicos = df.iloc[:, 1].dropna().unique()  # coluna 1 (segunda) é o PO

            for PO in pos_unicos:
                item_movimento = []
                item_saldo_final = []

                for item in msc_lista:
                    if item.startswith(f'{conta};{PO};PO;1;FP;') and item.endswith("period_change;D"):
                        item_movimento.append(item)
                    elif item.startswith(f'{conta};{PO};PO;1;FP;') and item.endswith('ending_balance;C'):
                        item_saldo_final.append(item)

                if not item_movimento or not item_saldo_final:
                    st.warning(f"Aba {conta}, PO {PO} não encontrou linhas correspondentes no CSV, pulando...")
                    continue

                valor_mov = float(item_movimento[0].split(";")[13])
                valor_end = float(item_saldo_final[0].split(";")[13])
                valor_mov_novo = valor_mov + valor_end

                partes = item_movimento[0].split(";")
                partes[13] = f"{valor_mov_novo:.2f}"
                itens_novos = []
                itens_novos.append(";".join(partes))

                # Filtra só as linhas desse PO
                linhas_po = df[df.iloc[:, 1] == PO].values.tolist()

                for linha in linhas_po:
                    fonte = linha[2]
                    valor = linha[3]
                    partes[5] = fonte
                    partes[6] = 'FR'
                    partes[13] = f'{float(valor):.2f}'
                    partes[14] = 'period_change'
                    partes[15] = 'C'
                    itens_novos.append(";".join(partes))
                    partes[14] = 'ending_balance'
                    itens_novos.append(";".join(partes))

                # Substitui no resultado final
                nova_lista = []
                for item in msc_nova:
                    if not item in item_saldo_final:
                        if item in item_movimento:
                            for item_novo in itens_novos:
                                nova_lista.append(item_novo)
                        else:
                            nova_lista.append(item)
                msc_nova = nova_lista

                itens_processados.append(f"{conta}/{PO}")

        # Gera o CSV atualizado em memória
        output = io.StringIO()
        output.write("\n".join(msc_nova))
        output.seek(0)

        st.success(f"Processamento concluído! Contas/POs processados: {', '.join(itens_processados)}")

        # Botão de download
        st.download_button(
            label="Baixar MSC atualizada",
            data=output.getvalue(),
            file_name="202508 MSC_atualizada.csv",
            mime="text/csv"
        )
