import streamlit as st
import pandas as pd
import io
import openpyxl

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

            # Definir a natureza da conta e o indicador FP
            if conta.startswith('1'):
                natureza_saldo_conta = 'D'
                natureza_mov_baixa = 'C'
                indicador_FP = "1;FP"
            elif conta.startswith('2'):
                natureza_saldo_conta = 'C'
                natureza_mov_baixa = 'D'
                indicador_FP = "1;FP"
            elif conta.startswith('8'):
                natureza_saldo_conta = 'C'
                natureza_mov_baixa = 'D'
                indicador_FP = ";"

            
            # POs únicos nessa aba
            pos_unicos = df.iloc[:, 1].dropna().unique()  # coluna 1 (segunda) é o PO

            for PO in pos_unicos:
                item_movimento = []
                item_saldo_final = []
                item_movimento_devedor = []

                for item in msc_lista:
                    if item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f"period_change;{natureza_mov_baixa}"):
                        item_movimento.append(item)
                    elif item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f'ending_balance;{natureza_saldo_conta}'):
                        item_saldo_final.append(item)
                    elif item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f'period_change;{natureza_saldo_conta}'):
                        if natureza_saldo_conta == 'D':
                            item_movimento_devedor.append(item)


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
                if item_movimento_devedor:
                    itens_novos.append(item_movimento_devedor[0])
                
                # Filtra só as linhas desse PO
                linhas_po = df[df.iloc[:, 1] == PO].values.tolist()
                linhas_po_df = df[df.iloc[:, 1] == PO]
                soma_valores = float(pd.to_numeric(linhas_po_df.iloc[:, 3], errors="coerce").sum())

                if round(soma_valores, 2) != round(valor_end, 2):
                    st.warning(f"Aba {conta}: O valor total do PO {PO} \(R\$ {soma_valores:.2f}\) não bate com o saldo final na MSC (R$ {valor_end})!")
                    

                for linha in linhas_po:
                    fonte = linha[2]
                    valor = linha[3]
                    partes[5] = fonte
                    partes[6] = 'FR'
                    partes[13] = f'{float(valor):.2f}'
                    partes[14] = 'period_change'
                    partes[15] = natureza_saldo_conta
                    itens_novos.append(";".join(partes))
                    partes[14] = 'ending_balance'
                    itens_novos.append(";".join(partes))

                # Substitui no resultado final
                nova_lista = []
                for item in msc_nova:
                    if not item in item_saldo_final and not item in item_movimento_devedor:
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
