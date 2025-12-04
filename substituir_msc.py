import streamlit as st
import pandas as pd
import io
import openpyxl

st.title("Atualizador de Arquivo MSC")

# Upload dos arquivos
uploaded_csv = st.file_uploader("Faça upload do arquivo da MSC no formato .CSV", type=["csv"])
# uploaded_csv_anterior = st.file_uploader("Faça o upload da MSC anterior no formato .CSV", type=["csv"])
uploaded_xlsx = st.file_uploader("Faça upload do arquivo de distribuição por fontes no formato .XLSX", type=["xlsx"])

# Só mostra o botão se os dois arquivos (MSC e distribuição) forem enviados, MSC anterior não é obrigatório
if uploaded_csv and uploaded_xlsx:
    # Lista todas as abas do XLSX
    xls = pd.ExcelFile(uploaded_xlsx)
    sheet_names = xls.sheet_names
    uploaded_xlsx.seek(0)  # reseta o ponteiro do arquivo
    
    st.success(f"Arquivos carregados com sucesso! ✅\n\nForam detectadas as abas: {', '.join(sheet_names)}")
    
    if st.button("Confirmar e processar"):
        
        lista_erros = []

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
                # Classifica todos os itens que atendem a combinação de conta e PO
                
                item_saldo_inicial = []
                item_movimento_baixa = []
                item_movimento_normal = []
                item_saldo_final = []

                for item in msc_lista:
                    if item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f'beginning_balance;{natureza_saldo_conta}'): # saldo inicial normal
                        item_saldo_inicial.append(item)
                    elif item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f'beginning_balance;{natureza_mov_baixa}'): # saldo inicial invertido
                        item_saldo_inicial.append(item)
                    elif item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f'period_change;{natureza_saldo_conta}'): # movimento normal
                        item_movimento_normal.append(item)
                    elif item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f"period_change;{natureza_mov_baixa}"): # movimento baixa
                        item_movimento_baixa.append(item)
                    elif item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f'ending_balance;{natureza_saldo_conta}'): # saldo final normal 
                        item_saldo_final.append(item)
                        invertido = False
                    elif item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f'ending_balance;{natureza_mov_baixa}'): # saldo final invertido
                        item_saldo_final.append(item)
                        invertido = True # Flag para indicar que o saldo final está invertido na MSC
                                      
                if not item_saldo_final:
                    erro = f"Aba {conta}, PO {PO} não encontrou linhas correspondentes no CSV, pulando..."
                    st.warning(erro)
                    lista_erros.append(erro)
                    continue

                # Calcula o novo valor da linha de movimento de baixa    
                if not item_movimento_baixa:
                    valor_mov_baixa = 0
                else:
                    valor_mov_baixa = float(item_movimento_baixa[0].split(";")[13])

                # Calcula o novo valor da linha de movimento normal
                if not item_movimento_normal:
                    valor_mov_normal = 0
                else:
                    valor_mov_normal = float(item_movimento_normal[0].split(";")[13])
                
                valor_end = float(item_saldo_final[0].split(";")[13])
                valor_mov_baixa_novo = valor_mov_baixa + valor_end
                valor_mov_normal_novo = valor_mov_normal + valor_end

                # Substitui o valor de movimento de baixa             
                partes = item_saldo_final[0].split(";")
                partes[14] = 'period_change'

                if partes[15] == natureza_saldo_conta: # Verifica a naturezas do saldo final, se for devedor, o movimento de baixa deve ser credor e vice-versa
                    partes[13] = f"{valor_mov_baixa_novo:.2f}"
                    partes[15] = natureza_mov_baixa
                else:
                    partes[13] = f"{valor_mov_normal_novo:.2f}"
                    partes[15] = natureza_saldo_conta

                itens_novos = []
                itens_novos.append(";".join(partes)) # Cria a linha de movimento com o valor novo
                
                # Filtra só as linhas desse PO
                linhas_po = df[df.iloc[:, 1] == PO].values.tolist()
                linhas_po_df = df[df.iloc[:, 1] == PO]
                soma_valores = abs(float(pd.to_numeric(linhas_po_df.iloc[:, 3], errors="coerce").sum()))

                if round(soma_valores, 2) != round(valor_end, 2):
                    erro = f"Aba {conta}: O valor total do PO {PO} \(R\$ {soma_valores:.2f}\) não bate com o saldo final na MSC (R$ {valor_end})!"
                    st.warning(erro)
                    lista_erros.append(erro)    

                for linha in linhas_po:
                    fonte = linha[2]
                    valor = abs(float(linha[3])) # considerar valores absolutos e ajustar apenas a natureza da conta
                    partes[5] = fonte
                    partes[6] = 'FR'
                    partes[13] = f'{float(valor):.2f}'
                    partes[14] = 'period_change'
                    if float(linha[3]) > 0:
                        partes[15] = natureza_saldo_conta
                    else:
                        partes[15] = natureza_mov_baixa # ajustar a natureza quando o valor for negativo
                    itens_novos.append(";".join(partes)) # Cria as linhas de movimento para gerar saldo em cada fonte
                    partes[14] = 'ending_balance'
                    itens_novos.append(";".join(partes)) # Cria as linhas de saldo final para cada fonte
                
                # Substitui no resultado final
                nova_lista = []

                for item in msc_nova:
                    # Saldo inicial → mantém como está
                    if item in item_saldo_inicial:
                        nova_lista.append(item)
                    # Saldo final → substitui por itens_novos
                    elif item in item_saldo_final:
                        for item_novo in itens_novos:
                            nova_lista.append(item_novo)
                    # Movimento baixa
                    elif item in item_movimento_baixa:
                        if invertido:
                            nova_lista.append(item)
                    # Movimento normal
                    elif item in item_movimento_normal:
                        if not invertido:
                            nova_lista.append(item)
                    # TODAS AS OUTRAS LINHAS
                    else:
                        nova_lista.append(item)

                msc_nova = nova_lista

                itens_processados.append(f"{conta}/{PO}")

        # Gera os arquivos em memória
        output = io.StringIO()
        output.write("\n".join(msc_nova))
        output.seek(0)

        erros = io.StringIO()
        erros.write("\n".join(lista_erros))
        erros.seek(0)

        st.success(f"Processamento concluído! Contas/POs processados: {', '.join(itens_processados)}")

        # Botão de download
        st.download_button(
            label="Baixar MSC atualizada",
            data=output.getvalue(),
            file_name="MSC_atualizada.csv",
            mime="text/csv"
        )

        if lista_erros:
            st.download_button(
                label="Baixar log de erros",
                data=erros.getvalue(),
                file_name="erros.txt",
                mime="text/csv"
            )
