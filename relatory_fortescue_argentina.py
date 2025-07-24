import uno
import os
import datetime
import pandas as pd
from pandas import ExcelWriter
import re # Importa o módulo de expressões regulares para parsear o ID_Tarea

def migrar_dados_para_adelino():
    # Obtém o contexto do LibreOffice
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    model = desktop.getCurrentComponent()

    if not model:
        exibir_mensagem_erro("Nenhum documento do LibreOffice está aberto. Por favor, abra qualquer documento para executar a macro.")
        return

    # Pasta FIXA conforme sua informação - ESTA LINHA É CRUCIAL PARA EVITAR O ERRO 'REASON 21'
    pasta = r"C:\Users\Paola\Desktop"

    # Caminhos dos arquivos
    arquivo_adelino = os.path.join(pasta, "adelino.xlsx")
    arquivo_pamela = os.path.join(pasta, "pamela.xlsx")

    # --- Verificação de existência dos arquivos ---
    if not os.path.exists(arquivo_adelino):
        exibir_mensagem_erro(f"Arquivo não encontrado: {arquivo_adelino}\nCertifique-se de que ele está na pasta especificada no código ({pasta}).")
        return
    if not os.path.exists(arquivo_pamela):
        exibir_mensagem_erro(f"Arquivo não encontrado: {arquivo_pamela}\nCertifique-se de que ele está na pasta especificada no código ({pasta}).")
        return

    # --- AVISO CRÍTICO AO USUÁRIO ---
    exibir_mensagem_aviso("ATENÇÃO: Certifique-se de que os arquivos 'adelino.xlsx' e 'pamela.xlsx' estão COMPLETAMENTE FECHADOS no LibreOffice Calc (e em qualquer outro programa) antes de continuar. Caso contrário, a cópia pode não ser visível ou causar erros.\n\nEste método SOBRESCREVERÁ o arquivo 'adelino.xlsx' e sua formatação original será perdida.")
    # --- FIM DO AVISO CRÍTICO ---

    try:
        # Lê as planilhas com pandas
        df_adelino = pd.read_excel(arquivo_adelino)
        df_pamela = pd.read_excel(arquivo_pamela)

        # --- Validação de colunas ---
        if "ASSUNTO" not in df_pamela.columns:
            exibir_mensagem_erro(f"A coluna 'ASSUNTO' (Coluna B) não foi encontrada em '{os.path.basename(arquivo_pamela)}'.")
            return
        if "DATA DA SOLUÇÃO" not in df_pamela.columns:
            exibir_mensagem_erro(f"A coluna 'DATA DA SOLUÇÃO' (Coluna A) não foi encontrada em '{os.path.basename(arquivo_pamela)}'. Não é possível filtrar ou copiar por data.")
            return
        if "SERVIÇO REALIZADO" not in df_pamela.columns:
            exibir_mensagem_erro(f"A coluna 'SERVIÇO REALIZADO' (Coluna C) não foi encontrada em '{os.path.basename(arquivo_pamela)}'.")
            return

        if "Descripción" not in df_adelino.columns:
            exibir_mensagem_erro(f"A coluna 'Descripción' não foi encontrada em '{os.path.basename(arquivo_adelino)}'. Não é possível colar os dados.")
            return
        if "Resolución" not in df_adelino.columns:
            exibir_mensagem_erro(f"A coluna 'Resolución' não foi encontrada em '{os.path.basename(arquivo_adelino)}'. Não é possível preencher esta coluna.")
            return
        if "Fecha_Finalización" not in df_adelino.columns:
            exibir_mensagem_erro(f"A coluna 'Fecha_Finalización' não foi encontrada em '{os.path.basename(arquivo_adelino)}'. Não é possível extrair a data de referência.")
            return
        if "ID_Tarea" not in df_adelino.columns:
            exibir_mensagem_erro(f"A coluna 'ID_Tarea' não foi encontrada em '{os.path.basename(arquivo_adelino)}'. Não é possível gerar a sequência de IDs.")
            return
        if "Responsable" not in df_adelino.columns:
            exibir_mensagem_erro(f"A coluna 'Responsable' não foi encontrada em '{os.path.basename(arquivo_adelino)}'. Não é possível preencher esta coluna.")
            return
        if "Estado" not in df_adelino.columns:
            exibir_mensagem_erro(f"A coluna 'Estado' não foi encontrada em '{os.path.basename(arquivo_adelino)}'. Não é possível preencher esta coluna.")
            return
        if "Fecha_Inicio" not in df_adelino.columns:
            exibir_mensagem_erro(f"A coluna 'Fecha_Inicio' não foi encontrada em '{os.path.basename(arquivo_adelino)}'. Não é possível preencher esta coluna.")
            return

        # --- Extrair a data da última linha de 'Fecha_Finalización' de adelino.xlsx e adicionar 1 dia ---
        data_minima_origem_texto = "a data atual do sistema + 1 dia"
        try:
            df_adelino['Fecha_Finalización'] = pd.to_datetime(df_adelino['Fecha_Finalización'], errors='coerce')
            ultima_data_valida_adelino_series = df_adelino['Fecha_Finalización'].dropna()
            if not ultima_data_valida_adelino_series.empty:
                ultima_data_adelino_date = ultima_data_valida_adelino_series.iloc[-1].date()
                data_minima = ultima_data_adelino_date + datetime.timedelta(days=1)
                data_minima_origem_texto = f"a última data da coluna 'Fecha_Finalización' de '{os.path.basename(arquivo_adelino)}' ({ultima_data_adelino_date.strftime('%d/%m/%Y')}) + 1 dia"
            else:
                exibir_mensagem_aviso(f"A coluna 'Fecha_Finalización' em '{os.path.basename(arquivo_adelino)}' está vazia ou não contém datas válidas. Usando a data atual do sistema + 1 dia ({ (datetime.date.today() + datetime.timedelta(days=1)).strftime('%d/%m/%Y') }) como data de referência.")
                data_minima = datetime.date.today() + datetime.timedelta(days=1)
        except Exception as e:
            exibir_mensagem_aviso(f"Erro ao extrair a última data da coluna 'Fecha_Finalización' de '{os.path.basename(arquivo_adelino)}': {e}. Usando a data atual do sistema + 1 dia ({ (datetime.date.today() + datetime.timedelta(days=1)).strftime('%d/%m/%Y') }) como data de referência.")
            data_minima = datetime.date.today() + datetime.timedelta(days=1)
        # --- FIM DA PERSONALIZAÇÃO DE DATA DE INÍCIO DE FILTRO ---

        # --- Filtrar pamela.xlsx por 'DATA DA SOLUÇÃO' a partir da 'data_minima' obtida ---
        df_pamela['DATA DA SOLUÇÃO'] = pd.to_datetime(df_pamela['DATA DA SOLUÇÃO'], errors='coerce', dayfirst=True)
        df_pamela_filtrado_data = df_pamela[
            (df_pamela['DATA DA SOLUÇÃO'].dt.date >= data_minima) & 
            (df_pamela['DATA DA SOLUÇÃO'].notna())
        ].copy()

        # --- Remover linhas que contêm "Visita Presencial" OU "Relatório Semanal de Atividades" da coluna ASSUNTO ---
        # Certifica-se de que os dados relevantes de pamela_assunto_data, pamela_data_solucao e pamela_servico_realizado são filtrados juntos
        
        # Condição para remover "Visita Presencial"
        cond_visita_presencial = df_pamela_filtrado_data["ASSUNTO"].astype(str).str.contains("Visita Presencial", case=False, na=False)
        # Condição para remover "Relatório Semanal de Atividades"
        cond_relatorio_semanal = df_pamela_filtrado_data["ASSUNTO"].astype(str).str.contains("Relatório Semanal de Atividades", case=False, na=False)

        # A máscara de filtro será VERDADEIRO para as linhas que NÃO contêm NENHUMA das condições
        filter_mask = ~(cond_visita_presencial | cond_relatorio_semanal)
        
        pamela_assunto_data = df_pamela_filtrado_data.loc[filter_mask, "ASSUNTO"]
        pamela_datas_solucao = df_pamela_filtrado_data.loc[filter_mask, "DATA DA SOLUÇÃO"]
        pamela_servico_realizado = df_pamela_filtrado_data.loc[filter_mask, "SERVIÇO REALIZADO"] 
        # --- FIM DA PERSONALIZAÇÃO POR TEXTO ---

        # --- VERIFICAÇÕES DE DADOS APÓS OS FILTROS ---
        num_new_rows = len(pamela_assunto_data)

        if num_new_rows == 0:
            exibir_mensagem_aviso(f"Após a filtragem por data (a partir de {data_minima.strftime('%d/%m/%Y')}, que é {data_minima_origem_texto}) e por texto ('Visita Presencial' e 'Relatório Semanal de Atividades'), a coluna 'ASSUNTO' em '{os.path.basename(arquivo_pamela)}' está vazia ou não contém dados válidos. Nenhuma informação será copiada.")
            return

        valid_data_count = pamela_assunto_data.dropna().astype(str).str.strip().astype(bool).sum()

        if valid_data_count == 0:
            exibir_mensagem_aviso(f"Após a filtragem por data e por texto, a coluna 'ASSUNTO' em '{os.path.basename(arquivo_pamela)}' contém {len(pamela_assunto_data)} linhas, mas todas parecem estar vazias ou conter apenas espaços em branco. Nenhuma informação significativa será copiada.")
            return
        # --- FIM DAS VERIFICAÇÕES ---
        
        # --- Gerar a sequência para a coluna 'ID_Tarea' (coluna A) ---
        last_id_tarea = ""
        last_valid_id_series = df_adelino['ID_Tarea'].dropna()
        if not last_valid_id_series.empty:
            last_id_tarea = str(last_valid_id_series.iloc[-1])
        
        start_number = 1 
        if last_id_tarea:
            match = re.search(r'ST-(\d+)', last_id_tarea)
            if match:
                try:
                    start_number = int(match.group(1)) + 1
                except ValueError:
                    exibir_mensagem_aviso(f"Não foi possível extrair o número da última 'ID_Tarea' '{last_id_tarea}'. A sequência de IDs começará de 1.")
            else:
                exibir_mensagem_aviso(f"O formato da última 'ID_Tarea' '{last_id_tarea}' não é 'ST-XXX'. A sequência de IDs começará de 1.")
        else:
             exibir_mensagem_aviso(f"A coluna 'ID_Tarea' em '{os.path.basename(arquivo_adelino)}' está vazia. A sequência de IDs começará de 1.")

        new_ids = []
        for i in range(num_new_rows):
            new_ids.append(f"ST-{start_number + i:03d}") 
        # --- FIM DA PERSONALIZAÇÃO DE ID_Tarea ---

        # Cria um DataFrame para as novas linhas, com as mesmas colunas de adelino.xlsx
        new_rows_df = pd.DataFrame(index=range(num_new_rows), columns=df_adelino.columns)
        
        # Atribui os dados gerados/filtrados
        new_rows_df['ID_Tarea'] = new_ids
        new_rows_df['Descripción'] = pamela_assunto_data.reset_index(drop=True)
        new_rows_df['Resolución'] = pamela_servico_realizado.reset_index(drop=True)
        new_rows_df['Responsable'] = "Adelino Silva" 
        new_rows_df['Estado'] = "Completado"
        
        # Preenche as colunas 'Fecha_Inicio' e 'Fecha_Finalización' com a data da Pamela
        new_rows_df['Fecha_Inicio'] = pamela_datas_solucao.dt.date.reset_index(drop=True)
        new_rows_df['Fecha_Finalización'] = pamela_datas_solucao.dt.date.reset_index(drop=True) 
        
        # Concatena o DataFrame existente com as novas linhas
        df_final_adelino = pd.concat([df_adelino, new_rows_df], ignore_index=True)

        # --- SALVAR O DATAFRAME COMPLETO SOBRESCRREVENDO O ARQUIVO ---
        try:
            df_final_adelino.to_excel(arquivo_adelino, index=False, sheet_name='Sheet1')
            
            exibir_mensagem_informacao(f"Exportação concluída com sucesso!\n{valid_data_count} registros válidos (a partir de {data_minima.strftime('%d/%m/%Y')}, que é {data_minima_origem_texto}, e excluindo 'Visita Presencial' e 'Relatório Semanal de Atividades') da coluna 'ASSUNTO' de '{os.path.basename(arquivo_pamela)}' foram copiados para a coluna 'Descripción' de '{os.path.basename(arquivo_adelino)}'.\n\nNovos IDs 'ID_Tarea' e as colunas 'Responsable', 'Estado', 'Fecha_Inicio', 'Fecha_Finalización' e 'Resolución' foram gerados/preenchidos.\n\nATENÇÃO: Este método SOBRESCREVEU o arquivo, então a formatação original (cores, fontes, etc.) pode ter sido perdida, mas garante que todos os dados foram copiados corretamente. Para que as mudanças sejam visíveis, certifique-se de que 'adelino.xlsx' estava FECHADO antes de executar a macro.")

        except Exception as e:
            exibir_mensagem_erro(f"Ocorreu um erro ao tentar sobrescrever o arquivo '{os.path.basename(arquivo_adelino)}':\n{e}\nIsso pode ocorrer se o arquivo estiver aberto ou se houver um problema de permissão.")

    except Exception as e:
        exibir_mensagem_erro(f"Ocorreu um erro durante o processamento:\n{e}\nVerifique se os arquivos Excel não estão abertos e se as colunas estão corretas.")

# --- Funções para exibir mensagens ---
def exibir_mensagem_informacao(mensagem):
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    parent_window = desktop.getCurrentFrame().getContainerWindow()
    box = toolkit.createMessageBox(parent_window, 1, 1, "Macro Python - Informação", mensagem)
    box.execute()

def exibir_mensagem_erro(mensagem):
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    parent_window = desktop.getCurrentFrame().getContainerWindow()
    box = toolkit.createMessageBox(parent_window, 2, 1, "Macro Python - Erro", mensagem)
    box.execute()

def exibir_mensagem_aviso(mensagem):
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    parent_window = desktop.getCurrentFrame().getContainerWindow()
    box = toolkit.createMessageBox(parent_window, 3, 1, "Macro Python - Aviso", mensagem)
    box.execute()
