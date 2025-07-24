import uno
import os
import datetime
import pandas as pd
from pandas import ExcelWriter

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
            exibir_mensagem_erro(f"A coluna 'DATA DA SOLUÇÃO' (Coluna A) não foi encontrada em '{os.path.basename(arquivo_pamela)}'. Não é possível filtrar por data.")
            return

        if "Descripción" not in df_adelino.columns:
            exibir_mensagem_erro(f"A coluna 'Descripción' não foi encontrada em '{os.path.basename(arquivo_adelino)}'. Não é possível colar os dados.")
            return

        # --- NOVA PERSONALIZAÇÃO: Filtrar por 'DATA DA SOLUÇÃO' a partir da DATA ATUAL do sistema ---
        # Converte a coluna 'DATA DA SOLUÇÃO' para datetime, tratando erros
        df_pamela['DATA DA SOLUÇÃO'] = pd.to_datetime(df_pamela['DATA DA SOLUÇÃO'], errors='coerce', dayfirst=True)
        
        # Define a data mínima para o filtro como a data atual do sistema
        data_minima = datetime.date.today() # Pega a data de hoje

        # Filtra as linhas onde 'DATA DA SOLUÇÃO' é maior ou igual à data_minima
        # E também garante que a data não seja NaT (Not a Time)
        df_pamela_filtrado_data = df_pamela[
            (df_pamela['DATA DA SOLUÇÃO'].dt.date >= data_minima) & 
            (df_pamela['DATA DA SOLUÇÃO'].notna())
        ].copy() # .copy() para evitar SettingWithCopyWarning

        # Extrai a coluna 'ASSUNTO' do DataFrame já filtrado pela data
        pamela_assunto_data = df_pamela_filtrado_data["ASSUNTO"]
        # --- FIM DA PERSONALIZAÃO POR DATA ---

        # --- PERSONALIZAÇÃO EXISTENTE: Remover linhas que contêm "Visita Presencial" ---
        # Converte para string e filtra, ignorando maiúsculas/minúsculas
        # O filtro de data já foi aplicado, agora aplicamos o filtro de texto
        pamela_assunto_data = pamela_assunto_data.astype(str)
        pamela_assunto_data = pamela_assunto_data[~pamela_assunto_data.str.contains("Visita Presencial", case=False, na=False)]
        # --- FIM DA PERSONALIZAÇÃO POR TEXTO ---

        # --- VERIFICAÇÕES DE DADOS APÓS OS FILTROS ---
        if pamela_assunto_data.empty:
            exibir_mensagem_aviso(f"Após a filtragem por data (a partir da data atual do sistema: {data_minima.strftime('%d/%m/%Y')}) e por texto ('Visita Presencial'), a coluna 'ASSUNTO' em '{os.path.basename(arquivo_pamela)}' está vazia ou não contém dados válidos. Nenhuma informação será copiada.")
            return

        valid_data_count = pamela_assunto_data.dropna().astype(str).str.strip().astype(bool).sum()

        if valid_data_count == 0:
            exibir_mensagem_aviso(f"Após a filtragem por data e por texto, a coluna 'ASSUNTO' em '{os.path.basename(arquivo_pamela)}' contém {len(pamela_assunto_data)} linhas, mas todas parecem estar vazias ou conter apenas espaços em branco. Nenhuma informação significativa será copiada.")
            return
        # --- FIM DAS VERIFICAÇÕES ---
        
        # Cria um DataFrame para as novas linhas, com as mesmas colunas de adelino.xlsx
        num_new_rows = len(pamela_assunto_data)
        new_rows_df = pd.DataFrame(index=range(num_new_rows), columns=df_adelino.columns)
        
        # Atribui os dados da coluna 'ASSUNTO' de Pamela para a coluna 'Descripción'
        # Usamos reset_index(drop=True) para alinhar corretamente os índices
        new_rows_df['Descripción'] = pamela_assunto_data.reset_index(drop=True)
        
        # Concatena o DataFrame existente com as novas linhas
        # `ignore_index=True` garante que o índice seja redefinido para o DataFrame final
        df_final_adelino = pd.concat([df_adelino, new_rows_df], ignore_index=True)

        # --- SALVAR O DATAFRAME COMPLETO SOBRESCRREVENDO O ARQUIVO ---
        try:
            df_final_adelino.to_excel(arquivo_adelino, index=False, sheet_name='Sheet1')
            
            exibir_mensagem_informacao(f"Exportação concluída com sucesso!\n{valid_data_count} registros válidos (a partir da data atual do sistema: {data_minima.strftime('%d/%m/%Y')} e excluindo 'Visita Presencial') da coluna 'ASSUNTO' de '{os.path.basename(arquivo_pamela)}' foram copiados para a coluna 'Descripción' de '{os.path.basename(arquivo_adelino)}'.\n\nATENÇÃO: Este método SOBRESCREVEU o arquivo, então a formatação original (cores, fontes, etc.) pode ter sido perdida, mas garante que todos os dados foram copiados corretamente. Para que as mudanças sejam visíveis, certifique-se de que 'adelino.xlsx' estava FECHADO antes de executar a macro.")

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
