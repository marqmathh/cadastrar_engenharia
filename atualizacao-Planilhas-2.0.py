from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import tkinter as tk
from tkinter import messagebox
import os, shutil, win32com.client, time, ctypes
from datetime import datetime
# Função para mover os arquivos para uma pasta específica
def mover_arquivos(pasta_origem, pasta_destino):
    for item in os.listdir(pasta_origem):
        item_path = os.path.join(pasta_origem, item)
        try:
            if os.path.isfile(item_path):
                shutil.move(item_path, pasta_destino)
        except Exception as e:
            print(f'Erro ao mover {item_path} para {pasta_destino}: {e}')
# Função para excluir os arquivos em uma pasta destino
def excluir_itens(pasta):
    for item in os.listdir(pasta):
        item_path = os.path.join(pasta, item)
        try:
            if os.path.isfile(item_path):
                os.remove(item_path) 
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path) 
        except Exception as e:
            print(f'Erro ao excluir {item_path}: {e}')
# Função para pressionar os botões de exportar
def exportar():
    time.sleep(30)
    export_button = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ZDBExportButton"]')))
    export_button.click()
    try: 
        embed_export_button = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="EmbedExportXLSMenuItem"]')))
        time.sleep(1)
        embed_export_button.click()
        
        final_button = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[32]/div/div[2]/div/footer/div/button[1]')))
        time.sleep(1)
        final_button.click()
        print('Planilha exportada com sucesso (modelo padrão)')
        time.sleep(1)
    except Exception:
        embed_export_button = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="SumAndPvtEmbedExportXLSMenuItem"]')))
        time.sleep(1)
        embed_export_button.click()
        
        final_button = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[34]/div/div[2]/div/footer/div/button[1]')))
        time.sleep(1)
        final_button.click()
        time.sleep(1)
        print('Planilha exportada com sucesso (modelo secundário)')
# Função para abrir uma nova guia no nomus
def mudar_Link(texto,i):
    navegador.execute_script("window.open('');")
    navegador.switch_to.window(navegador.window_handles[i])
    navegador.get(texto)
# Função para exportar as planilhas do nomus
def exporta_Nomus():
    urls = [
        ("https://reports.nomus.com.br/open-view/751489003498986462", 0),  # PC - Follow up
        ("https://reports.nomus.com.br/open-view/751489003495455916", 1),  # SC COMPLETA
        ("https://reports.nomus.com.br/open-view/751489003499869977", 2),  # Cotação
        ("https://reports.nomus.com.br/open-view/751489003499311067", 3),  # NF Tipo de movimentação
        ("https://reports.nomus.com.br/open-view/751489003497642590", 4),  # Fornecedores - TB-06
        ("https://reports.nomus.com.br/open-view/751489003385051213", 5),  # Ordens Lead
        ("https://reports.nomus.com.br/open-view/751489003418257074", 6),  # Necessidade
        ("https://reports.nomus.com.br/open-view/751489003348744953", 7),  # ID-04e
        ("https://reports.nomus.com.br/open-view/751489003473916593", 8),  # Tempo de produção
        ("https://reports.nomus.com.br/open-view/751489003396077915", 9),  # Entradas
        ("https://reports.nomus.com.br/open-view/751489003343390887", 10), # Pedido de compra
        ("https://reports.nomus.com.br/open-view/751489003375648746", 11), # Pv's
        ("https://reports.nomus.com.br/open-view/751489003343401232", 12), # SC's
        ("https://reports.nomus.com.br/open-view/751489003343402198", 13), # Entradas compras
        ("https://reports.nomus.com.br/open-view/751489003383877015", 14), # Faturamento periodo
        ("https://reports.nomus.com.br/open-view/751489003384381506", 15), # CRM
        ("https://reports.nomus.com.br/open-view/751489003381705320", 16), # Vendas
        ("https://reports.nomus.com.br/open-view/751489003396836362", 17), # Emissão de PV
        ("https://reports.nomus.com.br/open-view/751489003384833393", 18), # Status do PV
        ("https://reports.nomus.com.br/open-view/751489003501209517", 19), # Familia
        ("https://reports.nomus.com.br/open-view/751489003299922759", 20), # CONTROLE DA PRODUÇÃO
        ("https://reports.nomus.com.br/open-view/751489003399788047", 21), # ENTRADAS (LIVRO)
        ("https://reports.nomus.com.br/open-view/751489003415814176", 22), # Nfs emitidas
        ("https://reports.nomus.com.br/open-view/751489003399436517", 23), # Saída
        ("https://reports.nomus.com.br/open-view/751489003406446494", 24), # Dw Compras
        ("https://reports.nomus.com.br/open-view/751489003342815143", 25), # Modelo NF's
        ("https://reports.nomus.com.br/open-view/751489003415906765", 26), # Grupos
        ("https://reports.nomus.com.br/open-view/751489003432511966", 27), # movimentação - pvs
        ("https://reports.nomus.com.br/open-view/751489003502954549", 28), # Descrições
        ("https://reports.nomus.com.br/open-view/751489003502944448", 29), # Estoque - tempo real
        ("https://reports.nomus.com.br/open-view/751489003521923921", 30), # qtde_estoque
        ("https://reports.nomus.com.br/open-view/751489003423136168", 31), # ABC Vendas
    ]
    navegador.get(urls[0][0])
    time.sleep(10)
    exportar() 
    for url, index in urls[1:]:
        mudar_Link(url, index)
        exportar() 
    time.sleep(60)
    navegador.quit()
# Função para salvar os arquivos do download em outro lugar antes de começar a atualização
def salva_Arquivos():
    pasta_destino = r'Z:\PUBLICO\## Salva arquivos'
    pasta_downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
    mover_arquivos(pasta_downloads, pasta_destino)
# Função para mover os arquivos baixados para a pasta do servidor
def move_Arquivos():
    pasta_destino = r'Z:\PUBLICO\### Banco de dados'
    pasta_downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
    excluir_itens(pasta_destino)
    mover_arquivos(pasta_downloads, pasta_destino)
# Função para voltar os arquivos para a pasta Download
def voltar_Arquivos():
    pasta_destino = r'Z:\PUBLICO\## Salva arquivos'
    pasta_downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
    mover_arquivos(pasta_destino, pasta_downloads)
# Função para atulizar as planilhas padrões
def atualizar_Planilhas(local_Planilha):
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = True
    wb = excel.Workbooks.Open(local_Planilha)
    wb.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone() 
    wb.Close(SaveChanges=True)
    print(f"Planilha atualizada com sucesso: {local_Planilha}")
    excel.Quit()
# Função de atualizar as planilhas que não possuimos acesso
def atualizar_Planilhas_especiais(local_Planilha):
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = True
    wb = excel.Workbooks.Open(local_Planilha)
    wb.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone() 
    data_atual = datetime.now().strftime("%Y-%m-%d")
    nome_arquivo, extensao = os.path.splitext(local_Planilha)
    novo_nome = f"{nome_arquivo} - {data_atual}{extensao}"
    wb.SaveAs(novo_nome)
    wb.Close(SaveChanges=True)
    excel.Quit()
    
    print(f"Planilha atualizada e salva como: {novo_nome}")
# Atualiza as planilhas
def atualizar_Excel():
    planilhas = [
        # 'Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Lead time - Familia.xlsx',
        'Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Monitoramento de estoque.xlsx',
        'Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\OP_porcentagem.xlsx',
        'Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Vendas.xlsx',
        'Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Indicadores.xlsx',
        'Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Setores de estoque.xlsx'
    ]
    planilhas_especiais = [
        # 'Z:\\ISO 9000 - SGQ\\6-PROCESSO SUPRIMENTOS\\REGISTROS\\TB-06_AvalProvedoresExternos-Rev05.xlsx',
        'Z:\\ISO 9000 - SGQ\\9 - PPCP\\.PCP\\Controle\\Controle.xlsm',
        'Z:\\PUBLICO\\Araujo\\Planilhas-Indicadores\\Gráficos-IDs-PCP-2024 - IDs 04a 04b 13b.xlsx',
        'Z:\\PUBLICO\\Araujo\\Planilhas-Indicadores\\Gráfico-ID-Produção-2024 - ID-02-01-10-2024.xlsx',
        'Z:\\PUBLICO\\Araujo\\Planilhas-Indicadores\\Graficos-IDs-Compras-2024 - IDs 09 13a 13c.xlsx',
        'Z:\\PUBLICO\\Araujo\\Planilhas-Indicadores\\Graficos-IDs-Engenharia-2024 - IDs10a 10b 11 12 e 14.xlsx',
        'Z:\\ISO 9000 - SGQ\\5-PROCESSO PROJETOS\\REGISTROS\\Controle - AT.xlsm'
    ]
    for planilha in planilhas:
        atualizar_Planilhas(planilha)
    for planilha in planilhas_especiais:
        atualizar_Planilhas_especiais(planilha)
# Exibe as intruções
def instrucoes():
     messagebox.showinfo("Instrução de Atualização de Indicadores e Planilhas", "Olá\n\nEste programa é responsável pela atualização dos indicadores e das planilhas padrão.\n\nPlanilhas que são atualizadas automaticamente:\nLead time\nMonitoramento de estoque\nSetores de estoque\nTB-06\nControle\nControle AT\nIndicadores salvos no Publico/Araujo\n\nNo entanto, existem três planilhas que, devido à interdependência de suas atualizações, não podem ser automatizadas. Caso seja necessário utilizar a Curva ABC, será preciso atualizar manualmente as seguintes planilhas:\n\n'Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Curva ABC - 2021 - 2024.xlsx'\n\n'Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Curva ABC - fornecedores.xlsx'\n\n'Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Curva ABC - Ordens.xlsx'\n\nPara atualizar, basta entrar nela e pressionar em 'Atualizar Tudo'")
# Exibe a mensagem de inicialização do programa
def inicializacao():
    messagebox.showinfo("Inicialização", "A atividade terá início em breve. Por favor, não utilize o computador até que a atualização seja concluída.")
# Exibe a mensagem de finalização do programa
def finalizacao():
    messagebox.showinfo("Finalização", "Atividade concluída com sucesso!")
# Atualiza tudo
def atualizacao_geral():
    inicializacao()
    time.sleep(0.5)
    salva_Arquivos()
    time.sleep(0.5)
    exporta_Nomus()
    time.sleep(5)
    move_Arquivos()
    time.sleep(0.5)
    voltar_Arquivos()
    time.sleep(0.5)
    atualizar_Excel()
    time.sleep(0.5)
# Atualiza tudo e avisa
def atualizacao_nivel_1():
    atualizacao_geral()
    finalizacao()
# Atualiza tudo e bloqueia
def atualizacao_nivel_2():
    atualizacao_geral()
    ctypes.windll.user32.LockWorkStation()
# Atualiza tudo e suspende
def atualizacao_nivel_3():
    atualizacao_geral()
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
# Atualiza tudo e desliga
def atualizacao_nivel_4():
    atualizacao_geral()
    os.system("shutdown /s /t 10")
# Cria a interface da aplicação
def criar_interface():
    janela = tk.Tk()
    janela.title("Automação de Processos")
    janela.geometry("320x210")
    janela.config(bg="lightblue")  
    tk.Button(janela, text="Instruções", command=instrucoes, width=30, bg="blue", fg="white").pack(pady=(20,5))
    tk.Button(janela, text="Atualizar tudo", command=atualizacao_nivel_1, width=30, bg="purple1", fg="white").pack(pady=5)
    tk.Button(janela, text="Atualizar tudo e bloquear", command=atualizacao_nivel_2, width=30, bg="purple2", fg="white").pack(pady=5)
    tk.Button(janela, text="Atualizar tudo e suspender", command=atualizacao_nivel_3, width=30, bg="purple3", fg="white").pack(pady=5)
    tk.Button(janela, text="Atualizar tudo e desligar", command=atualizacao_nivel_4, width=30, bg="purple4", fg="white").pack(pady=5)
    janela.mainloop()
# Rodar a aplicação
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)
criar_interface()
