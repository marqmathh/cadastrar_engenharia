from selenium import webdriver
import pyautogui
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
import win32com.client
import os
from pywinauto.application import Application
import ctypes
import tkinter as tk
from tkinter import messagebox

pyautogui.PAUSE = 1

def caminho_Download():
    pyautogui.press('winleft')
    pyautogui.write('Downloads')
    pyautogui.press('enter')

def caminho_publico():
    pyautogui.press('winleft')
    pyautogui.write('Z:\PUBLICO')
    pyautogui.press('enter')
    pyautogui.press('pgup')
    pyautogui.press('enter')

def caminho_Iso():
    pyautogui.press('winleft')
    pyautogui.write('Z:\PUBLICO')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('down')
    pyautogui.press('enter')

def exportar():
    time.sleep(30)
    try:
        
        export_button = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="ZDBExportButton"]'))
        )
        export_button.click()

        embed_export_button = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="EmbedExportXLSMenuItem"]'))
        )
        time.sleep(1)
        embed_export_button.click()
        print("Botão 'EmbedExportXLSMenuItem' clicado com sucesso!")
        
        final_button = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[32]/div/div[2]/div/footer/div/button[1]'))
        )
        time.sleep(1)
        final_button.click()

        time.sleep(1)

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

def exportar_Diff():
    time.sleep(30)
    try:
        
        export_button = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="ZDBExportButton"]'))
        )
        export_button.click()

        embed_export_button = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="SumAndPvtEmbedExportXLSMenuItem"]'))
        )
        time.sleep(1)
        embed_export_button.click()
        print("Botão 'EmbedExportXLSMenuItem' clicado com sucesso!")
        
        final_button = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[34]/div/div[2]/div/footer/div/button[1]'))
        )
        time.sleep(1)
        final_button.click()

        time.sleep(1)

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

def mudar_Link(texto,i):
    navegador.execute_script("window.open('');")
    navegador.switch_to.window(navegador.window_handles[i])
    navegador.get(texto)

def exporta_Nomus():
    navegador.get('https://reports.nomus.com.br/open-view/751489003498986462') # PC - Follow up
    time.sleep(10)
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003495455916',1) # SC COMPLETA
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003499869977',2) # Cotação
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003499311067',3) # NF Tipo de movimentação
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003497642590',4) # Fornecedores - TB-06
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003385051213',5) # Ordens Lead
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003418257074',6) # Necessidade
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003348744953',7) # ID-04e
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003473916593',8) # Tempo de produção
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003396077915',9) # Entradas
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003343390887',10) # Pedido de compra
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003375648746',11) # Pv's
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003343401232',12) # SC's
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003343402198',13) # Entradas compras
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003383877015',14) # Faturamento periodo
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003384381506',15) # CRM
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003381705320',16) # Vendas
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003396836362',17) # Emissão de PV
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003384833393',18) # Status do PV
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003501209517',19) # Familia
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003299922759',20) # CONTROLE DA PRODUÇÃO
    exportar_Diff()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003399788047',21) # ENTRADAS (LIVRO)
    exportar_Diff()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003415814176',22) # Nfs emitidas
    exportar_Diff()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003399436517',23) # Saída 
    exportar_Diff()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003406446494',24) # Dw Compras
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003342815143',25) # Modelo NF's 
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003415906765',26) # Grupos
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003432511966',27) # movimentação - pvs
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003502954549',28) # Descrições
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003502944448',29) # Estoque - tempo real
    exportar_Diff()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003521923921',30) # qtde_estoque
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003423136168',31) # ABC Vendas
    exportar()
    time.sleep(60)
    navegador.quit()

def salva_Arquivos():
    pyautogui.press('winleft')
    pyautogui.write('Downloads')
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','a')
    pyautogui.hotkey('ctrl','x')
    pyautogui.press('winleft')
    pyautogui.write('Z:\PUBLICO')
    pyautogui.press('enter')
    pyautogui.press('pgup')
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','v')
    time.sleep(10)
    pyautogui.press('esc')
    pyautogui.hotkey('ctrl','w')
    pyautogui.hotkey('ctrl','w')
    time.sleep(0.5)

def move_Arquivos():
    caminho_Iso()
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','a')
    pyautogui.press('delete')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(5)
    caminho_Download()
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','a')
    pyautogui.hotkey('ctrl','x')
    pyautogui.hotkey('ctrl','w')
    pyautogui.hotkey('ctrl','v')
    time.sleep(10)
    pyautogui.hotkey('ctrl','w')
    time.sleep(0.5)

def voltar_Arquivos():
    caminho_publico()
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','a')
    pyautogui.hotkey('ctrl','x')
    caminho_Download()
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','v')
    time.sleep(10)
    pyautogui.press('esc')
    pyautogui.hotkey('ctrl','w')
    pyautogui.hotkey('ctrl','w')
    time.sleep(0.5)

def titulo():
    pyautogui.alert("""O código vai começar.

Favor não útilizar o PC até a finalização.""")

def finalizacao():
    pyautogui.alert("""O código foi finalizado.

Não esqueça de pressionar o botão de atualizar nas suas planilhas.

Obrigado por aguardar. """)

def atualizar_Planilhas(local_Planilha):
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = True

    try:
        wb = excel.Workbooks.Open(local_Planilha)
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        wb.Close(SaveChanges=True)
        
    except Exception as e:
        print(f"Erro ao abrir ou atualizar a planilha: {e}")

    finally:
        excel.Quit()

def atualizar_Excel():
    # atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\9 - PPCP\\Lead time\\Banco de dados\\Banco de dados - Lead.xlsx')
    # atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\9 - PPCP\\Lead time\\Lead time - Familia.xlsx')
    # atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\12-SISTEMA\Sistema\\planilhas\\Lead time - Familia.xlsx')
    # atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\6-PROCESSO SUPRIMENTOS\\REGISTROS\\Controle de inventários.xlsx')
    # atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\6-PROCESSO SUPRIMENTOS\\REGISTROS\\TB-06_AvalProvedoresExternos-Rev05.xlsx')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\9 - PPCP\\.PCP\\Controle\\Controle.xlsm')
    # atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\5-PROCESSO PROJETOS\\REGISTROS\\Controle - AT.xlsm')
    atualizar_Planilhas('Z:\\PUBLICO\Araujo\\Planilhas-Indicadores\\SCS vs CARTEIRAS.xlsx')
    atualizar_Planilhas('Z:\\PUBLICO\\Araujo\\Planilhas-Indicadores\\Gráficos-IDs-PCP-2024 - IDs 04a 04b 13b.xlsx')
    atualizar_Planilhas('Z:\\PUBLICO\\Araujo\\Planilhas-Indicadores\\Gráfico-ID-Produção-2024 - ID-02-01-10-2024.xlsx')
    atualizar_Planilhas('Z:\\PUBLICO\\Araujo\\Planilhas-Indicadores\\Graficos-IDs-Compras-2024 - IDs 09 13a 13c.xlsx')
    # atualizar_Planilhas('Z:\\PUBLICO\\Araujo\\Planilhas-Indicadores\\Graficos-IDs-Engenharia-2024 - IDs10a 10b 11 12 e 14.xlsx')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Lead time - Familia.xlsx')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Monitoramento de estoque.xlsx')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\OP_porcentagem.xlsx')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Vendas.xlsx')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Indicadores.xlsx')
    
def atualizar_aplicativos():
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Setores de estoque.xlsx')
    # atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Curva ABC - 2021 - 2024.xlsx')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Curva ABC - fornecedores.xlsx')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\planilhas\\Curva ABC - Ordens.xlsx')

def atualizar_engenharia():
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\5-PROCESSO PROJETOS\\REGISTROS\\Controle - AT.xlsm')
    atualizar_Planilhas('Z:\\PUBLICO\\Araujo\\Planilhas-Indicadores\\Graficos-IDs-Engenharia-2024 - IDs10a 10b 11 12 e 14.xlsx')

def atualizar_compras():
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\6-PROCESSO SUPRIMENTOS\\REGISTROS\\TB-06_AvalProvedoresExternos-Rev05.xlsx')

def bloquear_PC():
    ctypes.windll.user32.LockWorkStation()

def desligar_PC():
    os.system("shutdown /s /t 10")

def suspender_PC():
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

def atualizar_e_avisar():
    salva_Arquivos()
    time.sleep(10)
    exporta_Nomus()
    time.sleep(10)
    move_Arquivos()
    time.sleep(10)
    voltar_Arquivos()
    time.sleep(10)
    atualizar_Excel()
    time.sleep(10)
    finalizacao()

def atualizar_e_bloquear():
    salva_Arquivos()
    time.sleep(10)
    exporta_Nomus()
    time.sleep(10)
    move_Arquivos()
    time.sleep(10)
    voltar_Arquivos()
    time.sleep(10)
    atualizar_Excel()
    time.sleep(10)
    bloquear_PC()

def atualizar_e_suspender():
    salva_Arquivos()
    time.sleep(10)
    exporta_Nomus()
    time.sleep(10)
    move_Arquivos()
    time.sleep(10)
    voltar_Arquivos()
    time.sleep(10)
    atualizar_Excel()
    time.sleep(10)
    suspender_PC()

def atualizar_e_desligar():
    salva_Arquivos()
    time.sleep(10)
    exporta_Nomus()
    time.sleep(10)
    move_Arquivos()
    time.sleep(10)
    voltar_Arquivos()
    time.sleep(10)
    atualizar_Excel()
    time.sleep(10)
    desligar_PC()

def readme():
    pyautogui.alert("""
As primeiras opções que começam com "Atualizar" ela atulualiza e faz algo com o seu computador, ideal para atualizar no final do expediente ou no horario de almoço
As sequencias são passo a passo caso a pessoa queira executar tudo um por um""")
    
def criar_interface():
    janela = tk.Tk()
    janela.title("Automação de Processos")
    janela.geometry("400x600")
      
    tk.Label(janela, text="Selecione a ação desejada:", font=("Arial", 14)).pack(pady=10)

    tk.Button(janela, text="0 - Instruções", command=readme, width=30, bg="blue", fg="white").pack(pady=5)
    tk.Button(janela, text="Atualizar e Avisar", command=atualizar_e_avisar, width=30, bg="green", fg="white").pack(pady=5)
    tk.Button(janela, text="Atualizar e Bloquear", command=atualizar_e_bloquear, width=30, bg="green", fg="white").pack(pady=5)
    tk.Button(janela, text="Atualizar e Suspender", command=atualizar_e_suspender, width=30, bg="green", fg="white").pack(pady=5)
    tk.Button(janela, text="Atualizar e Desligar", command=atualizar_e_desligar, width=30, bg="green", fg="white").pack(pady=5)
    tk.Button(janela, text="1° - Salva os arquivos", command=salva_Arquivos, width=30, bg="yellow", fg="black").pack(pady=5)
    tk.Button(janela, text="2° - Baixar planilhas", command=exporta_Nomus, width=30, bg="yellow", fg="black").pack(pady=5)
    tk.Button(janela, text="3° - Mover arquivos", command=move_Arquivos, width=30, bg="yellow", fg="black").pack(pady=5)
    tk.Button(janela, text="4° - Voltar arquivos", command=voltar_Arquivos, width=30, bg="yellow", fg="black").pack(pady=5)
    tk.Button(janela, text="5° - Atualizar planilhas/indicadores", command=atualizar_Excel, width=30, bg="yellow", fg="black").pack(pady=5)
    tk.Button(janela, text="6° - Atualizar engenharia", command=atualizar_engenharia, width=30, bg="yellow", fg="black").pack(pady=5)
    tk.Button(janela, text="7° - Atualizar TB-06", command=atualizar_compras, width=30, bg="yellow", fg="black").pack(pady=5)
    tk.Button(janela, text="8° - Atualizar Aplicativos", command=atualizar_aplicativos, width=30, bg="yellow", fg="black").pack(pady=5)

    tk.Button(janela, text="Sair", command=janela.quit, bg="red", fg="white", width=30).pack(pady=10)

    janela.mainloop()

# titulo()

# salva_Arquivos()

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

criar_interface()


navegador.quit()


# exporta_Nomus()

# move_Arquivos()

# voltar_Arquivos()

# atualizar_Excel()

# finalizacao()

# bloquear_PC()

# desligar_PC()

# suspender_PC()