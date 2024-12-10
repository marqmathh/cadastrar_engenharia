from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import time
import pandas as pd
import os
import win32com.client

def extracao_lm_excel():
    pasta_saida = r'Z:\PUBLICO\Araujo\web_itens_lm'

    def entrar_lm_e_enviar_para_o_excel():
        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico) 
        usuario = entry_usuario.get() 
        senha = entry_senha.get()
        cod = entry_codigo.get()
        data_atual = datetime.now().strftime("%Y-%m-%d")

        def login(usuario, senha):
            tela_usuario = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="campologin"]')))
            tela_usuario.send_keys(usuario)
            tela_senha = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="login_form"]/div/main/div/section[3]/div[2]/input')))
            tela_senha.send_keys(senha)    
            tela_senha.send_keys(Keys.RETURN)

        def entrar_nomus():
            navegador.get('https://tspro.nomus.com.br/tspro/Login.do?metodo=PreLogin') # Tela inicial
            login(usuario, senha)

        def tela_produtos():
            navegador.get('https://tspro.nomus.com.br/tspro/Produto.do?metodo=Pesquisar') # Produtos

        def entra_produto(cod):
            produto_01 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@id, 'divMinimizavel')]/table/tbody/tr[2]/td[1]/input"))) # Código do produto
            produto_01.send_keys(cod)

            botao_buscar = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="botao_pesquisar"]'))) # Buscar produtos
            botao_buscar.click()

            botao_buscar = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, f"//*[text()='{cod}']"))) #Seleciona o elemento
            botao_buscar.click()

            xpaths = ['//*[@id="produtoAtivoAguardandoLiberacao_itemSubMenu_acessarListaMateriais"]', '//*[@id="produtoAtivoLiberado_itemSubMenu_acessarListaMateriais"]']
            for xpath in xpaths:
                try:
                    botao_entrar_lm = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    botao_entrar_lm.click()
                    break
                except Exception:
                    pass    

        def abrir_lm():
            botao_lm = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="botao_Funcoes_especiais"]'))) # Buscar produtos
            botao_lm.click()

            time.sleep(1)

            botao_lm2 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="botao_carregar_todos_niveis_lista"]'))) # Buscar produtos
            botao_lm2.click()

            time.sleep(1)

            # Espera até que a tabela seja carregada
            table_xpath = '/html/body/div[2]/div[4]/div/form/div[2]/table'
            table = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, table_xpath)))

            # Extrai os dados da tabela
            rows = table.find_elements(By.TAG_NAME, 'tr')
            data = []
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, 'td')
                row_data = [col.text for col in cols]
                data.append(row_data)

            # Cria um DataFrame com os dados
            df = pd.DataFrame(data)
            df.drop(index=df.index[:130], inplace=True)
            df.drop(df.columns[[0, 1]], axis=1, inplace=True)
            df.dropna(inplace=False)

            # Exporta os dados para um arquivo Excel
            df.to_excel(os.path.join(pasta_saida, f'{cod} - {data_atual}.xlsx'), index=False)

            # Fecha o navegador
            navegador.quit()

        def planilha():
            caminho_arquivo = r"Z:\ISO 9000 - SGQ\9 - PPCP\.PCP\Extrair LM\Extração.xlsx"
            excel = win32com.client.Dispatch("Excel.Application")
            workbook = excel.Workbooks.Open(caminho_arquivo)
            workbook.RefreshAll()
            time.sleep(1)
            nome_da_aba = "Consulta PAI"
            try:
                sheet = workbook.Worksheets(nome_da_aba)
            except Exception as e:
                print(f"Erro: A aba '{nome_da_aba}' não foi encontrada.")
                excel.Quit()
                exit()
            sheet.Range("F2").Value = cod
            excel.Visible = True
        entrar_nomus()
        tela_produtos()
        entra_produto(cod)
        abrir_lm()
        planilha()

    app = tk.Tk()
    app.title("Engenharia")
    app.geometry("400x400")
    app.config(bg="lightblue")  

    tk.Label(app, text="Selecione a opção desejada: ", font=("Arial", 14), bg="lightblue").pack(pady=5)
    label_codigo = tk.Label(app, text="Código Pai:", bg="lightblue")
    label_codigo.pack(padx=5, pady=5)

    entry_codigo = tk.Entry(app, bg="white", bd=2, relief="solid", justify="center")
    entry_codigo.pack(padx=5, pady=5)

    label_usuario = tk.Label(app, text="Usuário:", bg="lightblue")
    label_usuario.pack(padx=5, pady=5)

    entry_usuario = tk.Entry(app, bg="white", bd=2, relief="solid", justify="center")
    entry_usuario.pack(padx=5, pady=5)

    label_senha = tk.Label(app, text="Senha:",  bg="lightblue")
    label_senha.pack(padx=5, pady=5)

    entry_senha = tk.Entry(app, show="*", bg="white", bd=2, relief="solid", justify="center")
    entry_senha.pack(padx=5, pady=5)

    tk.Button(app, text="Extrair Lista", command=entrar_lm_e_enviar_para_o_excel, width=20, bg="gold", fg="black").pack(pady=10)


    app.mainloop()
