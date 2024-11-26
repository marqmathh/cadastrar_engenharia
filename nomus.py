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



def criar_janela_cadastro_produtos():
    arquivo_excel = r'Z:\PUBLICO\Araujo\Cadastros\Cadastros.xlsx'
    df = pd.read_excel(arquivo_excel, sheet_name='Cadastro')

    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico)

    usuario = entry_usuario.get() 
    senha = entry_senha.get()
    origem = '0 - Nacional (exceto as indicadas nos códigos de 3 a 5)'
    hoje = datetime.today()
    dia = hoje.strftime("%d")
    mes = hoje.strftime("%m")
    ano = hoje.year
    dia_atual = f'{dia}{mes}{ano}'
    empresa_tspro = '//*[@id="empresa_1"]'
    empresa_phd = '//*[@id="empresa_3"]'
    estoque_processo = '//*[@id="setor_2"]'
    estoque_em_terceiros = '//*[@id="setor_5"]'
    esoque_retorno_terceiros = '//*[@id="setor_7"]'
    estoque_de_clientes = '//*[@id="setor_8"]'
    estoque_almoxarifado_central = '//*[@id="setor_9"]'
    estoque_expedicao_acabados = '//*[@id="setor_10"]'
    ativo_imobilizados = '//*[@id="setor_11"]'
    estoque_almoxarifado_pre_producao = '//*[@id="setor_16"]'
    estoque_uso_e_consumo = '//*[@id="setor_17"]'
    almoxarifado_phd = '//*[@id="setor_13"]'
    est_processo_phd = '//*[@id="setor_14"]'
    acabados_phd = '//*[@id="setor_15"]'

    index = 0

    def login(usuario, senha):
            tela_usuario = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="campologin"]')))
            tela_usuario.send_keys(usuario)
            tela_senha = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="login_form"]/div/main/div/section[3]/div[2]/input')))
            time.sleep(1)
            tela_senha.send_keys(senha)    
            tela_senha.send_keys(Keys.RETURN)

    def entrar_nomus():
        navegador.get('https://tspro.nomus.com.br/tspro/Login.do?metodo=PreLogin') # Tela inicial
        login(usuario, senha)

    def tela_produtos():
        navegador.get('https://tspro.nomus.com.br/tspro/Produto.do?metodo=Pesquisar') # Produtos

    def criar_produtos():
        criar_produtos_01 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="botao_criar_produto"]'))) # Criar produtos
        criar_produtos_01.click()

    def cadastrar_aba_inicial(cod, descricao, und, tipo, grupo, metodo):
        criar_produtos_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="nome"]'))) # Código do produto
        criar_produtos_02.send_keys(cod)

        criar_produtos_03 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//table/tbody/tr[7]/td[2]/textarea'))) # Descrição do produto
        criar_produtos_03.send_keys(descricao)


        unidade_medida_select = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idUnidadeMedida"]'))) # Unidade de medida
        select_unidade_medida = Select(unidade_medida_select)
        select_unidade_medida.select_by_visible_text(und)

        tipo_produto = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tipoProduto_id"]'))) # Tipo de produto
        select_tipo_produto = Select(tipo_produto)
        select_tipo_produto.select_by_visible_text(tipo)

        grupo_produto = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="nomeGrupoProduto_id_select"]'))) # Grupo de produto
        select_grupo_produto = Select(grupo_produto)
        select_grupo_produto.select_by_visible_text(grupo)

        metodo_produto = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idPadraoSuprimento"]'))) # Método de suprimento
        select_metodo_produto = Select(metodo_produto)
        select_metodo_produto.select_by_visible_text(metodo)

    def cadastrar_empresas():
        empresa = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ui-id-28"]'))) # Empresas
        empresa.click()

        if tipo == 'EMBALAGENS':
            empresa_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_uso_e_consumo))) # Seleciona o estoque
            empresa_02.click()

        if tipo == 'INSUMOS E MATERIAS PRIMAS':
            empresa_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_processo))) # Seleciona o estoque
            empresa_02.click()

            empresa_03 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_em_terceiros))) # Seleciona o estoque
            empresa_03.click()

            empresa_04 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, esoque_retorno_terceiros))) # Seleciona o estoque
            empresa_04.click()

            empresa_05 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_almoxarifado_central))) # Seleciona o estoque
            empresa_05.click()

            empresa_06 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_almoxarifado_pre_producao))) # Seleciona o estoque
            empresa_06.click()

        if tipo == 'MATERIAL DE TERCEIROS':
            empresa_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_almoxarifado_pre_producao))) # Seleciona o estoque
            empresa_02.click()

            empresa_03 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_de_clientes))) # Seleciona o estoque
            empresa_03.click()

        if tipo == 'MATERIAL DE USO E CONSUMO':
            empresa_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_uso_e_consumo))) # Seleciona o estoque
            empresa_02.click()

        if tipo == 'PRODUTO ACABADO':
            empresa_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_expedicao_acabados))) # Seleciona o estoque
            empresa_02.click()

        if tipo == 'PRODUTO SEMI ACABADO':
            empresa_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_processo))) # Seleciona o estoque
            empresa_02.click()

            empresa_03 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_almoxarifado_central))) # Seleciona o estoque
            empresa_03.click()

            empresa_04 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_almoxarifado_pre_producao))) # Seleciona o estoque
            empresa_04.click()
        
        if tipo == 'SERVICOS':
            empresa_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_processo))) # Seleciona o estoque
            empresa_02.click()

            empresa_03 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, esoque_retorno_terceiros))) # Seleciona o estoque
            empresa_03.click()

        if tipo == 'USO E CONSUMO COM ESTOQUE':
            empresa_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, estoque_uso_e_consumo))) # Seleciona o estoque
            empresa_02.click()  

        if tipo == 'ATIVO IMOBILIZADO':
            empresa_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, ativo_imobilizados))) # Seleciona o estoque
            empresa_02.click()

    def cadastrar_fiscal(descricao_fiscal):
        fiscal = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ui-id-29"]'))) # Fiscal
        fiscal.click()

        fiscal_01 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="modificarDescricaoNFe"]'))) # Fiscal
        fiscal_01.click()

        fiscal_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="descricaoNFe"]'))) # Descrição fiscal
        fiscal_02.send_keys(descricao_fiscal)
        
        fiscal_04_select = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//table/tbody/tr[9]/td[2]/select'))) # origem do produto
        select_fiscal_04 = Select(fiscal_04_select)
        select_fiscal_04.select_by_visible_text(origem)

    def cadastrar_pcp(pcp_ressuprimento):
        pcp_01 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ui-id-34"]'))) # PCP
        pcp_01.click()

        pcp_01_select = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tr_politica_ressuprimento"]/td[2]/select'))) # PCP ressuprimento
        select_pcp_01 = Select(pcp_01_select)
        select_pcp_01.select_by_visible_text(pcp_ressuprimento)

    def cadastrar_mrp():
        mrp = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ui-id-35"]'))) # MRP
        mrp.click()

        mrp_01 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tr_lote_multiplo"]/td[2]/input'))) # lote multiplo
        mrp_01.send_keys(0)

        mrp_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tr_lote_minimo"]/td[2]/input'))) # lote minimo
        mrp_02.send_keys(0)

        mrp_03 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tr_lote_maximo"]/td[2]/input'))) # lote maximo
        mrp_03.send_keys(0)

        mrp_04 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tr_estoque_de_seguranca"]/td[2]/input'))) # estoque de seguraça
        mrp_04.send_keys(0)

        mrp_05 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tr_estoque_maximo"]/td[2]/input'))) # estoque maximo
        mrp_05.send_keys(0)

    def salvar_produto():
        salvar = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="botao_salvar"]'))) # Salvar cadastro
        salvar.click()

    def cadastrar_custo(custo_padrao, dia_atual):
        if metodo == 'Comprado':
            custo = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ui-id-37"]'))) # custo
            custo.click()

            custo_01 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="custoPadraoCompra_id"]'))) # custo padrao
            custo_01.send_keys(custo_padrao)

            custo_02 = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="dataReferenciaCusto_id"]'))) # data referencia
            custo_02.send_keys(dia_atual)

            salvar_produto()
        else:
            salvar_produto()

    def cadastrar_produto(cod, descricao, und, tipo, grupo, metodo, descricao_fiscal, pcp_ressuprimento, custo_padrao, dia_atual):
        criar_produtos()
        cadastrar_aba_inicial(cod, descricao, und, tipo, grupo, metodo)
        cadastrar_empresas()
        cadastrar_fiscal(descricao_fiscal)
        cadastrar_pcp(pcp_ressuprimento)
        cadastrar_mrp()
        cadastrar_custo(custo_padrao, dia_atual)
       
    entrar_nomus()
    tela_produtos()
    
    for index, row in df.iterrows():
        cod = row['Código']
        descricao = row['Descrição']
        und = row['Unidade']
        tipo = row['Tipo do produto']
        grupo = row['Grupo do produto']
        metodo = row['Metodo']
        descricao_fiscal = row['Descrição Fiscal']
        pcp_ressuprimento = row['PCP Ressuprimento']
        custo_padrao = row['Custo padrao']
                    
        cadastrar_produto(cod, descricao, und, tipo, grupo, metodo, descricao_fiscal, pcp_ressuprimento, custo_padrao, dia_atual)

def criar_janela_planilha_cadastro():
    arquivo_excel = r'Z:\PUBLICO\Araujo\Cadastros\Cadastros.xlsx'
    os.startfile(arquivo_excel)

def criar_janela_lista_de_materiais():
    

app = tk.Tk()
app.title("Engenharia")
app.geometry("400x600")

tk.Label(app, text="Selecione a opção desejada: ", font=("Arial", 10)).pack(pady=5)
label_usuario = tk.Label(app, text="Usuário:")
label_usuario.pack(padx=5, pady=5)

entry_usuario = tk.Entry(app)
entry_usuario.pack(padx=5, pady=5)

label_senha = tk.Label(app, text="Senha:")
label_senha.pack(padx=5, pady=5)

entry_senha = tk.Entry(app, show="*")
entry_senha.pack(padx=5, pady=5)

tk.Button(app, text="Abrir planilha cadastro", command=criar_janela_planilha_cadastro, width=20, bg="green", fg="white").pack(pady=10)
tk.Button(app, text="Cadastrar Produtos", command=criar_janela_cadastro_produtos, width=20, bg="green", fg="white").pack(pady=10)

tk.Button(app, text="Sair", command=app.quit, bg="red", fg="white", width=20).pack(pady=10)

app.mainloop()
