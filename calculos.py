import tkinter as tk
from tkinter import ttk, messagebox
import math

# Funções para abrir as janelas específicas
def abrir_tabela_pesos():
    janela_tabela = tk.Toplevel(app)
    janela_tabela.title("Tabela de Pesos")
    janela_tabela.geometry("500x500")
    criar_janela_tabela_pesos(janela_tabela)

def abrir_mangueiras():
    janela_mangueiras = tk.Toplevel(app)
    janela_mangueiras.title("Mangueiras")
    janela_mangueiras.geometry("500x650")
    criar_janela_mangueira(janela_mangueiras)

# def abrir_bracos():
#     janela_bracos = tk.Toplevel(app)
#     janela_bracos.title("Braços")
#     janela_bracos.geometry("700x900")
#     criar_janela_bracos(janela_bracos)

# def abrir_succao():
#     janela_succao = tk.Toplevel(app)
#     janela_succao.title("Sucção")
#     janela_succao.geometry("400x300")
#     tk.Label(janela_succao, text="Funcionalidade de Sucção em desenvolvimento.", pady=20).pack()

# def abrir_perda_de_carga():
    # janela_succao = tk.Toplevel(app)
    # janela_succao.title("Perda de Carga")
    # janela_succao.geometry("400x300")
    # tk.Label(janela_succao, text="Funcionalidade de Perda de carga em desenvolvimento.", pady=20).pack()


# Função para criar a janela de Tabela de Pesos
def criar_janela_tabela_pesos(janela):
    # Dicionários para tipos de cálculo e materiais
    tipo_calculo = {
        "Tubo Redondo": "TR",
        "Tubo Quadrado/Retangular": "TQR",
        "Cantoneira": "CT",
        "Barra Redonda": "BR",
        "Barra Retangular/Quadrada": "BRQ",
        "Barra Sextavada": "BS"
    }

    material_aplicacao = {
        "Aço": 7.85,
        "Alumínio": 2.71,
        "Alumínio Fundido": 2.6,
        "Bronze": 8.8,
        "Latão": 8.55,
        "Chumbo": 11.31,
        "Chumbo fundido": 10.37,
        "Teflon": 2.2,
        "PVC": 1.45,
        "Policetal": 1.425,
        "Polietileno (PEAD)": 0.92,
        "Polipropileno": 0.91,
        "Nylon": 1.14,
        "HMW": 0.942,
    }

    def atualizar_campos(event):
        for widget in campos_frame.winfo_children():
            widget.destroy()
        
        calculo = tipo_calculo.get(tipo_combobox.get())

        if calculo == "TR":
            criar_campo("Diâmetro Externo (mm)", "di_ext")
            criar_campo("Espessura (mm)", "esp")
            criar_campo("Comprimento (mm)", "comp")
        elif calculo == "TQR":
            criar_campo("Lado 1 (mm)", "l1")
            criar_campo("Lado 2 (mm)", "l2")
            criar_campo("Espessura (mm)", "esp")
            criar_campo("Comprimento (mm)", "comp")
        elif calculo == "CT":
            criar_campo("Lado 1 (mm)", "l1")
            criar_campo("Lado 2 (mm)", "l2")
            criar_campo("Espessura (mm)", "esp")
            criar_campo("Comprimento (mm)", "comp")
        elif calculo == "BR":
            criar_campo("Diâmetro (mm)", "di")
            criar_campo("Comprimento (mm)", "comp")
        elif calculo == "BRQ":
            criar_campo("Lado 1 (mm)", "l1")
            criar_campo("Lado 2 (mm)", "l2")
            criar_campo("Comprimento (mm)", "comp")
        elif calculo == "BS":
            criar_campo("Fator A (mm)", "fa")
            criar_campo("Comprimento (mm)", "comp")

    def criar_campo(label_text, field_name):
        frame = tk.Frame(campos_frame)
        frame.pack(anchor="w", pady=2)
        tk.Label(frame, text=label_text, width=20, anchor="w").pack(side="left")
        entry = tk.Entry(frame, width=15)
        entry.pack(side="right")
        campos[field_name] = entry

    def converter_valor(valor):
        try:
            return float(valor.replace(",", "."))
        except ValueError:
            raise ValueError(f"Valor inválido: {valor}")

    def calcular():
        tipo_escolhido = tipo_combobox.get()
        material_escolhido = material_combobox.get()
        
        calculo = tipo_calculo.get(tipo_escolhido)
        material = material_aplicacao.get(material_escolhido)

        try:
            if calculo == "TR":
                di_ext = converter_valor(campos["di_ext"].get())
                esp = converter_valor(campos["esp"].get())
                comp = converter_valor(campos["comp"].get())
                kg_metro = ((di_ext**2) * (0.7854/1000) * material) - (((di_ext - (esp * 2))**2) * (0.7854/1000) * material)
                kg_total = (kg_metro * comp) / 1000
            elif calculo == "TQR":
                l1 = converter_valor(campos["l1"].get())
                l2 = converter_valor(campos["l2"].get())
                esp = converter_valor(campos["esp"].get())
                comp = converter_valor(campos["comp"].get())
                kg_metro = ((l1 * l2) / 1000) * material - (((l1 - esp * 2) * (l2 - esp * 2)) / 1000) * material
                kg_total = (kg_metro * comp) / 1000
            elif calculo == "CT":
                l1 = converter_valor(campos["l1"].get())
                l2 = converter_valor(campos["l2"].get())
                esp = converter_valor(campos["esp"].get())
                comp = converter_valor(campos["comp"].get())
                kg_metro = (((l1 * esp) + ((l2 - esp) * esp)) * material * 1000) / 1000000000
                kg_total = (kg_metro * comp) / 1000
            elif calculo == "BR":
                di = converter_valor(campos["di"].get())
                comp = converter_valor(campos["comp"].get())
                kg_metro = ((di / 2000)**2) * 3.1416 * material * 1000
                kg_total = (kg_metro * comp) / 1000
            elif calculo == "BRQ":
                l1 = converter_valor(campos["l1"].get())
                l2 = converter_valor(campos["l2"].get())
                comp = converter_valor(campos["comp"].get())
                kg_metro = (l1 / 1000) * (l2 / 1000) * comp * material
                kg_total = (kg_metro * comp) / 1000
            elif calculo == "BS":
                fa = converter_valor(campos["fa"].get())
                comp = converter_valor(campos["comp"].get())
                kg_metro = (fa / 1000) * 0.57735 * (fa / 2000) * 6 * comp * material
                kg_total = (kg_metro * comp) / 1000
            else:
                raise ValueError("Selecione um tipo de cálculo válido.")
            
            resultado_label.config(text=f"Peso total: {kg_total:.3f} Kg")
        except ValueError as ve:
            resultado_label.config(text=f"Erro: {ve}")
        except Exception as e:
            resultado_label.config(text=f"Erro inesperado: {e}")

    # Interface gráfica para Tabela de Pesos
    tk.Label(janela, text="Tabela de pesos").pack(pady=5)
    tipo_combobox = ttk.Combobox(janela, values=list(tipo_calculo.keys()), width=30)
    tipo_combobox.pack()
    tipo_combobox.bind("<<ComboboxSelected>>", atualizar_campos)

    tk.Label(janela, text="Material").pack(pady=5)
    material_combobox = ttk.Combobox(janela, values=list(material_aplicacao.keys()), width=30)
    material_combobox.pack()

    campos_frame = tk.Frame(janela)
    campos_frame.pack(pady=10)

    tk.Button(janela, text="Calcular", command=calcular, width=30).pack(pady=10)
    resultado_label = tk.Label(janela, text="", justify="left", bg="white", width=40, height=5, relief="sunken", anchor="nw", padx=5, pady=5)
    resultado_label.pack(pady=10)
    campos = {}

# Função para criar a janela de Mangueiras
def criar_janela_mangueira(janela):
    def calcular_mangueira():
        try:
             # Obtendo valores do formulário
            diametro = diametro_var.get()
            comprimento = float(entry_comprimento.get())
            material_arame = material_arame_var.get().lower()
            tipo_mangueira = int(tipo_mangueira_var.get())
            material_mangueira = material_mangueira_var.get()

            diametros = {
                "1/4": 6.33,
                "3/8": 9.525,
                "1/2": 12.7,
                "3/4": 19.05,
                "1": 25.4,
                "1.1/2": 38.1,
                "2": 50.8,
                "2.1/2": 63.5,
                "3": 76.2,
                "4": 101.6,
                "6": 152.4,
                "8": 203.2,
                }
            materiais = {
                "aco galvanizado": 7.2,
                "aco revestido": 8.8,
                "aco inox 304": 8.4,
                "aco inox 316": 8.1,
                "aco inox 304 revestido": 7,
                }
            diametro_interno = diametros.get(diametro)
            if diametro_interno is None:
                messagebox.showerror("Erro", "Diâmetro inválido. Escolha um válido.")
                return

            mat_arame = materiais.get(material_arame)
            if mat_arame is None:
                messagebox.showerror("Erro", "Material de arame inválido. Escolha um válido.")
                return

            # Cálculo inicial
            esp_ate_4, esp_ate_6, esp_ate_8 = 15, 22, 28

            if diametro_interno <= 25.4:  # Menor que 1 pol
                diametro_externo = diametro_interno + esp_ate_4
                arame_interno = 2
                arame_externo = 2
                passo_arame = 10
                di_corda = '1/4"'
                di_corda_value = 6.35
                gm_corda = 20
                pead_porcent = 1
            elif diametro_interno == 38.1 or diametro_interno == 50.8:  # 1.1/2 e 2 pol
                diametro_externo = diametro_interno + esp_ate_4
                arame_interno = 2.5
                arame_externo = 2.5
                passo_arame = 11
                di_corda = '1/4"'
                di_corda_value = 6.35
                gm_corda = 20
                pead_porcent = 1
            elif diametro_interno == 63.5:  # 2.1/2 pol
                diametro_externo = diametro_interno + esp_ate_4
                arame_interno = 3
                arame_externo = 3
                passo_arame = 14
                di_corda = '1/4"'
                di_corda_value = 6.35
                gm_corda = 20
                pead_porcent = 1
            elif diametro_interno == 76.2:  # 3 pol
                diametro_externo = diametro_interno + esp_ate_4
                arame_interno = 3.4
                arame_externo = 3.4
                passo_arame = 14
                di_corda = '1/4"'
                di_corda_value = 6.35
                gm_corda = 20
                pead_porcent = 1
            elif diametro_interno == 101.6: # 4 pol
                diametro_externo = diametro_interno + esp_ate_4
                arame_interno = 3.4
                arame_externo = 3.4
                passo_arame = 14
                di_corda = '3/8"'
                di_corda_value = 9.525
                gm_corda = 48
                pead_porcent = 1
            elif diametro_interno == 152.4: # 6 pol
                diametro_externo = diametro_interno + esp_ate_6
                arame_interno = 5
                arame_externo = 6
                passo_arame = 18
                di_corda = '1/2"'
                di_corda_value = 12.7
                gm_corda = 77
                pead_porcent = 1.5
            elif diametro_interno == 203.2: # 8 pol
                diametro_externo = diametro_interno + esp_ate_8
                arame_interno = 6
                arame_externo = 6
                passo_arame = 20
                di_corda = '1/2"'
                di_corda_value = 12.7
                gm_corda = 77
                pead_porcent = 1.5
            if tipo_mangueira == 1:
                if material_mangueira == "PP":
                    if diametro_interno <= 25.4:  # 1 pol
                        camadas = 4
                        rafia = 3
                        pelicula_ptfe = 0
                        larg_rafia = 150
                        gm2_rafia = 160
                        larg_teflon = 0
                    elif diametro_interno == 38.1 or  diametro_interno == 50.8:  # 1.1/2 e 2 pol
                        camadas = 6
                        rafia = 3
                        pelicula_ptfe = 0
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 0
                    elif diametro_interno == 63.5 or diametro_interno == 76.2: #  2.1/2 3 pol
                        camadas = 6
                        rafia = 3
                        pelicula_ptfe = 0
                        larg_rafia = 500
                        gm2_rafia = 160
                        larg_teflon = 0
                    elif diametro_interno == 101.6: # 4 pol
                        camadas = 12
                        rafia = 5
                        pelicula_ptfe = 0
                        larg_rafia = 500
                        gm2_rafia = 160
                        larg_teflon = 0
                    elif diametro_interno == 152.4: # 6 pol
                        camadas = 24
                        rafia = 10
                        pelicula_ptfe = 0
                        larg_rafia = 800
                        gm2_rafia = 220
                        larg_teflon = 0
                    elif diametro_interno == 203.2: # 8 pol
                        camadas = 28
                        rafia = 10
                        pelicula_ptfe = 0
                        larg_rafia = 800
                        gm2_rafia = 220
                        larg_teflon = 0
            if tipo_mangueira == 2:
                if material_mangueira == "PP":
                    if diametro_interno <= 25.4:  # 1 pol
                        camadas = 5
                        rafia = 3
                        pelicula_ptfe = 0
                        larg_rafia = 150
                        gm2_rafia = 160
                        larg_teflon = 0
                    elif diametro_interno == 38.1 or  diametro_interno == 50.8:  # 1.1/2 e 2 pol
                        camadas = 6
                        rafia = 4
                        pelicula_ptfe = 0
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 0
                    elif diametro_interno == 63.5 or diametro_interno == 76.2: #  2.1/2 3 pol
                        camadas = 10
                        rafia = 3
                        pelicula_ptfe = 0
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 0
                    elif diametro_interno == 101.6: # 4 pol
                        camadas = 12
                        rafia = 5
                        pelicula_ptfe = 0
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 0
                    elif diametro_interno == 152.4: # 6 pol
                        camadas = 24
                        rafia = 10
                        pelicula_ptfe = 0
                        larg_rafia = 800
                        gm2_rafia = 220
                        larg_teflon = 0
                    elif diametro_interno == 203.2: # 8 pol
                        camadas = 28
                        rafia = 14
                        pelicula_ptfe = 0
                        larg_rafia = 800
                        gm2_rafia = 220
                        larg_teflon = 0
            if tipo_mangueira == 3:
                if material_mangueira == "PP":
                    if diametro_interno <= 25.4:  # 1 pol
                        camadas = 5
                        rafia = 3
                        pelicula_ptfe = 0
                        larg_rafia = 150
                        gm2_rafia = 160
                        larg_teflon = 0
                    elif diametro_interno == 38.1 or  diametro_interno == 50.8:  # 1.1/2 e 2 pol
                        camadas = 6
                        rafia = 3
                        pelicula_ptfe = 0
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 0
                    elif diametro_interno == 63.5 or diametro_interno == 76.2: #  2.1/2 3 pol
                        camadas = 10
                        rafia = 5
                        pelicula_ptfe = 0
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 0
                    elif diametro_interno == 101.6: # 4 pol
                        camadas = 12
                        rafia = 7
                        pelicula_ptfe = 0
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 0
                    elif diametro_interno == 152.4: # 6 pol
                        camadas = 24
                        rafia = 10
                        pelicula_ptfe = 0
                        larg_rafia = 800
                        gm2_rafia = 220
                        larg_teflon = 0
                    elif diametro_interno == 203.2: # 8 pol
                        camadas = 28
                        rafia = 18
                        pelicula_ptfe = 0
                        larg_rafia = 800
                        gm2_rafia = 220
                        larg_teflon = 0
            if tipo_mangueira == 1:
                if material_mangueira == "PTFE":
                    if diametro_interno <= 25.4:  # 1 pol
                        camadas = 4
                        rafia = 3
                        pelicula_ptfe = 1
                        larg_rafia = 150
                        gm2_rafia = 220
                        larg_teflon = 150
                    elif diametro_interno == 38.1 or  diametro_interno == 50.8:  # 1.1/2 e 2 pol
                        camadas = 6
                        rafia = 3
                        pelicula_ptfe = 1
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 250
                    elif diametro_interno == 63.5 or diametro_interno == 76.2: #  2.1/2 3 pol
                        camadas = 6
                        rafia = 3
                        pelicula_ptfe = 1
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 400
                    elif diametro_interno == 101.6: # 4 pol
                        camadas = 12
                        rafia = 5
                        pelicula_ptfe = 1
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 400
                    elif diametro_interno == 152.4: # 6 pol
                        camadas = 24
                        rafia = 10
                        pelicula_ptfe = 1
                        larg_rafia = 800
                        gm2_rafia = 220
                        larg_teflon = 800
                    elif diametro_interno == 203.2: # 8 pol
                        camadas = 28
                        rafia = 10
                        pelicula_ptfe = 1
                        larg_rafia = 800
                        gm2_rafia = 220
                        larg_teflon = 800
            if tipo_mangueira == 2:
                if material_mangueira == "PTFE":
                    if diametro_interno <= 25.4:  # 1 pol
                        camadas = 5
                        rafia = 3
                        pelicula_ptfe = 1
                        larg_rafia = 150
                        gm2_rafia = 160
                        larg_teflon = 150
                    elif diametro_interno == 38.1 or  diametro_interno == 50.8:  # 1.1/2 e 2 pol
                        camadas = 6
                        rafia = 4
                        pelicula_ptfe = 1
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 250
                    elif diametro_interno == 63.5 or diametro_interno == 76.2: #  2.1/2 3 pol
                        camadas = 10
                        rafia = 3
                        pelicula_ptfe = 1
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 400
                    elif diametro_interno == 101.6: # 4 pol
                        camadas = 12
                        rafia = 5
                        pelicula_ptfe = 1
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 400
                    elif diametro_interno == 152.4: # 6 pol
                        camadas = 24
                        rafia = 10
                        pelicula_ptfe = 1
                        larg_rafia = 800
                        gm2_rafia = 220
                        larg_teflon = 800
                    elif diametro_interno == 203.2: # 8 pol
                        camadas = 28
                        rafia = 14
                        pelicula_ptfe = 1
                        larg_rafia = 800
                        gm2_rafia = 220
                        larg_teflon = 800
            if tipo_mangueira == 3:
                if material_mangueira == "PTFE":
                    if diametro_interno <= 25.4:  # 1 pol
                        camadas = 5
                        rafia = 3
                        pelicula_ptfe = 1
                        larg_rafia = 150
                        gm2_rafia = 160
                        larg_teflon = 150
                    elif diametro_interno == 38.1 or  diametro_interno == 50.8:  # 1.1/2 e 2 pol
                        camadas = 6
                        rafia = 3
                        pelicula_ptfe = 1
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 250
                    elif diametro_interno == 63.5 or diametro_interno == 76.2: #  2.1/2 3 pol
                        camadas = 10
                        rafia = 5
                        pelicula_ptfe = 1
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 400
                    elif diametro_interno == 101.6: # 4 pol
                        camadas = 12
                        rafia = 7
                        pelicula_ptfe = 1
                        larg_rafia = 400
                        gm2_rafia = 160
                        larg_teflon = 400
                    elif diametro_interno == 152.4: # 6 pol
                        camadas = 24
                        rafia = 10
                        pelicula_ptfe = 1
                        larg_rafia = 800
                        gm2_rafia = 220
                        larg_teflon = 800
                    elif diametro_interno == 203.2: # 8 pol
                        camadas = 28
                        rafia = 18
                        pelicula_ptfe = 1
                        larg_rafia = 800
                        gm2_rafia = 220
                        larg_teflon = 800
            if tipo_mangueira == 4:
                if material_mangueira == "Nylon":
                    if diametro_interno <= 25.4:  # 1 pol
                        camadas = 5
                        rafia = 0
                        pelicula_ptfe = 2
                        larg_rafia = 0
                        gm2_rafia = 0
                        larg_teflon = 150
                    elif diametro_interno == 38.1 or  diametro_interno == 50.8:  # 1.1/2 e 2 pol
                        camadas = 5
                        rafia = 0
                        pelicula_ptfe = 2
                        larg_rafia = 0
                        gm2_rafia = 0
                        larg_teflon = 250
                    elif diametro_interno == 63.5 or diametro_interno == 76.2: #  2.1/2 3 pol
                        camadas = 6
                        rafia = 0
                        pelicula_ptfe = 2
                        larg_rafia = 0
                        gm2_rafia = 0
                        larg_teflon = 400
                    elif diametro_interno == 101.6: # 4 pol
                        camadas = 7
                        rafia = 0
                        pelicula_ptfe = 2
                        larg_rafia = 0
                        gm2_rafia = 0
                        larg_teflon = 400
                    elif diametro_interno == 152.4: # 6 pol
                        camadas = 8
                        rafia = 0
                        pelicula_ptfe = 2
                        larg_rafia = 0
                        gm2_rafia = 0
                        larg_teflon = 800
                    elif diametro_interno == 203.2: # 8 pol
                        camadas = 9
                        rafia = 0
                        pelicula_ptfe = 2
                        larg_rafia = 0
                        gm2_rafia = 0
                        larg_teflon = 800
            gab_mm = passo_arame + arame_interno
            raio_int_mm = (diametro_interno/2) + arame_interno/2
            comp_int_mm = (1.1*comprimento*math.sqrt(abs(((2*3.14*raio_int_mm) ** 2) + (gab_mm**2)))/gab_mm)/1000
            cm3_int_mm = ((3.14*((arame_interno/10)**2))/4) * (comp_int_mm*100)
            peso_arame_int = (mat_arame*cm3_int_mm)/1000
            raio_ext_mm = (diametro_externo/2) - arame_externo/2
            comp_ext_mm = (math.sqrt(abs(((2*3.14*raio_ext_mm)**2)+(gab_mm**2)))*comprimento/gab_mm)/1000
            cm3_ext_mm = ((3.14*((arame_externo/10)**2))/4) * (comp_ext_mm*100)
            peso_arame_ext = ((mat_arame*cm3_ext_mm)/1000)*1.1
            comp_filme = (math.sqrt(abs(((2*3.14*diametro_externo)**2)+((800*1.8)**2)))*comprimento/(800*1.8))/1000
            area_filme = comp_filme*0.8
            peso_filme = ((area_filme*46)/1000)*camadas*1.2
            try:
                comp_rafia = (math.sqrt(abs(((2*3.14*((diametro_interno+diametro_externo)/2))**2)+((larg_rafia*1.8)**2)))*comprimento/(larg_rafia*1.8))/1000
                area_rafia = comp_rafia*(larg_rafia/1000)
                peso_rafia = ((area_rafia*gm2_rafia)/1000*rafia)*1.2
            except ZeroDivisionError:
                comp_rafia = 0
                area_rafia = 0
                peso_rafia = 0
            try:
                comp_teflon = (math.sqrt(abs(((2*3.14*diametro_externo)**2)+((larg_teflon*1.8)**2)))*comprimento/(larg_teflon*1.8))/1000
                area_teflon = comp_teflon*(larg_teflon/1000)
                peso_teflon = ((area_teflon*235)/1000)*pelicula_ptfe*1.2
            except ZeroDivisionError:
                comp_teflon = 0
                area_teflon = 0
                peso_teflon = 0
            comp_corda = (math.sqrt(abs(((2*3.14*raio_ext_mm)**2)+di_corda_value**2))*comprimento/di_corda_value)/1000
            corda = (comp_corda*gm_corda)/1000
            pead = pead_porcent*comprimento

            resultado_box.delete("1.0", tk.END)  # Limpa o conteúdo anterior
            resultado = f' - Diamentro interno: {diametro_interno}\n - Diâmetro externo: {diametro_externo}\n\n - Arame interno : {round(peso_arame_int,2)} Kg\n - Arame externo : {round(peso_arame_ext,2)} Kg\n\n - Comprimento : {int(comprimento)} mm\n - Diametro do arame : {passo_arame}\n\n - QTDE. Camadas filme : {camadas}\n - Filme : {round(peso_filme,2)} Kg\n\n - QTDE. Camadas rafia : {rafia}\n - Rafia : {round(peso_rafia,2)} Kg\n\n - QTDE. Pelicula teflon : {pelicula_ptfe}\n - Pelicula teflon : {round(peso_teflon)} Kg\n\n - QTDE. PEAD : {int(pead)} mm\n\n - Corda de : {di_corda}\n - QTDE. Corda: {round(corda,2)} Kg'
            resultado_box.insert(tk.END, resultado)

        except ValueError:
            resultado_box.delete("1.0", tk.END)
            resultado_box.insert(tk.END, "Erro: Certifique-se de que todos os campos estão preenchidos corretamente.")

    # Entrada de Dados
    tk.Label(janela, text="Diâmetro:").grid(row=0, column=0, pady=5, sticky="e")
    global diametro_var, entry_comprimento, material_arame_var, tipo_mangueira_var, material_mangueira_var, resultado_box
    diametro_var = tk.StringVar()
    diametro_menu = ttk.Combobox(janela, textvariable=diametro_var, values=[
        "1/4", "3/8", "1/2", "3/4", "1", "1.1/2", "2", "2.1/2", "3", "4", "6", "8"
    ])
    diametro_menu.grid(row=0, column=1, pady=5)

    tk.Label(janela, text="Comprimento (mm):").grid(row=1, column=0, pady=5, sticky="e")
    entry_comprimento = tk.Entry(janela)
    entry_comprimento.grid(row=1, column=1, pady=5)

    tk.Label(janela, text="Material do Arame:").grid(row=2, column=0, pady=5, sticky="e")
    material_arame_var = tk.StringVar()
    material_arame_menu = ttk.Combobox(janela, textvariable=material_arame_var, values=[
        "aco galvanizado", "aco revestido", "aco inox 304", "aco inox 316", "aco inox 304 revestido"
    ])
    material_arame_menu.grid(row=2, column=1, pady=5)

    tk.Label(janela, text="Tipo de Mangueira:").grid(row=3, column=0, pady=5, sticky="e")
    tipo_mangueira_var = tk.StringVar()
    tipo_mangueira_menu = ttk.Combobox(janela, textvariable=tipo_mangueira_var, values=["1", "2", "3"])
    tipo_mangueira_menu.grid(row=3, column=1, pady=5)

    tk.Label(janela, text="Material da Mangueira:").grid(row=4, column=0, pady=5, sticky="e")
    material_mangueira_var = tk.StringVar()
    material_mangueira_menu = ttk.Combobox(janela, textvariable=material_mangueira_var, values=["PP", "PTFE"])
    material_mangueira_menu.grid(row=4, column=1, pady=5)

    # Botão de Cálculo
    btn_calcular = tk.Button(janela, text="Calcular", command=calcular_mangueira)
    btn_calcular.grid(row=5, columnspan=2, pady=20)
    
    # Caixa de Resultado
    tk.Label(janela, text="Resultado:").grid(row=6, column=0, columnspan=2)
    resultado_box = tk.Text(janela, height=22, width=60)
    resultado_box.grid(row=7, column=0, columnspan=2, pady=10)

# Função para calcular os braços
# def criar_janela_bracos(janela):
    # def calcular_bc():
    #     inicio = entry_inicio_var
    #     tipo_de_braco = entry_tipo_de_braco_var
    #     terminacao = entry_terminacao_var
    #     diametro = entry_diametro_var
    #     material_bc = entry_material_bc_var
    #     primario = entry_primario_var
    #     secundario = entry_secundario_var
    #     mergulhador = entry_mergulhador_var
    #     acessorio_1 = entry_acessorio_1_var
    #     acessorio_2 = entry_acessorio_2_var
    #     primeira_flange_primario = entry_primeira_flange_primario_var
    #     tubo_primario = entry_tubo_primario_var
    #     segunda_flange_primario = entry_segunda_flange_primario_var
    #     primeira_flange_secundario = entry_primeira_flange_secundario_var
    #     tubo_secundario = entry_tubo_secundario_var
    #     segunda_flange_secundario = entry_segunda_flange_secundario_var
    #     flange_mergulhador = entry_flange_mergulhador_var
    #     tubo_mergulhador = entry_tubo_mergulhador_var
    #     primeira_junta = entry_primeira_junta_var
    #     qtde_junta_1 = entry_qtde_junta_1_var
    #     segunda_junta = entry_segunda_junta_var
    #     qtde_junta_2 = entry_qtde_junta_2_var
    #     mola_selecao = entry_mola_selecao_var
    #     furos = entry_furos_var
    #     modelo =""
    #     material_braco=""
    #     peso_total=0
    #     momento_total=0
    #     terminacao_do_braco=""
    #     mola_selecionada = ""
    #     ajuste_traseiro=0
    #     voltas = 0

    #     pesos = {
    #         "4 POL SCH 10 AL":2.79,
    #         "4 POL SCH 40 AL":5.35666666666667,
    #         "4 POL SCH 10 ACO":8.37,
    #         "4 POL SCH 40 ACO":16.07,
    #         "3 POL SCH 10 AL":2.15333333333333,
    #         "3 POL SCH 40 AL":3.76333333333333, 
    #         "3 POL SCH 10 ACO":6.46,
    #         "3 POL SCH 40 ACO":11.29,
    #         "FLG TTMA 4 POL ACO":1.5,
    #         "FLG TTMA 4 POL AL":0.5,
    #         "FLG TTMA 3 POL ACO":1.1,
    #         "FLG TTMA 3 POL AL":0.37,
    #         "FLG ANSI 4 POL 150 AC SO":5.9,
    #         "FLG ANSI 3 POL 150 AC SO":3.6,
    #         "FLG ANSI 4 POL 150 AL SO":2.95,
    #         "FLG ANSI 3 POL 150 AL SO":1.8,
    #         "FLG JG 4 POL ACO":1.8,
    #         "FLG JG 3 POL ACO":1.5,
    #         "SEM FLANGE":0,
    #         "JM40TT-42A":9,
    #         "JM40TT-32A":6.8,
    #         "JM40SS-42A":8.5,
    #         "JM40SS-32A":6.3,
    #         "DEFLETOR 4 POL AL":1.6,
    #         "DEFLETOR 3 POL AL":1,
    #         "DEFLETOR 4 POL AC":1.2,
    #         "DEFLETOR 3 POL AC":1,
    #         "CHANFRO":0,
    #         "VALVULA DEADMAN":12.3,
    #         "VALVULA ESFERA BP 4 POL":26.5,
    #         "VALVULA ESFERA BP 3 POL":15.2,
    #         "VALVULA ESFERA WF 4 POL":9.8,
    #         "VALVULA ESFERA WF 3 POL":5.5,
    #         "SEM VALVULA":0,
    #         "OUTROS ITENS":3,
    #         "CJ. MOLA":30.43,
    #         "CURVA RC 4 POL AC SCH 10":1.3,
    #         "CURVA RC 3 POL AC SCH 10":0.7,
    #         "MODULO 4 AC":6.19,
    #         "MODULO 3 AC":5.2,
    #         "CJ. ENTRADA 4 POL AC":8.56,
    #         "CJ. ENTRADA 3 POL AC":6.35,
    #         "CURVA RL FLG JG 4 POL AC":7.5,
    #         "CURVA RL FLG JG 3 POL AC":5.1,
    #         "PRIMÁRIO 4 POL AC":14.75,
    #         "PRIMÁRIO 3 POL AC":14.7,
    #     }
    #     comprimentos = {
    #         "VALVULA DEADMAN":0.34,
    #         "VALVULA ESFERA BP 4 POL":0.229,
    #         "VALVULA ESFERA BP 3 POL":0.203,
    #         "VALVULA ESFERA WF 4 POL":0.124,
    #         "VALVULA ESFERA WF 3 POL":0.103,
    #         "CJ. MOLA":609,
    #         "PRIMÁRIO 4 POL AC":0.42,
    #         "PRIMÁRIO 3 POL AC":0.394,
    #     }
    #     tipo_de_bc = {
    #         "BTA":"BTA",
    #         "BTC":"BTC",
    #         "BTE":"BTE",
    #         "BTG":"BTG",
    #         "BBI":"BBI",
    #         "BBJ":"BBJ",
    #         "BBJ":"BBJ",
    #         "BBK":"BBK",
    #         "BBL":"BBL",
    #         "BBN":"BBN",
    #         "BBP":"BBP",
    #     }
    #     material = {
    #         "2CA":"2CA",
    #         "3CA":"3CA",
    #         "2CC":"2CC",
    #         "3CC":"3CC",
    #         "2I4":"2I4",
    #         "3I4":"3I4",
    #         "2I6":"2I6",
    #         "3I6":"3I6",
    #         "24A":"24A",
    #         "34A":"34A",
    #     }
    #     entradas_1 = {
    #         'CJ. ENTRADA 4 POL AC':'MODULO 4 AC',
    #         'CJ. ENTRADA 3 POL AC':'MODULO 3 AC',
    #         'MODULO 4 AC':'CURVA RL FLG JG 4 POL AC',
    #         'MODULO 3 AC':'CURVA RL FLG JG 3 POL AC',
    #     }
    #     pri = {
    #         'PRIMÁRIO 4 POL AC':210,
    #         'PRIMÁRIO 3 POL AC':242,
    #     }
    #     mola_4_12_max = {
    #         1:539.9,
    #         2:519.9,
    #         3:498.4,
    #         4:475,
    #         5:450.13,
    #     }
    #     mola_4_max = {
    #         1:123.4,
    #         2:148.4,
    #         3:170,
    #         4:188.6,
    #         5:204.4,
    #     }
    #     mola = {
    #         "MOLA 4 E 12":'MOLA 4 E 12',
    #         "MOLA 4":'MOLA 4',
    #     }
                       

        
    #     if inicio == "TOP":
    #        ###########################################################################
    #         mola_selecionada = mola.get(mola_selecao)
    #         terminacao_do_braco = f'.0{terminacao}'
    #         peso_acessorio_1 = pesos.get(acessorio_1)
    #         peso_acessorio_2 = pesos.get(acessorio_2)
    #         modelo = tipo_de_bc.get(tipo_de_braco)
    #         material_braco = material.get(material_bc)
    #         comprimento_1 = comprimentos.get(acessorio_1,0)
    #         comprimento_2 = comprimentos.get(acessorio_2,0)
    #         qtde_padrao = 1
    #         qtde_modulo_entrada = 2
    #         # # Primeira parte
    #         conj_entrada = f'CJ. ENTRADA {diametro} POL AC'
    #         modulo_entrada = entradas_1.get(conj_entrada)
    #         curva_entrada = entradas_1.get(modulo_entrada)
         

    #         peso_conj_entrada = pesos.get(conj_entrada)
    #         peso_modulo_entrada = pesos.get(modulo_entrada) * qtde_modulo_entrada
    #         peso_curva_entrada = pesos.get(curva_entrada)

    #         d_mola_conj_entrada = 0
    #         d_mola_modulo_entrada = 0
    #         d_mola_curva_entrada = 0

    #         momento_mola_conj_entrada = (peso_conj_entrada * d_mola_conj_entrada * 9.8) / 1000
    #         momento_mola_modulo_entrada = (peso_modulo_entrada * d_mola_modulo_entrada * 9.8) / 1000
    #         momento_mola_curva_entrada = (peso_curva_entrada * d_mola_curva_entrada * 9.8) / 1000

    #         soma_pesos_1 = peso_conj_entrada + peso_modulo_entrada + peso_curva_entrada
    #         soma_momentos_1 = momento_mola_conj_entrada + momento_mola_curva_entrada + momento_mola_modulo_entrada
    #         ###########################################################################
    #         # # Segunda parte
    #         primario_bc = f'PRIMÁRIO {diametro} POL AC'

    #         distancia_tubo = comprimentos[primario_bc]

    #         qtde_tubo = primario - (comprimento_1 + distancia_tubo)

    #         peso_primario = pesos.get(primario_bc)
    #         peso_flange_1 = pesos.get(primeira_flange_primario)
    #         peso_tubo = pesos.get(tubo_primario)*qtde_tubo
    #         peso_flange_2 = pesos.get(segunda_flange_primario)

    #         d_mola_primario = pri.get(primario_bc)
    #         d_mola_acessorio_1 = ((comprimento_1*1000)/2) + (comprimentos.get(primario_bc)*1000)
    #         d_mola_flange_1 = (comprimento_1*1000) + (comprimentos.get(primario_bc)*1000)
    #         d_mola_tubo = (((primario * 1000) - d_mola_flange_1)/2) + d_mola_flange_1
    #         d_mola_flange_2 = primario * 1000

    #         momento_mola_primario= (peso_primario * d_mola_primario * 9.8) / 1000
    #         momento_mola_acessorio_1 = (peso_acessorio_1  * d_mola_acessorio_1 * 9.8) / 1000
    #         momento_mola_flange_1 = (peso_flange_1 * d_mola_flange_1 * 9.8) / 1000
    #         momento_mola_tubo = (peso_tubo * d_mola_tubo * 9.8) / 1000
    #         momento_mola_flange_2 = (peso_flange_2 * d_mola_flange_2 * 9.8) / 1000

    #         soma_pesos_2 = peso_primario + peso_acessorio_1 + peso_flange_1 + peso_tubo + peso_flange_2
    #         soma_momentos_2 = momento_mola_primario + momento_mola_acessorio_1 + momento_mola_flange_1 + momento_mola_tubo + momento_mola_flange_2
    #         ###########################################################################
    #         # # Terceira parte
    #         qtde_tubo_secundario = secundario - 0.21

    #         peso_flange_3 = pesos.get(primeira_flange_secundario)
    #         peso_tubo_secundario = pesos.get(tubo_secundario)*qtde_tubo_secundario
    #         peso_flange_4 = pesos.get(segunda_flange_secundario)

    #         d_mola_secundario = primario

    #         momento_mola_flange_3 = (peso_flange_3 * d_mola_secundario * 9.8) / 1000
    #         momento_mola_tubo_secundario = (peso_tubo_secundario * d_mola_secundario * 9.8) / 1000
    #         momento_mola_flange_4 = (peso_flange_2 * d_mola_secundario * 9.8) / 1000

    #         soma_pesos_3 = peso_flange_3 + peso_tubo_secundario + peso_flange_4
    #         soma_momentos_3 = momento_mola_flange_3 + momento_mola_tubo_secundario + momento_mola_flange_4
    #         ###########################################################################
    #         # # Quarta parte
    #         qtde_mergulhador = mergulhador - 0.11

    #         peso_flange_mergulhador = pesos.get(flange_mergulhador)
    #         peso_tubo_mergulhador = pesos.get(tubo_mergulhador)*qtde_mergulhador

    #         d_mola_mergulhador = primario * 1000

    #         momento_mola_flange_mergulhador = (peso_flange_mergulhador * d_mola_mergulhador * 9.8) / 1000
    #         momento_mola_tubo_mergulhador = (peso_tubo_mergulhador * d_mola_mergulhador * 9.8) / 1000
    #         momento_mola_acessorio_2 = (peso_acessorio_2 * d_mola_mergulhador * 9.8) / 1000

    #         soma_momentos_6 = momento_mola_flange_mergulhador + momento_mola_tubo_mergulhador + momento_mola_acessorio_2
    #         soma_pesos_6 = peso_flange_mergulhador + peso_tubo_mergulhador + peso_acessorio_2
    #         ###########################################################################
    #         # # Quinta parte
    #         peso_primeira_junta = pesos.get(primeira_junta) * qtde_junta_1
    #         peso_segunda_junta = pesos.get(segunda_junta) * qtde_junta_2

    #         d_mola_junta = d_mola_mergulhador

    #         momento_mola_primeira_junta = (peso_primeira_junta * d_mola_junta * 9.8) / 1000
    #         momento_mola_segunda_junta = (peso_segunda_junta * d_mola_junta * 9.8) / 1000

    #         soma_pesos_4 = peso_primeira_junta + peso_segunda_junta
    #         soma_momentos_4 = momento_mola_primeira_junta + momento_mola_segunda_junta
    #         ###########################################################################
    #         ###########################################################################
    #         # # Finalização
    #         peso_total = soma_pesos_1 + soma_pesos_2 + soma_pesos_3 + soma_pesos_4  + soma_pesos_6
    #         momento_total = soma_momentos_1 + soma_momentos_2 + soma_momentos_3 + soma_momentos_4  + soma_momentos_6
    #         max_2 = mola_4_12_max.get(furos)
    #         max_1 = mola_4_max.get(furos)
    #     if inicio == "BOT":
    #          ###########################################################################
    #         mola_selecionada = mola.get(mola_selecao)
    #         terminacao_do_braco = f'.0{terminacao}'
    #         peso_acessorio_1 = pesos.get(acessorio_1)
    #         peso_acessorio_2 = pesos.get(acessorio_2)
    #         modelo = tipo_de_bc.get(tipo_de_braco)
    #         material_braco = material.get(material_bc)
    #         comprimento_1 = comprimentos.get(acessorio_1,0)
    #         comprimento_2 = comprimentos.get(acessorio_2,0)
    #         # # Primeira parte
    #         conj_entrada = f'CJ. ENTRADA {diametro} POL AC'
    #         modulo_entrada = entradas_1.get(conj_entrada)
    #         curva_entrada = entradas_1.get(modulo_entrada)

    #         qtde_padrao = 1
    #         qtde_modulo_entrada = 2

    #         peso_conj_entrada = pesos.get(conj_entrada)
    #         peso_modulo_entrada = pesos.get(modulo_entrada) * 2
    #         peso_curva_entrada = pesos.get(curva_entrada)

    #         d_mola_conj_entrada = 0
    #         d_mola_modulo_entrada = 0
    #         d_mola_curva_entrada = 0

    #         momento_mola_conj_entrada = (peso_conj_entrada * d_mola_conj_entrada * 9.8) / 1000
    #         momento_mola_modulo_entrada = (peso_modulo_entrada * d_mola_modulo_entrada * 9.8) / 1000
    #         momento_mola_curva_entrada = (peso_curva_entrada * d_mola_curva_entrada * 9.8) / 1000

    #         soma_pesos_1 = peso_conj_entrada + peso_modulo_entrada + peso_curva_entrada
    #         soma_momentos_1 = momento_mola_conj_entrada + momento_mola_curva_entrada + momento_mola_modulo_entrada
    #         ###########################################################################
    #         # # Segunda parte
    #         primario_bc = f'PRIMÁRIO {diametro} POL AC'

    #         distancia_tubo = comprimentos[primario_bc]

    #         qtde_tubo = primario - (comprimento_1 + distancia_tubo)

    #         peso_primario = pesos.get(primario_bc)
    #         peso_flange_1 = pesos.get(primeira_flange_primario)
    #         peso_tubo = pesos.get(tubo_primario)*qtde_tubo
    #         peso_flange_2 = pesos.get(segunda_flange_primario)

    #         d_mola_primario = pri.get(primario_bc)
    #         d_mola_acessorio_1 = ((comprimento_1*1000)/2) + (comprimentos.get(primario_bc)*1000)
    #         d_mola_flange_1 = (comprimento_1*1000) + (comprimentos.get(primario_bc)*1000)
    #         d_mola_tubo = (((primario * 1000) - d_mola_flange_1)/2) + d_mola_flange_1
    #         d_mola_flange_2 = primario * 1000

    #         momento_mola_primario= (peso_primario * d_mola_primario * 9.8) / 1000
    #         momento_mola_acessorio_1 = (peso_acessorio_1  * d_mola_acessorio_1 * 9.8) / 1000
    #         momento_mola_flange_1 = (peso_flange_1 * d_mola_flange_1 * 9.8) / 1000
    #         momento_mola_tubo = (peso_tubo * d_mola_tubo * 9.8) / 1000
    #         momento_mola_flange_2 = (peso_flange_2 * d_mola_flange_2 * 9.8) / 1000

    #         soma_pesos_2 = peso_primario + peso_acessorio_1 + peso_flange_1 + peso_tubo + peso_flange_2
    #         soma_momentos_2 = momento_mola_primario + momento_mola_acessorio_1 + momento_mola_flange_1 + momento_mola_tubo + momento_mola_flange_2
    #         ###########################################################################
    #         # # Terceira parte
    #         qtde_tubo_secundario = secundario - 0.21

    #         peso_flange_3 = pesos.get(primeira_flange_secundario)
    #         peso_tubo_secundario = pesos.get(tubo_secundario)*qtde_tubo_secundario
    #         peso_flange_4 = pesos.get(segunda_flange_secundario)

    #         d_mola_secundario = primario

    #         momento_mola_flange_3 = (peso_flange_3 * d_mola_secundario * 9.8) / 1000
    #         momento_mola_tubo_secundario = (peso_tubo_secundario * d_mola_secundario * 9.8) / 1000
    #         momento_mola_flange_4 = (peso_flange_2 * d_mola_secundario * 9.8) / 1000

    #         soma_pesos_3 = peso_flange_3 + peso_tubo_secundario + peso_flange_4
    #         soma_momentos_3 = momento_mola_flange_3 + momento_mola_tubo_secundario + momento_mola_flange_4
    #         ###########################################################################
    #         # # Quarta parte
    #         qtde_mergulhador = mergulhador - 0.11

    #         peso_flange_mergulhador = pesos.get(flange_mergulhador)
    #         peso_tubo_mergulhador = pesos.get(tubo_mergulhador)*qtde_mergulhador

    #         d_mola_mergulhador = primario * 1000

    #         momento_mola_flange_mergulhador = (peso_flange_mergulhador * d_mola_mergulhador * 9.8) / 1000
    #         momento_mola_tubo_mergulhador = (peso_tubo_mergulhador * d_mola_mergulhador * 9.8) / 1000
    #         momento_mola_acessorio_2 = (peso_acessorio_2 * d_mola_mergulhador * 9.8) / 1000

    #         soma_momentos_6 = momento_mola_flange_mergulhador + momento_mola_tubo_mergulhador + momento_mola_acessorio_2
    #         soma_pesos_6 = peso_flange_mergulhador + peso_tubo_mergulhador + peso_acessorio_2
    #         ###########################################################################
    #         # # Quinta parte
    #         peso_primeira_junta = pesos.get(primeira_junta) * qtde_junta_1
    #         peso_segunda_junta = pesos.get(segunda_junta) * qtde_junta_2

    #         d_mola_junta = d_mola_mergulhador

    #         momento_mola_primeira_junta = (peso_primeira_junta * d_mola_junta * 9.8) / 1000
    #         momento_mola_segunda_junta = (peso_segunda_junta * d_mola_junta * 9.8) / 1000

    #         soma_pesos_4 = peso_primeira_junta + peso_segunda_junta
    #         soma_momentos_4 = momento_mola_primeira_junta + momento_mola_segunda_junta
    #         ###########################################################################
    #         ###########################################################################
    #         # # Finalização
    #         peso_total = soma_pesos_1 + soma_pesos_2 + soma_pesos_3 + soma_pesos_4  + soma_pesos_6
    #         momento_total = soma_momentos_1 + soma_momentos_2 + soma_momentos_3 + soma_momentos_4 + soma_momentos_6
    #         max_2 = mola_4_12_max.get(furos)
    #         max_1 = mola_4_max.get(furos)
    #     if mola_selecao == "MOLA 4 E 12":
    #         ajuste_traseiro = 300-(244-(((momento_total/9.8*1000)/((max_1)*(1.82)))-1012+(max_2)+244))-15.8
    #         voltas = ajuste_traseiro/3
    #     if mola_selecao == "MOLA 4":
    #         ajuste_traseiro = 300-(244-(((momento_total/9.8*1000)/((max_1)*(1.13)))-1012+(max_2)+244))-15.8
    #         voltas = ajuste_traseiro/3

    #     resultado_box.delete("1.0", tk.END)
    #     resultado = '-------------------------------------------------------------------','Entradas',f'Braço: {inicio}',f'Tipo de braço: {modelo}',f'Diâmetro: {diametro} Polegadas',f'Material: {material_braco}',f'Primario: {primario}',f'Secundário: {secundario}',f'Mergulhador: {mergulhador}',f'Válvula: {acessorio_1}',f'Terminal do mergulhador: {acessorio_2}','-------------------------------------------------------------------','Dimensionamento da Mola',f'PESO TOTAL [kg]: {round(peso_total,2)}',f'MOMENTO MOLA [Nm]: {round(momento_total,2)}',f'MOMENTO ENTRADA [Nm]: {round(momento_total,2)}',f'CÓDIGO DO BC: {modelo}-{diametro}{material_bc}{terminacao_do_braco}',f'Mola Selecionada: {mola_selecionada}',f'Furos: {furos}',f'AJUSTE TRASEIRO (cota a) [mm]: {round(ajuste_traseiro,2)}',f'VOLTAS DE AJUSTE: {round(voltas,2)}','-------------------------------------------------------------------',"1. CJ. DE ENTRADA DO BC",f'{"Descrição":<20} | {"QTDE.":<6} | {"Peso (Kg)":<12} | {"Distância da mola":<20} | {"Momento da mola":<16} | {"Distância da entrada":<20} | {"Momento da entrada":<16}',f'{conj_entrada:<20} | {qtde_padrao:<6} | {peso_conj_entrada:<12} | {d_mola_conj_entrada:<20} | {momento_mola_conj_entrada:<16} | {d_mola_conj_entrada:<20} | {momento_mola_conj_entrada:<16}',f'{modulo_entrada:<20} | {qtde_modulo_entrada:<6} | {peso_modulo_entrada:<12} | {d_mola_modulo_entrada:<20} | {momento_mola_modulo_entrada:<16} | {d_mola_modulo_entrada:<20} | {momento_mola_modulo_entrada:<16}',f'{curva_entrada:<20} | {qtde_padrao:<6} | {peso_curva_entrada:<12} | {d_mola_curva_entrada:<20} | {momento_mola_curva_entrada:<16} | {d_mola_curva_entrada:<20} | {momento_mola_curva_entrada:<16}','-------------------------------------------------------------------',"2. PRIMÁRIO DO BC",f'{"Descrição":<20} | {"QTDE.":<6} | {"Peso (Kg)":<12} | {"Distância da mola":<20} | {"Momento da mola":<16} | {"Distância da entrada":<20} | {"Momento da entrada":<16}',f'{primario_bc:<20} | {qtde_padrao:<6} | {peso_primario:<12} | {d_mola_primario:<20} | {round(momento_mola_primario,2):<16} | {d_mola_primario:<20} | {round(momento_mola_primario,2):<16}',f'{acessorio_1:<20} | {qtde_padrao:<6} | {peso_acessorio_1:<12} | {d_mola_acessorio_1:<20} | {round(momento_mola_acessorio_1,2):<16} | {d_mola_acessorio_1:<20} | {round(momento_mola_acessorio_1,2):<16}',f'{primeira_flange_primario:<20} | {qtde_padrao:<6} | {peso_flange_1:<12} | {d_mola_flange_1:<20} | {round(momento_mola_flange_1,2):<16} | {d_mola_flange_1:<20} | {round(momento_mola_flange_1,2):<16}',f'{tubo_primario:<20} | {round(qtde_tubo,2):<6} | {round(peso_tubo,2):<12} | {d_mola_tubo:<20} | {round(momento_mola_tubo):<16} | {d_mola_tubo:<20} | {round(momento_mola_tubo):<16}',f'{segunda_flange_primario:<20} | {qtde_padrao:<6} | {peso_flange_2:<12} | {d_mola_flange_2:<20} | {round(momento_mola_flange_2,2):<16} | {d_mola_flange_2:<20} | {round(momento_mola_flange_2,2):<16}','-------------------------------------------------------------------','3. SECUNDÁRIO DO BC',f'{"Descrição":<20} | {"QTDE.":<6} | {"Peso (Kg)":<12} | {"Distância da mola":<20} | {"Momento da mola":<16} | {"Distância da entrada":<20} | {"Momento da entrada":<16}',f'{primeira_flange_secundario:<20} | {qtde_padrao:<6} | {peso_flange_3:<12} | {d_mola_secundario:<20} | {round(momento_mola_flange_3):<16} | {d_mola_secundario:<20} | {round(momento_mola_flange_3):<16}',f'{tubo_secundario:<20} | {round(qtde_tubo_secundario,2):<6} | {round(peso_tubo_secundario,2):<12} | {d_mola_secundario:<20} | {round(momento_mola_tubo_secundario):<16} | {d_mola_secundario:<20} | {round(momento_mola_tubo_secundario):<16}',f'{segunda_flange_secundario:<20} | {qtde_padrao:<6} | {peso_flange_4:<12} | {d_mola_secundario:<20} | {round(momento_mola_flange_4):<16} | {d_mola_secundario:<20} | {round(momento_mola_flange_4):<16}','-------------------------------------------------------------------','4. MERGULHADOR DO BC',f'{"Descrição":<20} | {"QTDE.":<6} | {"Peso (Kg)":<12} | {"Distância da mola":<20} | {"Momento da mola":<16} | {"Distância da entrada":<20} | {"Momento da entrada":<16}',f'{flange_mergulhador:<20} | {qtde_padrao:<6} | {peso_flange_mergulhador:<12} | {d_mola_mergulhador:<20} | {round(momento_mola_flange_mergulhador,2):<16} | {d_mola_mergulhador:<20} | {round(momento_mola_flange_mergulhador,2):<16}',f'{tubo_mergulhador:<20} | {round(qtde_mergulhador,2):<6} | {peso_tubo_mergulhador:<12} | {d_mola_mergulhador:<20} | {round(momento_mola_tubo_mergulhador,2):<16} | {d_mola_mergulhador:<20} | {round(momento_mola_tubo_mergulhador,2):<16}',f'{acessorio_2:<20} | {qtde_padrao:<6} | {peso_acessorio_2:<12} | {d_mola_mergulhador:<20} | {round(momento_mola_acessorio_2,2):<16} | {d_mola_mergulhador:<20} | {round(momento_mola_acessorio_2,2):<16}','-------------------------------------------------------------------','5. JGs INTERMEDIÁRIAS',f'{"Descrição":<20} | {"QTDE.":<6} | {"Peso (Kg)":<12} | {"Distância da mola":<20} | {"Momento da mola":<16} | {"Distância da entrada":<20} | {"Momento da entrada":<16}',f'{primeira_junta:<20} | {qtde_junta_1:<6} | {d_mola_junta:<12} | {round(momento_mola_primeira_junta,2):<20} | {d_mola_junta:<16} | {round(momento_mola_primeira_junta,2):<20} | {d_mola_junta:<16}',f'{segunda_junta:<20} | {qtde_junta_2:<6} | {d_mola_junta:<12} | {round(momento_mola_segunda_junta,2):<20} | {d_mola_junta:<16} | {round(momento_mola_segunda_junta,2):<20} | {d_mola_junta:<16}','-------------------------------------------------------------------'
    #     resultado_box.insert(tk.END, resultado)



    # tk.Label(janela, text="TOP ou BOT? ").grid(row=1, column=0, pady=5, sticky="e")
    # entry_inicio_var = tk.StringVar()
    # entry_inicio_menu = ttk.Combobox(janela, textvariable=entry_inicio_var,values=["TOP", "BOT"],width=30)
    # entry_inicio_menu.grid(row=1, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe o modelo do braço: ").grid(row=2, column=0, pady=5, sticky="e")
    # entry_tipo_de_braco_var = tk.StringVar()
    # entry_tipo_de_braco_menu = ttk.Combobox(janela, textvariable=entry_tipo_de_braco_var,values=["BTA","BTC","BTE","BTG","BBI","BBJ","BBJ","BBK","BBL","BBN", "BBP"],width=30)
    # entry_tipo_de_braco_menu.grid(row=2, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe a terminação do braço: ").grid(row=3, column=0, pady=5, sticky="e")
    # entry_terminacao_var = tk.Entry(janela, width=30)
    # entry_terminacao_var.grid(row=3, column=1, pady=5)
    # tk.Label(janela, text="Informe o diâmetro do braço: ").grid(row=4, column=0, pady=5, sticky="e")
    # entry_diametro_var = tk.Entry(janela, width=30)
    # entry_diametro_var.grid(row=4, column=1, pady=5)
    # tk.Label(janela, text="Informe o material do braço: ").grid(row=5, column=0, pady=5, sticky="e")
    # entry_material_bc_var = tk.StringVar()
    # entry_material_bc_menu = ttk.Combobox(janela, textvariable=entry_material_bc_var,values=["2CA","3CA","2CC","3CC","2I4","3I4","2I6","3I6","24A","34A"], width=30)
    # entry_material_bc_menu.grid(row=5, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe o tamanho do primário em metros (separador = . ): ").grid(row=6, column=0, pady=5, sticky="e")
    # entry_primario_var = tk.Entry(janela, width=30)
    # entry_primario_var.grid(row=6, column=1, pady=5)
    # tk.Label(janela, text="Informe o tamanho do secundário em metros (separador = . ): ").grid(row=7, column=0, pady=5, sticky="e")
    # entry_secundario_var = tk.Entry(janela, width=30)
    # entry_secundario_var.grid(row=7, column=1, pady=5)
    # tk.Label(janela, text="Informe o tamanho do mergulhador em metros (separador = . ): ").grid(row=8, column=0, pady=5, sticky="e")
    # entry_mergulhador_var = tk.Entry(janela, width=30)
    # entry_mergulhador_var.grid(row=8, column=1, pady=5)
    # tk.Label(janela, text="Informe o acessório do braço: ").grid(row=9, column=0, pady=5, sticky="e")
    # entry_acessorio_1_var = tk.StringVar()
    # entry_acessorio_1_menu = ttk.Combobox(janela, textvariable=entry_acessorio_1_var,values=["VALVULA DEADMAN", "VALVULA ESFERA BP 4 POL", "VALVULA ESFERA BP 3 POL", "VALVULA ESFERA WF 4 POL", "VALVULA ESFERA WF 3 POL", "SEM VALVULA"], width=30)
    # entry_acessorio_1_menu.grid(row=9, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe o terminal do mergulhador: ").grid(row=10, column=0, pady=5, sticky="e")
    # entry_acessorio_2_var = tk.StringVar()
    # entry_acessorio_2_menu = ttk.Combobox(janela, textvariable=entry_acessorio_2_var,values=["DEFLETOR 4 POL AL", "DEFLETOR 3 POL AL", "DEFLETOR 4 POL AC", "DEFLETOR 3 POL AC", "CHANFRO"], width=30)
    # entry_acessorio_2_menu.grid(row=10, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe a primeira flange do primário: ").grid(row=11, column=0, pady=5, sticky="e")
    # entry_primeira_flange_primario_var = tk.StringVar()
    # entry_primeira_flange_primario_menu = ttk.Combobox(janela, textvariable=entry_primeira_flange_primario_var,values=["FLG TTMA 4 POL ACO", "FLG TTMA 4 POL AL", "FLG TTMA 3 POL ACO", "FLG TTMA 3 POL AL", "FLG ANSI 4 POL 150 AC SO", "FLG ANSI 3 POL 150 AC SO", "FLG ANSI 4 POL 150 AL SO", "FLG ANSI 3 POL 150 AL SO", "FLG JG 4 POL ACO", "FLG JG 3 POL ACO", "SEM FLANGE"], width=30)
    # entry_primeira_flange_primario_menu.grid(row=11, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe o tubo do primário: ").grid(row=12, column=0, pady=5, sticky="e")
    # entry_tubo_primario_var = tk.StringVar()
    # entry_tubo_primario_menu = ttk.Combobox(janela, textvariable=entry_tubo_primario_var,values=["4 POL SCH 10 AL", "4 POL SCH 40 AL", "4 POL SCH 10 ACO", "4 POL SCH 40 ACO", "3 POL SCH 10 AL", "3 POL SCH 40 AL", "3 POL SCH 10 ACO","3 POL SCH 40 ACO"], width=30)
    # entry_tubo_primario_menu.grid(row=12, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe a segunda flange do primário: ").grid(row=13, column=0, pady=5, sticky="e")
    # entry_segunda_flange_primario_var = tk.StringVar()
    # entry_segunda_flange_primario_menu = ttk.Combobox(janela, textvariable=entry_segunda_flange_primario_var,values=["FLG TTMA 4 POL ACO", "FLG TTMA 4 POL AL", "FLG TTMA 3 POL ACO", "FLG TTMA 3 POL AL", "FLG ANSI 4 POL 150 AC SO", "FLG ANSI 3 POL 150 AC SO", "FLG ANSI 4 POL 150 AL SO", "FLG ANSI 3 POL 150 AL SO", "FLG JG 4 POL ACO", "FLG JG 3 POL ACO", "SEM FLANGE"], width=30)
    # entry_segunda_flange_primario_menu.grid(row=13, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe a primeira flange do secundário: ").grid(row=14, column=0, pady=5, sticky="e")
    # entry_primeira_flange_secundario_var = tk.StringVar()
    # entry_primeira_flange_secundario_menu = ttk.Combobox(janela, textvariable=entry_primeira_flange_secundario_var,values=["FLG TTMA 4 POL ACO", "FLG TTMA 4 POL AL", "FLG TTMA 3 POL ACO", "FLG TTMA 3 POL AL", "FLG ANSI 4 POL 150 AC SO", "FLG ANSI 3 POL 150 AC SO", "FLG ANSI 4 POL 150 AL SO", "FLG ANSI 3 POL 150 AL SO", "FLG JG 4 POL ACO", "FLG JG 3 POL ACO", "SEM FLANGE"], width=30)
    # entry_primeira_flange_secundario_menu.grid(row=14, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe o tubo do secundário: ").grid(row=15, column=0, pady=5, sticky="e")
    # entry_tubo_secundario_var = tk.StringVar()
    # entry_tubo_secundario_menu = ttk.Combobox(janela, textvariable=entry_tubo_secundario_var,values=["4 POL SCH 10 AL", "4 POL SCH 40 AL", "4 POL SCH 10 ACO", "4 POL SCH 40 ACO", "3 POL SCH 10 AL", "3 POL SCH 40 AL", "3 POL SCH 10 ACO","3 POL SCH 40 ACO"], width=30)
    # entry_tubo_secundario_menu.grid(row=15, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe a segunda flange do secundário: ").grid(row=16, column=0, pady=5, sticky="e")
    # entry_segunda_flange_secundario_var = tk.StringVar()
    # entry_segunda_flange_secundario_menu = ttk.Combobox(janela, textvariable=entry_segunda_flange_secundario_var,values=["FLG TTMA 4 POL ACO", "FLG TTMA 4 POL AL", "FLG TTMA 3 POL ACO", "FLG TTMA 3 POL AL", "FLG ANSI 4 POL 150 AC SO", "FLG ANSI 3 POL 150 AC SO", "FLG ANSI 4 POL 150 AL SO", "FLG ANSI 3 POL 150 AL SO", "FLG JG 4 POL ACO", "FLG JG 3 POL ACO", "SEM FLANGE"], width=30)
    # entry_segunda_flange_secundario_menu.grid(row=16, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe a flange do mergulhador: ").grid(row=17, column=0, pady=5, sticky="e")
    # entry_flange_mergulhador_var = tk.StringVar()
    # entry_flange_mergulhador_menu = ttk.Combobox(janela, textvariable=entry_flange_mergulhador_var,values=["FLG TTMA 4 POL ACO", "FLG TTMA 4 POL AL", "FLG TTMA 3 POL ACO", "FLG TTMA 3 POL AL", "FLG ANSI 4 POL 150 AC SO", "FLG ANSI 3 POL 150 AC SO", "FLG ANSI 4 POL 150 AL SO", "FLG ANSI 3 POL 150 AL SO", "FLG JG 4 POL ACO", "FLG JG 3 POL ACO", "SEM FLANGE"], width=30)
    # entry_flange_mergulhador_menu.grid(row=17, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe o tubo do mergulhador: ").grid(row=18, column=0, pady=5, sticky="e")
    # entry_tubo_mergulhador_var = tk.StringVar()
    # entry_tubo_mergulhador_menu = ttk.Combobox(janela, textvariable=entry_tubo_mergulhador_var,values=["4 POL SCH 10 AL", "4 POL SCH 40 AL", "4 POL SCH 10 ACO", "4 POL SCH 40 ACO", "3 POL SCH 10 AL", "3 POL SCH 40 AL", "3 POL SCH 10 ACO","3 POL SCH 40 ACO"], width=30)
    # entry_tubo_mergulhador_menu.grid(row=18, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe a primeira junta do braço: ").grid(row=19, column=0, pady=5, sticky="e")
    # entry_primeira_junta_var = tk.StringVar()
    # entry_primeira_junta_menu = ttk.Combobox(janela, textvariable=entry_primeira_junta_var,values=["JM40TT-42A", "JM40TT-32A", "JM40SS-42A", "JM40SS-32A"], width=30)
    # entry_primeira_junta_menu.grid(row=19, column=1, padx=5, pady=5)
    # tk.Label(janela, text="Informe a quantidade: ").grid(row=20, column=0, pady=5, sticky="e")
    # entry_qtde_junta_1_var = tk.Entry(janela, width=30)
    # entry_qtde_junta_1_var.grid(row=20, column=1, pady=5)
    # tk.Label(janela, text="Informe a segunda junta do braço: ").grid(row=21, column=0, pady=5, sticky="e")
    # entry_segunda_junta_var = tk.StringVar()
    # entry_segunda_junta_menu = ttk.Combobox(janela, textvariable=entry_segunda_junta_var,values=["JM40TT-42A", "JM40TT-32A", "JM40SS-42A", "JM40SS-32A"], width=30)
    # entry_segunda_junta_menu.grid(row=21, column=1, pady=5)
    # tk.Label(janela, text="Informe a quantidade: ").grid(row=22, column=0, pady=5, sticky="e")
    # entry_qtde_junta_2_var = tk.Entry(janela, width=30)
    # entry_qtde_junta_2_var.grid(row=22, column=1, pady=5)
    # tk.Label(janela, text="Selecione a mola: ").grid(row=23, column=0, pady=5, sticky="e")
    # entry_mola_selecao_var = tk.StringVar()
    # entry_mola_selecao_menu = ttk.Combobox(janela, textvariable=entry_mola_selecao_var,values=["MOLA 4 E 12","MOLA 4"], width=30)
    # entry_mola_selecao_menu.grid(row=23, column=1, pady=5)
    # tk.Label(janela, text="Selecione a quantidade de furos (Até 5): ").grid(row=24, column=0, pady=5, sticky="e")
    # entry_furos_var = tk.Entry(janela, width=30)
    # entry_furos_var.grid(row=24, column=1, pady=5)

    # btn_calcular = tk.Button(janela, text="Calcular", command=calcular_bc)
    # btn_calcular.grid(row=25, columnspan=2, pady=20)

    # # Caixa de Resultado
    # tk.Label(janela, text="Resultado:").grid(row=26, column=0, columnspan=2)
    # resultado_box = tk.Text(janela, height=60, width=60)
    # resultado_box.grid(row=26, column=0, columnspan=2, pady=10)

# Tela principal
app = tk.Tk()
app.title("Menu Principal")
app.geometry("400x400")

# Botões da tela inicial
tk.Label(app, text="Selecione o calculo desejado:", font=("Arial", 14)).pack(pady=10)

tk.Button(app, text="Calcular Pesos", command=abrir_tabela_pesos, width=20, bg="green", fg="white").pack(pady=10)
tk.Button(app, text="Calcular Mangueiras", command=abrir_mangueiras, width=20, bg="blue", fg="white").pack(pady=10)
# tk.Button(app, text="Calcular Braços", command=abrir_bracos, width=20, bg="pink", fg="black").pack(pady=10)
# tk.Button(app, text="Calcular Sucção", command=abrir_succao, width=20).pack(pady=10)
# tk.Button(app, text="Calcular Perda de carga", command=abrir_perda_de_carga, width=20).pack(pady=10)

tk.Button(app, text="Sair", command=app.quit, bg="red", fg="white", width=20).pack(pady=10)

app.mainloop()
