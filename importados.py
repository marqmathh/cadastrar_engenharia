import tkinter as tk
from tkinter import ttk, messagebox
from sqlalchemy import create_engine, Column, Integer, String, Float
from sqlalchemy.orm import sessionmaker, declarative_base
import requests

# Configuração do banco de dados
db = create_engine('sqlite:///Z:\\ISO 9000 - SGQ\\12-SISTEMA\\Sistema\\banco de dados\\Importados.db')
Session = sessionmaker(bind=db)
db_session = Session()
Base1 = declarative_base()

# Definição do modelo do banco de dados
class Produtos(Base1):
    __tablename__ = 'Produtos'
    Id = Column('Id', Integer, primary_key=True, autoincrement=True)
    Empresa = Column('Empresa', String)
    Produto = Column('Produto', String)
    Descricao = Column('Descrição', String)
    preco = Column('preço', Float)

# Função para obter cotações
def obter_cotacoes():
    cotacoes = requests.get('https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL')
    cotacoes = cotacoes.json()
    return float(cotacoes['USDBRL']['bid']), float(cotacoes['EURBRL']['bid'])

cotacao_dolar, cotacao_euro = obter_cotacoes()

# Função para buscar produtos no banco de dados
def buscar_produtos():
    nome_produto = entrada_nome.get().strip()  # Remover espaços em branco extras
    if not nome_produto:
        messagebox.showwarning("Aviso", "Por favor, insira ao menos uma parte do nome do produto.")
        return

    filtro = filtro_combobox.get()
    nome_final_produto = f'%{nome_produto}%'
    produtos_encontrados = []

    if filtro == 'Código':
        produtos_encontrados = db_session.query(Produtos).filter(Produtos.Produto.like(nome_final_produto)).all()
    elif filtro == 'Empresa':
        produtos_encontrados = db_session.query(Produtos).filter(Produtos.Empresa.like(nome_final_produto)).all()
    else:  # Filtro por Descrição
        produtos_encontrados = db_session.query(Produtos).filter(Produtos.Descricao.like(nome_final_produto)).all()

    if not produtos_encontrados:
        messagebox.showerror("Erro", f"Nenhum produto encontrado com o termo '{nome_produto}'.")
        return

    # Preenchendo os resultados na tabela
    for row in resultados_tree.get_children():
        resultados_tree.delete(row)

    for produto in produtos_encontrados:
        preco_final = calcular_preco(produto)
        resultados_tree.insert('', tk.END, values=(
            produto.Id, produto.Empresa, produto.Produto, produto.Descricao, f"R$ {preco_final:.2f}"
        ))

def calcular_preco(produto):
    if produto.Empresa == 'OPW':
        return round(produto.preco * cotacao_dolar * 1.10 * 1.65 * 0.534, 2)
    return round(produto.preco * cotacao_dolar * 1.10 * 1.65, 2)

# Interface gráfica com Tkinter
app = tk.Tk()
app.title("Consulta de Produtos Importados")

# Layout da aplicação
frame_filtro = tk.Frame(app)
frame_filtro.pack(pady=10)

tk.Label(frame_filtro, text="Nome do Produto:").grid(row=0, column=0, padx=5, pady=5)
entrada_nome = tk.Entry(frame_filtro, width=30)
entrada_nome.grid(row=0, column=1, padx=5, pady=5)

tk.Label(frame_filtro, text="Filtro:").grid(row=0, column=2, padx=5, pady=5)
filtro_combobox = ttk.Combobox(frame_filtro, values=["Código", "Empresa", "Descrição"])
filtro_combobox.grid(row=0, column=3, padx=5, pady=5)
filtro_combobox.set("Código")

tk.Button(frame_filtro, text="Buscar", command=buscar_produtos).grid(row=0, column=4, padx=5, pady=5)

# Tabela para exibir resultados
resultados_tree = ttk.Treeview(app, columns=("ID", "Empresa", "Produto", "Descrição", "Preço"), show='headings')
resultados_tree.heading("ID", text="ID")
resultados_tree.heading("Empresa", text="Empresa")
resultados_tree.heading("Produto", text="Produto")
resultados_tree.heading("Descrição", text="Descrição")
resultados_tree.heading("Preço", text="Preço")

resultados_tree.column("ID", width=50, anchor=tk.CENTER)
resultados_tree.column("Empresa", width=100, anchor=tk.W)
resultados_tree.column("Produto", width=150, anchor=tk.W)
resultados_tree.column("Descrição", width=250, anchor=tk.W)
resultados_tree.column("Preço", width=100, anchor=tk.E)

resultados_tree.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# Rodar o aplicativo
app.mainloop()
