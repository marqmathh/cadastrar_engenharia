Criar ambiente virtual:
python -m venv venv

Ativar ambiente virtual:
.\venv\Scripts\activate

Criar aplicação:
pyinstaller --onefile -w atualizacao-Planilhas-tinker.py

Deletar a pasta build e o arquivo que gerou .spec
dentro da pasta dist está o nosso executavel