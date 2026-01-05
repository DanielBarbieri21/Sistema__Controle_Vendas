"""
Script para gerar executável do Sistema de Controle de Vendas
"""
import PyInstaller.__main__
import os
import sys

# Nome do executável
app_name = "SistemaControleVendas"

# Arquivos e pastas a incluir
arquivos_adicional = [
    ("icon.png", "."),  # Incluir o ícone
]

# Comando PyInstaller
args = [
    "app.py",  # Arquivo principal
    "--name", app_name,
    "--onefile",  # Gerar um único arquivo executável
    "--windowed",  # Sem console (interface gráfica)
    "--icon", "icon.png",  # Ícone do executável
    "--add-data", "icon.png;.",  # Incluir ícone nos dados
    "--hidden-import", "PySide6",
    "--hidden-import", "PySide6.QtCore",
    "--hidden-import", "PySide6.QtGui",
    "--hidden-import", "PySide6.QtWidgets",
    "--hidden-import", "pandas",
    "--hidden-import", "openpyxl",
    "--hidden-import", "sqlalchemy",
    "--hidden-import", "reportlab",
    "--hidden-import", "sqlalchemy.dialects.sqlite",
    "--hidden-import", "sqlalchemy.dialects.sqlite.base",
    "--collect-all", "openpyxl",  # Incluir todos os módulos do openpyxl
    "--collect-all", "pandas",  # Incluir todos os módulos do pandas
    "--noconfirm",  # Sobrescrever arquivos existentes
    "--clean",  # Limpar cache antes de construir
]

# Adicionar arquivos adicionais
for arquivo, destino in arquivos_adicional:
    args.extend(["--add-data", f"{arquivo};{destino}"])

print("=" * 60)
print("Gerando executável do Sistema de Controle de Vendas")
print("=" * 60)
print(f"\nComando: pyinstaller {' '.join(args)}\n")

try:
    PyInstaller.__main__.run(args)
    print("\n" + "=" * 60)
    print("Executavel gerado com sucesso!")
    print(f"Localizacao: dist\\{app_name}.exe")
    print("=" * 60)
except Exception as e:
    print(f"\nErro ao gerar executavel: {e}")
    sys.exit(1)

