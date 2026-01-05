# Instruções para Gerar o Executável

## Pré-requisitos

1. Python 3.8 ou superior instalado
2. Todas as dependências instaladas:
   ```bash
   pip install -r requirements.txt
   ```

## Gerar o Executável

### Opção 1: Usando o script batch (Windows)
```bash
build.bat
```

### Opção 2: Usando o script Python
```bash
python build.py
```

### Opção 3: Comando manual
```bash
pyinstaller --name SistemaControleVendas --onefile --windowed --icon=icon.png --add-data "icon.png;." app.py
```

## Resultado

O executável será gerado na pasta `dist` com o nome:
- **SistemaControleVendas.exe**

## Observações

- O executável aceita **qualquer arquivo .xlsx**, não apenas "Vendas Mercado.xlsx"
- O ícone `icon.png` será incluído no executável
- O banco de dados `vendas.db` será criado na mesma pasta do executável quando você importar dados pela primeira vez

## Distribuição

Para distribuir o aplicativo, você precisa apenas do arquivo `.exe` gerado. Não é necessário incluir os arquivos Python ou as dependências.

