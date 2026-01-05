@echo off
echo ============================================================
echo Gerando executável do Sistema de Controle de Vendas
echo ============================================================
echo.

REM Verificar se PyInstaller está instalado
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller não encontrado. Instalando...
    pip install pyinstaller
)

REM Executar o script de build
python build.py

if errorlevel 1 (
    echo.
    echo Erro ao gerar executável!
    pause
    exit /b 1
)

echo.
echo ============================================================
echo Executável gerado com sucesso!
echo Localização: dist\SistemaControleVendas.exe
echo ============================================================
pause

