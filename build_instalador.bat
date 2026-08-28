@echo off
REM Gera o ClassicSports_DIFAL_Instalador.exe.
REM O instalador baixa os arquivos do GitHub (Marcus-Carvalho/classicsports-difal),
REM entao o repositorio precisa estar PUBLICO e com os arquivos publicados antes.
cd /d "%~dp0"

python -m pip install --upgrade pyinstaller

python -m PyInstaller --onefile --windowed --clean --icon=icon.ico ^
  --name=ClassicSports_DIFAL_Instalador ^
  instalador_gui.py

echo.
echo Pronto: dist\ClassicSports_DIFAL_Instalador.exe
echo Antes de distribuir, confira se este link abre no navegador:
echo https://raw.githubusercontent.com/Marcus-Carvalho/classicsports-difal/main/versao.json
pause
