@echo off
setlocal

set VENV_DIR=build_env

echo Creating isolated build environment...
if not exist "%VENV_DIR%\Scripts\python.exe" (
    python -m venv "%VENV_DIR%"
)

echo Installing dependencies into build environment...
"%VENV_DIR%\Scripts\python.exe" -m pip install --upgrade pip
"%VENV_DIR%\Scripts\python.exe" -m pip install -r requirements.txt pyinstaller

echo Building GUI version (recommended for double-click users)...
"%VENV_DIR%\Scripts\python.exe" -m PyInstaller --noconfirm --onefile --windowed ^
  --name feixiaoqiu-to-paimonmoe-gui ^
  --add-data "data;data" ^
  gui.py

echo Building Paimon.moe Excel converter console exe...
"%VENV_DIR%\Scripts\python.exe" -m PyInstaller --noconfirm --onefile --console ^
  --name feixiaoqiu-to-paimonmoe ^
  --add-data "data;data" ^
  converter.py

echo Building UIGF exporter console exe...
"%VENV_DIR%\Scripts\python.exe" -m PyInstaller --noconfirm --onefile --console ^
  --name feixiaoqiu-to-uigf ^
  --add-data "data;data" ^
  to_uigf.py

echo.
echo Done. Check the dist folder:
echo   dist\feixiaoqiu-to-paimonmoe-gui.exe   (double-click this one)
echo   dist\feixiaoqiu-to-paimonmoe.exe
echo   dist\feixiaoqiu-to-uigf.exe
endlocal
