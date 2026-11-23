@echo off
cd /d "%~dp0"
echo ===== Build exe (PyInstaller) =====
echo.
echo   1 = Folder version (faster start, recommended)
echo   2 = Single exe version (-F, slower start)
echo.
set /p MODE=Choose 1 or 2 and press Enter: 
set OPT=
if "%MODE%"=="2" set OPT=--onefile
echo.
echo [1/2] Installing / updating PyInstaller (Python 3.15 needs 6.21 or newer)
python -m pip install --upgrade pyinstaller
if errorlevel 1 goto fail
echo.
echo [2/2] Building. The first time takes a while, please wait.
python -m PyInstaller --noconfirm --clean --windowed %OPT% --name ScienceFairTool --collect-submodules modes --collect-data sv_ttk --collect-all plotly main.py
if errorlevel 1 goto fail
echo.
if "%MODE%"=="2" (
  echo Done. Your exe is dist\ScienceFairTool.exe
) else (
  echo Done. Copy the WHOLE folder dist\ScienceFairTool to another PC and run ScienceFairTool.exe
)
pause
exit /b 0
:fail
echo.
echo Build failed. Please copy the error messages above and send them to Claude.
pause
exit /b 1
