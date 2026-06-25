@echo off
title Installation - QHSE Manager Pro
echo Installation des dependances Python...
python --version
if errorlevel 1 (
    echo Python n'est pas installe ou n'est pas accessible.
    echo Installe Python depuis https://www.python.org/downloads/
    pause
    exit /b
)
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
echo.
echo Installation terminee.
pause
