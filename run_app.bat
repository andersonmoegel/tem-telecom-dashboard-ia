@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

REM Evita pergunta inicial de e-mail/telemetria do Streamlit
set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

title Telecom Dashboard IA - Inicializador

echo ============================================================
echo  Telecom Dashboard IA - Inicializador
echo ============================================================
echo.

REM ------------------------------------------------------------
REM 1) Localiza Python instalado
REM ------------------------------------------------------------
set "PYTHON_CMD="

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 --version >nul 2>nul
    if !errorlevel!==0 set "PYTHON_CMD=py -3"
)

if not defined PYTHON_CMD (
    where python >nul 2>nul
    if !errorlevel!==0 (
        python --version >nul 2>nul
        if !errorlevel!==0 set "PYTHON_CMD=python"
    )
)

if not defined PYTHON_CMD (
    echo [INFO] Python nao encontrado neste computador.
    echo [INFO] Tentando instalar Python automaticamente via winget...
    echo.

    where winget >nul 2>nul
    if errorlevel 1 (
        echo [ERRO] O winget nao foi encontrado no Windows.
        echo.
        echo Instale o Python manualmente em:
        echo https://www.python.org/downloads/
        echo.
        echo Durante a instalacao, marque a opcao: Add python.exe to PATH
        pause
        exit /b 1
    )

    winget install --id Python.Python.3.12 -e --source winget --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo.
        echo [ERRO] Nao foi possivel instalar o Python automaticamente.
        echo Instale manualmente em: https://www.python.org/downloads/
        echo Marque a opcao: Add python.exe to PATH
        pause
        exit /b 1
    )

    echo.
    echo [INFO] Python instalado. Atualizando ambiente da sessao...
    set "PATH=%LocalAppData%\Programs\Python\Python312;%LocalAppData%\Programs\Python\Python312\Scripts;%PATH%"

    where py >nul 2>nul
    if !errorlevel!==0 (
        py -3 --version >nul 2>nul
        if !errorlevel!==0 set "PYTHON_CMD=py -3"
    )

    if not defined PYTHON_CMD (
        where python >nul 2>nul
        if !errorlevel!==0 (
            python --version >nul 2>nul
            if !errorlevel!==0 set "PYTHON_CMD=python"
        )
    )
)

if not defined PYTHON_CMD (
    echo [ERRO] Python ainda nao foi localizado apos a instalacao.
    echo Feche esta janela, abra novamente o run_app.bat e tente outra vez.
    pause
    exit /b 1
)

echo [OK] Python encontrado:
%PYTHON_CMD% --version

echo.
echo [INFO] Atualizando pip...
%PYTHON_CMD% -m pip install --upgrade pip
if errorlevel 1 (
    echo [AVISO] Nao foi possivel atualizar o pip. Continuando mesmo assim...
)

echo.
echo [INFO] Instalando dependencias do projeto...
%PYTHON_CMD% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao instalar dependencias do requirements.txt.
    echo Verifique sua conexao com a internet e tente novamente.
    pause
    exit /b 1
)

echo.
echo [INFO] Garantindo pacote requests...
%PYTHON_CMD% -m pip install requests
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao instalar o pacote requests.
    pause
    exit /b 1
)

echo.
echo [INFO] Iniciando o sistema...
echo.
%PYTHON_CMD% -m streamlit run app.py

pause
