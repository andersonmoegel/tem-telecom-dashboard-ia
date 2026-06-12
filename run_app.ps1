$env:STREAMLIT_BROWSER_GATHER_USAGE_STATS = "false"
Set-Location -Path $PSScriptRoot
$ErrorActionPreference = "Stop"

Write-Host "============================================================"
Write-Host " Telecom Dashboard IA - Inicializador"
Write-Host "============================================================"
Write-Host ""

function Get-PythonCommand {
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        try { & py -3 --version | Out-Null; return "py -3" } catch {}
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        try { & python --version | Out-Null; return "python" } catch {}
    }

    return $null
}

$pythonCmd = Get-PythonCommand

if (-not $pythonCmd) {
    Write-Host "[INFO] Python nao encontrado. Tentando instalar via winget..."
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        Write-Host "[ERRO] winget nao encontrado. Instale o Python manualmente em https://www.python.org/downloads/"
        Read-Host "Pressione Enter para sair"
        exit 1
    }

    winget install --id Python.Python.3.12 -e --source winget --accept-package-agreements --accept-source-agreements
    $env:Path = "$env:LocalAppData\Programs\Python\Python312;$env:LocalAppData\Programs\Python\Python312\Scripts;$env:Path"
    $pythonCmd = Get-PythonCommand
}

if (-not $pythonCmd) {
    Write-Host "[ERRO] Python ainda nao foi localizado. Abra novamente o script apos a instalacao."
    Read-Host "Pressione Enter para sair"
    exit 1
}

Write-Host "[OK] Python encontrado: $pythonCmd"
Invoke-Expression "$pythonCmd --version"

Write-Host "[INFO] Atualizando pip..."
try { Invoke-Expression "$pythonCmd -m pip install --upgrade pip" } catch { Write-Host "[AVISO] Nao foi possivel atualizar pip. Continuando..." }

Write-Host "[INFO] Instalando dependencias..."
Invoke-Expression "$pythonCmd -m pip install -r requirements.txt"

Write-Host "[INFO] Garantindo pacote requests..."
Invoke-Expression "$pythonCmd -m pip install requests"

Write-Host "[INFO] Iniciando sistema..."
Invoke-Expression "$pythonCmd -m streamlit run app.py"
