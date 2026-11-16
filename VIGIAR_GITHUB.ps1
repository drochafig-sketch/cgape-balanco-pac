# Vigia de sincronizacao com o GITHUB como referencia (pedido do usuario em
# 16/11/2026). Substitui VIGIAR_SINCRONIZACAO.bat. Roda escondido a cada
# login (chave Run, via lancar_vigia_oculto.ps1) e repete para sempre:
#
#   1) PLANILHA PAC: se a planilha da pasta "PANORAMA\2026" (editada a mao)
#      estiver diferente da que esta no GitHub, copia para esta pasta, faz
#      um commit SO dela e envia ao GitHub.
#   2) GITHUB -> ESTA PASTA (a cada 5 min): se o GitHub tiver commits novos
#      (ex.: da outra maquina), atualiza esta pasta (so avanco simples,
#      "fast-forward" -- nunca apaga alteracao local em andamento). Se veio
#      codigo novo, gera o .exe de novo e o copia para ca quando o programa
#      estiver fechado.
#   3) GITHUB -> ONEDRIVE e COAHS: os destinos recebem EXATAMENTE o que esta
#      no GitHub (origin/master) + o .exe. Alteracoes locais ainda nao
#      enviadas ao GitHub NAO vao para la. Espelho (/MIR): o que nao estiver
#      no GitHub e apagado nos destinos (fora as pastas excluidas abaixo).
#
# Uso manual:
#   powershell -ExecutionPolicy Bypass -File VIGIAR_GITHUB.ps1 -UmaVez     (uma passada e sai)
#   powershell -ExecutionPolicy Bypass -File VIGIAR_GITHUB.ps1 -Simular    (mostra o que o espelho faria, sem mudar nada)
param([switch]$UmaVez, [switch]$Simular)

$ErrorActionPreference = "Continue"
$Repo = $PSScriptRoot
$Pasta2026 = (Resolve-Path (Join-Path $Repo "..\..")).Path
$PadraoPlanilha = "PANORAMA - PAC ORIGINAL*.xlsx"
$Log = Join-Path $Repo "_vigia_sincronizacao.log"
$Python = Join-Path $Repo ".venv\Scripts\python.exe"
$Spec = Join-Path $Repo "CGAPE - BALANCO PAC.spec"
$NomeExe = "CGAPE - BALANCO PAC.exe"
$Espelho = Join-Path $env:LOCALAPPDATA "CGAPE_espelho_github"
$Estado = Join-Path $env:LOCALAPPDATA "CGAPE_espelho_github.estado"
$Destinos = @(
    @{ Pasta = "C:\Users\diego.figueiredo\OneDrive\2025 2026 Onedrive\MONITORAMENTO\PANORAMA\Ambiente Virtual"; Log = "_sync_onedrive.log" },
    @{ Pasta = "S:\00 NOVA REDE - COAM\venv"; Log = "_sync_coam.log" }
)
$IntervaloGithub = 300   # segundos entre consultas ao GitHub
$IntervaloLaco = 5

$Git = Join-Path $env:LOCALAPPDATA "Programs\PortableGit\cmd\git.exe"
if (-not (Test-Path $Git)) { $Git = "git" }

function Registrar([string]$texto) {
    $linha = "[{0:dd/MM/yyyy HH:mm:ss}] {1}" -f (Get-Date), $texto
    Add-Content -Path $Log -Value $linha -Encoding UTF8
    if ($UmaVez -or $Simular) { Write-Output $linha }
}

function G {
    # git dentro do repositorio; devolve a saida (texto) e deixa $LASTEXITCODE
    $saida = & $Git -c safe.directory=* -C $Repo @args 2>&1
    return ($saida | ForEach-Object { "$_" }) -join "`n"
}

function Arquivo-Livre([string]$caminho) {
    try { $f = [IO.File]::Open($caminho, 'Open', 'Read', 'Read'); $f.Close(); return $true } catch { return $false }
}

# ---- 2) GitHub -> esta pasta ------------------------------------------------
$script:PrecisaExe = $false
function Atualizar-DoGithub {
    G fetch -q origin | Out-Null
    if ($LASTEXITCODE -ne 0) { return }
    $atras = [int](G rev-list --count "HEAD..origin/master")
    if ($atras -le 0) { return }
    $antes = G rev-parse HEAD
    $msg = G merge --ff-only -q origin/master
    if ($LASTEXITCODE -ne 0) {
        Registrar "GitHub tem $atras commit(s) novo(s), mas a atualizacao automatica nao foi possivel (alteracao local em conflito ou historico divergente): $msg"
        return
    }
    $mudados = G diff --name-only $antes HEAD
    Registrar "Pasta atualizada com $atras commit(s) novo(s) do GitHub."
    if ($mudados -match '\.(py|spec)(\n|$)') { $script:PrecisaExe = $true }
}

# ---- 1) planilha da pasta 2026 -> GitHub ------------------------------------
function Enviar-Planilha {
    $origem = Get-ChildItem -Path $Pasta2026 -Filter $PadraoPlanilha -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike "~$*" } | Select-Object -First 1
    if (-not $origem) { return }
    $nome = $origem.Name
    $hashOrigem = G hash-object -- $origem.FullName
    $hashGithub = G rev-parse "origin/master:$nome"
    if ($hashOrigem -eq $hashGithub) {
        # GitHub ja tem esta versao; garante que a copia daqui tambem tem
        $local = Join-Path $Repo $nome
        if ((Test-Path $local) -and ((G hash-object -- $local) -ne $hashOrigem) -and (Arquivo-Livre $origem.FullName)) {
            Copy-Item $origem.FullName $local -Force
        }
        return
    }
    if (-not (Arquivo-Livre $origem.FullName)) { return }   # ainda sendo salva/aberta em modo exclusivo
    Atualizar-DoGithub
    Copy-Item $origem.FullName (Join-Path $Repo $nome) -Force
    G add -- $nome | Out-Null
    G commit -q -m "Planilha PAC atualizada (vigia, copia da pasta 2026)" -- $nome | Out-Null
    if ($LASTEXITCODE -ne 0) { return }
    $msg = G push -q origin HEAD:master
    if ($LASTEXITCODE -eq 0) { Registrar "Planilha PAC da pasta 2026 enviada ao GitHub." }
    else { Registrar "Planilha PAC commitada, mas o envio ao GitHub falhou (tenta de novo no proximo ciclo): $msg" }
}

function Enviar-Pendentes {
    # commit da planilha que nao subiu no ciclo anterior (ex.: sem rede)
    $frente = [int](G rev-list --count "origin/master..HEAD")
    $atras = [int](G rev-list --count "HEAD..origin/master")
    if ($frente -gt 0 -and $atras -eq 0) {
        $ultimo = G log -1 --format=%s
        if ($ultimo -like "Planilha PAC atualizada (vigia*") {
            G push -q origin HEAD:master | Out-Null
            if ($LASTEXITCODE -eq 0) { Registrar "Commit pendente da planilha enviado ao GitHub." }
        }
    }
}

# ---- .exe ---------------------------------------------------------------------
$script:ExePendente = $false
function Gerar-Exe {
    if (-not (Test-Path $Python)) { Registrar "Nao gerei o .exe: .venv nao encontrado."; $script:PrecisaExe = $false; return }
    Registrar "Codigo novo do GitHub: gerando o .exe..."
    $p = Start-Process -FilePath $Python -ArgumentList @("-m", "PyInstaller", "--noconfirm", "--clean", "`"$Spec`"") `
        -WorkingDirectory $Repo -WindowStyle Hidden -Wait -PassThru `
        -RedirectStandardOutput (Join-Path $env:TEMP "cgape_vigia_pyi.out") -RedirectStandardError (Join-Path $env:TEMP "cgape_vigia_pyi.err")
    $script:PrecisaExe = $false
    if ($p.ExitCode -ne 0) { Registrar "ERRO ao gerar o .exe (codigo $($p.ExitCode)); ver %TEMP%\cgape_vigia_pyi.err"; return }
    $script:ExePendente = $true
}
function Copiar-Exe {
    $novo = Join-Path $Repo "dist\$NomeExe"
    if (-not (Test-Path $novo)) { $script:ExePendente = $false; return }
    if (Get-Process | Where-Object { $_.ProcessName -like "CGAPE*" }) { return }   # programa aberto: tenta depois
    Copy-Item $novo (Join-Path $Repo $NomeExe) -Force
    $script:ExePendente = $false
    Registrar ".exe novo copiado para a pasta."
}

# ---- 3) GitHub -> OneDrive / COAHS ---------------------------------------------
function Espelhar {
    $commit = G rev-parse origin/master
    if ($LASTEXITCODE -ne 0) { return }
    $exe = Join-Path $Repo $NomeExe
    $marca = $commit + "|" + $(if (Test-Path $exe) { (Get-Item $exe).LastWriteTimeUtc.Ticks } else { "sem-exe" })
    if (-not $Simular -and (Test-Path $Estado) -and ((Get-Content $Estado -Raw).Trim() -eq $marca)) { return }

    # copia limpa do que esta no GitHub: uma "worktree" do Git (preserva os
    # nomes com acento -- o tar do Windows os corrompia) sempre no
    # origin/master, sem nada local
    if (-not (Test-Path (Join-Path $Espelho ".git"))) {
        if (Test-Path $Espelho) { Remove-Item $Espelho -Recurse -Force }
        G worktree prune | Out-Null
        G worktree add -q --detach $Espelho origin/master | Out-Null
    } else {
        # safe.directory: a copia fica em C:, o repositorio em \\tuba -- sem
        # isso o Git recusa ("dubious ownership")
        & $Git -c safe.directory=* -C $Espelho checkout -q -f --detach origin/master 2>&1 | Out-Null
        & $Git -c safe.directory=* -C $Espelho clean -fdxq 2>&1 | Out-Null
    }
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path (Join-Path $Espelho ".git"))) { Registrar "ERRO ao preparar a copia do GitHub para o espelho."; return }
    if (Test-Path $exe) { Copy-Item $exe (Join-Path $Espelho $NomeExe) -Force }

    $ok = $true
    foreach ($d in $Destinos) {
        $rc = @($Espelho, $d.Pasta, "/MIR",
                "/XD", "Lib", "Scripts", "Include", "share", "build", "dist", "__pycache__", ".git", ".claude", ".venv",
                "/XF", ".git", "Thumbs.db", "pyvenv.cfg", "desktop.ini", "_sync_onedrive.log", "_sync_coam.log", "_vigia_sincronizacao.log",
                "/XA:SH", "/FFT", "/R:2", "/W:5", "/NP", "/NJH", "/NJS")
        if ($Simular) {
            Write-Output "===== SIMULACAO: $($d.Pasta) ====="
            & robocopy.exe @rc /L /NDL
        } else {
            & robocopy.exe @rc /NFL /NDL "/LOG:$(Join-Path $Repo $d.Log)" | Out-Null
            if ($LASTEXITCODE -ge 8) { $ok = $false; Registrar "ERRO no espelho para $($d.Pasta) (robocopy $LASTEXITCODE)." }
        }
    }
    if (-not $Simular -and $ok) {
        Set-Content -Path $Estado -Value $marca -Encoding ASCII
        Registrar "OneDrive e COAHS espelhados com o GitHub ($($commit.Substring(0,7)))."
    }
}

# ---- laco ---------------------------------------------------------------------
if ($Simular) { G fetch -q origin | Out-Null; Espelhar; return }

Registrar "Vigia GitHub iniciado (GitHub como referencia: planilha 2026 -> GitHub -> pasta/OneDrive/COAHS)."
$ultimaConsulta = [datetime]::MinValue
while ($true) {
    try {
        if (((Get-Date) - $ultimaConsulta).TotalSeconds -ge $IntervaloGithub) {
            Atualizar-DoGithub
            Enviar-Pendentes
            $ultimaConsulta = Get-Date
        }
        Enviar-Planilha
        if ($script:PrecisaExe) { Gerar-Exe }
        if ($script:ExePendente) { Copiar-Exe }
        Espelhar
    } catch {
        Registrar "ERRO no vigia: $($_.Exception.Message)"
    }
    if ($UmaVez) { break }
    Start-Sleep -Seconds $IntervaloLaco
}
