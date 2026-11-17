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
#   3) ONEDRIVE -> GITHUB (pedido do usuario em 16/11/2026): antes de
#      espelhar no OneDrive (e a cada 5 min), confere se algum arquivo foi
#      editado LA depois do ultimo espelho. Se foi, e o mesmo arquivo nao
#      mudou no GitHub nem esta sendo editado nesta pasta, traz para ca, faz
#      commit e envia ao GitHub (dai segue o caminho normal para esta pasta,
#      o COAHS e o .exe). Se houver CONFLITO (o arquivo tambem mudou no
#      GitHub ou aqui), NAO sobrescreve nenhum dos dois: para de espelhar
#      so no OneDrive, guarda copia em _conflitos_onedrive\ e deixa o aviso
#      _CONFLITO_SINCRONIZACAO.txt na pasta do OneDrive. Arquivos APAGADOS no
#      OneDrive nao sao apagados no GitHub (voltam no proximo espelho). A
#      planilha PAC e o .exe nao sao recolhidos (a planilha tem fonte propria,
#      a pasta "2026"; o .exe e gerado).
#   1b) ACOES DO CRONOGRAMA (pedido do usuario em 17/11/2026): a cada hora
#      baixa a planilha do servico publico da CONDER (importar_cronograma.py,
#      mesmo conteudo do botao "Exportar para planilha" do site
#      app.conder.ba.gov.br/cronogramaacoes) e, se o conteudo mudou, faz
#      commit SO dela e envia ao GitHub. Nao e recolhida do OneDrive.
#   4) GITHUB -> ONEDRIVE e COAHS: os destinos recebem EXATAMENTE o que esta
#      no GitHub (origin/master) + o .exe. Alteracoes locais ainda nao
#      enviadas ao GitHub NAO vao para la. Espelho (/MIR): o que nao estiver
#      no GitHub e apagado nos destinos (fora as pastas excluidas abaixo).
#
# Uso manual:
#   powershell -ExecutionPolicy Bypass -File VIGIAR_GITHUB.ps1 -UmaVez     (uma passada e sai)
#   powershell -ExecutionPolicy Bypass -File VIGIAR_GITHUB.ps1 -Simular    (mostra o que faria, sem mudar nada)
#   ... -UmaVez -AceitarOneDrive   (conflito: vale a versao do OneDrive, que vai para o GitHub)
#   ... -UmaVez -AceitarGithub     (conflito: vale a do GitHub; a do OneDrive fica so na copia em _conflitos_onedrive)
param([switch]$UmaVez, [switch]$Simular, [switch]$AceitarOneDrive, [switch]$AceitarGithub)

$ErrorActionPreference = "Continue"
$Repo = $PSScriptRoot
$Pasta2026 = (Resolve-Path (Join-Path $Repo "..\..")).Path
$PadraoPlanilha = "PANORAMA - PAC ORIGINAL*.xlsx"
$Log = Join-Path $Repo "_vigia_sincronizacao.log"
$Python = Join-Path $Repo ".venv\Scripts\python.exe"
$Spec = Join-Path $Repo "CGAPE - BALANCO PAC.spec"
$NomeExe = "CGAPE - BALANCO PAC.exe"
$Espelho = Join-Path $env:LOCALAPPDATA "CGAPE_espelho_github"
$EstadoAntigo = Join-Path $env:LOCALAPPDATA "CGAPE_espelho_github.estado"
$PastaConflitos = Join-Path $Repo "_conflitos_onedrive"
$NomeAviso = "_CONFLITO_SINCRONIZACAO.txt"
$Destinos = @(
    @{ Nome = "OneDrive"; Pasta = "C:\Users\diego.figueiredo\OneDrive\2025 2026 Onedrive\MONITORAMENTO\PANORAMA\Ambiente Virtual"; Log = "_sync_onedrive.log"; Recolher = $true },
    @{ Nome = "COAHS"; Pasta = "S:\00 NOVA REDE - COAM\venv"; Log = "_sync_coam.log"; Recolher = $false }
)
$ExcluirPastas = @("Lib", "Scripts", "Include", "share", "build", "dist", "__pycache__", ".git", ".claude", ".venv", "_conflitos_onedrive")
$ExcluirArquivos = @(".git", "Thumbs.db", "pyvenv.cfg", "desktop.ini", "_sync_onedrive.log", "_sync_coam.log", "_vigia_sincronizacao.log", $NomeAviso)
$IntervaloGithub = 300   # segundos entre consultas ao GitHub (e conferencias do OneDrive)
$IntervaloLaco = 5
# "Acoes do Cronograma.xlsx" baixada do servico da CONDER (importar_cronograma.py)
# a cada hora -- pedido do usuario em 17/11/2026. Nome montado com [char]
# porque o PowerShell 5.1 le este arquivo sem BOM como ANSI.
$NomeCronograma = "A" + [char]0x00E7 + [char]0x00F5 + "es do Cronograma.xlsx"
$IntervaloCronograma = 3600
$EsperaFalhaCronograma = 600

$Git = Join-Path $env:LOCALAPPDATA "Programs\PortableGit\cmd\git.exe"
if (-not (Test-Path $Git)) { $Git = "git" }

function Registrar([string]$texto) {
    $linha = "[{0:dd/MM/yyyy HH:mm:ss}] {1}" -f (Get-Date), $texto
    Add-Content -Path $Log -Value $linha -Encoding UTF8
    # Write-Host (e nao Write-Output): a mensagem nao pode se misturar ao
    # valor devolvido pelas funcoes (ex.: o $true/$false de Recolher-Destino)
    if ($UmaVez -or $Simular) { Write-Host $linha }
}

function G {
    # git dentro do repositorio; devolve a saida (texto) e deixa $LASTEXITCODE
    $saida = & $Git -c safe.directory=* -c core.quotepath=false -C $Repo @args 2>&1
    return ($saida | ForEach-Object { "$_" }) -join "`n"
}

function Arquivo-Livre([string]$caminho) {
    try { $f = [IO.File]::Open($caminho, 'Open', 'Read', 'Read'); $f.Close(); return $true } catch { return $false }
}

# ---- estado de cada destino: "commit|marca do exe|hora do ultimo espelho" ----
function Arquivo-Estado($d) { Join-Path $env:LOCALAPPDATA ("CGAPE_espelho_github_" + $d.Nome + ".estado") }
function Ler-Estado($d) {
    $arq = Arquivo-Estado $d
    if (-not (Test-Path $arq)) { $arq = $EstadoAntigo }   # formato antigo (um so para todos)
    $r = @{ Commit = ""; Exe = ""; Hora = $null }
    if (Test-Path $arq) {
        $p = (Get-Content $arq -Raw).Trim().Split("|")
        $r.Commit = $p[0]
        if ($p.Count -gt 1) { $r.Exe = $p[1] }
        if ($p.Count -gt 2 -and $p[2]) { $r.Hora = [datetime]::new([long]$p[2], [DateTimeKind]::Utc) }
    }
    return $r
}
function Gravar-Estado($d, $commit, $exe) {
    Set-Content -Path (Arquivo-Estado $d) -Value ("{0}|{1}|{2}" -f $commit, $exe, (Get-Date).ToUniversalTime().Ticks) -Encoding ASCII
}
function Marca-Exe {
    $exe = Join-Path $Repo $NomeExe
    if (Test-Path $exe) { return [string](Get-Item $exe).LastWriteTimeUtc.Ticks } else { return "sem-exe" }
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
    # commit do vigia (planilha ou OneDrive) que nao subiu no ciclo anterior (ex.: sem rede)
    $frente = [int](G rev-list --count "origin/master..HEAD")
    $atras = [int](G rev-list --count "HEAD..origin/master")
    if ($frente -gt 0 -and $atras -eq 0) {
        $ultimo = G log -1 --format=%s
        if ($ultimo -like "Planilha PAC atualizada (vigia*" -or $ultimo -like "Alteracoes feitas no OneDrive (vigia*" -or
            $ultimo -like "Acoes do Cronograma atualizada (vigia*") {
            G push -q origin HEAD:master | Out-Null
            if ($LASTEXITCODE -eq 0) { Registrar "Commit pendente do vigia enviado ao GitHub." }
        }
    }
}

# ---- 1b) Acoes do Cronograma: servico da CONDER -> GitHub (a cada hora) -----
$script:ProximoCronograma = [datetime]::MinValue
function Importar-Cronograma {
    if ((Get-Date) -lt $script:ProximoCronograma) { return }
    $script:ProximoCronograma = (Get-Date).AddSeconds($IntervaloCronograma)
    $script_ = Join-Path $Repo "importar_cronograma.py"
    if (-not (Test-Path $Python) -or -not (Test-Path $script_)) { return }
    Atualizar-DoGithub
    $arq = Join-Path $Repo $NomeCronograma
    $saida = (& $Python $script_ --se-mudou $arq 2>$null) -join " "
    $codigo = $LASTEXITCODE
    if ($codigo -ne 0 -and $codigo -ne 3) {
        # sem rede, servico fora do ar ou planilha aberta no Excel: tenta de novo em 10 min
        $script:ProximoCronograma = (Get-Date).AddSeconds($EsperaFalhaCronograma)
        Registrar "Acoes do Cronograma: nao consegui atualizar pelo servico da CONDER (tenta de novo em 10 min): $saida"
        return
    }
    # sem mudanca no servico, mas a copia daqui pode estar diferente do GitHub
    # (ex.: baixada a mao pelo site) -- nos dois casos o GitHub fica com a daqui
    $sujo = G status --porcelain -- $NomeCronograma
    if (-not $sujo) { return }
    G add -- $NomeCronograma | Out-Null
    G commit -q -m "Acoes do Cronograma atualizada (vigia, servico CONDER)" -- $NomeCronograma | Out-Null
    if ($LASTEXITCODE -ne 0) { return }
    $msg = G push -q origin HEAD:master
    if ($LASTEXITCODE -eq 0) { Registrar "Acoes do Cronograma atualizada pelo servico da CONDER e enviada ao GitHub ($saida)." }
    else { Registrar "Acoes do Cronograma commitada, mas o envio ao GitHub falhou (tenta de novo no proximo ciclo): $msg" }
}

# ---- 3) OneDrive -> GitHub ------------------------------------------------------
function Listar-Arquivos([string]$pasta, [string]$raiz, $desde) {
    # arquivos do destino que entram no espelho (mesmas exclusoes do robocopy)
    $saida = @()
    $itens = Get-ChildItem -LiteralPath $pasta -Force -ErrorAction SilentlyContinue
    foreach ($i in $itens) {
        if ($i.Attributes -band ([IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System)) { continue }
        if ($i.PSIsContainer) {
            if ($ExcluirPastas -contains $i.Name) { continue }
            $saida += Listar-Arquivos $i.FullName $raiz $desde
            continue
        }
        if ($ExcluirArquivos -contains $i.Name) { continue }
        if ($i.Name -like "~$*" -or $i.Name -like "*.tmp" -or $i.Name -eq $NomeExe -or $i.Name -like $PadraoPlanilha -or
            $i.Name -eq $NomeCronograma) { continue }   # o cronograma tem fonte propria (servico da CONDER)
        if ($desde -and $i.LastWriteTimeUtc -le $desde) { continue }
        $saida += [pscustomobject]@{ Arquivo = $i.FullName; Rel = $i.FullName.Substring($raiz.Length).TrimStart("\").Replace("\", "/") }
    }
    return $saida
}

$script:UltimoAvisoConflito = ""
function Recolher-Destino($d) {
    # devolve $true se pode espelhar neste destino (nada pendente vindo dele)
    if (-not (Test-Path -LiteralPath $d.Pasta)) { return $true }
    $est = Ler-Estado $d
    G fetch -q origin | Out-Null
    $base = $est.Commit
    if ($base) { G cat-file -e "$base^{commit}" | Out-Null; if ($LASTEXITCODE -ne 0) { $base = "" } }
    if (-not $base) { $base = G rev-parse origin/master }
    $candidatos = Listar-Arquivos $d.Pasta $d.Pasta $est.Hora
    $mudadosGithub = @((G diff --name-only $base origin/master) -split "`n" | Where-Object { $_ })
    # o Windows nao diferencia maiusculas ("GOVERNO.png" no OneDrive e o
    # mesmo arquivo que "GOVERNO.PNG" no GitHub), o Git sim: usa o nome como
    # esta no GitHub/ultimo espelho, senao criaria um arquivo duplicado
    $nomesGit = @{}
    foreach ($ref in @($base, "origin/master")) {
        foreach ($n in ((G ls-tree -r --name-only $ref) -split "`n")) { if ($n -and -not $nomesGit.ContainsKey($n.ToLower())) { $nomesGit[$n.ToLower()] = $n } }
    }
    foreach ($c in $candidatos) { if ($nomesGit.ContainsKey($c.Rel.ToLower())) { $c.Rel = $nomesGit[$c.Rel.ToLower()] } }
    $novos = @(); $conflitos = @()
    foreach ($c in $candidatos) {
        $hDestino = G hash-object "--path=$($c.Rel)" -- $c.Arquivo
        if ($LASTEXITCODE -ne 0) { continue }
        $hBase = G rev-parse -q --verify "$($base):$($c.Rel)"
        if ($hDestino -eq $hBase) { continue }                       # so "tocado", sem mudar conteudo
        $hGithub = G rev-parse -q --verify "origin/master:$($c.Rel)"
        if ($hDestino -eq $hGithub) { continue }                     # ja igual ao GitHub
        $sujo = G status --porcelain -- $c.Rel
        if ($sujo -or ($mudadosGithub -contains $c.Rel)) { $conflitos += $c } else { $novos += $c }
    }
    if ($AceitarOneDrive) { $novos += $conflitos; $conflitos = @() }

    if ($Simular) {
        Write-Host "===== SIMULACAO: recolher do $($d.Nome) ====="
        foreach ($c in $novos) { Write-Host "  enviaria ao GitHub: $($c.Rel)" }
        foreach ($c in $conflitos) { Write-Host "  CONFLITO (mudou tambem no GitHub ou nesta pasta): $($c.Rel)" }
        if (-not $novos -and -not $conflitos) { Write-Host "  nada editado no $($d.Nome) desde o ultimo espelho" }
        return $true
    }

    $aviso = Join-Path $d.Pasta $NomeAviso
    if ($conflitos) {
        $carimbo = "{0:yyyy-MM-dd_HHmmss}" -f (Get-Date)
        $lista = ($conflitos | ForEach-Object { $_.Rel }) -join ", "
        foreach ($c in $conflitos) {
            $destinoCopia = Join-Path (Join-Path $PastaConflitos $carimbo) ($c.Rel.Replace("/", "\"))
            New-Item -ItemType Directory -Force -Path (Split-Path $destinoCopia) | Out-Null
            Copy-Item -LiteralPath $c.Arquivo -Destination $destinoCopia -Force
        }
        if ($AceitarGithub) {
            Registrar "Conflito no $($d.Nome) resolvido pela versao do GitHub ($lista); a versao do $($d.Nome) ficou guardada em _conflitos_onedrive\$carimbo."
            $conflitos = @()
        } else {
            if ($script:UltimoAvisoConflito -ne $lista) {
                Registrar "CONFLITO no $($d.Nome): $lista mudou la e tambem no GitHub (ou nesta pasta). Nada foi sobrescrito; espelho do $($d.Nome) parado. Copia guardada em _conflitos_onedrive\$carimbo. Resolva e rode SINCRONIZAR_BACKUPS.bat, ou o vigia com -AceitarOneDrive / -AceitarGithub."
                $script:UltimoAvisoConflito = $lista
            }
            $texto = @(
                "CONFLITO DE SINCRONIZACAO ($('{0:dd/MM/yyyy HH:mm}' -f (Get-Date)))",
                "",
                "Estes arquivos foram editados aqui no OneDrive e TAMBEM mudaram no GitHub",
                "(ou estao sendo editados na pasta X:). Nada foi sobrescrito, e o OneDrive",
                "nao recebe atualizacoes ate o conflito ser resolvido:",
                "")
            $texto += @($conflitos | ForEach-Object { "  - " + $_.Rel })
            $texto += @("", "Copia destas versoes do OneDrive: _conflitos_onedrive\$carimbo (na pasta X:).")
            Set-Content -LiteralPath $aviso -Encoding UTF8 -Value $texto
            return $false
        }
    }
    $script:UltimoAvisoConflito = ""
    if (Test-Path -LiteralPath $aviso) { Remove-Item -LiteralPath $aviso -Force }

    if (-not $novos) {
        if (-not $est.Hora -and $est.Commit) { Gravar-Estado $d $est.Commit $est.Exe }   # marca a hora p/ nao conferir tudo de novo
        return $true
    }

    # envia as edicoes do OneDrive ao GitHub
    Atualizar-DoGithub
    if ([int](G rev-list --count "HEAD..origin/master") -gt 0) {
        Registrar "Edicoes no $($d.Nome) aguardando: esta pasta nao conseguiu se atualizar com o GitHub (ver acima). Espelho do $($d.Nome) parado ate resolver."
        return $false
    }
    $rels = @()
    foreach ($c in $novos) {
        $destinoLocal = Join-Path $Repo ($c.Rel.Replace("/", "\"))
        New-Item -ItemType Directory -Force -Path (Split-Path $destinoLocal) | Out-Null
        Copy-Item -LiteralPath $c.Arquivo -Destination $destinoLocal -Force
        $rels += $c.Rel
    }
    G add -- @rels | Out-Null
    G commit -q -m "Alteracoes feitas no OneDrive (vigia): $($rels.Count) arquivo(s)" -- @rels | Out-Null
    if ($LASTEXITCODE -ne 0) { Registrar "Nao consegui fazer o commit das edicoes do $($d.Nome): $($rels -join ', ')"; return $false }
    $msg = G push -q origin HEAD:master
    if ($LASTEXITCODE -ne 0) {
        Registrar "Edicoes do $($d.Nome) commitadas, mas o envio ao GitHub falhou (tenta de novo): $msg"
        return $false
    }
    Registrar "Edicoes feitas no $($d.Nome) enviadas ao GitHub: $($rels -join ', ')."
    if (($rels -join "`n") -match '\.(py|spec)(\n|$)') { $script:PrecisaExe = $true }
    return $true
}

# ---- .exe ---------------------------------------------------------------------
$script:ExePendente = $false
function Gerar-Exe {
    if (-not (Test-Path $Python)) { Registrar "Nao gerei o .exe: .venv nao encontrado."; $script:PrecisaExe = $false; return }
    Registrar "Codigo novo: gerando o .exe..."
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
    try { Copy-Item $novo (Join-Path $Repo $NomeExe) -Force -ErrorAction Stop } catch { return }   # em uso em outra maquina
    $script:ExePendente = $false
    Registrar ".exe novo copiado para a pasta."
}

# ---- 4) GitHub -> OneDrive / COAHS ---------------------------------------------
$script:CopiaCommit = ""
function Preparar-Copia([string]$commit) {
    # copia limpa do que esta no GitHub: uma "worktree" do Git (preserva os
    # nomes com acento -- o tar do Windows os corrompia), sem nada local
    if ($script:CopiaCommit -eq $commit -and (Test-Path (Join-Path $Espelho ".git"))) { return $true }
    if (-not (Test-Path (Join-Path $Espelho ".git"))) {
        if (Test-Path $Espelho) { Remove-Item $Espelho -Recurse -Force }
        G worktree prune | Out-Null
        G worktree add -q --detach $Espelho $commit | Out-Null
    } else {
        # safe.directory: a copia fica em C:, o repositorio em \\tuba -- sem
        # isso o Git recusa ("dubious ownership")
        & $Git -c safe.directory=* -C $Espelho checkout -q -f --detach $commit 2>&1 | Out-Null
        & $Git -c safe.directory=* -C $Espelho clean -fdxq 2>&1 | Out-Null
    }
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path (Join-Path $Espelho ".git"))) { Registrar "ERRO ao preparar a copia do GitHub para o espelho."; return $false }
    $exe = Join-Path $Repo $NomeExe
    if (Test-Path $exe) { Copy-Item $exe (Join-Path $Espelho $NomeExe) -Force }
    $script:CopiaCommit = $commit
    return $true
}

function Espelhar {
    foreach ($d in $Destinos) {
        $commit = G rev-parse origin/master
        if ($LASTEXITCODE -ne 0) { return }
        $exeMarca = Marca-Exe
        $est = Ler-Estado $d
        if (-not $Simular -and $est.Commit -eq $commit -and $est.Exe -eq $exeMarca) { continue }
        if ($d.Recolher) {
            if (-not (Recolher-Destino $d)) { continue }   # edicao pendente/conflito: nao sobrescreve
            $commit = G rev-parse origin/master            # pode ter mudado (edicoes enviadas)
        }
        if (-not (Preparar-Copia $commit)) { return }
        $rc = @($Espelho, $d.Pasta, "/MIR",
                "/XD") + $ExcluirPastas + @("/XF") + $ExcluirArquivos + @("/XA:SH", "/FFT", "/R:2", "/W:5", "/NP", "/NJH", "/NJS")
        if ($Simular) {
            Write-Output "===== SIMULACAO: espelho $($d.Pasta) ====="
            & robocopy.exe @rc /L /NDL
            continue
        }
        & robocopy.exe @rc /NFL /NDL "/LOG:$(Join-Path $Repo $d.Log)" | Out-Null
        if ($LASTEXITCODE -ge 8) { Registrar "ERRO no espelho para $($d.Pasta) (robocopy $LASTEXITCODE)."; continue }
        Gravar-Estado $d $commit $exeMarca
        Registrar "$($d.Nome) espelhado com o GitHub ($($commit.Substring(0,7)))."
    }
}

# ---- laco ---------------------------------------------------------------------
if ($Simular) { G fetch -q origin | Out-Null; Espelhar; return }

Registrar "Vigia GitHub iniciado (GitHub como referencia: planilha 2026 e edicoes do OneDrive -> GitHub -> pasta/OneDrive/COAHS)."
$ultimaConsulta = [datetime]::MinValue
while ($true) {
    try {
        if (((Get-Date) - $ultimaConsulta).TotalSeconds -ge $IntervaloGithub -or $AceitarOneDrive -or $AceitarGithub) {
            Atualizar-DoGithub
            Enviar-Pendentes
            foreach ($d in $Destinos) { if ($d.Recolher) { Recolher-Destino $d | Out-Null } }
            $ultimaConsulta = Get-Date
        }
        Enviar-Planilha
        Importar-Cronograma
        if ($script:PrecisaExe) { Gerar-Exe }
        if ($script:ExePendente) { Copiar-Exe }
        Espelhar
    } catch {
        Registrar "ERRO no vigia: $($_.Exception.Message)"
    }
    if ($UmaVez) { break }
    Start-Sleep -Seconds $IntervaloLaco
}
