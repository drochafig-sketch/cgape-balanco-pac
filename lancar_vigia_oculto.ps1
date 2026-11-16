# Dispara VIGIAR_GITHUB.ps1 escondido (sem janela), a cada login do
# Windows (chave Run criada por LIGAR_SINCRONIZACAO_AO_MUDAR.bat).
# GitHub como referencia (16/11/2026) -- ver o cabecalho de VIGIAR_GITHUB.ps1.
#
# Nao usa .vbs: nesta maquina o Windows Script Host esta bloqueado por
# politica; o objeto COM WScript.Shell criado pelo PowerShell funciona.
# $PSScriptRoot = pasta onde ESTE arquivo esta salvo.
$shell = New-Object -ComObject "WScript.Shell"
$null = $shell.Run("powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$PSScriptRoot\VIGIAR_GITHUB.ps1`"", 0, $false)
