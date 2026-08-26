# Dispara VIGIAR_SINCRONIZACAO.bat escondido (sem nenhuma janela de
# console aparecendo), via WScript.Shell.Run com janela oculta (estilo 0).
#
# Nao usa um .vbs (como o padrao antigo de outros projetos): nesta
# maquina, rodar arquivos .vbs pelo wscript.exe fica sem efeito (o
# Windows Script Host esta bloqueado por politica), mas criar o mesmo
# objeto COM (WScript.Shell) direto de dentro do PowerShell funciona
# normalmente -- e o caminho usado aqui.
#
# $PSScriptRoot resolve sozinho para a pasta onde ESTE arquivo esta
# salvo, nao precisa de caminho fixo.
$shell = New-Object -ComObject "WScript.Shell"
$null = $shell.Run("cmd.exe /c `"$PSScriptRoot\VIGIAR_SINCRONIZACAO.bat`"", 0, $false)
