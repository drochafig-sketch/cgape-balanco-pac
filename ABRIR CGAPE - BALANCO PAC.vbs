' Abre o CGAPE - BALANCO PAC a partir de uma copia local do .exe.
'
' Copia "CGAPE - BALANCO PAC.exe" desta pasta (rede) para
' %LOCALAPPDATA%\CGAPE\ quando ele for diferente da copia local, e roda a
' copia local apontando CGAPE_PASTA para esta pasta (onde ficam as
' planilhas). Assim ninguem prende o .exe da rede e ele pode ser atualizado
' a qualquer hora. Pedido do usuario em 17/11/2026.
Option Explicit
Dim fso, sh, pasta, origem, destDir, destino, copiar, fo, fd
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
pasta = fso.GetParentFolderName(WScript.ScriptFullName)
origem = pasta & "\CGAPE - BALANCO PAC.exe"
' Se o .exe de dist\ (build mais recente) for mais novo, usa ele: assim a
' versao nova vale mesmo com o .exe da pasta preso por alguem que o abriu
' direto (pedido do usuario em 17/11/2026).
Dim origemDist
origemDist = pasta & "\dist\CGAPE - BALANCO PAC.exe"
If fso.FileExists(origemDist) Then
  If Not fso.FileExists(origem) Then
    origem = origemDist
  ElseIf fso.GetFile(origemDist).DateLastModified > fso.GetFile(origem).DateLastModified Then
    origem = origemDist
  End If
End If
destDir = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\CGAPE"
destino = destDir & "\CGAPE - BALANCO PAC.exe"

If Not fso.FileExists(origem) Then
  MsgBox "Nao encontrei o programa:" & vbCrLf & origem, vbExclamation, "CGAPE - BALANCO PAC"
  WScript.Quit 1
End If
If Not fso.FolderExists(destDir) Then fso.CreateFolder destDir

copiar = True
If fso.FileExists(destino) Then
  Set fo = fso.GetFile(origem)
  Set fd = fso.GetFile(destino)
  copiar = (fo.Size <> fd.Size) Or (fo.DateLastModified <> fd.DateLastModified)
End If
If copiar Then
  On Error Resume Next
  fso.CopyFile origem, destino, True
  If Err.Number <> 0 Then
    ' copia local em uso (programa ja aberto neste computador): segue com ela
    Err.Clear
    If Not fso.FileExists(destino) Then
      On Error GoTo 0
      MsgBox "Nao consegui copiar o programa para este computador.", vbExclamation, "CGAPE - BALANCO PAC"
      WScript.Quit 1
    End If
  End If
  On Error GoTo 0
End If

sh.Environment("PROCESS")("CGAPE_PASTA") = pasta
sh.CurrentDirectory = pasta
sh.Run """" & destino & """", 1, False
