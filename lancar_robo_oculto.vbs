Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' Pega o caminho da pasta onde este script (.vbs) está localizado
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Monta o caminho completo para o arquivo .bat que está na mesma pasta
' (Altere "run_robot.bat" se o seu arquivo .bat tiver outro nome)
strBatPath = objFSO.BuildPath(strScriptPath, "run_robot.bat")

' Executa o .bat de forma oculta, usando o caminho completo e garantido
WshShell.Run chr(34) & strBatPath & Chr(34), 0

Set WshShell = Nothing
Set objFSO = Nothing