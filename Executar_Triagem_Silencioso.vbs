' ==============================================================================
' Lançador 100% Silencioso da Triagem VRE (Sem Janela Preta de Terminal)
' Detecta automaticamente o Python portátil embutido ou o do sistema Windows
' ==============================================================================

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
currentDir = fso.GetParentFolderName(WScript.ScriptFullName)

' 1. Localiza o executável pythonw.exe (prioriza embutido)
pywBin = ""
If fso.FileExists(currentDir & "\Configurações\python\pythonw.exe") Then
    pywBin = """" & currentDir & "\Configurações\python\pythonw.exe"""
ElseIf fso.FileExists(currentDir & "\Configuracoes\python\pythonw.exe") Then
    pywBin = """" & currentDir & "\Configuracoes\python\pythonw.exe"""
ElseIf fso.FileExists(currentDir & "\python\pythonw.exe") Then
    pywBin = """" & currentDir & "\python\pythonw.exe"""
Else
    pywBin = "pythonw.exe"
End If

' 2. Localiza o script de processamento
scriptPath = ""
If fso.FileExists(currentDir & "\Configurações\src\runner\main_seco.py") Then
    scriptPath = currentDir & "\Configurações\src\runner\main_seco.py"
ElseIf fso.FileExists(currentDir & "\Configuracoes\src\runner\main_seco.py") Then
    scriptPath = currentDir & "\Configuracoes\src\runner\main_seco.py"
ElseIf fso.FileExists(currentDir & "\src\runner\main_seco.py") Then
    scriptPath = currentDir & "\src\runner\main_seco.py"
End If

If scriptPath <> "" Then
    args = ""
    For Each arg In WScript.Arguments
        args = args & " """ & arg & """"
    Next
    ' Executa com janela oculta (0) e de forma assíncrona/rápida (False)
    WshShell.Run pywBin & " """ & scriptPath & """" & args, 0, False
Else
    MsgBox "Arquivo de execução da triagem não encontrado na pasta Configurações\src.", 16, "Erro - SMAURB VRE"
End If
