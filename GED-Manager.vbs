' GED-Manager.vbs
' Lance GED-Manager sans fenetre noire (utilise pythonw.exe)
' Double-cliquer sur ce fichier pour demarrer l'application.

Dim oShell, oFSO, sDir, sPythonw, sScript

Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")

' Dossier contenant ce fichier VBS = dossier GED-Manager
sDir    = oFSO.GetParentFolderName(WScript.ScriptFullName)
sScript = sDir & "\main.py"

' Chercher pythonw.exe dans le PATH systeme
sPythonw = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python313\pythonw.exe"
If Not oFSO.FileExists(sPythonw) Then
    ' Fallback : chercher via commande where
    sPythonw = "pythonw"
End If

' Lancer sans fenetre (0 = hidden)
oShell.Run """" & sPythonw & """ """ & sScript & """", 0, False

Set oShell = Nothing
Set oFSO   = Nothing
