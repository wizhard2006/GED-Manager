@echo off
REM creer_raccourci.bat
REM Cree un raccourci "GED-Manager" sur le bureau pointant vers GED-Manager.vbs
REM A executer UNE SEULE FOIS apres avoir deplace le dossier GED-Manager.

setlocal

SET SCRIPT_DIR=%~dp0
SET VBS_PATH=%SCRIPT_DIR%GED-Manager.vbs

echo Creation du raccourci sur le bureau...

REM Recuperer le chemin reel du bureau via PowerShell (fonctionne en francais et en anglais)
FOR /F "usebackq delims=" %%D IN (`powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"`) DO SET DESKTOP=%%D

SET SHORTCUT_PATH=%DESKTOP%\GED-Manager.lnk

powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; " ^
  "$s = $ws.CreateShortcut('%SHORTCUT_PATH%'); " ^
  "$s.TargetPath = 'wscript.exe'; " ^
  "$s.Arguments = '\"%VBS_PATH%\"'; " ^
  "$s.WorkingDirectory = '%SCRIPT_DIR%'; " ^
  "$s.Description = 'GED-Manager - Gestion Electronique de Documents'; " ^
  "$s.Save()"

IF EXIST "%SHORTCUT_PATH%" (
    echo Raccourci cree avec succes : %SHORTCUT_PATH%
) ELSE (
    echo ERREUR : le raccourci n'a pas pu etre cree.
)

pause
