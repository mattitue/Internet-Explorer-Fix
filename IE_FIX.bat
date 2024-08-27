@echo off
setlocal

:: Execute the VBScript using mshta.exe with embedded script
mshta.exe "javascript:var ie = new ActiveXObject('InternetExplorer.Application'); ie.Visible = true; close();"

endlocal
