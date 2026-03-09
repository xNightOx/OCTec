@echo off
setlocal
set "SVC=OCTecService"
set "BASE=%~dp0.."
set "NSSM=%BASE%\bin\nssm.exe"

if exist "%NSSM%" (
  "%NSSM%" stop "%SVC%" confirm >nul 2>&1
  "%NSSM%" remove "%SVC%" confirm >nul 2>&1
) else (
  rem tenta via sc se nssm nao estiver mais no disco
  sc stop "%SVC%" >nul 2>&1
  sc delete "%SVC%" >nul 2>&1
)

echo Servico "%SVC%" removido (se existia).
endlocal
exit /b 0
