@echo off
setlocal
chcp 65001 >nul

set "SVC=OCTecService"
set "BASE=%~dp0.."
set "EXE=%BASE%\OCTec.exe"
set "NSSM=%BASE%\bin\nssm.exe"

if not exist "%NSSM%" (
  echo ERRO: nssm.exe nao encontrado em "%NSSM%".
  exit /b 2
)

rem Garante pasta de logs
if not exist "%ProgramData%\OCTec" mkdir "%ProgramData%\OCTec" >nul 2>&1

rem Para/remove instalacao anterior (idempotente)
"%NSSM%" stop "%SVC%" confirm >nul 2>&1
"%NSSM%" remove "%SVC%" confirm >nul 2>&1

rem Instala o servico apontando para o MESMO EXE com --service
"%NSSM%" install "%SVC%" "%EXE%" --service

rem Configuracoes uteis
"%NSSM%" set "%SVC%" AppDirectory "%BASE%"
"%NSSM%" set "%SVC%" ObjectName LocalSystem

rem Logs com rotacao
"%NSSM%" set "%SVC%" AppStdout "%ProgramData%\OCTec\service.log"
"%NSSM%" set "%SVC%" AppStderr "%ProgramData%\OCTec\service.err.log"
"%NSSM%" set "%SVC%" AppRotateFiles 1
"%NSSM%" set "%SVC%" AppRotateOnline 1
"%NSSM%" set "%SVC%" AppRotateBytes 10485760

rem Tentativas de parada suave; se nao der, NSSM finaliza
"%NSSM%" set "%SVC%" AppStopMethodConsole 2000
"%NSSM%" set "%SVC%" AppStopMethodWindow 2000
"%NSSM%" set "%SVC%" AppStopMethodThreads 0

rem (Opcional mas recomendado) restart policy
"%NSSM%" set "%SVC%" AppRestartDelay 5000
sc failure "%SVC%" actions= restart/5000/restart/5000/restart/5000 reset= 86400
sc failureflag "%SVC%" 1

rem Descricao visivel no services.msc
sc description "%SVC%" "OCTec - motor de OCR/monitoramento (sem UI)"

rem Tipo de inicializacao (escolha um): delayed-auto ou auto
sc config "%SVC%" start= delayed-auto

rem Sobe o servico
"%NSSM%" start "%SVC%"

echo Servico "%SVC%" instalado e iniciado.
endlocal
exit /b 0
