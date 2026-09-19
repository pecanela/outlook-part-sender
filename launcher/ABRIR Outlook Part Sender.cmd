@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Outlook Part Sender PORTABLE

set "VERSION=3.1.0"
set "APPDIR=%LOCALAPPDATA%\OutlookPartSenderPortable"
set "SRC=%~dp0OutlookPartSenderPortable.cs"
if not exist "%SRC%" set "SRC=%~dp0..\src\OutlookPartSenderPortable.cs"
set "EXE=%APPDIR%\OutlookPartSenderPortable.exe"
set "LOG=%APPDIR%\compile.log"
set "VERFILE=%APPDIR%\version.txt"

if not exist "%SRC%" (
  echo.
  echo ERRO: arquivo OutlookPartSenderPortable.cs nao encontrado.
  echo Fonte nao encontrado no pacote nem no diretorio ..\src do repositorio.
  echo.
  pause
  exit /b 1
)

if not exist "%APPDIR%" mkdir "%APPDIR%" >nul 2>&1

set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"

if not exist "%CSC%" (
  echo.
  echo O .NET Framework 4 do Windows nao foi encontrado.
  echo Este PC pode ter uma politica corporativa ou recurso do Windows desabilitado.
  echo.
  pause
  exit /b 2
)

set "NEEDBUILD=1"
if exist "%EXE%" if exist "%VERFILE%" (
  set /p OLDVER=<"%VERFILE%"
  if "!OLDVER!"=="%VERSION%" set "NEEDBUILD=0"
)

if "%NEEDBUILD%"=="1" (
  echo Preparando Outlook Part Sender pela primeira vez...
  copy /Y "%SRC%" "%APPDIR%\OutlookPartSenderPortable.cs" >nul
  "%CSC%" /nologo /target:winexe /optimize+ /platform:anycpu /main:OutlookPartSender.PortableProgram ^
    /out:"%EXE%" ^
    /reference:System.dll ^
    /reference:System.Core.dll ^
    /reference:System.Drawing.dll ^
    /reference:System.Windows.Forms.dll ^
    /reference:System.Web.Extensions.dll ^
    /reference:Microsoft.CSharp.dll ^
    "%APPDIR%\OutlookPartSenderPortable.cs" >"%LOG%" 2>&1

  if errorlevel 1 (
    echo.
    echo Nao foi possivel preparar o programa neste PC.
    echo O computador da empresa pode estar bloqueando compilacao local.
    echo.
    echo Detalhes: "%LOG%"
    echo.
    pause
    exit /b 3
  )
  >"%VERFILE%" echo %VERSION%
)

if not exist "%EXE%" (
  echo ERRO: executavel nao foi criado.
  pause
  exit /b 4
)

start "" "%EXE%"
exit /b 0
