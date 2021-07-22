@echo off
:: BatchGotAdmin
::-------------------------------------
REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"="
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B
:gotAdmin
    pushd "%CD%"
    CD /D "%~dp0"
Regsvr32.exe /s c:\sts\ChamalButton.ocx
Regsvr32.exe /s c:\sts\HookMenu.ocx
Regsvr32.exe /s c:\sts\TABCTL32.OCX
Regsvr32.exe /s c:\sts\MSCOMCT2.OCX
Regsvr32.exe /s c:\sts\MyEllipticButton.ocx
Regsvr32.exe /s c:\sts\vkUserControlsXP.ocx
Regsvr32.exe /s c:\sts\AniGif.ocx
Regsvr32.exe /s c:\sts\MSFLXGRD.ocx
Regsvr32.exe /s c:\sts\MSCOMCTL.OCX
Regsvr32.exe /s c:\sts\Comdlg32.ocx
Regsvr32.exe /s c:\sts\RICHTX32.ocx