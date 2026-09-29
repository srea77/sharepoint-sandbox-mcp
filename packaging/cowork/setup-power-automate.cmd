@echo off
setlocal
title Set up Power Automate for SharePoint Sandbox

set "ROOT=%~dp0"
set "SERVER=%ROOT%runtime\sharepoint-mcp.exe"
set "INSTALLER=%ROOT%install-azure-cli-user.ps1"

if not exist "%SERVER%" (
  echo The bundled SharePoint MCP runtime was not found:
  echo   %SERVER%
  pause
  exit /b 1
)

where az.cmd >nul 2>nul
if errorlevel 1 (
  if not exist "%INSTALLER%" (
    echo Azure CLI is missing and the reviewed installer was not found.
    pause
    exit /b 1
  )
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%INSTALLER%"
  if errorlevel 1 goto :failed
)

echo.
echo Signing the current Windows user into Power Automate through Microsoft Azure CLI...
"%SERVER%" power-automate-login
if errorlevel 1 goto :failed

echo.
echo Discovering available Power Platform environments for trusted policy setup...
"%SERVER%" discover-power-automate
if errorlevel 1 goto :failed

echo.
echo Power Automate authentication is ready.
echo Next, use the displayed environment ID with:
echo   sharepoint-mcp.exe discover-power-automate --environment-id ID
echo Then add only approved flow IDs and aliases to a reviewed runtime policy and run:
echo   sharepoint-mcp.exe configure-scope --policy PATH
echo.
pause
exit /b 0

:failed
echo.
echo Power Automate setup did not complete. Existing SharePoint configuration was preserved.
pause
exit /b 1
