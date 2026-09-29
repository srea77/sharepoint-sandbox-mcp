@echo off
setlocal
title Set up SharePoint Sandbox for Claude Cowork

set "ROOT=%~dp0"
set "SERVER=%ROOT%runtime\sharepoint-mcp.exe"
set "PLUGIN=%ROOT%sharepoint-sandbox.plugin"

if not exist "%SERVER%" (
  echo The bundled SharePoint MCP runtime was not found:
  echo   %SERVER%
  echo Re-extract the complete distribution ZIP and try again.
  pause
  exit /b 1
)

echo.
echo SharePoint Sandbox trusted setup
echo --------------------------------
echo Paste the canonical URL of the disposable site that Claude may control.
echo Example: https://contoso.sharepoint.com/sites/ai-agent-sandbox
echo Leave it blank only if this Windows user was already configured.
echo.
set /p "SITE_URL=Sandbox site URL: "

if not "%SITE_URL%"=="" (
  "%SERVER%" sharepoint-browser-login --site-url "%SITE_URL%"
  if errorlevel 1 goto :failed
  "%SERVER%" configure --site-url "%SITE_URL%"
  if errorlevel 1 goto :failed
)

echo.
echo Checking authentication, immutable configuration, and site connectivity...
"%SERVER%" doctor
if errorlevel 1 goto :failed

if not exist "%PLUGIN%" (
  echo.
  echo Setup succeeded, but the Cowork plugin file was not found:
  echo   %PLUGIN%
  echo Upload the .plugin file from Cowork manually.
  pause
  exit /b 0
)

echo.
echo Trusted setup succeeded. Opening the plugin in Claude Desktop...
start "" "%PLUGIN%"
echo If Claude does not open, use Cowork ^> Customize ^> Plugins ^> Add or Upload
echo and select sharepoint-sandbox.plugin beside this script.
echo.
pause
exit /b 0

:failed
echo.
echo Setup did not complete. No extension was installed.
echo Resolve the message above, then run this file again.
pause
exit /b 1
