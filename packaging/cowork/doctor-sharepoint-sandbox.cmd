@echo off
setlocal
title Diagnose SharePoint Sandbox for Claude Cowork

set "SERVER=%~dp0runtime\sharepoint-mcp.exe"
if not exist "%SERVER%" (
  echo The bundled SharePoint MCP runtime was not found:
  echo   %SERVER%
  echo Re-extract the complete distribution ZIP and try again.
  pause
  exit /b 1
)

"%SERVER%" doctor
set "RESULT=%ERRORLEVEL%"
echo.
if "%RESULT%"=="0" (
  echo SharePoint Sandbox is ready for Claude Cowork.
) else (
  echo Doctor found a configuration, authentication, dependency, or connectivity problem.
  echo Run setup-sharepoint-sandbox.cmd to renew trusted configuration and sign-in.
)
pause
exit /b %RESULT%
