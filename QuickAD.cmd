REM 
@echo off

SET ScriptDirectory=%~dp0
SET PowerShellScriptPath=%ScriptDirectory%QuickAD.ps1
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '%PowerShellScriptPath%'";
