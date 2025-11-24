@echo off
rem ============================================================
rem ===        1-2. TemplatePayloadUtils.bat                 ===
rem === Utilidades comunes para resolver la base de payload ===
rem === y detectar si contiene archivos de plantilla        ===
rem ============================================================

:ResolveBaseDirectory
setlocal
set "RBD_INPUT=%~1"
set "RBD_OUTPUT_VAR=%~2"

if "%RBD_INPUT:~-1%" NEQ "\\" set "RBD_INPUT=%RBD_INPUT%\\"

set "RBD_FOUND="
for %%D in ("%RBD_INPUT%" "%RBD_INPUT%payload\\" "%RBD_INPUT%templates\\" "%RBD_INPUT%extracted\\") do (
    for %%F in ("%%~D*.dot*" "%%~D*.pot*" "%%~D*.xlt*" "%%~D*.thmx") do (
        if exist "%%~fF" set "RBD_FOUND=%%~D"
    )
    if defined RBD_FOUND goto :RBD_Found
)

:RBD_Found
if not defined RBD_FOUND set "RBD_FOUND=%RBD_INPUT%"

endlocal & set "%RBD_OUTPUT_VAR%=%RBD_FOUND%"
exit /b 0

:HasTemplatePayload
setlocal enabledelayedexpansion
set "HP_PATH=%~1"
set "HP_OUT=%~2"
set "HP_FOUND=0"

if not defined HP_PATH goto :HasTemplatePayloadEnd
if "!HP_PATH:~-1!" NEQ "\\" set "HP_PATH=!HP_PATH!\\"

for %%F in ("!HP_PATH!*.dot*" "!HP_PATH!*.pot*" "!HP_PATH!*.xlt*" "!HP_PATH!*.thmx") do (
    if exist "%%~fF" set "HP_FOUND=1"
)

:HasTemplatePayloadEnd
set "HP_RESULT=!HP_FOUND!"
endlocal & if not "%HP_OUT%"=="" set "%HP_OUT%=%HP_RESULT%"
exit /b 0
