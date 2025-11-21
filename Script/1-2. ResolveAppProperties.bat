@echo off
rem ============================================================
rem ===            1-2. ResolveAppProperties.bat             ===
rem ===           Biblioteca de propiedades de la app        ===
rem ===  Uso: call "1-2. ResolveAppProperties.bat" :ResolveAppProperties APP_NAME  ===
rem ===       Devuelve: PROP_REG_NAME, PROP_MRU_VAR,         ===
rem ===                 PROP_COUNTER_VAR                     ===
rem ============================================================

goto :ResolveAppProperties

:ResolveAppProperties
rem APP_NAME esperado:
rem   WORD
rem   POWERPOINT
rem   EXCEL

set "APP_UP=%~1"

set "PROP_REG_NAME="
set "PROP_MRU_VAR="
set "PROP_COUNTER_VAR="

if /I "%APP_UP%"=="WORD" (
    set "PROP_REG_NAME=Word"
    set "PROP_MRU_VAR=WORD_MRU_PATH"
    set "PROP_COUNTER_VAR=GLOBAL_ITEM_COUNT_WORD"
) else if /I "%APP_UP%"=="POWERPOINT" (
    set "PROP_REG_NAME=PowerPoint"
    set "PROP_MRU_VAR=PPT_MRU_PATH"
    set "PROP_COUNTER_VAR=GLOBAL_ITEM_COUNT"
) else if /I "%APP_UP%"=="EXCEL" (
    set "PROP_REG_NAME=Excel"
    set "PROP_MRU_VAR=EXCEL_MRU_PATH"
    set "PROP_COUNTER_VAR=GLOBAL_ITEM_COUNT_EXCEL"
) else (
    rem Si no existe, devolver vacío
    set "PROP_REG_NAME="
    set "PROP_MRU_VAR="
    set "PROP_COUNTER_VAR="
)

exit /b 0
