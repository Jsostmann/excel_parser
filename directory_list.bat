@echo off
setlocal enabledelayedexpansion

if "%1"=="" (
    echo Usage: %0 ^<number_of_directories_to_check^>
    exit /b 1
)

set "dirCount=0"
set "maxDirs=%1"

for /d %%D in (*) do (
    if !dirCount! geq !maxDirs! (
        goto :done
    )

    set "fileCount=0"

    for %%F in ("%%D\*") do (
        if not "%%~aF"=="d" (
            set /a fileCount+=1
        )
    )

    if !fileCount! neq 2 (
        echo %%D
    )

    set /a dirCount+=1
)

:done
endlocal
