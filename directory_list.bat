@echo off
setlocal enabledelayedexpansion

:: Iterate over directories in the current directory
for /d %%D in (*) do (
    set "fileCount=0"
    
    :: Count the number of files in the directory
    for %%F in ("%%D\*") do (
        if not "%%~aF"=="d" (
            set /a fileCount+=1
        )
    )

    :: Check if the file count is not equal to 2
    if !fileCount! neq 2 (
        echo %%D
    )
)

endlocal