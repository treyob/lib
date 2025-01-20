@echo off
setlocal enabledelayedexpansion

echo Checking the component store for corruption...

REM Check the component store health
DISM /online /cleanup-image /CheckHealth

REM Get the error level returned by CheckHealth
set errorlevel_checkhealth=%errorlevel%

echo CheckHealth completed with error level: !errorlevel_checkhealth!

REM Check if the component store is repairable
if !errorlevel_checkhealth! equ 0 (
    echo The component store is healthy. Proceeding to detailed scan...
) else (
    echo The component store might be corrupted. Performing a detailed scan...
    
    REM Scan the component store health to get detailed information
    DISM /online /cleanup-image /ScanHealth

    REM Get the error level returned by ScanHealth
    set errorlevel_scanhealth=%errorlevel%

    echo ScanHealth completed with error level: !errorlevel_scanhealth!

    REM Check if detailed scan found issues
    if !errorlevel_scanhealth! equ 0 (
        echo No issues found during detailed scan. Proceeding to system file check...
    ) else (
        echo Issues found during detailed scan. Attempting to repair the component store...
        
        REM Attempt to repair the component store
        DISM /online /cleanup-image /RestoreHealth

        REM Get the error level returned by RestoreHealth
        set errorlevel_restorehealth=%errorlevel%

        echo RestoreHealth completed with error level: !errorlevel_restorehealth!

        REM Check if the RestoreHealth was successful
        if !errorlevel_restorehealth! equ 0 (
            echo Component store repair successful. Proceeding to system file check...
        ) else (
            echo Failed to repair the component store. Exiting...
            pause
            exit /b !errorlevel_restorehealth!
        )
    )
)

REM Run the System File Checker to repair system files
echo Running System File Checker...
sfc /scannow

echo All operations completed.
pause
