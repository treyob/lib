# Dental Software Fixes
Function Invoke-EZDentDLLFix {
    [CmdletBinding()]
    param()
    $dllSourcePath = "C:\EzSensor\libiomp5md.dll"
    $dllDestinationPath = "C:\Program Files (x86)\VATECH\EZDent-i\bin"
    
    # Check if the DLL file exists
    if (Test-Path -Path $dllSourcePath) {
        Write-Host "Copying the IO Library to fix EzDent" -ForegroundColor Blue
        try {
            # Attempt to copy the DLL
            Copy-Item -Path $dllSourcePath -Destination $dllDestinationPath -Force
            Write-Host "Operation completed. DLL copied successfully." -ForegroundColor Blue
        }
        catch {
            Write-Host "An error occurred while copying the DLL: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Error: The DLL 'libiomp5md.dll' was not found at $dllSourcePath." -ForegroundColor Red
    }
    Write-Host "Press Enter to continue." -ForegroundColor Blue
    Read-Host
}
Function Invoke-ESServConnectFix {
    [CmdletBinding()]
    param()
    $mcSourcePath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Config\machine.config"
    $mcdSourcePath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Config\machine.config.default"
    $mcDestinationPath = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\Config\machine.config"
    
    # Check if the machine.config file exists
    if (Test-Path -Path $mcSourcePath) {
        Write-Host "Copying the 'machine.config' file to fix Eaglesoft" -ForegroundColor Blue
        try {
            # Attempt to copy the machine.config
            Copy-Item -Path $mcSourcePath -Destination $mcDestinationPath -Force > $null
            Write-Host "Operation completed. 'machine.config' copied successfully." -ForegroundColor DarkGreen
        }
        catch {
            Write-Host "An error occurred while copying the 'machine.config': $_" -ForegroundColor Red
        }
    }
    else {
        Write-Warning "The 'machine.config' file was not found at $mcSourcePath."
        Write-Host "Attempting to use the machine.config.default instead"
        if (Test-Path -Path $mcdSourcePath) {
            try {
                # Attempt to copy the machine.config.default
                Copy-Item -Path $mcdSourcePath -Destination $mcDestinationPath -Force > $null
                Copy-Item -Path $mcdSourcePath -Destination $mcSourcePath -Force > $null
                Write-Host "Operation completed. 'machine.config' copied from 'machine.config.default' successfully." -ForegroundColor DarkGreen
            }
            catch {
                Write-Host "An error occurred while copying the 'machine.config.default': $_" -ForegroundColor Red
            }    
        }
        else { Write-Host "Did not find 'machine.config' or 'machine.config.default'. Operation failed." -ForegroundColor Red}

    }
    Write-Host "Press Enter to continue." -ForegroundColor Blue
    Read-Host
}
Function Invoke-ESHexDecFix {
    [CmdletBinding()]
    param()
    $userProfile = [System.Environment]::GetFolderPath('UserProfile')
    $pattersonPath = Join-Path -Path $userProfile -ChildPath 'AppData\Local\Patterson_Companies'
    # Check if the folder exists
    if (Test-Path $pattersonPath) {
        # Get all the folders in the Patterson_Companies directory
        $folders = Get-ChildItem -Path $pattersonPath -Directory

        # Loop through each folder and rename it
        foreach ($folder in $folders) {
            $newFolderName = "_$($folder.Name)"  # Add _ at the front of the folder name
            # Check if the new folder name already exists
            $counter = 1
            $newFolderPath = Join-Path -Path $pattersonPath -ChildPath $newFolderName
            while (Test-Path $newFolderPath) {
                # If the folder exists, append another underscore to the name
                $newFolderName = "_$($folder.Name)_$counter"
                $newFolderPath = Join-Path -Path $pattersonPath -ChildPath $newFolderName
                $counter++
            }
            # Rename the folder
            Rename-Item -Path $folder.FullName -NewName $newFolderName
            Write-Host "Renamed $($folder.Name) to $newFolderName"
        }
        # Relaunch Eaglesoft 
        $eaglesoftPath = "C:\EagleSoft\Shared Files\Eaglesoft.exe"  # Adjust the path if needed
        if (Test-Path $eaglesoftPath) {
            Start-Process $eaglesoftPath
            Write-Host "Relaunched Eaglesoft."
        } else {
            Write-Host "Eaglesoft executable not found. Open manually"
        }
    } else {
        Write-Host "Patterson_Companies folder not found at $pattersonPath"
    }
}
Function Clear-EZCache {
    [CmdletBinding()]
    param()
    $ezCachePath = "C:\Program Files (x86)\VATECH\EzDent-i\Cache"
    $firstLoop = $true  # Flag to track the first loop iteration
    # Make sure the folder exists
    if (-Not (Test-Path -PathType Container -Path $ezCachePath)) {
        Write-Warning "EZDent cache folder not found! Exiting"; Read-Host "Press Enter to continue"
        return
    } else {
        Write-Host "Found EZDent cache folder, continuing"
    }

    # Wait for EZDent processes to be closed
    do {
        Clear-Host
        Invoke-PounceCat "Making sure Ez3d-i64 and VTE232 processes are not running"
        $ez3DRunning = Get-Process -Name "Ez3d-i64" -ErrorAction SilentlyContinue
        $ezDentIRunning = Get-Process -Name "VTE232" -ErrorAction SilentlyContinue
        # If the processes are running and it's not the first loop, display a message
        if ($ez3DRunning -and -not $firstLoop) {
            Write-Host "EZ3D (Process: Ez3d-i64) is running. Waiting..."
        }
        if ($ezDentIRunning -and -not $firstLoop) {
            Write-Host "EZDent-i (Process: VTE232) is running. Waiting..."
        }
        $firstLoop = $false
        Start-Sleep -Seconds 1
    } while ($ez3DRunning -or $ezDentIRunning)
    Remove-Item -Recurse $ezCachePath -Force
    if (-Not (Test-Path -PathType Container -Path $ezCachePath)) {
        Write-Host "EZDent cache successfully cleared" -ForegroundColor DarkGreen; Read-Host "Press Enter to continue"
    } else {
        Write-Warning "Something went wrong. EZDent cache could not be deleted. You can try manually. `nPath: $ezCachePath"; pause
    }
}
# Helper functions for Dental Software Section
Function Invoke-SmartDocScannerFix {
    [CmdletBinding()]
    param()
    Start-Process "cmd.exe" -ArgumentList "/c", '"C:\Eaglesoft\Shared Files\setupdocmgr64.bat"' -WorkingDirectory "C:\Eaglesoft\Shared Files"
    Write-Host "Remember to set the scanner to the corresponding WIA-Printer in"
    Write-Host "Eaglesoft > File > Preferences > X-Ray Tab > Scanner" -ForegroundColor DarkCyan
    Read-Host "Press Enter to dismiss"
}
Function Install-MouthwatchDrivers {
    [CmdletBinding()]
    param()
    Clear-Host
    Write-Host "Downloading mouthwatch to C:\mouthwatch.exe"
    Start-BitsTransfer "https://mouthwatch.com/wp-content/uploads/downloads/setupmouthwatch.exe" -Destination "C:\mouthwatch.exe"
    Invoke-UltraCat "Running Mouthwatch. Install files will be deleted" " when mouthwatch is closed"; Start-Process "C:\mouthwatch.exe"
    do {
        $process = Get-Process -Name "mouthwatch" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force "C:\mouthwatch.exe"



}
function Install-ESDexisSensorIntegration {
    param (
        [string]$DownloadUrl = "https://obtoolbox-public.s3.us-east-2.amazonaws.com/3rd-party-tools/Gendex_Dexis_Integration.reg",
        [string]$DestinationPath = "$env:TEMP\obsoftware\Gendex_Dexis_Integration.reg"
    )

    try {
        Write-Host "Downloading .reg file from $DownloadUrl..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $DestinationPath -UseBasicParsing
        Write-Host "Download complete. File saved to $DestinationPath" -ForegroundColor Green

        Write-Host "Importing registry file..." -ForegroundColor Cyan
        Start-Process -FilePath "reg.exe" -ArgumentList "import `"$DestinationPath`"" -Verb RunAs -Wait

        Write-Host "Registry file successfully imported." -ForegroundColor Green
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Error "An error occurred: $_"
    }
}

Function Invoke-DentalSoftwareSection {
    do {
        Clear-Host
        Write-Host "Dental Software Section`n" -ForegroundColor DarkCyan
        Write-Host "Patterson Dental" -ForegroundColor DarkYellow
        Write-Host @"
1. Eaglesoft machine.config Fix: Unable to contact ES Server
2. Eaglesoft Hexdec Fix: White Screen in Pt. acct 
3. Eaglesoft SmartDoc fix: Only scanning 1 page at a time or
    DriverInit Failed when scanning                     _._     _,-'""```-._
4. Eaglesoft Download                                 (,-.```._,'(       |\```-/| 
5. Schick with Eaglesoft 24                                ```-.-' \ )-```( , o o)
6. Eaglesoft Dexis/Gendex sensor integration                      ```-    \```_`"'-"
"@
    Write-Host "VATECH" -ForegroundColor DarkYellow                                                
    Write-Host @" 
7. EzDent-i Fix: libiomp5md.dll error during IO acquisition
8. Clear EzDent Cache

"@
    Write-Host "Sirona" -ForegroundColor DarkYellow                                                
    Write-Host @" 
9. Schick Drivers 
10. Schick sensor Curve Hero documentation
11. Sidexis Migration Section

"@
    Write-Host "Others" -ForegroundColor DarkYellow                                                
    Write-Host @" 

12. Install Mouthwatch Drivers                                    
13 TDO XDR Sensor Fix for Win11: TDO crashes when taking intraoral X-Rays
14. Dentrix Installation and Migration Tool

00. Exit Dental Software Section                               
"@
$dentalChoice = Read-Host "Enter the number of your choice"
    

        switch ($dentalChoice) {
            "1"  { Invoke-ESServConnectFix }
            "2"  { Invoke-ESHexDecFix }
            "3"  { Invoke-SmartDocScannerFix }
            "4"  { Open-ExternalLink 'https://pattersonsupport.custhelp.com/app/answers/detail/a_id/23400#New%20Server' }
            "5"  { Open-ExternalLink 'https://pattersonsupport.custhelp.com/app/answers/detail/a_id/44313/kw/44313' }
            "6"  { Install-ESDexisSensorIntegration }
            "7"  { Invoke-EZDentDLLFix }
            "8"  { Clear-EZCache }
            "9"  { Open-ExternalLink 'https://www.dentsplysironasupport.com/en-us/user_section/user_section_imaging/schick_brand_software.html' }
            "10"  { Open-ExternalLink 'https://curvebulkdata-live.s3.amazonaws.com/35a4ab0e-5b82-4543-9eb9-4d1884e041b2/stored_files/843_68752747c6b9b_Schick_Legacy_Driver_Installation.pdf?response-content-disposition=inline%3B%20filename%2A%3DUTF-8%27%27Schick%2520Legacy%2520Driver%2520Installation.pdf&X-Amz-Content-Sha256=UNSIGNED-PAYLOAD&X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential=AKIAWM5JH3YWMSWI32JV%2F20250714%2Fus-east-1%2Fs3%2Faws4_request&X-Amz-Date=20250714T155053Z&X-Amz-SignedHeaders=host&X-Amz-Expires=43200&X-Amz-Signature=2e1353fb90eeb0374c10f1043ba5c0b9e2b79fb8fd5be6b4b35ef1f239cb490b' }
            "11" { Open-SidexisMigrationSection }
            "12" { Install-MouthwatchDrivers }
            "13" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-TDOXDRW11Fix; Exit } > $null }
            "14" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-DentrixInstallMigrateTool; Exit } > $null 
        Read-Host "Dentrix Installation and Migration tool is downloading in the background and will start in a moment" }
        }
    } while ($dentalChoice -ne "00")
}
