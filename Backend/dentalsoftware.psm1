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
        else { Write-Host "Did not find 'machine.config' or 'machine.config.default'. Operation failed." -ForegroundColor Red }

    }
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
    }
    else {
        Write-Host "Patterson_Companies folder not found at $pattersonPath"
    }
}
Function Invoke-Eaglesoft {
    # Relaunch Eaglesoft 
    $eaglesoftPath = "C:\EagleSoft\Shared Files\Eaglesoft.exe"  # Adjust the path if needed
    if (Test-Path $eaglesoftPath) {
        Start-Process $eaglesoftPath
        Write-Host "Relaunched Eaglesoft."
    }
    else {
        Write-Host "Eaglesoft executable not found. Open manually"
    }
}
Function Invoke-OCXREGFix {
    [CmdletBinding()]
    Param()
    $ocxreg = "C:\Eaglesoft\Shared Files\OcxReg.exe"
    if (Test-Path -Path $ocxreg) {
        Write-Host "Found OCXREG.exe"
    } else {
        Write-Error "OCXREG.exe not found in $ocxreg"
        return
    }
    Write-Host "Running OCXREG.exe" -ForegroundColor Blue
    Start-Process -FilePath $ocxreg -Verb RunAs -Wait
    Write-Host "OCXREG completed" 
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
    }
    else {
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
    }
    else {
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
Function Disable-MemoryIntegrity {
    # Define registry paths
    $mciKey = "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity"
    $vbsKey = "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard"

    # Try to read the current values
    try {
        $mciValue = (Get-ItemProperty -Path $mciKey -Name Enabled -ErrorAction Stop).Enabled
    }
    catch {
        Write-Warning "Memory Integrity registry key not found. It may not be enabled on this system."
        Start-Sleep -Seconds 2
        return
    }

    try {
        $vbsValue = (Get-ItemProperty -Path $vbsKey -Name EnableVirtualizationBasedSecurity -ErrorAction Stop).EnableVirtualizationBasedSecurity
    }
    catch {
        $vbsValue = 0
    }

    # Check if already disabled
    if ($mciValue -eq 0 -and $vbsValue -eq 0) {
        Write-Host "✅ Memory Integrity is already disabled. No changes made."
        Start-Sleep -Seconds 2
        return
    }

    # Disable Memory Integrity
    Set-ItemProperty -Path $mciKey -Name Enabled -Value 0
    Set-ItemProperty -Path $vbsKey -Name EnableVirtualizationBasedSecurity -Value 0

    Write-Warning "⚠️ Memory Integrity has been disabled. Please restart your computer for changes to take effect."
}
function Stop-EaglesoftProcesses {
    Write-Host "Checking for Eaglesoft-related processes..." -ForegroundColor Blue
    $processesToStop = @("Eaglesoft", "esmessenger", "esiconnect", "AutoDetectServer", "esinetconnect")
    foreach ($processName in $processesToStop) {
        $process = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($process) {
            Write-Host "$processName is running. Stopping process..."
            Stop-Process -Name $processName -Force
        }
    }
}
Function Install-MSXML4 {
    Clear-Host
    Write-Host "Downloading MSXML4 install"
    $destinationDir = "$env:TEMP\obsoftware\msxml" 
    if (-not (Test-Path -Path $destinationDir)) {
        mkdir $destinationDir
    }
    $exists = Get-ChildItem -Path $destinationDir -Filter "msxml4-*.exe" | Where-Object { -not $_.PSIsContainer }
    if ($exists) {
        Write-Host "Found MSXML 4.0 installer, skipping download"
    }
    else {
        Start-BitsTransfer -Source "https://pattersonsupport.custhelp.com/euf/assets/Answers/20595/MSXML_4.0.zip" -Destination "$destinationDir\MSXML4.0.zip"
        Expand-Archive -Path "$destinationDir\MSXML4.0.zip" -DestinationPath $destinationDir
        Remove-Item "$destinationDir\MSXML4.0.zip" > $null
    }
    $file = Get-ChildItem -Path $destinationDir -Filter "msxml4-*.exe" | Select-Object -First 1
    if ($file) {
        Start-Process $file.FullName -Verb RunAs -Wait
        do {
            $process = Get-Process -Name "$($file.BaseName)" -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 1
        } while ($process)
        Remove-Item $destinationDir -Force -Recurse > $null 
    }
    else {
        Write-Host "No MSXML4 installer found." -ForegroundColor Red
    }
}
Function Install-IOSS {
    $destinationDir = "$env:TEMP\obsoftware\IOSS" 
    $downloadLink = "https://www.dentsplysironasupport.com/content/dam/master/product-procedure-brand-categories/imaging/product-categories/software/schick-software/IOSS_v3.2.zip"
    if (-not (Test-Path -Path $destinationDir)) {
        mkdir $destinationDir
    }
    Start-BitsTransfer -Source $downloadLink -Destination "$destinationDir\IOSS.zip"
    Expand-Archive -Path "$destinationDir\IOSS.zip" -DestinationPath $destinationDir
    Remove-Item "$destinationDir\IOSS.zip" -Force -ErrorAction SilentlyContinue
    Start-Process "$destinationDir\IOSS_v3.2\Autorun.exe" -Verb RunAs -Wait
    do {
        $process = Get-Process -Name "Autorun" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item $destinationDir -Force -Recurse -ErrorAction SilentlyContinue | Out-Null
}
# The goal of this function is to automatically restart the service on failure and change "logon as" to local system account
Function Optimize-IOSSService {
    [CmdletBinding()]
    Param()
    Start-Process -FilePath "sc.exe" -ArgumentList 'config', 'SironaUSBService', 'obj=', 'LocalSystem' -NoNewWindow -Wait
    Write-Host "Set LogOn user to LocalSystem"
    Start-Process -FilePath "sc.exe" -ArgumentList 'failure', 'SironaUSBService', 'reset=', '0', 'actions=', 'restart/5000/restart/5000/restart/5000' -NoNewWindow -Wait
    Write-Host "Set to restart service on failure"

    try {
        Write-Host "Setting StartupType to Automatic (Delayed)"
        Set-Service -Name SironaUSBService -StartupType AutomaticDelayed -ErrorAction Stop
        Write-Host "IOSS Service set to Automatic (Delayed) startup" -ForegroundColor Green
    }
    catch {
        Write-Error "Unable to set service `"Sirona Intraoral Sensor Software`" StartupType to `"Automatic Delayed.`" Please do so manually."
        Write-Error "$_"
        Pause
    }
    try {
        Write-Host "Restarting IOSS Service"
        Restart-Service -Name SironaUSBService -ErrorAction Stop
        Write-Host "IOSS Service restarted" -ForegroundColor Green
        Start-Sleep -Seconds 1
    }
    catch {
        Write-Error "Failed to restart the `"Sirona Intraoral Sensor Software`" service. Please do so manually"
        Write-Error "$_"
        pause
    }
}
function Install-ESSchickSensorIntegration {
    param (
        [string]$DownloadUrl = "https://obtoolbox-public.s3.us-east-2.amazonaws.com/3rd-party-tools/Schick_Sensor_Integration.reg",
        [string]$DestinationPath = "$env:TEMP\obsoftware\Schick_Sensor_Integration.reg"
    )

    try {
        Write-Host "Downloading .reg file from $DownloadUrl..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $DestinationPath -UseBasicParsing
        Write-Host "Download complete. File saved to $DestinationPath" -ForegroundColor Green

        Write-Host "Importing registry file..." -ForegroundColor Cyan
        Start-Process -FilePath "reg.exe" -ArgumentList "import `"$DestinationPath`"" -Verb RunAs -Wait

        Write-Host "Schick Sensor Registry file successfully imported." -ForegroundColor Green
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Error "An error occurred: $_"
    }
}

Function Install-CDREliteDriver {
    $destinationDir = "$env:TEMP\obsoftware\CDRElite" 
    $downloadLink = "https://www.dentsplysironasupport.com/content/dam/websites/dentsplysironasupport/schick-brand-software/CDRElite5_16.zip"
    $zipFileName = "CDR.zip"
    $processName = "CDR Elite Setup"
    if ((Test-Path -Path $destinationDir)) {
        Remove-Item $destinationDir -Force -Recurse > $null
    }
    mkdir $destinationDir
    Start-BitsTransfer -Source $downloadLink -Destination "$destinationDir\$zipFileName"
    Expand-Archive -Path "$destinationDir\$zipFileName" -DestinationPath $destinationDir
    Remove-Item "$destinationDir\$zipFileName"
    Start-Process "$destinationDir\CDRElite\CDR Elite Setup.exe" -Verb RunAs -Wait
    do { 
        $process = Get-Process -Name $processName -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
}
Function Install-CDRPatch {
    $patchMsiPath = "$env:TEMP\obsoftware\CDRElite\CDRElite\Patch\CDRPatch-2808.msi"
    Stop-EaglesoftProcesses
    $path = "C:\Program Files (x86)\Schick Technologies\Shared Files"
    $excludedFiles = @("CDRData.dll", "OMEGADLL.dll")
    # Delete all files except the excluded ones
    Get-ChildItem -Path $path -File | Where-Object { $excludedFiles -notcontains $_.Name } | Remove-Item -Force
    # Delete all subfolders
    Get-ChildItem -Path $path -Directory | Remove-Item -Recurse -Force
    if (Test-Path $patchMsiPath) {
        Write-Host "Installing CDRPatch"
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$patchMsiPath`" /qn /norestart" -Wait -Verb RunAs
    }
    else {
        Write-Warning "Patch MSI not found: $patchMsiPath"
    }
}
Function Invoke-ESIOSSSection {
    do {
        Clear-Host
        Write-Host "Eaglesoft 24 Schick Stuff`n" -ForegroundColor DarkCyan
        $es24SchickTable = @(
            [PSCustomObject]@{ Step = '1'; Script = 'Disable Memory Integrity' }
            [PSCustomObject]@{ Step = '2'; Script = 'MSXML 4.0 Install' }
            [PSCustomObject]@{ Step = '3'; Script = 'Install IOSS' }
            [PSCustomObject]@{ Step = '4'; Script = 'Optimize the IOSS Service' }
            [PSCustomObject]@{ Step = '5'; Script = 'Install CDRElite Driver' }
            [PSCustomObject]@{ Step = '6'; Script = 'CDRElite Patching and Configuring ES24.20+' }
            [PSCustomObject]@{ Step = '7'; Script = 'Schick Sensor Integration for OPs' }
        )
        $es24SchickTable | Format-Table -AutoSize
        $es24SchickChoice = Read-Host "Enter the number of your choice, or press enter to exit"
        switch ($es24SchickChoice) {
            "1" { Disable-MemoryIntegrity }
            "2" { Install-MSXML4 }
            "3" { Install-IOSS }
            "4" { Optimize-IOSSService }
            "5" { Install-CDREliteDriver }
            "6" { Install-CDRPatch }
            "7" { Install-ESSchickSensorIntegration }
            "" { Write-Host "Exiting" }
            Default { Write-Host "Invalid, please try again" -ForegroundColor DarkYellow; Start-Sleep -Seconds 1 }
        }
    } while ($es24SchickChoice -ne "")

}

Function Invoke-DentalSoftwareSection {
    do {
        Clear-Host
        Write-Host "Dental Software Section`n" -ForegroundColor DarkCyan
        Write-Host "Patterson Dental" -ForegroundColor DarkYellow
        Write-Host @"
1. Eaglesoft General Fix (machine.config & hexadec & ocxreg)
2. Eaglesoft SmartDoc fix: Only scanning 1 page at a time or
    DriverInit Failed when scanning                     _._     _,-'""```-._
3. Eaglesoft Download                                 (,-.```._,'(       |\```-/| 
4. Schick with Eaglesoft 24                                ```-.-' \ )-```( , o o)
5. Eaglesoft Dexis/Gendex sensor integration                      ```-    \```_`"'-"
6. ES24.20 IOSS Configuration
"@
        Write-Host "VATECH" -ForegroundColor DarkYellow                                                
        Write-Host @" 
7. EzDent-i Fix: libiomp5md.dll error during IO acquisition
8. Clear EzDent Cache

"@
        Write-Host "Sirona" -ForegroundColor DarkYellow                                                
        Write-Host @" 
9. Schick Drivers 
10. Sidexis Migration Section for 4.3

"@
        Write-Host "Others" -ForegroundColor DarkYellow                                                
        Write-Host @" 

11. Install Mouthwatch Drivers                                    
12 TDO XDR Sensor Fix for Win11: TDO crashes when taking intraoral X-Rays
13. Dentrix Installation and Migration Tool

00. Exit Dental Software Section                               
"@
        $dentalChoice = Read-Host "Enter the number of your choice"
    

        switch ($dentalChoice) {
            # Patterson
            "1" { 
                Stop-EaglesoftProcesses
                Invoke-ESServConnectFix 
                Invoke-ESHexDecFix
                Invoke-OCXREGFix
                Invoke-Eaglesoft
                Write-Host "Press Enter to continue." -ForegroundColor Blue
                Read-Host
            }
            "2" { Invoke-SmartDocScannerFix }
            "3" { Open-ExternalLink 'https://pattersonsupport.custhelp.com/app/answers/detail/a_id/23400#New%20Server' }
            "4" { Open-ExternalLink 'https://pattersonsupport.custhelp.com/app/answers/detail/a_id/44313/kw/44313' }
            "5" { Install-ESDexisSensorIntegration }
            "6" { Invoke-ESIOSSSection }
            # VATECH
            "7" { Invoke-EZDentDLLFix }
            "8" { Clear-EZCache }
            # Sirona
            "9" { Open-ExternalLink 'https://www.dentsplysironasupport.com/en-us/user_section/user_section_imaging/schick_brand_software.html' }
            "10" { Open-SidexisMigrationSection }
            # Others
            "11" { Install-MouthwatchDrivers }
            "12" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-TDOXDRW11Fix; Exit } > $null }
            "13" {
                Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-DentrixInstallMigrateTool; Exit } > $null 
                Read-Host "Dentrix Installation and Migration tool is downloading in the background and will start in a moment" 
            }
        }
    } while ($dentalChoice -ne "00")
}
