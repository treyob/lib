function Install-Office {
    param (
        [string]$version,
        [int]$arch
    )
    $xmlPath = "C:\officeodt\officeodt-$version-x$arch.xml"
    $officeODTlink = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_18623-20156.exe"
    if (-Not (Test-Path -PathType Container "C:\officeodt")) {
        mkdir c:\officeodt > $null
        Write-Host "Created the C:\officeodt directory"
    }
    # Define your variables
    switch ($version) {
        "365B" { $version = "O365ProPlusRetail" }
        "2016" { $version = "ProPlus2016Retail"}
    }
    if ((Read-Host "Would you like to exclude OneDrive? (Y/n)") -notin @("n","N","no")) {
        $excludeOneDrive = '<ExcludeApp ID="OneDrive" />'
    }

    # Create the XML content with variable substitution
    $xmlContent = @"
<Configuration>
  <Add SourcePath="C:\officeodt\" OfficeClientEdition="$arch" Channel="Current">
    <Product ID="$version">
      <Language ID="en-us" />
      $excludeOneDrive
    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DisplayLevel" Value="Full" />

</Configuration>
"@
    Write-Host "XML file Contents are as follows:`n$xmlContent"
    # Output the XML to a file
    $xmlContent | Out-File -FilePath $xmlPath -Encoding UTF8
    
    Start-BitsTransfer -Source $officeODTlink -Destination "C:\officeodt\officeodt.exe"
    Write-Host "Please accept the Microsoft License Terms and choose location " -NoNewline
    Write-Host "C:\officeodt" -ForegroundColor Yellow
    & C:\officeodt\officeodt.exe
    # Waiting for officeodt to finish
    do {
        $process = Get-Process -Name "officeodt" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Start-Sleep -Seconds 1

    if (Test-Path -PathType Leaf -Path C:\officeodt\setup.exe) {
        Clear-Host
        Write-Host "Found office setup.exe! Continuing..."
    } else {
        Write-Error "Did not find the setup file in the c:\officeodt folder. Please try again."
        Write-Host "Cleaning up..."
        Remove-Item -Recurse C:\officeodt
        while (Test-Path C:\officeodt) {
            Write-Host "C:\officeodt could not the removed. Trying again in 3 seconds"...
            Remove-Item -Recurse C:\officeodt
        }
        return "Mission failed, we'll get them next time"
    }
    Write-Host "Configuring setup files..."
    & C:\officeodt\setup.exe /download $xmlPath
    Write-Host "Installing..."
    & C:\officeodt\setup.exe /configure $xmlPath
    Write-Host "Setup has finished running."
    Write-Host "Cleaning up..."
    Remove-Item -Recurse C:\officeodt
    while (Test-Path C:\officeodt) {
        Write-Host "C:\officeodt could not the removed. Trying again in 3 seconds"...
        Remove-Item -Recurse C:\officeodt
    }
    Write-Host "Mission complete"
}
function Invoke-OfficeInstall {
    # Architecture selection
    while (($arch = Read-Host "Architecture? (32 or 64)") -notin @("32", "64")) {
        Write-Host "Invalid input. Please enter 32 or 64." -ForegroundColor Red
    }

    # Version selection
    while (($version = Read-Host "Office version? (2016 or 365B or other)") -notin @("2016", "365B","other")) {
        Write-Host "Invalid input. Please enter 2016 or 365B." -ForegroundColor Red
    }
    if ($version -in @("other")) {
        $versionArray = @(
            "AccessRetail",
            "Access2019Retail",
            "Access2021Retail",
            "Access2024Retail",
            "Access2019Volume",
            "Access2021Volume",
            "Access2024Volume",
            "ExcelRetail",
            "Excel2019Retail",
            "Excel2021Retail",
            "Excel2024Retail",
            "Excel2019Volume",
            "Excel2021Volume",
            "Excel2024Volume",
            "HomeBusinessRetail",
            "HomeBusiness2019Retail",
            "HomeBusiness2021Retail",
            "HomeBusiness2024Retail",
            "HomeStudentRetail",
            "HomeStudent2019Retail",
            "HomeStudent2021Retail",
            "Home2024Retail",
            "O365HomePremRetail",
            "OneNoteFreeRetail",
            "OneNoteRetail",
            "OneNote2021Volume",
            "OneNote2024Volume",
            "OutlookRetail",
            "Outlook2019Retail",
            "Outlook2021Retail",
            "Outlook2024Retail",
            "Outlook2019Volume",
            "Outlook2021Volume",
            "Outlook2024Volume",
            "Personal2019Retail",
            "Personal2021Retail",
            "PowerPointRetail",
            "PowerPoint2019Retail",
            "PowerPoint2021Retail",
            "PowerPoint2024Retail",
            "PowerPoint2019Volume",
            "PowerPoint2021Volume",
            "PowerPoint2024Volume",
            "ProfessionalRetail",
            "Professional2019Retail",
            "Professional2021Retail",
            "Professional2024Retail",
            "ProjectProXVolume",
            "ProjectPro2019Retail",
            "ProjectPro2021Retail",
            "ProjectPro2024Retail",
            "ProjectPro2019Volume",
            "ProjectPro2021Volume",
            "ProjectPro2024Volume",
            "ProjectStdRetail",
            "ProjectStdXVolume",
            "ProjectStd2019Retail",
            "ProjectStd2021Retail",
            "ProjectStd2024Retail",
            "ProjectStd2019Volume",
            "ProjectStd2021Volume",
            "ProjectStd2024Volume",
            "ProPlus2019Volume",
            "ProPlus2021Volume",
            "ProPlus2024Volume",
            "ProPlusSPLA2021Volume",
            "ProPlus2019Retail",
            "ProPlus2021Retail",
            "ProPlus2024Retail",
            "PublisherRetail",
            "Publisher2019Retail",
            "Publisher2021Retail",
            "Publisher2019Volume",
            "Publisher2021Volume",
            "Standard2019Volume",
            "Standard2021Volume",
            "StandardSPLA2021Volume",
            "Standard2024Volume",
            "VisioProXVolume",
            "VisioPro2019Retail",
            "VisioPro2021Retail",
            "VisioPro2024Retail",
            "VisioPro2019Volume",
            "VisioPro2021Volume",
            "VisioPro2024Volume",
            "VisioStdRetail",
            "VisioStdXVolume",
            "VisioStd2019Retail",
            "VisioStd2021Retail",
            "VisioStd2024Retail",
            "VisioStd2019Volume",
            "VisioStd2021Volume",
            "VisioStd2024Volume",
            "WordRetail",
            "Word2019Retail",
            "Word2021Retail",
            "Word2024Retail",
            "Word2019Volume",
            "Word2024Volume",
            "Word2021Volume",
            "O365ProPlusEEANoTeamsRetail",
            "O365ProPlusRetail",
            "O365BusinessEEANoTeamsRetail",
            "O365BusinessRetail"
        )
        # Creating a hashtable from the array
        $versionsHT = @{}
        $validInputs = @()
        for ($i = 0; $i -lt $versionArray.Count; $i++) {
            $key = ($i + 1).ToString()
            $versionsHT[$key] = $versionArray[$i]
            $validInputs += $key
        }
        $validInputs += "exit"
        
        # User selection loop
        $selectedKey = ""
        while ($selectedKey -notin $validInputs) {
            Clear-Host
            $versionsHT.GetEnumerator() | Sort-Object { [int]$_.Key } | ForEach-Object {
                Write-Host "$($_.Key): $($_.Value)"
            }
        
            $selectedKey = Read-Host "Please enter a number from the list above or type 'exit' to cancel"
            if ($selectedKey -notin $validInputs) {
                Read-Host "Invalid input. Press Enter to try again"
            }
        }
        
        if ($selectedKey -ne "exit") {
            $version = $versionsHT[$selectedKey]
        } else {
            return "Exiting"
        }
    }

    Clear-Host
    Install-Office $version $arch
}

Function Invoke-ReleaseRenew {
    Write-Host "Releasing and Renewing IP Address..." -ForegroundColor Magenta
    ipconfig /release
    ipconfig /renew
    Write-Host "Operation completed. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Windows repair. Runs sfc and repair-windowsimage (DISM)
Function Invoke-WindowsRepair {
    Write-Host "Repairing Windows..." -ForegroundColor Magenta
    Write-Host "Running DISM to repair Windows image..." -ForegroundColor Magenta
    Repair-WindowsImage -Online -RestoreHealth
    Write-Host "Running System File Checker (sfc /scannow)..." -ForegroundColor Magenta
    sfc /scannow
    Write-Host "Repair completed. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Reset the printer spooler
Function Invoke-ResetSpooler {
    Write-Host "Resetting Print Spooler..." -ForegroundColor Magenta
    Stop-Service -Name spooler -Force
    Start-Service -Name spooler
    Write-Host "Print Spooler has been reset. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Get Device Serial Number
Function Get-DeviceInfo {
    Write-Host "Getting Device Info..." -ForegroundColor Magenta
    Get-WmiObject -Class Win32_BIOS | Select-Object -ExpandProperty SerialNumber
    Write-Host "Operation completed. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Group Policy update
Function Update-GroupPolicy {
    Write-Host "Running GPUpdate..." -ForegroundColor Magenta
    gpupdate /force
    Write-Host "Group Policy Update completed. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Flush and Register DNS
Function Invoke-DNSFlushRegister {
    Write-Host "Flushing and Registering DNS..." -ForegroundColor Magenta
    ipconfig /flushdns
    ipconfig /registerdns
    Write-Host "DNS operations completed. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Create Local Admin User
Function New-LocalAdmin {
    Write-Host "Creating a Local User and Adding to Administrators..." -ForegroundColor Magenta
    $username = Read-Host "Enter the username for the new user"
    $password = Read-Host "Enter the password for the new user" -AsSecureString
    New-LocalUser -Name $username -Password $password -FullName $username -Description "Local user created via script"
    Add-LocalGroupMember -Group "Administrators" -Member $username
    Write-Host "User $username has been created and added to the Administrators group. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Enable PS Remoting
Function Invoke-PSRemoting {
    Write-Host "Enabling PSRemoting..." -ForegroundColor Magenta
    Enable-PSRemoting -Force
    Write-Host "PSRemoting has been enabled. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Disable IPv6
Function Disable-IPv6 {
    Write-Host "Disabling IPv6..." -ForegroundColor Magenta
    Get-NetAdapter | ForEach-Object {
        Write-Host "Disabling IPv6 on adapter: $($_.Name)" -ForegroundColor Magenta
        Disable-NetAdapterBinding -Name $_.Name -ComponentID "ms_tcpip6"
    }
    Write-Host "IPv6 Disabled. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Enable Remote Desktop
Function Enable-RDP {
    Write-Host "Enabling Remote Desktop..." -ForegroundColor Magenta
    Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0
    Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
    Write-Host "Remote Desktop has been enabled. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Allow all Windows Defender Firewall
Function Unlock-FirewallAll {
    Write-Host "Setting Firewall Profiles to Allow" -ForegroundColor Magenta
    Set-NetFirewallProfile -Profile Domain,Private,Public -DefaultInboundAction Allow
    Write-Host "Firewall configuration completed successfully. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Run chkdsk on C
Function Invoke-chkdskC {
    Write-Host "Running CheckDisk for C:. A reboot will be required, can take up to 1 hour." -ForegroundColor Magenta
    chkdsk C: /r  
    Write-Host "Restart PC when it's best. Press Enter to continue." -ForegroundColor Magenta
    Read-Host     
}
# Download C++ Redistributables
Function Get-CppRedist {
    # Function to download and execute the redistributable installer
    function Install-Redistributable {
        param ($FileName, $Url)
        $filePath = "C:\$FileName"
        Invoke-WebRequest $Url -OutFile $filePath
        Write-Host "Running $FileName installer..." -ForegroundColor Magenta
        Start-Process $filePath
    }
    # Install both x64 and x86 C++ redistributables
    Install-Redistributable "c++.exe" "https://aka.ms/vs/17/release/vc_redist.x64.exe"
    Install-Redistributable "c++x86.exe" "https://aka.ms/vs/17/release/vc_redist.x86.exe"
    Write-Host "Press Enter to go back (this will delete the app, make sure it's closed)" -ForegroundColor Magenta
    Read-Host
    Remove-Item C:\c++* -Force
}
# Run Advanced IP Scanner
Function Get-IPScanner {
    $zipUrl = "https://github.com/treyob/lib/releases/download/v0.2/ipscanner.zip"
    $targetDir = "C:\ipscanner"
    Clear-Host; Write-Host "Downloading Advanced IP Scanner"
    Start-BitsTransfer -Source $zipUrl -Destination "$env:TEMP\obsoftware\ipscanner.zip"
    Clear-Host; Write-Host "Extracting Advanced IP Scanner"
    Expand-Archive -Path "$env:TEMP\obsoftware\ipscanner.zip" -DestinationPath $targetDir -Force
    Clear-Host; Write-Host "Running ipscanner"
    Start-Process -FilePath "$targetDir\Advanced_IP_Scanner_2.5.4594.1.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "Advanced IP scanner should now be running." "This screen will exit when you close it."
    do {
        $process = Get-Process -Name "Advanced_IP_Scanner_2.5.4594.1" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir, "$env:TEMP\obsoftware\ipscanner.zip"
}


# This function can determine the state of Bitlocker
Function Get-BitlockerStatus {
    # Get Protection Status
    $protectionOn = manage-bde -status C: | Select-String "Protection Status"
    
    if ($protectionOn -match "Protection Status:\s*Protection On") {
        $protectionOn = $true
    } elseif ($protectionOn -match "Protection Status:\s*Protection Off \(1 reboots left\)") {
        $protectionOn = "suspendreboot"
    } elseif ($protectionOn -match "Protection Status:\s*Protection Off") {
        $protectionOn = $false
    }

    # Get Encryption Status
    $fullyEncrypted = manage-bde -status C: | Select-String "Percent Encrypted"
    if ($fullyEncrypted -match "Percent Encrypted:\s* 100%\s*") {
        $fullyEncrypted = $true
    } else {
        $fullyEncrypted = $false
    }

    # Get Encryption Method
    $encryptionEnabled = manage-bde -status C: | Select-String "Encryption Method"
    if ($encryptionEnabled -notmatch "\s*Encryption Method:\s*None\s*") {
        $encryptionEnabled = $true
    } else {
        $encryptionEnabled = $false
    }

    # Determine BitLocker Status
    if ($protectionOn -eq $true -and $encryptionEnabled) {
        $bitlockerStatus = "protected"
    } elseif ($protectionOn -eq $false -and $encryptionEnabled) {
        $bitlockerStatus = "suspended"
    } elseif ($protectionOn -eq "suspendreboot") {
        $bitlockerStatus = "suspendreboot"
    } elseif ($fullyEncrypted -eq $false -and $protectionOn -eq $false -and $encryptionEnabled -eq $false) {
        $bitlockerStatus = "disabled"
    } else {
        $bitlockerStatus = "unexpected"
        Write-Host "BitLocker is in an unexpected state. See below"
        manage-bde -status C:
    }
    
    return $bitlockerStatus
}
# This function is used for the server power options with power prevention if drive is protected
Function Set-BitLockerState {
    param (
        [string]$action
    )
    if ($bitlockerStatus -in @("suspended","suspendreboot","disabled") -and $action -in @("Shutdown", "Reboot")) {
        Write-Host "BitLocker is suspended, proceeding with $action." -ForegroundColor Green
        $cmd = if ($action -eq "Shutdown") { shutdown /s /t 0 /c '$action by OverBytesTech script' /f /d p:4:1 } else { shutdown /r /t 0 /c "$action by OverBytesTech script" /f /d p:4:1 }
        $cmd
        Write-Host "Server will now $action." -ForegroundColor Yellow
        exit
    } elseif ($action -eq "Suspend") {
        $bitlockerStatus = Get-BitlockerStatus
        if ($bitlockerStatus -eq "suspendreboot") {Write-Host "Resuming first"; Resume-BitLocker -MountPoint "C:"}
        Write-Host "Suspending BitLocker for C: until manually enabled."
        Suspend-BitLocker -MountPoint "C:" -RebootCount 0
        $bitlockerStatus = Get-BitlockerStatus
        if ($bitlockerStatus -eq "suspended") {
            Write-Host "BitLocker successfully suspended for C:." -ForegroundColor Green
        } else {
            Write-Host "Bitlocker Suspend Failed!" -ForegroundColor Red
        }
        Read-Host "Press Enter to dismiss"
    } elseif ($action -eq "SuspendReboot") {
        $bitlockerStatus = Get-BitlockerStatus
        if ($bitlockerStatus -eq "suspended") {Write-Host "Resuming first"; Resume-BitLocker -MountPoint "C:"}
        Write-Host "Suspending BitLocker for C: for 1 reboot."
        Suspend-BitLocker -MountPoint "C:" -RebootCount 1
        $bitlockerStatus = Get-BitlockerStatus
        if ($bitlockerStatus -eq "suspendreboot") {
            Write-Host "BitLocker successfully suspended for C:." -ForegroundColor Green
        } else {
            Write-Host "Bitlocker Suspend Failed!`nBitlockerStatus:" -ForegroundColor Red;Get-BitlockerStatus
        }
        Read-Host "Press Enter to dismiss"
    } elseif ($action -eq "Resume") {
        Write-Host "Resuming BitLocker for C:."
        Resume-BitLocker -MountPoint "C:"
        $bitlockerStatus = Get-BitlockerStatus
        if ($bitlockerStatus -eq "protected") {
            Write-Host "BitLocker successfully resumed for C:." -ForegroundColor Green
        } else {
            Write-Host "Failed to resume BitLocker for C:.`nBitlockerStatus: $bitlockerStatus" -ForegroundColor Red
        }
        Read-Host "Press Enter to dismiss"
    } elseif ($action -eq "Status") {
        $bitlockerStatus = Get-BitlockerStatus
        Write-Host = "Status variable: $bitlockerStatus`n`n"
        manage-bde -status C:
        Read-Host "Press Enter to dismiss"
    } else {
        Write-Host "BitLocker is not suspended. Cannot proceed with $action." -ForegroundColor Red
        Read-Host "Press Enter to dismiss"
    }
}
# Setting Adobe Security Settings
Function Set-AdobeSettings {
    $acrobatPath32 = "Software\Adobe\Acrobat Reader"
    $acrobatPath64 = "Software\Adobe\Adobe Acrobat"

    function Get-UserFromSID {
        param ($sid)
        try {
            $securityIdentifier = New-Object System.Security.Principal.SecurityIdentifier($sid)
            $account = $securityIdentifier.Translate([System.Security.Principal.NTAccount])
        } catch {
            Write-Warning "Could not resolve SID: $sid"
        }
        return $account.Value
    }
    function Set-SecuritySettings {
        param ($path, $key, $arch)
        # Check if the path exists, if not, create the path
        if (-not (Test-Path $path)) {
            New-Item -Path $path -Force
            Write-Host "Created registry path: $path" -ForegroundColor Yellow
        }
        # Check if the key exists, if not, create the key
        if ($null -eq (Get-ItemProperty -Path $path -Name $key -ErrorAction SilentlyContinue))
        {
            Set-ItemProperty -Path $path -Name $key -Value 0
            Write-Host "Created $key key for $username for Acrobat $version $arch" -ForegroundColor DarkGreen
        } else {
            #Write-Host "Working on path $path"
            if ($path -match "TrustManager") {$version = $path -replace '^.*\\([^\\]+)\\TrustManager$', '$1'}
            else {$version = $path -replace '^.*\\([^\\]+)\\Privileged$', '$1'}
            Write-Host "Disabling $key for $username for Acrobat $version $arch"
            Set-ItemProperty -Path $path -Name $key -Value 0
        }
    } 
    # Iterate over every user registry profile
    Clear-Host
    Get-ChildItem -Path "Registry::HKEY_USERS\" | Where-Object { $_.Name -match '^HKEY_USERS\\S-1-5-21-' -and $_.Name -notmatch '_Classes$'} | ForEach-Object {
        $sid = $_.Name -replace '^HKEY_USERS\\', ''
        $username = Get-UserFromSID $sid
        Write-Host "Found user: $username" -ForegroundColor DarkGreen
        # Ensure the correct path to Acrobat 32-Bit and exclude Volatile Environment
        if (Test-Path -Path "Registry::HKEY_USERS\$sid\$acrobatPath32") {
            # Ensure the registry path is not part of Volatile Environment before proceeding
            if ($sid -notmatch '\\Volatile Environment$') {
                Get-ChildItem -Path "Registry::HKEY_USERS\$sid\$acrobatPath32" | ForEach-Object {
                    # Set security settings for each matching user profile
                    Set-SecuritySettings "Registry::$_\TrustManager" "32-Bit"
                    Set-SecuritySettings "Registry::$_\TrustManager" "32-Bit"
                    Set-SecuritySettings "Registry::$_\Privileged" "32-Bit"
                    Write-Host
                }
            } else {
                Write-Host "Volatile Environment detected for SID $sid, skipping this profile."
            }
        } else {
            Write-Host "Acrobat 32-Bit not found"
        }

        # Check for Acrobat 64-Bit and exclude Volatile Environment
        if (Test-Path -Path "Registry::HKEY_USERS\$sid\$acrobatPath64") {
            # Ensure the registry path is not part of Volatile Environment before proceeding
            if ($sid -notmatch '\\Volatile Environment$') {
                Get-ChildItem -Path "Registry::HKEY_USERS\$sid\$acrobatPath64" | ForEach-Object {
                    # Set security settings for each matching user profile
                    Write-Host
                    Set-SecuritySettings "Registry::$_\TrustManager" "bEnhancedSecurityStandalone" "64-bit"
                    Set-SecuritySettings "Registry::$_\TrustManager" "bEnhancedSecurityInBrowser" "64-Bit"
                    Set-SecuritySettings "Registry::$_\Privileged" "bProtectedMode" "64-Bit"
                }
            } else {
                Write-Host "Volatile Environment detected for SID $sid, skipping this profile."
            }
        } else {
            Write-Host "Acrobat 64-Bit not found"
        }

    }
}
# Verbose Cats
Function Invoke-PounceCat {
    param ($msg, $msg2)
    Write-Host @"
   ____
  (.   \
    \  |  
     \ |___(\--/)
   __/    (  . . )   -- $msg
  "'._.    '-.O.'       $msg2
       '-.  \ "|\
          '.,,/'.,,
"@ -ForegroundColor Magenta
}
Function Invoke-UltraCat {
    param ($msg, $msg2)
    Write-Host @"

         /\_/\  
        ( o.o )  -- $msg
         > ^ <      $msg2 
       /       \
      /  |   |  \
     (   |   |   )
      \  |   |  /
       \_|___|_/
"@ -ForegroundColor Magenta
}

Function Get-ADDeviceInfo {
    # Get all computers from AD
    $computers = Get-ADComputer -Filter * | Select-Object -ExpandProperty Name
    # Create an array to store results
    $results = @()
    # Loop through each computer and gather system information
    foreach ($computer in $computers) {
        try {
            # Check if the computer is online
            if (Test-Connection -ComputerName $computer -Count 1 -Quiet) {
                # Get System Information
                $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $computer
                $cpu = Get-CimInstance -ClassName Win32_Processor -ComputerName $computer
                $ram = [math]::Round(($os.TotalVisibleMemorySize / 1MB), 2)  # Convert KB to GB
                $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" -ComputerName $computer
                $bios = Get-CimInstance -ClassName Win32_BIOS -ComputerName $computer

                # Store the results
                $results += [PSCustomObject]@{
                    ComputerName = $computer
                    OSVersion    = "$($os.Caption) ($($os.Version))"
                    CPU          = $cpu.Name
                    RAM_GB       = $ram
                    DiskSize_GB  = [math]::Round(($disk.Size / 1GB), 2)
                    SerialNumber = $bios.SerialNumber
                }
            } else {
                Write-Warning "$computer is offline or unreachable."
            }
        } catch {
            Write-Warning "Failed to get data from $computer. Error: $_"
        }
    }

    # Export results to CSV
    $results | Export-Csv -Path "C:\DomainComputersInfo.csv" -NoTypeInformation

    # Display results
    $results
}

# Server Power and Bitlocker Options
Function Invoke-ServerPowerOptions {
    do {
        Clear-Host
        Write-Host "Server Power Options" -ForegroundColor Magenta
        $powerChoice = Read-Host @"
1. Shutdown now 
2. Reboot now 
3. Suspend BitLocker for C: for 1 reboot 
4. Suspend BitLocker for C: until manually enabled
5. Resume BitLocker for C:
6. Check C: BitLocker Status
7. Exit 

Enter the number of your choice
"@
        $bitlockerStatus = Get-BitlockerStatus
        switch ($powerChoice) {
            "1" { Set-BitLockerState -action "Shutdown" }
            "2" { Set-BitLockerState -action "Reboot" }
            "3" { Set-BitLockerState -action "SuspendReboot" }
            "4" { Set-BitLockerState -action "Suspend" }
            "5" { Set-BitLockerState -action "Resume" }
            "6" { Set-BitLockerState -action "Status" }
            "7" { Write-Host "Exiting Server Power Options" }
            Default {
                Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
                Read-Host "Press Enter to try again"
            }
        }
        } while ($powerChoice -ne "7")
    
}
# Download, extract, and run wiztree
Function Invoke-WizTree {
    $zipUrl = "https://www.diskanalyzer.com/files/wiztree_4_24_portable.zip"
    $targetDir = "C:\WizTree"
    Clear-Host; Write-Host "Downloading WizTree"
    Invoke-WebRequest -Uri $zipUrl -OutFile "$env:TEMP\wiztree.zip"
    Clear-Host; Write-Host "Extracting WizTree"
    Expand-Archive -Path "$env:TEMP\wiztree.zip" -DestinationPath $targetDir
    Clear-Host; Write-Host "Running WizTree."
    Start-Process -FilePath "$targetDir\WizTree64.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "WizTree should now be running." "This screen will exit when you close WizTree."
    do {
        $process = Get-Process -Name "WizTree64" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir, "$env:TEMP\wiztree.zip"
}
Function Invoke-HWInfo {
    $zipUrl = "https://cytranet-dal.dl.sourceforge.net/project/hwinfo/Windows_Portable/hwi_820.zip?viasf=1"
    $targetDir = "C:\HWInfo"
    Clear-Host; Write-Host "Downloading hwinfo"
    Invoke-WebRequest -Uri $zipUrl -OutFile "$env:TEMP\hwinfo.zip"
    Clear-Host; Write-Host "Extracting hwinfo"
    New-Item -ItemType Directory $targetDir
    Expand-Archive -Path "$env:TEMP\hwinfo.zip" -DestinationPath $targetDir
    Invoke-WebRequest -Uri https://github.com/treyob/lib/releases/download/v0.2/HWiNFO64.INI -OutFile "$targetDir\HWiNFO64.INI"
    Clear-Host; Write-Host "Running hwinfo."
    Start-Process -FilePath "$targetDir\hwinfo64.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "hwinfo should now be running." "This screen will exit when you close hwinfo."
    do {
        $process = Get-Process -Name "hwinfo64" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir, "$env:TEMP\hwinfo.zip"
}

# Dot Net Repair Tool
Function Invoke-DotNetRepair {
    $installUrl = "https://github.com/treyob/lib/releases/download/v0.2/NetFxRepairTool.exe"
    $targetDir = "C:\DotNetRepair"
    Clear-Host; Write-Host "Installing and Starting .Net Repair Tool"
    if (-Not (Test-Path -Path $targetDir)) {New-Item -ItemType Directory -Path $targetDir -Force | Out-Null}
    Invoke-WebRequest -Uri $installUrl -OutFile "$targetDir\NetFxRepairTool.exe"
    Clear-Host; Write-Host "Running .Net Repair Tool."
    Start-Process -FilePath "$targetDir\NetFxRepairTool.exe" -Verb RunAs
    Clear-Host; Invoke-UltraCat ".Net Repair Tool should now be running." "This screen will exit when you close the tool."
    do {
        $process = Get-Process -Name "NetFxRepairTool" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir
}
Function Invoke-TDOXDRW11Fix {
    $installUrl = "https://github.com/treyob/lib/releases/download/v0.2/TDOXDRWin11Fix.exe"
    $targetDir = "C:\TDOXDRWin11Fix"
    Clear-Host; Write-Host "Installing and Starting .Net Repair Tool"
    if (-Not (Test-Path -Path $targetDir)) {New-Item -ItemType Directory -Path $targetDir -Force | Out-Null}
    Start-BitsTransfer -Source $installUrl -Destination "$targetDir\TDOXDRWin11Fix.exe"
    Clear-Host; Write-Host "Downloading the script"
    Start-Process -FilePath "$targetDir\TDOXDRWin11Fix.exe" -Verb RunAs
    Clear-Host; Invoke-UltraCat "The fix should now be running." "This screen will exit when you close the installer."
    Start-Sleep -Seconds 2
    do {
        $process = Get-Process -Name "TDOXDRWin11Fix" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir
}
Function Invoke-NetScan {
    Clear-Host
    $Subnet = Read-Host "Enter subnet (e.g., 192.168.1)"
    $FirstIP = Read-Host "Enter the first IP in the range (e.g., 1)"
    $LastIP = Read-Host "Enter the first IP in the range (e.g., 254)"
    $OutputFile = "C:\ipscan.txt"
    $Results = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, [System.Environment]::ProcessorCount * 2)
    $RunspacePool.Open()
    $Runspaces = $FirstIP..$LastIP | ForEach-Object {
        $IP = "$Subnet.$_"
        $Runspace = [powershell]::Create().AddScript({
            param ($IP, $Results)
            
            # Test connectivity and retrieve details if online
            if (Test-Connection -ComputerName $IP -Count 1 -Quiet) {
                $DeviceName = (Resolve-DnsName -Name $IP -ErrorAction SilentlyContinue).NameHost
                $MAC = (Get-NetNeighbor -InterfaceAlias "*" | Where-Object {$_.IPAddress -eq $IP}).LinkLayerAddress

                # Store online device details
                $Results.Add([PSCustomObject]@{
                    "IP Address"  = $IP
                    "Device Name" = $DeviceName
                    "MAC Address" = $MAC
                }) | Out-Null
            }
        }).AddArgument($IP).AddArgument($Results)

        $Runspace.RunspacePool = $RunspacePool
        [PSCustomObject]@{ Pipe = $Runspace; Status = $Runspace.BeginInvoke() }
    }
    $Runspaces | ForEach-Object {
        $_.Pipe.EndInvoke($_.Status)
        $_.Pipe.Dispose()
    }
    $RunspacePool.Close()
    $RunspacePool.Dispose()
    $SortedResults = $Results | Sort-Object { [int]($_."IP Address" -split "\.")[3] }
    $SortedResults | Format-Table -AutoSize
    $SaveToFile = Read-Host "Save results to file? (Y/N)"
    if ($SaveToFile -match "^[Yy]$") {
        $SortedResults | Format-Table -AutoSize | Out-File -Encoding UTF8 -FilePath $OutputFile
        Write-Host "Results saved to $OutputFile" -ForegroundColor Green
    } else {
        Write-Host "Results NOT saved." -ForegroundColor Yellow
    }
    Invoke-PounceCat "Press enter to continue"; Read-Host
}
Function Invoke-TeamViewerQS {
    $installUrl = "https://www.teamviewer.com/link/?url=505374"
    $targetDir = "C:\TeamViewerQS"
    Clear-Host; Write-Host "Downloading and Starting TeamViewer"
    if (-Not (Test-Path -Path $targetDir)) {New-Item -ItemType Directory -Path $targetDir -Force | Out-Null}
    Invoke-WebRequest -Uri $installUrl -OutFile "$targetDir\TeamViewerQS.exe"
    Clear-Host; Write-Host "Running TeamViewer"
    Start-Process -FilePath "$targetDir\TeamViewerQS.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "TeamViewer should now be running."
    do {
        $process = Get-Process | Get-Process | Where-Object { $_.Name -like "TeamViewer*" }
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir
}

Function Invoke-NewTeamViewerQS {
    $installUrl = "https://www.teamviewer.com/link/?url=505374"
    $targetDir = "C:\TeamViewerQS"
    Clear-Host; Write-Host "Downloading and Starting TeamViewer"
    if (-Not (Test-Path -Path $targetDir)) {New-Item -ItemType Directory -Path $targetDir -Force | Out-Null}
    Start-BitsTransfer -Source $installUrl -Destination "$targetDir\TeamViewerQS.exe"
    Clear-Host; Write-Host "Running TeamViewer"
    Start-Process -FilePath "$targetDir\TeamViewerQS.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "TeamViewer should now be running."
    do {
        $process = Get-Process | Get-Process | Where-Object { $_.Name -like "TeamViewer*" }
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir
}

Function Invoke-DentrixInstallMigrateTool {
    $installUrl = "https://dentrix.com/support/core/MigrationAndInstallTool.exe"
    $targetDir = "C:\DentrixInstallMigrateTool"
    Clear-Host; Write-Host "Downloading and Starting Dentrix Installation and Migration tool"
    if (-Not (Test-Path -Path $targetDir)) {New-Item -ItemType Directory -Path $targetDir -Force | Out-Null}
    Start-BitsTransfer -Source $installUrl -Destination "$targetDir\DentrixInstallMigrateTool.exe"
    Clear-Host; Write-Host "Running the tool"
    Start-Process -FilePath "$targetDir\DentrixInstallMigrateTool.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "Dentrix Installation and Migration tool should now be running."
    do {
        $process = Get-Process | Get-Process | Where-Object { $_.Name -like "DentrixInstallMigrateTool" }
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir
}

Function Invoke-RevoUninstaller {
    Clear-Host
    $zipUrl = "https://81081bb5bd2a290e19d0-73f958688cc73e14784b0be099708265.ssl.cf1.rackcdn.com/RevoUninstaller_Portable.zip"
    $targetDir = "C:\RevoUninstaller"
    Clear-Host; Write-Host "Downloading RevoUninstaller"
    Invoke-WebRequest -Uri $zipUrl -OutFile "$env:TEMP\obsoftware\RevoUninstaller.zip"
    Clear-Host; Write-Host "Extracting RevoUninstaller"
    New-Item -ItemType Directory $targetDir -Force
    Expand-Archive -Path "$env:TEMP\obsoftware\RevoUninstaller.zip" -DestinationPath $targetDir -Force
    Clear-Host; Write-Host "Running RevoUninstaller."
    Start-Process -FilePath "$targetDir\RevoUninstaller_Portable\RevoUPort.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "RevoUninstaller should now be running." "This screen will exit when you close RevoUninstaller."
    Start-Sleep -Seconds 3
    do {
        $process = Get-Process -Name "Revo*" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir, "$env:TEMP\obsoftware\RevoUninstaller.zip"
}

#Function Start-HoldMusic {
#    param ()
#    if (-Not (Test-Path -Path "$env:temp\obsoftware\hold.wav")) {
#        Write-Host "Downloading hold music"
#        Start-BitsTransfer -Source "https://github.com/treyob/lib/releases/download/v0.2/opus.no.1.wav" -Destination "$env:temp\obsoftware\hold.wav" > $null
#    }
#    $global:sound = New-Object System.Media.SoundPlayer "$env:temp\obsoftware\hold.wav"
#    $global:sound.PlayLooping()
#}
#
#Function Stop-HoldMusic {
#    if ($global:sound -ne $null) {
#        $global:sound.Stop()
#        $global:sound.Dispose()
#        $global:sound = $null
#    }
#}
Function Start-HoldMusic {
    param (
        [string]$link,
        [string]$songName
    )
    Stop-HoldMusic
    if (-Not (Test-Path -Path "$env:temp\obsoftware\$songName.wav")) {
        Write-Host "Downloading hold music"
        Start-BitsTransfer -Source "$link" -Destination "$env:temp\obsoftware\$songName.wav" > $null
    }
    $global:sound = New-Object System.Media.SoundPlayer "$env:temp\obsoftware\$songName.wav"
    $global:sound.PlayLooping()
}

Function Stop-HoldMusic {
    if ($null -ne $global:sound) {
        $global:sound.Stop()
        $global:sound.Dispose()
        $global:sound = $null
    }
}
function Invoke-HoldMusicSection {
    param ()
    do {
        Clear-Host
        Write-Host "Hold Music Section`n" -ForegroundColor DarkCyan
        $holdMusicTable = @(
            [PSCustomObject]@{ Option = '1. Other'; Song = 'Opus No. 1 - Tim Carleton'}
            [PSCustomObject]@{ Option = '2. Dexis'; Song = '12 Spanish Dances in G Major Arr. for Guitar'}
            [PSCustomObject]@{ Option = '3. Stop Music'}
        )
        $holdMusicTable | Format-Table -AutoSize

        switch ($musicChoice) {
            "1" { Start-HoldMusic -link 'https://github.com/treyob/lib/releases/download/v0.2/opus.no.1.wav' -songName 'Opus No. 1 - Tim Carleton' }
            "2" { Start-HoldMusic -link 'https://github.com/treyob/lib/releases/download/v0.2/12.span.dance.wav' -songName '12 Spanish Dances in G Major Arr. for Guitar' }
            "3" { Stop-HoldMusic }
        }
        $musicChoice = Read-Host "Enter the number of your choice, or press enter to exit"
    } while ($musicChoice -ne "")
}
