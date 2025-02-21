# Dental Software Fixes
Function Invoke-EZDentDLLFix {
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
    $mcSourcePath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Config\machine.config"
    $mcDestinationPath = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\Config\machine.config"
    
    # Check if the machine.config file exists
    if (Test-Path -Path $mcSourcePath) {
        Write-Host "Copying the 'machine.config' file to fix Eaglesoft" -ForegroundColor Blue
        try {
            # Attempt to copy the machine.config
            Copy-Item -Path $mcSourcePath -Destination $mcDestinationPath -Force
            Write-Host "Operation completed. 'machine.config' copied successfully." -ForegroundColor Blue
        }
        catch {
            Write-Host "An error occurred while copying the 'machine.config': $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Error: The 'machine.config' file was not found at $mcSourcePath." -ForegroundColor Red
    }
    Write-Host "Press Enter to continue." -ForegroundColor Blue
    Read-Host
}
Function Invoke-ESHexDecFix {
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
Function Invoke-DentalSoftwareSection {
    do {
        Clear-Host
        Write-Host "Dental Software Section`n" -ForegroundColor DarkCyan
        Write-Host "Patterson Dental" -ForegroundColor DarkYellow
        Write-Host @"
1. Eaglesoft machine.config Fix: Unable to contact ES Server
2. Eaglesoft Hexdec Fix: White Screen in Pt. acct 
3. (Experimental) Eaglesoft SmartDoc fix: Only scanning 1 page at a time or
    DriverInit Failed when scanning                     _._     _,-'""```-._
4. Eaglesoft Download                                 (,-.```._,'(       |\```-/| 
5. Schick with Eaglesoft 24                                ```-.-' \ )-```( , o o)
                                                                  ```-    \```_`"'-"
"@                                                  
    Write-Host "Others" -ForegroundColor DarkYellow                                                
    $dentalChoice = Read-Host @" 
6. Schick Drivers                                   
7. EzDent-i Fix: DLL error during IO acquisition                
8. Install Mouthwatch Drivers                                    
9. Exit Dental Software Section                               

Enter the number of your choice
"@
    
# Function to run dental software section
        function Get-DentalSoftwareFix {
            param (
                [string]$choice
            )
            if ($choice -match "ezdent-i") { & Invoke-EZDentDLLFix }
            elseif ($choice -in @("schick","eaglesoft","schickes24")) {
                $link = if ($choice -eq "schick") { 'https://www.dentsplysironasupport.com/en-us/user_section/user_section_imaging/schick_brand_software.html' } 
                elseif ($choice -eq "eaglesoft") { 'https://pattersonsupport.custhelp.com/app/answers/detail/a_id/23400#New%20Server' }
                else {'https://pattersonsupport.custhelp.com/app/answers/detail/a_id/44313/kw/44313'}
                Start-Process msedge.exe "--new-window $link"
            }
            elseif ($choice -in @("eaglesoft machine.config","eaglesoft hexdec","smartdoc1scan")) {
                if ($choice -eq "eaglesoft machine.config") {& Invoke-ESServConnectFix} 
                elseif ($choice -eq "smartdoc1scan") { 
                    Start-Process "cmd.exe" -ArgumentList "/c", '"C:\Eaglesoft\Shared Files\setupdocmgr64.bat"' -WorkingDirectory "C:\Eaglesoft\Shared Files"
                    Write-Host "Remember to set the scanner to the corresponding WIA-Printer in"
                    Write-Host "Eaglesoft > File > Preferences > X-Ray Tab > Scanner" -ForegroundColor DarkCyan
                    Read-Host "Press Enter to dismiss"
                }
                else {& Invoke-ESHexDecFix; Read-Host "Press enter to dismiss"}
            }
            # Mouthwatch fix
            else {
                    Clear-Host
                    Write-Host "Downloading mouthwatch to C:\mouthwatch.exe"
                    Invoke-RestMethod "https://mouthwatch.com/wp-content/uploads/downloads/setupmouthwatch.exe" -OutFile "C:\mouthwatch.exe"
                    Invoke-UltraCat "Running Mouthwatch. Install files will be deleted" " when mouthwatch is closed"; Start-Process "C:\mouthwatch.exe"
                    do {
                        $process = Get-Process -Name "mouthwatch" -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 1
                    } while ($process)
                    Remove-Item -Recurse -Force "C:\mouthwatch.exe"
            }
        }
        switch ($dentalChoice) {
            "1" { Get-DentalSoftwareFix -choice "eaglesoft machine.config" }
            "2" { Get-DentalSoftwareFix -choice "eaglesoft hexdec" }
            "3" { Get-DentalSoftwareFix -choice "smartdoc1scan" }
            "4" { Get-DentalSoftwareFix -choice "eaglesoft" }
            "5" { Get-DentalSoftwareFix -choice "schickes24" }
            "6" { Get-DentalSoftwareFix -choice "schick" }
            "7" { Get-DentalSoftwareFix -choice "ezdent-i" }
            "8" { Get-DentalSoftwareFix -choice "mouthwatch" }
            "9" { Write-Host "Exiting Dental Software Section" }
            
            Default {
                Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
                Read-Host "Press Enter to try again"
            }
        }
    } while ($dentalchoice -ne "9") 
}
