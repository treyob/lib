do {
    Clear-Host
    Write-Host "Choose an option:" -ForegroundColor Magenta
    Write-Host "1. Release and Renew IP Address" -ForegroundColor Magenta
    Write-Host "2. Repair Windows (System File Checker and DISM)" -ForegroundColor Magenta
    Write-Host "3. Reset Print Spooler" -ForegroundColor Magenta
    Write-Host "4. Get Device Serial Number" -ForegroundColor Magenta
    Write-Host "5. Run GPUpdate" -ForegroundColor Magenta
    Write-Host "6. Flush and Register DNS" -ForegroundColor Magenta
    Write-Host "7. Create Local User and Add to Administrators" -ForegroundColor Magenta
    Write-Host "8. Enable PSRemoting" -ForegroundColor Magenta
    Write-Host "9. Disable IPv6" -ForegroundColor Magenta
    Write-Host "10. Enable Remote Desktop" -ForegroundColor Magenta
    Write-Host "11. Set Firewall Profiles to Allow All" -ForegroundColor Magenta
    Write-Host "12. CheckDisk C:\" -ForegroundColor Magenta
    Write-Host "13. Imaging Software Fixes" -ForegroundColor Magenta
    Write-Host "14. Download and Install Microsoft Visual C++ 2015-2022" -ForegroundColor Magenta
    Write-Host "15. Exit" -ForegroundColor Magenta

    $choice = Read-Host "Enter the number of your choice"

    switch ($choice) {
        "1" {
            Write-Host "Releasing and Renewing IP Address..." -ForegroundColor Magenta
            ipconfig /release
            ipconfig /renew
            Write-Host "Operation completed. Press Enter to continue." -ForegroundColor Magenta
            Read-Host
        }
        "2" {
            Write-Host "Repairing Windows..." -ForegroundColor Magenta
            Write-Host "Running System File Checker (sfc /scannow)..." -ForegroundColor Magenta
            sfc /scannow
            Write-Host "Running DISM to repair Windows image..." -ForegroundColor Magenta
            DISM /Online /Cleanup-Image /RestoreHealth
            Write-Host "Repair completed. Press Enter to continue." -ForegroundColor Magenta
            Read-Host
        }
        "3" {
            Write-Host "Resetting Print Spooler..." -ForegroundColor Magenta
            Stop-Service -Name spooler -Force
            Start-Service -Name spooler
            Write-Host "Print Spooler has been reset. Press Enter to continue." -ForegroundColor Magenta
            Read-Host
        }
        "4" {
            Write-Host "Getting Device Serial Number..." -ForegroundColor Magenta
            Get-WmiObject -Class Win32_BIOS | Select-Object -ExpandProperty SerialNumber
            Write-Host "Operation completed. Press Enter to continue." -ForegroundColor Magenta
            Read-Host
        }
        "5" {
            Write-Host "Running GPUpdate..." -ForegroundColor Magenta
            gpupdate /force
            Write-Host "Group Policy Update completed. Press Enter to continue." -ForegroundColor Magenta
            Read-Host
        }
        "6" {
            Write-Host "Flushing and Registering DNS..." -ForegroundColor Magenta
            ipconfig /flushdns
            ipconfig /registerdns
            Write-Host "DNS operations completed. Press Enter to continue." -ForegroundColor Magenta
            Read-Host
        }
        "7" {
            Write-Host "Creating a Local User and Adding to Administrators..." -ForegroundColor Magenta
            $username = Read-Host "Enter the username for the new user"
            $password = Read-Host "Enter the password for the new user" -AsSecureString
            New-LocalUser -Name $username -Password $password -FullName $username -Description "Local user created via script"
            Add-LocalGroupMember -Group "Administrators" -Member $username
            Write-Host "User $username has been created and added to the Administrators group. Press Enter to continue." -ForegroundColor Magenta
            Read-Host
        }
        "8" {
            Write-Host "Enabling PSRemoting..." -ForegroundColor Magenta
            Enable-PSRemoting -Force
            Write-Host "PSRemoting has been enabled. Press Enter to continue." -ForegroundColor Magenta
            Read-Host
        }
        "9" {
            Write-Host "Disabling IPv6..." -ForegroundColor Magenta
            $adapters = Get-NetAdapter

foreach ($adapter in $adapters) {
    Write-Host "Disabling IPv6 on adapter: $($adapter.Name)" -ForegroundColor Magenta
    Disable-NetAdapterBinding -Name $adapter.Name -ComponentID "ms_tcpip6"
}
            Write-Host "IPv6 Disabled. Press Enter to continue." -ForegroundColor Magenta
            Read-Host
        }
        "10" {
            Write-Host "Enabling Remote Desktop..." -ForegroundColor Magenta
            Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0
            Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
            Write-Host "Remote Desktop has been enabled. Press Enter to continue." -ForegroundColor Magenta
            Read-Host
        }
        "11" {
            Write-Host "Setting Firewall Profiles to Allow" -ForegroundColor Magenta
            Set-NetFirewallProfile -Profile Domain,Private,Public -DefaultInboundAction Allow
            Write-Host "Firewall configuration completed successfully. Press Enter to continue." -ForegroundColor Magenta
            Read-Host
        }
        "12" {
            Write-Host "Running CheckDisk for C:. A reboot will be required, can take up to 1 hour." -ForegroundColor Magenta
            chkdsk C: /r  
            Write-Host "Restart PC when its best. Press Enter to Continue." -ForegroundColor Magenta
            Read-Host     
        }
        "13" {
            Write-Host "Entering Imaging Software Fixes Section..." -ForegroundColor Blue
            do {
                Clear-Host
                $imagingchoice = Read-Host @"
1. EzDent-i DLL error during IO acquisition
2. Exit Imaging Software Fixes Section
Enter the number of your choice
"@
                switch ($imagingchoice) {
                    "1" {
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
                    "2" {
                        Write-Host "Exiting Imaging Software Fixes Section..." -ForegroundColor Blue
                    }
                    Default {
                        Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
                        Read-Host "Press Enter to try again"
                    }
                }
            } while ($imagingchoice -ne "2")  # Loop until the user selects to exit
        }
        "14" {
            Write-Host "Downloading C++" -ForegroundColor Magenta
            md C:\obsoftware;iwr https://aka.ms/vs/17/release/vc_redist.x64.exe -OutFile C:\obsoftware\c++.exe;iwr https://aka.ms/vs/17/release/vc_redist.x86.exe -OutFile C:\obsoftware\C++x86.exe;C:\obsoftware\C++.exe;C:\obsoftware\C++x86.exe;Write-Host "Press Enter to Continue" -ForegroundColor Magenta;Read-Host
          
        }
        "15" {
            Write-Host "Exiting the program. Goodbye!" -ForegroundColor Magenta
        }
        
        
    }
} while ($choice -ne "15")