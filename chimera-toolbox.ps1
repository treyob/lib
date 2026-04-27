# Setting temporary execution policy
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process -Force
# Setting TLS settings for script execution
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$tuiUrl   = "https://github.com/treyob/lib/releases/download/v1.01/chimera-toolbox.zip"
$tempPath = "$env:TEMP\chimera-toolbox"
$zipPath  = Join-Path $tempPath "chimera-toolbox.zip"
$exePath = Join-Path $tempPath "chimera-toolbox.exe"

# Ensure temp directory exists
if (-not (Test-Path $tempPath)) {
    New-Item -Path $tempPath -ItemType Directory -Force | Out-Null
}

# If on PS 7, use curl. If on Windows PowerShell, try BITS first and fall back to Invoke-WebRequest if it fails
try {
    if ($PSVersionTable.PSEdition -eq "Core") {
        curl.exe -s -S -L $tuiUrl -o $zipPath
        if ($LASTEXITCODE -ne 0) {
            throw "curl failed with exit code $LASTEXITCODE"
        }
    }
    else {
        try {
            Start-BitsTransfer -Source $tuiUrl -Destination $zipPath -ErrorAction Stop
        }
        catch {
            Write-Warning "BITS failed, falling back to Invoke-WebRequest..."
            Invoke-WebRequest -Uri $tuiUrl -OutFile $zipPath -ErrorAction Stop
        }
    }
}
catch {
    Write-Warning "Primary method failed, attempting Invoke-WebRequest fallback..."

    try {
        Invoke-WebRequest -Uri $tuiUrl -OutFile $zipPath -ErrorAction Stop
    }
    catch {
        Write-Error "All download methods failed: $_"
    }
}

# Wait until ZIP exists
while (-not (Test-Path $zipPath)) {
    Start-Sleep -Milliseconds 200
}

# Extract the ZIP file
Expand-Archive -Path $zipPath -DestinationPath $tempPath -Force


# Wait until EXE exists (ensures extraction finished)
while (-not (Test-Path $exePath)) {
    Start-Sleep -Milliseconds 200
}


try { # Run TUI in current terminal 
    Remove-Item -Path $zipPath -Force 
    & $exePath
}
finally { 
    # Delete the file after the program exits 
    if (Test-Path $tempPath) { 
        Remove-Item -Path $tempPath -Force -Recurse 
    } 
}