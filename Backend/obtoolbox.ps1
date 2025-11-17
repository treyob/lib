`Clear-Host
# Importing the modules
Write-Host "Setting temporary execution policy"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
Write-Host "Setting TLS settings for script execution"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$testing = $false
$modulearray = @(
    "https://github.com/treyob/lib/releases/download/v0.2/dentalsoftware.psm1",
    "https://github.com/treyob/lib/releases/download/v0.2/ob.psm1"
)
$tempPath = "$env:temp\obsoftware"
Function Import-OBModules {
    param($psm1Url)
    # Extract the file name from the URL
    $fileName = [System.IO.Path]::GetFileName($psm1Url)
    # Define the temp path and file path
    $psm1FilePath = Join-Path -Path $tempPath -ChildPath $fileName
    # Create the directory if it doesn't exist
    do {
        New-Item -Path $tempPath -ItemType Directory -Force > $null
    } while (-not (Test-Path -Path $tempPath))
    Invoke-WebRequest -Uri $psm1Url -OutFile $psm1FilePath 
    Import-Module -Name $psm1FilePath
    Write-Host "Imported Module $fileName"
}
if (-Not $testing) {
    ForEach ($psm1Url in $modulearray) {
        Import-OBModules $psm1Url
    }
}
else {
    # If Testing is enabled, pull modules from testing repository
    Write-Warning "Testing: Enabled"
    pause
    Import-OBModules "https://github.com/obmmoreno/obwinrepair/releases/download/obtoolbox-TUI/dentalsoftware.psm1"
    Import-OBModules "https://github.com/obmmoreno/obwinrepair/releases/download/obtoolbox-TUI/ob.psm1"
}

$global:sound = $null

$tuiUrl = "https://github.com/treyob/lib/releases/download/v0.2/obtoolbox_tui.exe"
$fileName = [System.IO.Path]::GetFileName($tuiUrl)
$filePath = Join-Path -Path $tempPath -ChildPath $fileName

Start-BitsTransfer -Source $tuiUrl -Destination $filePath `
                   -Priority Foreground `
                   -TransferType Download `
                   -Description "TUI for OverBytesToolbox"
Write-Host "Downloading $fileName"
Write-Host "Running OverBytesToolbox"




try {
    # Run TUI in current terminal
    & $filePath
}
finally {
    # Delete the file after the program exits
    if (Test-Path $filePath) {
        Remove-Item -Path $tempPath -Force -Recurse
        Write-Host "Deleted temporary folder: $tempPath"
    }
}