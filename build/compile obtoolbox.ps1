# Compile
# Resolve the repo root relative to this script's own location instead of a
# path hardcoded to one machine/user, so this works wherever the repo is
# checked out (this script lives in <repo>\build, so the repo root is one
# level up).
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

nuitka `
  --msvc=latest `
  --standalone `
  --python-flag=no_site `
  --noinclude-default-mode=error `
  --output-filename="chimera-toolbox.exe" `
  --include-data-file=.\src\app\obtoolbox.tcss=obtoolbox.tcss `
  --include-data-file=.\src\backend\data\passwords.xml=passwords.xml `
  --windows-icon-from-ico=.\src\app\overbytes.ico `
  --include-package=rich._unicode_data.unicode17-0-0 `
  --noinclude-pytest-mode=nofollow `
  --noinclude-setuptools-mode=nofollow `
  --noinclude-unittest-mode=nofollow `
  --lto=yes `
  --clang `
  <#--remove-output #>`
  ".\src\app\textual_login.py"

# Locate upx.exe instead of assuming a hardcoded machine-specific path:
# prefer it on PATH, fall back to the conventional C:\upx install, and skip
# compression (with a warning) rather than failing the whole build if
# neither is found.
$upxCmd = Get-Command upx.exe -ErrorAction SilentlyContinue
if ($upxCmd) {
    $upxPath = $upxCmd.Source
} elseif (Test-Path "C:\upx\upx.exe") {
    $upxPath = "C:\upx\upx.exe"
} else {
    $upxPath = $null
}

if ($upxPath) {
    & $upxPath --best --lzma "textual_login.dist\chimera-toolbox.exe"
} else {
    Write-Warning "upx.exe not found (checked PATH and C:\upx\upx.exe) - skipping exe compression."
}

# Make a directory called backend in the same directory as the compiled exe, and copy the ob.psm1, dentalsoftware.psm1, and fun.psm1 modules into it
$exe_dir = ".\textual_login.dist"
$backend_dir = Join-Path $exe_dir "backend"
New-Item -ItemType Directory -Force -Path $backend_dir
Copy-Item -Path ".\src\backend\modules\ob.psm1" -Destination $backend_dir -Force
Copy-Item -Path ".\src\backend\modules\dentalsoftware.psm1" -Destination $backend_dir -Force
Copy-Item -Path ".\src\backend\modules\fun.psm1" -Destination $backend_dir -Force

# Take all files and directories in the dist folder and zip them into a file called chimera-toolbox.zip in the working directory
Compress-Archive -Path (Join-Path $exe_dir "*") -DestinationPath .\chimera-toolbox.zip -Force

# Clean up build artifacts now that everything is packaged into chimera-toolbox.zip.
# textual_login.build is Nuitka's intermediate C build (never needed after compiling),
# and textual_login.dist is fully captured inside chimera-toolbox.zip, so both are safe to remove.
Write-Host "Cleaning up build artifacts..."
$build_dir = ".\textual_login.build"
Remove-Item -Recurse -Force -Path $build_dir -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force -Path $exe_dir -ErrorAction SilentlyContinue
Write-Host "Done. Output: chimera-toolbox.zip"