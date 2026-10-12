# Compile (PUBLIC build - no login, no authentication)
#
# This produces the build that is safe to publish to the public GitHub
# "lib" releases folder that chimera-toolbox.ps1 downloads from (anyone who
# runs the public "irm | iex" one-liner gets this exact .zip). It compiles
# obtoolbox_tui.py directly - skipping textual_login.py and its Login
# screen entirely - and does NOT bundle passwords.xml, since there is no
# login step here to check credentials against and shipping password
# hashes in a public download serves no purpose.
#
# For the password-gated build (which must stay OUT of the public lib
# repo), use "compile-private-chimera-toolbox.ps1" instead.

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
  --windows-icon-from-ico=.\src\app\overbytes.ico `
  --include-package=rich._unicode_data.unicode17-0-0 `
  --noinclude-pytest-mode=nofollow `
  --noinclude-setuptools-mode=nofollow `
  --noinclude-unittest-mode=nofollow `
  --lto=yes `
  --clang `
  <#--remove-output #>`
  ".\src\app\obtoolbox_tui.py"

# Locate upx.exe: prefer it on PATH, then fall back to a copy under this
# repo's own build\ folder (e.g. build\upx-5.2.1-win64\upx.exe) instead of
# a machine-specific C:\upx install - any upx.exe found anywhere under
# build\ is used, so dropping in a newer UPX release there keeps working
# without editing this script. Skip compression (with a warning) rather
# than failing the whole build if neither is found.
$upxCmd = Get-Command upx.exe -ErrorAction SilentlyContinue
if ($upxCmd) {
    $upxPath = $upxCmd.Source
} else {
    $upxPath = Get-ChildItem -Path "$repoRoot\build" -Filter "upx.exe" -Recurse -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
}

if ($upxPath) {
    & $upxPath --best --lzma "obtoolbox_tui.dist\chimera-toolbox.exe"
} else {
    Write-Warning "upx.exe not found (checked PATH and $repoRoot\build) - skipping exe compression."
}

# Make a directory called backend in the same directory as the compiled exe, and copy the ob.psm1, dentalsoftware.psm1, and fun.psm1 modules into it
$exe_dir = ".\obtoolbox_tui.dist"
$backend_dir = Join-Path $exe_dir "backend"
New-Item -ItemType Directory -Force -Path $backend_dir
Copy-Item -Path ".\src\backend\modules\ob.psm1" -Destination $backend_dir -Force
Copy-Item -Path ".\src\backend\modules\dentalsoftware.psm1" -Destination $backend_dir -Force
Copy-Item -Path ".\src\backend\modules\fun.psm1" -Destination $backend_dir -Force

# Take all files and directories in the dist folder and zip them into a file called chimera-toolbox.zip in the working directory.
# This filename matches exactly what chimera-toolbox.ps1 downloads from the
# public GitHub release, so this .zip is the one meant to be uploaded there.
Compress-Archive -Path (Join-Path $exe_dir "*") -DestinationPath .\chimera-toolbox.zip -Force

# Clean up build artifacts now that everything is packaged into chimera-toolbox.zip.
# obtoolbox_tui.build is Nuitka's intermediate C build (never needed after compiling),
# and obtoolbox_tui.dist is fully captured inside chimera-toolbox.zip, so both are safe to remove.
Write-Host "Cleaning up build artifacts..."
$build_dir = ".\obtoolbox_tui.build"
Remove-Item -Recurse -Force -Path $build_dir -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force -Path $exe_dir -ErrorAction SilentlyContinue
Write-Host "Done. Output: chimera-toolbox.zip (no login, no passwords.xml - safe for the public lib repo)"
