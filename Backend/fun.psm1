Function Invoke-Battleship {
    param(
    [int]$Size = 10,
    [switch]$Color
)

# =============================
# Battleship (PowerShell)
# - Single player vs CPU
# - Random placement
# - Smart-ish CPU targeting after a hit
# =============================

# ----------- Helpers -----------

function Assert-BoardSize {
    param([int]$n)
    if ($n -lt 5 -or $n -gt 26) {
        throw "Board size must be between 5 and 26."
    }
}

function Get-FleetDefinition {
    @(
        @{ Name='Carrier'    ; Size=5 }
        @{ Name='Battleship' ; Size=4 }
        @{ Name='Cruiser'    ; Size=3 }
        @{ Name='Submarine'  ; Size=3 }
        @{ Name='Destroyer'  ; Size=2 }
    )
}

function New-EmptyShotsMap { @{} }

function New-Queue { New-Object System.Collections.Queue }

function New-Random {
    if (-not $script:Rng) { $script:Rng = [System.Random]::new() }
    $script:Rng
}

function In-Bounds {
    param([int]$x, [int]$y, [int]$size)
    return ($x -ge 0 -and $x -lt $size -and $y -ge 0 -and $y -lt $size)
}

function Get-CoordKey {
    param([int]$x, [int]$y)
    "$x,$y"
}

function Parse-Coord {
    param(
        [string]$Text,
        [int]$Size
    )
    # Accept forms like: A5, a5, J10, 10J (letter + number in any order)
    $t = ($Text -replace '\s','').ToUpper()
    if (-not $t) { return $null }

    $letters = ($t -replace '[^A-Z]','')
    $digits  = ($t -replace '[^0-9]','')

    if (-not $letters -or -not $digits) { return $null }
    if ($letters.Length -ne 1) { return $null }
    $colLetter = $letters[0]
    $rowNum = 0
    if (-not [int]::TryParse($digits, [ref]$rowNum)) { return $null }

    $x = ([byte][char]$colLetter) - ([int][char]'A')  # 0-based column
    $y = $rowNum - 1                             # 0-based row

    if (-not (In-Bounds -x $x -y $y -size $Size)) { return $null }

    [pscustomobject]@{ X = $x; Y = $y }
}

function Format-Coord {
    param([int]$x, [int]$y)
    $colChar = [char]([int][char]'A' + $x)
    $row = $y + 1
    "$colChar$row"
}

# ----------- Fleet Placement -----------

function Test-PlacementClear {
    param(
        [hashtable]$Taken,
        [int]$size,
        [int]$x,
        [int]$y,
        [int]$length,
        [char]$orientation # 'H' or 'V'
    )
    for ($i=0; $i -lt $length; $i++) {
        $cx = $x
        $cy = $y
        if ($orientation -eq 'H') { $cx += $i } else { $cy += $i }
        if (-not (In-Bounds -x $cx -y $cy -size $size)) { return $false }
        $key = Get-CoordKey -x $cx -y $cy
        if ($Taken.ContainsKey($key)) { return $false }
    }
    return $true
}

function Place-Ship {
    param(
        [hashtable]$Taken,
        [int]$size,
        [string]$name,
        [int]$length
    )
    $rng = New-Random
    for ($attempt=0; $attempt -lt 5000; $attempt++) {
        $orientation = if ($rng.Next(2) -eq 0) {'H'} else {'V'}
        $x = $rng.Next(0, $size)
        $y = $rng.Next(0, $size)

        if (Test-PlacementClear -Taken $Taken -size $size -x $x -y $y -length $length -orientation $orientation) {
            $cells = New-Object System.Collections.Generic.List[object]
            for ($i=0; $i -lt $length; $i++) {
                $cx = $x
                $cy = $y
                if ($orientation -eq 'H') { $cx += $i } else { $cy += $i }
                $key = Get-CoordKey -x $cx -y $cy
                $Taken[$key] = $true
                $cells.Add([pscustomobject]@{ X=$cx; Y=$cy; Key=$key })
            }
            return [pscustomobject]@{
                Name = $name
                Size = $length
                Cells = $cells
            }
        }
    }
    throw "Failed to place ship: $name"
}

function Place-FleetRandom {
    param([int]$Size)
    $taken = @{}
    $fleetDef = Get-FleetDefinition
    $fleet = New-Object System.Collections.Generic.List[object]
    foreach ($s in $fleetDef) {
        $fleet.Add( (Place-Ship -Taken $taken -size $Size -name $s.Name -length $s.Size) )
    }
    $occupied = @{}
    foreach ($ship in $fleet) {
        foreach ($c in $ship.Cells) { $occupied[$c.Key] = $ship }
    }
    [pscustomobject]@{
        Fleet    = $fleet
        Occupied = $occupied  # "x,y" -> ship object
    }
}

# ----------- Game Mechanics -----------

function Test-ShipSunk {
    param(
        $Ship,
        [hashtable]$Shots
    )
    foreach ($c in $Ship.Cells) {
        if (-not $Shots.ContainsKey($c.Key)) { return $false }
    }
    return $true
}

function Test-AllSunk {
    param($Fleet, [hashtable]$Shots)
    foreach ($ship in $Fleet) {
        if (-not (Test-ShipSunk -Ship $ship -Shots $Shots)) { return $false }
    }
    return $true
}

# Using Invoke-Shot as the shot executor
function Invoke-Shot {
    param(
        [hashtable]$Shots,
        [hashtable]$OppOccupied,
        $OppFleet,
        [int]$x,
        [int]$y
    )
    $key = Get-CoordKey -x $x -y $y
    if ($Shots.ContainsKey($key)) {
        return [pscustomobject]@{ Result='Repeat'; Ship=$null; X=$x; Y=$y; Key=$key }
    }
    $Shots[$key] = $true
    if ($OppOccupied.ContainsKey($key)) {
        $ship = $OppOccupied[$key]
        $sunk = Test-ShipSunk -Ship $ship -Shots $Shots
        if ($sunk) {
            if (Test-AllSunk -Fleet $OppFleet -Shots $Shots) {
                return [pscustomobject]@{ Result='Win'; Ship=$ship; X=$x; Y=$y; Key=$key }
            } else {
                return [pscustomobject]@{ Result='Sunk'; Ship=$ship; X=$x; Y=$y; Key=$key }
            }
        } else {
            return [pscustomobject]@{ Result='Hit'; Ship=$ship; X=$x; Y=$y; Key=$key }
        }
    } else {
        return [pscustomobject]@{ Result='Miss'; Ship=$null; X=$x; Y=$y; Key=$key }
    }
}

# ----------- Rendering -----------

function Write-Cell {
    param(
        [string]$text,
        [ConsoleColor]$fg = [ConsoleColor]::Gray,
        [switch]$Color
    )
    if ($Color) {
        Write-Host $text -NoNewline -ForegroundColor $fg
    } else {
        Write-Host $text -NoNewline
    }
}

function Show-Board {
    param(
        [string]$Title,
        [int]$Size,
        $Fleet,
        [hashtable]$Shots,
        [switch]$Reveal,
        [switch]$Color
    )
    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan

    # Column headers
    Write-Host "    " -NoNewline
    for ($x=0; $x -lt $Size; $x++) {
        $col = [char]([int][char]'A' + $x)
        Write-Host ("{0} " -f $col) -NoNewline
    }
    Write-Host ""

    # Quick lookup for occupied cells
    $occupied = @{}
    foreach ($ship in $Fleet) {
        foreach ($c in $ship.Cells) { $occupied[$c.Key] = $true }
    }

    for ($y=0; $y -lt $Size; $y++) {
        $rowLabel = "{0,2}" -f ($y+1)
        Write-Host ("{0} | " -f $rowLabel) -NoNewline
        for ($x=0; $x -lt $Size; $x++) {
            $key = Get-CoordKey -x $x -y $y
            $hasShot = $Shots.ContainsKey($key)
            $hasShip = $occupied.ContainsKey($key)

            if ($Reveal) {
                if ($hasShot -and $hasShip) {
                    Write-Cell -text "X " -fg Red -Color:$Color
                } elseif ($hasShot -and -not $hasShip) {
                    Write-Cell -text "o " -fg DarkGray -Color:$Color
                } elseif ($hasShip) {
                    Write-Cell -text "S " -fg Green -Color:$Color
                } else {
                    Write-Cell -text ". " -fg DarkGray -Color:$Color
                }
            } else {
                if ($hasShot -and $hasShip) {
                    Write-Cell -text "X " -fg Red -Color:$Color
                } elseif ($hasShot) {
                    Write-Cell -text "o " -fg DarkGray -Color:$Color
                } else {
                    Write-Cell -text ". " -fg DarkGray -Color:$Color
                }
            }
        }
        Write-Host ""
    }
}

# ----------- CPU AI -----------

function Get-AdjacentCells {
    param([int]$x,[int]$y,[int]$size)
    $dirs = @(
        @{dx= 1; dy= 0},
        @{dx=-1; dy= 0},
        @{dx= 0; dy= 1},
        @{dx= 0; dy=-1}
    )
    $res = New-Object System.Collections.Generic.List[object]
    foreach ($d in $dirs) {
        $nx = $x + $d.dx
        $ny = $y + $d.dy
        if (In-Bounds -x $nx -y $ny -size $size) {
            $res.Add([pscustomobject]@{X=$nx; Y=$ny})
        }
    }
    $res
}

function CPU-ChooseShot {
    param(
        [int]$Size,
        [hashtable]$Shots,  # CPU's shots taken at the player
        [System.Collections.Queue]$Targets  # prioritized neighbors
    )
    # 1) Use queued targets first
    while ($Targets.Count -gt 0) {
        $cell = $Targets.Dequeue()
        $key = Get-CoordKey -x $cell.X -y $cell.Y
        if (-not $Shots.ContainsKey($key)) {
            return $cell
        }
    }
    # 2) Otherwise pick a random unseen cell
    $rng = New-Random
    for ($attempt=0; $attempt -lt 5000; $attempt++) {
        $x = $rng.Next(0,$Size)
        $y = $rng.Next(0,$Size)
        $key = Get-CoordKey -x $x -y $y
        if (-not $Shots.ContainsKey($key)) {
            return [pscustomobject]@{X=$x; Y=$y}
        }
    }
    # Fallback (should not happen)
    for ($yy=0; $yy -lt $Size; $yy++) {
        for ($xx=0; $xx -lt $Size; $xx++) {
            $key = Get-CoordKey -x $xx -y $yy
            if (-not $Shots.ContainsKey($key)) {
                return [pscustomobject]@{X=$xx; Y=$yy}
            }
        }
    }
    return $null
}

# ----------- Input / Turn -----------

function Read-PlayerShot {
    param([int]$Size, [hashtable]$Shots)
    while ($true) {
        $input = Read-Host "Enter target (e.g., A5)"
        if (-not $input) { continue }
        $coord = Parse-Coord -Text $input -Size $Size
        if (-not $coord) {
            Write-Host "Invalid coordinate. Use something like A5 or J10." -ForegroundColor Yellow
            continue
        }
        $key = Get-CoordKey -x $coord.X -y $coord.Y
        if ($Shots.ContainsKey($key)) {
            Write-Host ("You already fired at {0}." -f (Format-Coord $coord.X $coord.Y)) -ForegroundColor Yellow
            continue
        }
        return $coord
    }
}

# ----------- Main Game -----------

try {
    Assert-BoardSize -n $Size

    Clear-Host
    Write-Host "=== Battleship (PowerShell) ===" -ForegroundColor Cyan
    Write-Host "Board: ${Size}x$Size"
    Write-Host "Fleet: Carrier(5), Battleship(4), Cruiser(3), Submarine(3), Destroyer(2)"
    if ($Color) { Write-Host "Color mode: ON" -ForegroundColor Green }

    # Place Fleets
    $playerPlacement = Place-FleetRandom -Size $Size
    $cpuPlacement    = Place-FleetRandom -Size $Size

    $playerFleet     = $playerPlacement.Fleet
    $cpuFleet        = $cpuPlacement.Fleet
    $cpuOccupied     = $cpuPlacement.Occupied
    $playerOccupied  = @{}
    foreach ($ship in $playerFleet) { foreach ($c in $ship.Cells) { $playerOccupied[$c.Key] = $ship } }

    # Shot maps
    $playerShots   = New-EmptyShotsMap   # shots at CPU
    $cpuShots      = New-EmptyShotsMap   # shots at Player
    $cpuTargets    = New-Queue           # target neighbors after hit

    # Stats
    $turn = 1

    while ($true) {
        # Render boards
        Show-Board -Title "Your Board (ships revealed)" -Size $Size -Fleet $playerFleet -Shots $cpuShots -Reveal -Color:$Color
        Show-Board -Title "Enemy Waters (fog of war)"   -Size $Size -Fleet $cpuFleet    -Shots $playerShots -Color:$Color

        # ---- Player Turn ----
        Write-Host ""
        Write-Host ("Turn {0} - Your shot" -f $turn) -ForegroundColor Cyan
        $p = Read-PlayerShot -Size $Size -Shots $playerShots
        $result = Invoke-Shot -Shots $playerShots -OppOccupied $cpuOccupied -OppFleet $cpuFleet -x $p.X -y $p.Y

        switch ($result.Result) {
            'Repeat' { Write-Host "You already fired there. (This shouldn't show due to input guard.)" -ForegroundColor Yellow }
            'Miss'   {
                Write-Host ("You fired at {0}: MISS." -f (Format-Coord $result.X $result.Y)) -ForegroundColor DarkGray
            }
            'Hit'    {
                Write-Host ("You fired at {0}: HIT!" -f (Format-Coord $result.X $result.Y)) -ForegroundColor Red
            }
            'Sunk'   {
                Write-Host ("You sunk the enemy {0} at {1}!" -f $result.Ship.Name, (Format-Coord $result.X $result.Y)) -ForegroundColor Green
            }
            'Win'    {
                Write-Host ("You sunk the enemy {0}!" -f $result.Ship.Name) -ForegroundColor Green
                Write-Host "ALL ENEMY SHIPS SUNK. YOU WIN!" -ForegroundColor Cyan
                break
            }
        }

        if ($result.Result -eq 'Win') { break }

        # ---- CPU Turn ----
        Start-Sleep -Milliseconds 500
        Write-Host ""
        Write-Host ("Turn {0} - Enemy fires..." -f $turn) -ForegroundColor Cyan

        $cpuShot = CPU-ChooseShot -Size $Size -Shots $cpuShots -Targets $cpuTargets
        $cRes = Invoke-Shot -Shots $cpuShots -OppOccupied $playerOccupied -OppFleet $playerFleet -x $cpuShot.X -y $cpuShot.Y
        $coordStr = Format-Coord -x $cpuShot.X -y $cpuShot.Y

        switch ($cRes.Result) {
            'Repeat' {
                # Shouldn't happen due to chooser; skip and try again quickly
                continue
            }
            'Miss' {
                Write-Host ("Enemy fires at {0}: MISS." -f $coordStr) -ForegroundColor DarkGray
            }
            'Hit' {
                Write-Host ("Enemy fires at {0}: HIT!" -f $coordStr) -ForegroundColor Red
                foreach ($n in (Get-AdjacentCells -x $cpuShot.X -y $cpuShot.Y -size $Size)) {
                    $cpuTargets.Enqueue($n)
                }
            }
            'Sunk' {
                Write-Host ("Enemy fires at {0}: Your {1} is SUNK!" -f $coordStr, $cRes.Ship.Name) -ForegroundColor Yellow
            }
            'Win' {
                Write-Host ("Enemy fires at {0}: Your {1} is SUNK!" -f $coordStr, $cRes.Ship.Name) -ForegroundColor Yellow
                Write-Host "ALL YOUR SHIPS ARE SUNK. CPU WINS." -ForegroundColor Cyan
                break
            }
        }

        if ($cRes.Result -eq 'Win') { break }

        $turn++
        Start-Sleep -Milliseconds 300
        Clear-Host
    }

    # Final boards
    Show-Board -Title "FINAL - Your Board" -Size $Size -Fleet $playerFleet -Shots $cpuShots -Reveal -Color:$Color
    Show-Board -Title "FINAL - Enemy Board (revealed)" -Size $Size -Fleet $cpuFleet -Shots $playerShots -Reveal -Color:$Color

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
}

Function Get-LibreMines {
    Start-BitsTransfer -Source "https://github.com/Bollos00/LibreMines/releases/download/v2.3.0/libremines-v2.3.0-windows-qt6.zip" -Destination "$env:TEMP\obsoftware\LibreMines.zip"
    Expand-Archive -Path "$env:TEMP\obsoftware\LibreMines.zip" -DestinationPath "$env:TEMP\obsoftware\LibreMines" -Force
    Start-Process -FilePath "$env:TEMP\obsoftware\LibreMines\libremines\libremines.exe"
}