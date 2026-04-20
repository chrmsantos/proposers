param ([string]$FilePath)
$lines = Get-Content $FilePath
$currentSection = "Unknown"
$sections = @{}

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i].Trim()
    if ($line.StartsWith("'") -and $line.Contains("====")) {
        if ($i + 1 -lt $lines.Count -and -not $lines[$i+1].Trim().StartsWith("'====") -and $lines[$i+1].Trim().StartsWith("'")) {
            $secName = $lines[$i+1].Trim().Substring(1).Trim()
            if ($secName -match "^[A-Z0-9 _-]+$" -and $secName.Length -gt 3) {
                $currentSection = $secName
                if (-not $sections.ContainsKey($currentSection)) {
                    $sections[$currentSection] = @()
                }
            }
        }
    } elseif ($line.StartsWith("Public ") -or $line.StartsWith("Private ")) {
        if ($line.Contains(" Sub ") -or $line.Contains(" Function ")) {
            if (-not $sections.ContainsKey($currentSection)) {
                $sections[$currentSection] = @()
            }
            $sections[$currentSection] += $line
        }
    }
}

foreach ($sec in $sections.Keys) {
    Write-Host "-- $sec ($($sections[$sec].Count) items)"
    foreach ($item in $sections[$sec]) {
        Write-Host "   $item"
    }
}
