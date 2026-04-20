param ([string]$FilePath)
$lines = Get-Content $FilePath
$sections = @()
$currentSection = "HEADER"

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i].Trim()
    if ($line.StartsWith("'") -and $line.Contains("=====")) {
        if ($i + 1 -lt $lines.Count -and -not $lines[$i+1].Trim().StartsWith("'=====") -and $lines[$i+1].Trim().StartsWith("'")) {
            $secName = $lines[$i+1].Trim().Substring(1).Trim()
            if ($secName -match "^[A-Z0-9 _-]+$" -or $secName -match "^[A-Z].*$") {
                $currentSection = $secName
                $sections += $currentSection
            }
        }
    }
}
Write-Host "Sections found in order:"
$sections | Out-Host
