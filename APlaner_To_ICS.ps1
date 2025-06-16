clear

# ─── KONFIGURATION ───────────────────────────────────────────────
$config = @{
    jahr = 2025
    repoPath = "C:\Development\aplaner-to-ics"        # Lokales Git-Repo
    icsDateiname = "docs\calendar.ics"    # Name der .ics-Datei
    csvDatei = "export.csv"       # CSV im selben Ordner wie Skript
}

# ─── VERZEICHNIS WECHSELN ────────────────────────────────────────
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptDir

$csvPath = Join-Path $scriptDir $config.csvDatei
$icsPath = Join-Path $scriptDir $config.icsDateiname
$repoIcsPath = Join-Path $config.repoPath $config.icsDateiname

# ─── ICS GENERIEREN ──────────────────────────────────────────────
$kuerzelMap = @{
    "A"="Büro"; "O"="HomeOffice"; "U"="Urlaub"
    "P"="Urlaub geplant"; "H"="½ Urlaub geplant"; "½"="½ Urlaub"
    "S"="Sonderurlaub"; "D"="Dienstreise"; "Z"="Zusatzarbeitstag"
}
$monate = @{
    "Januar"=1; "Februar"=2; "März"=3; "April"=4; "Mai"=5; "Juni"=6
    "Juli"=7; "August"=8; "September"=9; "Oktober"=10; "November"=11; "Dezember"=12
}
$zeilen = Get-Content $csvPath -Encoding Default
$daten = $zeilen | Select-Object -Skip 1
$ics = @(
    "BEGIN:VCALENDAR"
    "VERSION:2.0"
    "PRODID:-//Kalender Export//EN"
    "CALSCALE:GREGORIAN"
)

foreach ($zeile in $daten) {
Write-Host "`nAnalysiere Zeile: $zeile"
Write-Host "→ Erster Eintrag (Monat?): '$($teile[0])'"
    $teile = $zeile -split ";"
    $monat = $teile[0]
    if (-not $monate.ContainsKey($monat)) { continue }
    $monatNum = $monate[$monat]

    for ($i = 1; $i -lt $teile.Length; $i++) {
        $inhalt = $teile[$i].Trim()
        if ($inhalt -eq "") { continue }

        try {
            $datum = Get-Date -Year $config.jahr -Month $monatNum -Day $i -ErrorAction Stop
        } catch {
            continue
        }

        $summary = if ($kuerzelMap.ContainsKey($inhalt)) {
            "D - $($kuerzelMap[$inhalt])"
        } else {
            "D - Unbekannt ($inhalt)"
        }

        $ics += @(
            "BEGIN:VEVENT"
            "DTSTART;VALUE=DATE=$($datum.ToString('yyyyMMdd'))"
            "DTEND;VALUE=DATE=$($datum.AddDays(1).ToString('yyyyMMdd'))"
            "SUMMARY:$summary"
            "TRANSP:OPAQUE"
            "END:VEVENT"
        )
    }
}

$ics += "END:VCALENDAR"
$ics -join "`r`n" | Set-Content -Path $icsPath -Encoding UTF8

# ─── GIT PUSH ────────────────────────────────────────────────────
try {
    Set-Location $config.repoPath
    git add -A
    git commit -m "Automatisches Update: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    git push

    # ─── Löschung bei Erfolg ─────────────────────────────────────
    Remove-Item -Path $icsPath -Force
    Write-Host "`n✅ Erfolgreich gepusht und lokale Datei gelöscht."
} catch {
    Write-Error "❌ Fehler beim Push oder Dateizugriff: $_"
    exit 1
}
