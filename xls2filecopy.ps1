# Einfach machen, nicht fragen... wenn man diese Variable auf "nuke" setzt
$nuke = "better not touch"

# Definieren, welche XLS-Datei, wo die Originaldateien liegen und wo die Kopien hin sollen
$ExcelFile = "C:\Users\Lenovo\Downloads\Thomas.xlsx"
$FilePathSource = "C:\Users\Lenovo\Downloads\"
$FilePathDestination = "C:\Users\Lenovo\Downloads\new\"

# Welche Spalte im XLS beinhaltet den bisherigen Dateinamen und welche den Text
$ROW_ID = 1
$ROW_DESCRIPTION = 2
# XLS-File oeffnen
$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open($ExcelFile)
$workSheet = $Workbook.Sheets.Item(1)
# Die erste Zeile auslassen wegen Ueberschrift
$RowNum = 2
While ($workSheet.Cells.Item($RowNum, 1).Text -ne "") {
    $ORIGINAL_SOURCE = $FilePathSource + $workSheet.Cells.Item($RowNum, $ROW_ID).Text

    # Dateiendung loswerden fuer den zukuenftigen Namen
    $ID =   $workSheet.Cells.Item($RowNum, 1).Text.REPLACE(".pdf", "")
    # Leerzeichen und anderen Kram loswerden
    $NEW_DESTINATION_END1 = $ID + "_" + $workSheet.Cells.Item($RowNum, $ROW_DESCRIPTION).Text -replace '[\W]', '_'
    # Dateiendung wieder dran
    $NEW_DESTINATION_END2 = $NEW_DESTINATION_END1 + ".pdf"
    # Kompletten Pfad ergänzen
    $NEW_DESTINATION = $FilePathDestination + $NEW_DESTINATION_END2
    
    # CMD Version ausgeben / das hier wird passieren
    "COPY " + $ORIGINAL_SOURCE + " " + $DESTINATION

    $nuke_em = ""
    if ( $nuke -ne "nuke" ) {
        # Abfragen, ob Du das wirklich kopieren willst
        $reallycopy = Read-Host -Prompt 'Shall I copy this (y/n)?'
        if ( $reallycopy -eq "y" ) {
            $nuke_em = "nuke"
        } else { 
            $nuke_em = "Oh no!"
        }
    } else {
        $nuke_em = "nuke"
    }

    if ( $nuke_em -eq "nuke" ) {
        # Kopieren
        Copy-Item $ORIGINAL_SOURCE -Destination $DESTINATION
        # Ausgabe "Copied"
        "Copied"
    } else {
        # Ausgabe "Not copied" bei allem ausser "y"
        "Not copied"
    }
    # Naechste Zeile
    $RowNum++
}

