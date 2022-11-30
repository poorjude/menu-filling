# constants
$XL_ALIGNMENT_CENTER = -4108
$WEEKDAYS = @("Понедельник", "Вторник", "Среда", "Четверг", "Пятница")
$MENU_NAMES = @("Правильное питание", "Сытное меню")

try {
$excel = New-Object -comobject Excel.Application
$excel.displayalerts = $false
$excel.screenUpdating = $false
# opening an existing "raw menu" file
$rawWorkbook = $excel.workbooks.open("$pwd\raw.xlsx")
# the upper line may make $excel.visible == $true
$excel.visible = $false
$raw = $rawWorkbook.worksheets.item(1)
# creating new "menu sample"
$sampleWorkbook = $excel.workbooks.add()
$sampleWorkbook.worksheets.item(3).delete()
$sampleWorkbook.worksheets.item(2).delete()
$sample = $sampleWorkbook.worksheets.item(1)
$sample.name = $raw.name

# finding two submenu in $raw to define their columns
$submenuCol = @(-1, -1)
$submenuCounter = 0
for ($row = 1; $row -lt 500; $row++) {
    for ($col = 1; $col -lt 100; $col++) {
        $currCellText = $raw.cells.item($row, $col).text;
        if ($currCellText -like ("*" + $MENU_NAMES[0] + "*")) {
            $submenuCol[0] = $col
            $submenuCounter++
        }
        elseif ($currCellText -like ("*" + $MENU_NAMES[1] + "*")) {
            $submenuCol[1] = $col
            $submenuCounter++
        }
        if ($submenuCounter -eq 2) { break }
    }
    if ($submenuCounter -eq 2) { break }
}

# extracting data from $raw and filling in $sample with it
$sampleRow = 1
for ($rawRow = 1; $rawRow -lt 500; $rawRow++) {
    $currCellText = $raw.cells.item($rawRow, $submenuCol[0]).text;
    # filtering empty cells and non-weekdays
    if ( $currCellText -eq "" -or $currCellText -ne ($WEEKDAYS -eq $currCellText) ) {
        continue
    }
    # entering weekday-name into $sample + formatting the cell
    $sample.cells.item($sampleRow, 1) = $WEEKDAYS -eq $currCellText
    $sample.cells.item($sampleRow, 1).font.bold = $true
    $sample.cells.item($sampleRow, 1).horizontalAlignment = $XL_ALIGNMENT_CENTER
    $mergeCells = $sample.range($sample.cells($sampleRow, 1), $sample.cells($sampleRow, 2)) 
    $mergeCells.mergeCells = $true

    for ($submenu_i = 0; $submenu_i -lt 2; $submenu_i++) {
        # entering menu name into $sample + formatting the cell
        $sample.cells($sampleRow + 1, 1 + $submenu_i) = $MENU_NAMES[$submenu_i]
        $sample.cells($sampleRow + 1, 1 + $submenu_i).font.bold = $true
        $sample.cells($sampleRow + 1, 1 + $submenu_i).horizontalAlignment = $XL_ALIGNMENT_CENTER

        $sample_i = 2;
        for ($raw_i = 2; $raw_i -lt 9; $raw_i++) {
            $dishCellText = $raw.cells.item($rawRow + $raw_i, $submenuCol[$submenu_i]).text
            # filtering empty cells
            if ($dishCellText -eq "") { continue }
            # entering dishes with their prices into $sample
            $dishPrice = $raw.cells.item($rawRow + $raw_i, $submenuCol[$submenu_i] + 2).text
            $sample.cells.item($sampleRow + $sample_i, 1 + $submenu_i) = $dishCellText + " — " + $dishPrice # (long dash is used here)
            $sample_i++
        }
    }
    # moving $sampleRow + moving it lower if the cells above contain any text
    $sampleRow += 7
    if ($sample.cells.item($sampleRow - 1, 1).text -ne "" -or $sample.cells.item($sampleRow - 1, 2).text -ne "") {
        $sampleRow++
    }
    # moving $rawRow
    $rawRow += 5
}

# formatting $sample table
$sample.usedRange.columns.font.size = 16
$sample.usedRange.columns.autofit()
$excel.activeWindow.zoom = 73

$sampleWorkbook.saveAs("$pwd\sample\sample.xlsx")
$sampleWorkbook.close()
$rawWorkbook.close()
$excel.screenUpdating = $true
$excel.displayalerts = $true
$excel.quit()

# making many copies of $sample
$namesList = (Get-Content -Path ".\list.txt")
foreach ($name in $namesList) {
    Copy-Item -Path ".\sample\sample.xlsx" -Destination "..\$name.xlsx"
}
} catch {
    $errorLogsPath = ".\ERROR_LOGS.txt"
    Add-Content -Path $errorLogsPath -Value (Get-Date)
    Add-Content -Path $errorLogsPath -Value "An error in the 1st script occurred:"
    Add-Content -Path $errorLogsPath -Value $_
    Add-Content -Path $errorLogsPath -Value $_.ScriptStackTrace
    Add-Content -Path $errorLogsPath -Value ""
    Write-Host "An error in the 1st script occurred:"
    Write-Host $_
    Write-Host $_.ScriptStackTrace
}