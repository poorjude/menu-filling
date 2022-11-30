# constants
$XL_LINESTYLE_NONE = -4142
$XL_LINESTYLE_CONTINUOUS = 1
$XL_ALIGNMENT_CENTER = -4108
$XL_ALIGNMENT_LEFT = -4131
$XL_COLORINDEX_YELLOW = 6
$XL_COLORINDEX_GREEN_BRIGHT = 14
$XL_COLORINDEX_GREEN_PALE = 43
$XL_COLORINDEX_PURPLE = 47
$FROM_PS_TO_XL_COLORINDEX_PURPLE_PALE = 39
$FROM_PS_TO_XL_COLORINDEX_YELLOW = 27
$WEEKDAYS = @("Понедельник", "Вторник", "Среда", "Четверг", "Пятница")
$COMBINED_ANSWER_TEXT = "Объединенный заказ"
$PARTIAL_ANSWER_TEXT = "Неполный заказ"

try {
$excel = New-Object -comobject Excel.Application
$excel.displayAlerts = $false
$excel.screenUpdating = $false
#$excel.visible = $false

# getting menu "skeleton" from .json
# it has the next structure: 5 days -> 2 submenu in each of them -> 5 dishes in each of them
$menu = (Get-Content -Raw ".\json\menu.json"| ConvertFrom-Json)
# opening an existing "sample" file that was copied to many people
$sampleWorkbook = $excel.workbooks.open("$pwd\sample\sample.xlsx")
$sample = $sampleWorkbook.sheets.item(1)



# 1. filling in $menu with data from $sample

$dayRow = 1
$dayCounter = 0
# loop for days
while ($dayRow -lt 200) {
    $dayCellText = $sample.cells.item($dayRow, 1).text
    # filtering empty cells and non-weekdays
    if ( $dayCellText -eq "" -or $dayCellText -ne ($WEEKDAYS -eq $dayCellText) ) {
        $dayRow++
        continue
    }
    $day = $menu.day[$dayCounter]
    $day.name = $dayCellText
    $day.resultList = $dayCounter + 1

    # loop for submenu
    $dishResultCol = 2
    for($submenuCounter = 0; $submenuCounter -lt 2; $submenuCounter++) {
        $submenu = $day.submenu[$submenuCounter]
        $submenu.name = $sample.cells.item($dayRow + 1, 1 + $submenuCounter).text

        # loop for dishes
        $dishCol = 1 + $submenuCounter
        for ($dishCounter = 0; $dishCounter -lt 5; $dishCounter++) {
            $dish = $submenu.dish[$dishCounter]
            $dishRow = $dayRow + 2 + $dishCounter
            $dishCellText = $sample.cells.item($dishRow, $dishCol).text
            # filtering empty cells (in case of having 4 dishes instead of 5)
            if ($dishCellText -eq "") { 
                break 
            }
            $dish.name = $dishCellText.split("—")[0] # long dash is used here
            $dish.price = $dishCellText.split("—")[1].split(",")[0] # long dash is used here
            $dish.sampleRow = $dishRow
            $dish.sampleCol = $dishCol
            $dish.resultCol = $dishResultCol
            $dishResultCol++
        }
        $dishResultCol++ # needed for separation of two submenu
    }
    $dayRow += 6
    $dayCounter++
}
$sampleWorkbook.close()



# 2. creating the "result" table

$resultWorkbook = $excel.workbooks.add()

# creating right amount of worksheets
$worksheetCounter = 1
for ($i = 0; $i -lt 5; $i++) {
    if ($menu.day[$i] -ne "") {
        $worksheetCounter++;
    }
}
while ($resultWorkbook.sheets.count -lt $worksheetCounter) {
    $tmp = $resultWorkbook.sheets.add()
}

# creating table headers
for ($dayCounter = 0; $dayCounter -lt 5; $dayCounter++) {
    $day = $menu.day[$dayCounter]
    # breaking the loop if only non-filled days left
    if ($day.name -eq "") {
        break;
    }
    $currSheet = $resultWorkbook.sheets.item($day.resultList)
    $currSheet.name = $day.name

    # filling in 1st row with day.name
    $currSheet.cells.item(1, 1) = $day.name

    # filling in 2nd row with two submenu.name
    $dayLength = 4 # needed for merging cells
    for ($submenuCounter = 0; $submenuCounter -lt 2; $submenuCounter++) {
        $submenu = $day.submenu[$submenuCounter]
        $submenuCol = $submenu.dish[0].resultCol
        $currSheet.cells.item(2, $submenuCol) = $submenu.name

        # filling in 3rd row with many dish.name
        $dishLength = -1 # needed for merging cells
        for ($dishCounter = 0; $dishCounter -lt 5; $dishCounter++) {
            $dish = $submenu.dish[$dishCounter]
            # breaking the loop of only non-filled dishes left
            if ($dish.name -eq "") {
                break;
            }
            $currSheet.cells.item(3, $dish.resultCol) = $dish.name
            $dishLength++
        }
        $dayLength += $dishLength
        # merging cells that will have submenu.name
        $mergeCells = $currSheet.range($currSheet.cells(2, $submenuCol), $currSheet.cells(2, $submenuCol + $dishLength)) 
        $mergeCells.mergeCells = $true
    }
    # merging cells that will contain day.name
    $mergeCells = $currSheet.range($currSheet.cells(1, 1), $currSheet.cells(1, $dayLength)) 
    $mergeCells.mergeCells = $true

    # formatting table headers
    for ($i = 1; $i -le 3; $i++) {
        $currSheet.rows.item($i).font.bold = $true
    }
    $currSheet.rows.item(3).wrapText = $true
    # formatting whole sheet
    $currSheet.rows.borders.lineStyle = $XL_LINESTYLE_CONTINUOUS
    $currSheet.usedRange.columns.columnWidth = 16
    $currSheet.rows.font.size = 12
    # formatting text alignment in columns
    $currSheet.rows.horizontalAlignment = $XL_ALIGNMENT_CENTER
    $currSheet.rows.verticalAlignment = $XL_ALIGNMENT_CENTER
    for($i = 0; $i -lt 2; $i++) {
        $col = $day.submenu[$i].dish[0].resultCol - 1
        $currSheet.columns($col).horizontalAlignment = $XL_ALIGNMENT_LEFT
    }
}
    



# 3. filling in each day-sheet in $resultWorkbook

$namesList = (Get-Content -Path ".\list.txt")

# opening every file with answers
$resultRow = 4
foreach ($name in $namesList) {
    $answerWorkbook = $excel.workbooks.open("$pwd\..\$name.xlsx")
    $answer = $answerWorkbook.worksheets.item(1)

    # choosing a day
    for ($dayCounter = 0; $dayCounter -lt 5; $dayCounter++) {
        $day = $menu.day[$dayCounter]
        # breaking the loop if only non-filled days left
        if ($day.name -eq "") {
            break;
        }
        $resultSheet = $resultWorkbook.sheets.item($day.resultList)

        # choosing a submenu
        for ($submenuCounter = 0; $submenuCounter -lt 2; $submenuCounter++) {
            $submenu = $day.submenu[$submenuCounter]
            # entering a name in $result
            $resultSheet.cells.item($resultRow, $submenu.dish[0].resultCol - 1) = $name
            
            # filling in dishes
            for ($dishCounter = 0; $dishCounter -lt 5; $dishCounter++) {
                $dish = $submenu.dish[$dishCounter]
                # breaking the loop of only non-filled dishes left
                if ($dish.name -eq "") {
                    break;
                }
                # extracting a cell from $answer and entering it in $result if needed
                $answerColor = $answer.cells.item($dish.sampleRow, $dish.sampleCol).interior.colorIndex
                if ($answerColor -eq $XL_COLORINDEX_YELLOW) {
                    $resultSheet.cells.item($resultRow, $dish.resultCol) = 1
                } 
                elseif ($answerColor -eq $XL_COLORINDEX_GREEN_BRIGHT -or $answerColor -eq $XL_COLORINDEX_GREEN_PALE) {
                    $resultSheet.cells.item($resultRow, $dish.resultCol) = 2
                }
                elseif ($answerColor -eq $XL_COLORINDEX_PURPLE) {
                    $resultSheet.cells.item($resultRow, $dish.resultCol) = 3
                }
            }
        }
    }
    $answerWorkbook.close()
    $resultRow++
}



# 4. checking whether there are unusual answers - "combined" or "partial" ones
$sheetsAmount = $resultWorkbook.sheets.count
for ($sheetCounter = 1; $sheetCounter -lt $sheetsAmount; $sheetCounter++) {
    $currSheet = $resultWorkbook.sheets.item($sheetCounter)
    $day = $menu.day[$sheetCounter - 1]
    if ($day.submenu[1].dish[4].name -eq "") {
        $lastCol = $day.submenu[1].dish[3].resultCol + 1
    } else {
        $lastCol = $day.submenu[1].dish[4].resultCol + 1
    }
    $rowsAmount = $currSheet.usedRange.rows.count

    # 4.1. checking whether there are "partial" answers - only one dish among №3/№4 ([2]/[3]) in any submenu
    for ($submenuCounter = 0; $submenuCounter -lt 2; $submenuCounter++) {
        $submenu = $day.submenu[$submenuCounter]
        # filtering submenu with four dishes
        if ($submenu.dish[4].name -eq "") {
            continue
        }
        for ($row = 4; $row -le $rowsAmount; $row++) {
            $dish2 = $currSheet.cells($row, $submenu.dish[2].resultCol).text
            $dish3 = $currSheet.cells($row, $submenu.dish[3].resultCol).text
            if ( ($dish2 -ne "") -and ($dish3 -eq "") ) {
                $fullText = $PARTIAL_ANSWER_TEXT + ": " + $submenu.dish[2].name
                $currSheet.cells($row, $lastCol) = $fullText
                $currSheet.cells($row, $lastCol).interior.colorIndex = $FROM_PS_TO_XL_COLORINDEX_PURPLE_PALE
            } 
            elseif ( ($dish2 -eq "") -and ($dish3 -ne "") ) {
                # adding another dish if needed
                if ($currSheet.cells($row, $lastCol).text -ne "") {
                    $fullText = $currSheet.cells($row, $lastCol).text + ", " + $submenu.dish[3].name
                } else {
                    $fullText = $PARTIAL_ANSWER_TEXT + ": " + $submenu.dish[3].name
                }
                $currSheet.cells($row, $lastCol) = $fullText
                $currSheet.cells($row, $lastCol).interior.colorIndex = $FROM_PS_TO_XL_COLORINDEX_PURPLE_PALE
            }
        }
    }

    # 4.2. checking whether there are "combined" answers - dishes №3/№4 ([2]/[3]) both from different submenu
    # this overwrites comments about "partial" answer

    # filtering days with at least one submenu having four dishes
    if ($day.submenu[0].dish[4].name -eq "" -or $day.submenu[1].dish[4].name -eq "") {
        continue
    }
    for ($row = 4; $row -le $rowsAmount; $row++) {
        $S0D2 = $currSheet.cells($row, $day.submenu[0].dish[2].resultCol).text
        $S0D3 = $currSheet.cells($row, $day.submenu[0].dish[3].resultCol).text
        $S1D2 = $currSheet.cells($row, $day.submenu[1].dish[2].resultCol).text
        $S1D3 = $currSheet.cells($row, $day.submenu[1].dish[3].resultCol).text

        # if 3rd dish from 1st submenu and 4th dish from 2nd submenu were both ordered
        if ( ($S0D2 -ne "") -and ($S0D3 -eq "") -and ($S1D2 -eq "") -and ($S1D3 -ne "") ) {
            $fullText = $COMBINED_ANSWER_TEXT + ": " + $day.submenu[0].dish[2].name + ", " +  $day.submenu[1].dish[3].name
            $currSheet.cells($row, $lastCol) = $fullText
            $currSheet.cells($row, $lastCol).interior.colorIndex = $FROM_PS_TO_XL_COLORINDEX_YELLOW
        }
        # if 4th dish from 1st submenu and 3rd dish from 2nd submenu were both ordered
        elseif ( ($S0D2 -eq "") -and ($S0D3 -ne "") -and ($S1D2 -ne "") -and ($S1D3 -eq "") ) {
            $fullText = $COMBINED_ANSWER_TEXT + ": " + $day.submenu[0].dish[3].name + ", " +  $day.submenu[1].dish[2].name
            $currSheet.cells($row, $lastCol) = $fullText
            $currSheet.cells($row, $lastCol).interior.colorIndex = $FROM_PS_TO_XL_COLORINDEX_YELLOW
        }
    }
}



# 5. filling in the result-sheet in $resultWorkbook
$sheetsAmount = $resultWorkbook.sheets.count
$resultSheet = $resultWorkbook.sheets.item($sheetsAmount)
$resultSheet.name = "Итого"
$currRow = 1
for ($sheetCounter = 1; $sheetCounter -lt $sheetsAmount; $sheetCounter++) {
    $currDaySheet = $resultWorkbook.sheets.item($sheetCounter)
    $day = $menu.day[$sheetCounter - 1]
    # copying table headers from the chosen day sheet
    $rangeSource = $currDaySheet.range($currDaySheet.cells(1, 1), $currDaySheet.cells(3, 14))
    $rangeSource.copy()
    # pasting them to the result sheet
    $rangeDest = $resultSheet.range($resultSheet.cells($currRow, 1), $resultSheet.cells($currRow + 2, 14))
    $resultSheet.paste($rangeDest)

    $currRow += 3

    # entering row titles and formatting them
    $resultSheet.cells($currRow, 1) = "Цена"
    $resultSheet.rows($currRow).numberFormat = "# ##0,00 ₽"
    $resultSheet.cells($currRow + 1, 1) = "Кол-во"
    $resultSheet.cells($currRow + 2, 1) = "Сумма"
    $resultSheet.rows($currRow + 2).numberFormat = "# ##0,00 ₽"
    # entering formulas counting results of the first dish (then they are copied and pasted for all other dishes)
    $resultSheet.cells.item($currRow + 1, 2).formulaLocal = "=СУММ($($currDaySheet.name)!B4:B500)"
    $amountSource = $resultSheet.cells($currRow + 1, 2)
    $resultSheet.cells.item($currRow + 2, 2).formulaLocal = "=B$($currRow)*B$($currRow+1)"
    $sumSource = $resultSheet.cells($currRow + 2, 2)
    for ($submenuCounter = 0; $submenuCounter -lt 2; $submenuCounter++) {
        $submenu = $day.submenu[$submenuCounter]
        for ($dishCounter = 0; $dishCounter -lt 5; $dishCounter++) {
            $dish = $submenu.dish[$dishCounter]
            # filtering empty dishes
            if ($dish.name -eq "") { break }
            $currCol = $dish.resultCol
            # 1) entering price
            $resultSheet.cells($currRow, $currCol) = $dish.price
            # 2) entering amount
            $amountSource.copy()
            $amountDest = $resultSheet.cells($currRow + 1, $currCol)
            $resultSheet.paste($amountDest)
            # 3) entering sum
            $sumSource.copy()
            $sumDest = $resultSheet.cells($currRow + 2, $currCol)
            $resultSheet.paste($sumDest)
        }
    }

    $currRow += 4
}
# formatting the result-sheet
$resultSheet.usedRange.rows.borders.lineStyle = $XL_LINESTYLE_NONE
$resultSheet.usedRange.columns.columnWidth = 16
$resultSheet.rows.horizontalAlignment = $XL_ALIGNMENT_CENTER
$resultSheet.rows.verticalAlignment = $XL_ALIGNMENT_CENTER
$resultSheet.rows.font.size = 12
$resultSheet.columns(1).font.bold = $true
$resultSheet.activate()
$excel.activeWindow.zoom = 70



# 6. formatting the day-sheets
for ($i = 1; $i -lt $resultWorkbook.sheets.count; $i++) {
    $currSheet = $resultWorkbook.sheets($i)
    # formatting columns that contain names
    for ($k = 0; $k -lt 2; $k++) {
        $col = $menu.day[$i - 1].submenu[$k].dish[0].resultCol - 1
        $currSheet.columns($col).autofit()
    }
    # formatting columns that contain information about unusual answers
    $col = $currSheet.usedRange.columns.count
    $currSheet.columns($col).autofit()
    $currSheet.columns($col).font.bold = $true
    $currSheet.columns($col).horizontalAlignment = $XL_ALIGNMENT_LEFT
    # adding an empty column between two submenu
    $col = $menu.day[$i - 1].submenu[1].dish[0].resultCol - 1
    $currSheet.columns($col).insert()
    $currSheet.columns($col).columnWidth = 1
    # changing zoom level
    $currSheet.activate()
    $excel.activeWindow.zoom = 75
}
$resultWorkbook.sheets(1).activate()



$resultWorkbook.saveAs("$pwd\result.xlsx")
$resultWorkbook.close()
$excel.screenUpdating = $true
$excel.displayAlerts = $true
$excel.quit()
} catch {
    $errorLogsPath = ".\ERROR_LOGS.txt"
    Add-Content -Path $errorLogsPath -Value (Get-Date)
    Add-Content -Path $errorLogsPath -Value "An error in the 2nd script occurred:"
    Add-Content -Path $errorLogsPath -Value $_
    Add-Content -Path $errorLogsPath -Value $_.ScriptStackTrace
    Add-Content -Path $errorLogsPath -Value ""
    Write-Host "An error in the 2nd script occurred:"
    Write-Host $_
    Write-Host $_.ScriptStackTrace
}