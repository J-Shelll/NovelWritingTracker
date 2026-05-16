# Writing Tracker - Step 1
# This script opens a Word document and counts the words.
# Set the file paths

$projectFolder = $PSScriptRoot

$wordFile = Join-Path $projectFolder "Novel.docx"
$excelFile = Join-Path $projectFolder "NovelTracker.xlsx"

# Check whether the Word document exists
$wordFileExists = Test-Path $wordFile

# Check whether the Excel tracker file exists
$excelFileExists = Test-Path $excelFile



if ($fileExists -eq $false) {
    Write-Host "ERROR: Word file not found."
    exit
}

if ($excelFileExists -eq $false) {
    Write-Host "ERROR: Excel file not found."
    exit
}

# Show the file check results
Write-Host "Word file exists: $wordFileExists"
Write-Host "Excel file exists: $excelFileExists"

#Start Microisoft Word in the background
$word = New-Object  -ComObject Word.Application
$word.Visible = $false 

Write-Host "Microsoft Word started successfully." 

# Open the Word document
$document = $word.Documents.Open($wordFile)

Write-Host "Word document opened successfully."

# Count the words in the document using Microsoft Word's built-in word count
$wordCount = $document.Range().ComputeStatistics(0)

Write-Host "Current word count: $wordCount"

# Close the Word document
$document.Close()

# Close Microsoft Word
$word.Quit()

# Start Microsoft Excel in the background
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

Write-Host "Microsoft Excel started successfully."

# Open the Excel workbook
$workbook = $excel.Workbooks.Open($excelFile)

Write-Host "Excel workbook opened successfully."


# Select the first worksheet
$sheet = $workbook.Worksheets.Item(1)

if ($null -eq $sheet) {
    Write-Host "ERROR: Worksheet was not selected."
    exit
}

Write-Host "First worksheet selected successfully."
Write-Host "Worksheet name: $($sheet.Name)"

# Get today's date in the same format as the Excel table
$today = Get-Date -Format "dd.MM.yyyy"

Write-Host "Today's date is: $today" 

# Read the first date cell from the table: C3
$firstDateCell = $sheet.Cells.Item(3, 3).Text

Write-Host "First date in table is: $firstDateCell"

# This will store the column number where today's date is found
$todayColumn = $null

# Search row 3 for today's date
foreach ($column in 3..60) {
    $dateInCell = $sheet.Cells.Item(3, $column).Text

    Write-Host "Checking column $column. Cell contains: $dateInCell"

    if ($dateInCell -eq $today) {
        $todayColumn = $column
    }
}

# Show which column matched today's date
Write-Host "Today's column is: $todayColumn"

if ($todayColumn -eq $null) {
    Write-Host "ERROR: Today's date was not found in row 3 of the Excel table."
    exit
}

# Get yesterday's date in the same format as the Excel table
$yesterday = (Get-Date).AddDays(-1).ToString("dd.MM.yyyy")

Write-Host "Yesterday's date is: $yesterday"

# This will store the column number where yesterday's date is found
$yesterdayColumn = $null

# Search row 3 for yesterday's date
foreach ($column in 3..60) {
    $dateInCell = $sheet.Cells.Item(3, $column).Text

    if ($dateInCell -eq $yesterday) {
        $yesterdayColumn = $column
    }
}

Write-Host "Yesterday's column is: $yesterdayColumn"

if ($yesterdayColumn -eq $null) {
    Write-Host "ERROR: Yesterday's date was not found in row 3 of the Excel table."
    exit
}

# Row numbers in the Excel table
$wordsPerDayRow = 4
$totalRow = 5

# Read yesterday's total word count from Excel
$yesterdayTotal = $sheet.Cells.Item($totalRow, $yesterdayColumn).Value2

Write-Host "Yesterday's total word count was: $yesterdayTotal"

# Calculate words written today
$wordsWrittenToday = [int]$wordCount - [int]$yesterdayTotal

Write-Host "Words written today: $wordsWrittenToday"

# Write today's total word count into the Total row
$sheet.Cells.Item($totalRow, $todayColumn).Value2 = $wordCount

Write-Host "Wrote current word count to row $totalRow, column $todayColumn."

# Write words written today into the Words per Day row
$sheet.Cells.Item($wordsPerDayRow, $todayColumn).Value2 = [int]$wordsWrittenToday

Write-Host "Wrote words per day to row $wordsPerDayRow, column $todayColumn."

# Row number where the target result should be written
$targetResultRow = 6

# Daily word target
$dailyTarget = 500

# Check whether today's writing target was met
if ($wordsWrittenToday -ge $dailyTarget) {
    $targetResult = "Met"
} else {
    $targetResult = "Short"
}

# Write the target result into row 6 under today's date
$sheet.Cells.Item($targetResultRow, $todayColumn).Value = $targetResult

Write-Host "Target result: $targetResult"
Write-Host "Wrote target result to row $targetResultRow, column $todayColumn."

# Close the Excel workbook without saving
$workbook.Save()
$workbook.Close($true)

# Close Microsoft Excel
$excel.Quit()