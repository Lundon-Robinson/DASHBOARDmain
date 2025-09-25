param (
    [string]$SelectedMonth
)
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
Add-Type -AssemblyName System.Windows.Forms

# =============================================================================
# CONFIGURATION
# =============================================================================
# NOTE: This path should match the cardholder_list_file in src/core/config.py
$CARDHOLDER_LIST_PATH = "\\reiltys\iomgroot\DeptShare_DHSS_Nobles\Management\Director of Finance, Performance & Delivery\16. Manx Care\FAS DSC\Purchase Cards info\Card Holder List\Purchase cardholder list DSC.xls"
# =============================================================================



function Show-MessageBox {
    param(
        [string]$message,
        [string]$title = "Confirm",
        [System.Windows.Forms.MessageBoxButtons]$buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
    )
    return [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons)
}

# Validate SelectedMonth parameter before proceeding
if (-not $SelectedMonth) {
    Write-Error "No month selected. Please pass the month as a parameter."
    exit 1
}

$monthName = $SelectedMonth.Trim()
$validMonths = @("January","February","March","April","May","June","July","August","September","October","November","December")
if (-not ($validMonths -contains $monthName)) {
    Write-Error "Invalid month name passed: $monthName"
    exit 1
}

$monthNum = [datetime]::ParseExact($monthName, 'MMMM', $null).Month
$prevMonthNum = if ($monthNum -eq 1) { 12 } else { $monthNum - 1 }
$prevMonthName = (Get-Culture).DateTimeFormat.GetMonthName($prevMonthNum)

# Open Outlook Application COM
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# Get shared mailbox "Manx Care, Finance"
$sharedMailbox = $namespace.Folders | Where-Object { $_.Name -eq "Manx Care, Finance" }
if (-not $sharedMailbox) {
    Write-Error "Shared mailbox 'Manx Care, Finance' not found."
    exit 1
}

$inbox = $sharedMailbox.Folders | Where-Object { $_.Name -eq "Inbox" }
if (-not $inbox) {
    Write-Error "Inbox folder not found in 'Manx Care, Finance'."
    exit 1
}

# Destination folder "Done" under your personal Inbox
$personalInbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
$doneFolder = $personalInbox.Folders | Where-Object { $_.Name -eq "Done" }
if (-not $doneFolder) {
    Write-Error "'Done' folder not found under your personal Inbox."
    exit 1
}

# Purchase Cardholder List workbook path
$purchaseCardholderListPath = $CARDHOLDER_LIST_PATH
$purchaseCardholderListExcel = New-Object -ComObject Excel.Application
$purchaseCardholderListExcel.Visible = $false

function Test-FileInUse {
    param([string]$filePath)

    try {
        $stream = [System.IO.File]::Open($filePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Close()
        return $false # File is NOT in use
    }
    catch {
        return $true # File is in use
    }
}

# Before opening workbook, check if file is in use
if (Test-FileInUse -filePath $purchaseCardholderListPath) {
    Write-Host "The Purchase Cardholder List file is currently in use by another process. Please close it and try again."
    exit 1
}

$purchaseCardholderListWorkbook = $null
try {
    $purchaseCardholderListWorkbook = $purchaseCardholderListExcel.Workbooks.Open($purchaseCardholderListPath)
} catch {
    Write-Error "Failed to open Purchase cardholder list DSC.xls at path $purchaseCardholderListPath"
    exit 1
}

$outstandingLogsSheet = $purchaseCardholderListWorkbook.Sheets.Item("OUTSTANDING LOGS")
if (-not $outstandingLogsSheet) {
    Write-Error "'OUTSTANDING LOGS' sheet not found in Purchase cardholder list workbook."
    $purchaseCardholderListWorkbook.Close($false)
    $purchaseCardholderListExcel.Quit()
    exit 1
}

# Target folder to save attachments
$saveFolder = "\\reiltys\iomgroot\DeptShare_DHSS_Nobles\Management\Director of Finance, Performance & Delivery\16. Manx Care\FAS DSC\Purchase Cards info\Purchase Cards Logs\To be sent to Treasury"
if (-not (Test-Path $saveFolder)) {
    Write-Error "Save folder does not exist: $saveFolder"
    $purchaseCardholderListWorkbook.Close($false)
    $purchaseCardholderListExcel.Quit()
    exit 1
}

function Find-NameInOutstandingLogs {
    param([string]$fullName)

    $lastRow = $outstandingLogsSheet.UsedRange.Rows.Count
    $colAP = 42 # Column AP = 42

    # Exact match first
    for ($row=2; $row -le $lastRow; $row++) {
        $cellValue = $outstandingLogsSheet.Cells.Item($row, $colAP).Text
        if ($cellValue -eq $fullName) {
            return $row
        }
    }

    # Close match by substring and user confirmation
    for ($row=2; $row -le $lastRow; $row++) {
        $cellValue = $outstandingLogsSheet.Cells.Item($row, $colAP).Text
        if ($cellValue -and ($cellValue.ToLower().Contains($fullName.ToLower()) -or $fullName.ToLower().Contains($cellValue.ToLower()))) {
            $response = Show-MessageBox "Found a close name match: '$cellValue' for '$fullName'. Is this the correct match?" "Name confirmation" ([System.Windows.Forms.MessageBoxButtons]::YesNo)
            if ($response -eq [System.Windows.Forms.DialogResult]::Yes) {
                return $row
            }
        }
    }

    return $null
}

$updateLog = @()

# Excel COM for attachment workbooks
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $false
$excelApp.DisplayAlerts = $false

function Unprotect-MonthlyLogSheet {
    param($workbook)

    try {
        $sheet = $workbook.Sheets.Item("Monthly Log")
    } catch {
        return $null
    }
    if (-not $sheet) { return $null }
    try {
        $sheet.Unprotect("P123")
    } catch {
        # Ignore unprotect errors
    }
    return $sheet
}

$emailsToMove = @()
$emailsSkippedDueToPendingY = @()
$emailsSkippedDueToNoMatch = @()
$errors = @()

foreach ($mail in $inbox.Items) {
    try {
        if ($mail.Class -ne 43) { continue } # 43=MailItem

        $attachments = $mail.Attachments | Where-Object { $_.FileName -like "*.xlsx" }
        if (-not $attachments -or $attachments.Count -eq 0) {
            continue
        }

        foreach ($att in $attachments) {
            $tempFilePath = Join-Path $env:TEMP ([System.IO.Path]::GetRandomFileName() + ".xlsx")
            $att.SaveAsFile($tempFilePath)

            $wb = $null
            try {
                $wb = $excelApp.Workbooks.Open($tempFilePath)
            } catch {
                $errors += "Failed to open attachment '$($att.FileName)' in email '$($mail.Subject)'. Error: $_"
                Remove-Item $tempFilePath -ErrorAction SilentlyContinue
                continue
            }

            $monthlyLogSheet = Unprotect-MonthlyLogSheet -workbook $wb
            if (-not $monthlyLogSheet) {
                $errors += "Monthly Log sheet not found or cannot be unprotected in attachment '$($att.FileName)' from email '$($mail.Subject)'. Skipping this attachment."
                $wb.Close($false)
                Remove-Item $tempFilePath -ErrorAction SilentlyContinue
                continue
            }

            $d3Value = $monthlyLogSheet.Range("D3").Text
            if (-not $d3Value -or -not $d3Value.Contains(" ")) {
                $errors += "Invalid format in cell D3 in '$($att.FileName)' - cannot extract full name."
                $wb.Close($false)
                Remove-Item $tempFilePath -ErrorAction SilentlyContinue
                continue
            }
            $fullName = $d3Value.Substring($d3Value.IndexOf(" ") + 1).Trim()

            # Use $monthName and $prevMonthName from parameter, no extraction from Excel

            $yearShort = (Get-Date).ToString("yy")

            $rowFound = Find-NameInOutstandingLogs -fullName $fullName
            if (-not $rowFound) {
                $errors += "Name '$fullName' NOT found or confirmed in Purchase Cardholder list from '$($att.FileName)'. Email will NOT be moved."
                $emailsSkippedDueToNoMatch += $mail
                $wb.Close($false)
                Remove-Item $tempFilePath -ErrorAction SilentlyContinue
                continue
            }

            $colZValue = $outstandingLogsSheet.Cells.Item($rowFound, 26).Text.Trim()
            if ($colZValue -match '^LOG\s?-?PENDING$') {
                $outstandingLogsSheet.Cells.Item($rowFound, 26).Value2 = "Y"
                $updateLog += "${fullName}: changed row ${rowFound} to Y"

            } elseif ($colZValue -eq "Y") {
                $errors += "Row for '$fullName' in Purchase Cardholder list already marked 'Y'. Email will NOT be moved."
                $emailsSkippedDueToPendingY += $mail
                $wb.Close($false)
                Remove-Item $tempFilePath -ErrorAction SilentlyContinue
                continue
            } elseif ($colZValue -eq "") {
                $errors += "Row for '$fullName' in Purchase Cardholder list has empty status in column Z. Email will NOT be moved."
                $emailsSkippedDueToPendingY += $mail
                $wb.Close($false)
                Remove-Item $tempFilePath -ErrorAction SilentlyContinue
                continue
            }

            $purchaseCardholderListWorkbook.Save()

            $wb.Close($false)

            $newFileName = "$fullName for $prevMonthName $yearShort.xlsx"
            $newFilePath = Join-Path $saveFolder $newFileName

            if (Test-Path $newFilePath) { Remove-Item $newFilePath -Force }
            Copy-Item -Path $tempFilePath -Destination $newFilePath

            Remove-Item $tempFilePath -ErrorAction SilentlyContinue

            if (-not ($emailsSkippedDueToNoMatch -contains $mail) -and -not ($emailsSkippedDueToPendingY -contains $mail)) {
                $emailsToMove += $mail
            }
        }

    } catch {
        $errors += "Error processing email '$($mail.Subject)': $_"
    }
}

# Cleanup Excel COMs
$purchaseCardholderListWorkbook.Close($true)
$purchaseCardholderListExcel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($purchaseCardholderListWorkbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($purchaseCardholderListExcel) | Out-Null

$excelApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null

# Move emails to Done folder
foreach ($mailToMove in $emailsToMove | Select-Object -Unique) {
    try {
        $mailToMove.Move($doneFolder) | Out-Null
    } catch {
        $errors += "Failed to move email '$($mailToMove.Subject)' to Done folder: $_"
    }
}

# Output summary
Write-Host "Processing complete.`n"

if ($errors.Count -gt 0) {
    Write-Host "Errors / Warnings:"
    foreach ($err in $errors) { Write-Host "- $err" }
    Write-Host ""
}

if ($emailsSkippedDueToNoMatch.Count -gt 0) {
    Write-Host "Emails skipped due to name mismatch confirmation failure:"
    foreach ($mail in $emailsSkippedDueToNoMatch | Select-Object -Unique) { Write-Host "- $($mail.Subject)" }
    Write-Host ""
}

if ($emailsSkippedDueToPendingY.Count -gt 0) {
    Write-Host "Emails skipped due to existing 'Y' or empty status in Purchase cardholder list column Z:"
    foreach ($mail in $emailsSkippedDueToPendingY | Select-Object -Unique) { Write-Host "- $($mail.Subject)" }
    Write-Host ""
}

$updateLog | ForEach-Object { Write-Output $_ }
Write-Host "Emails moved to Done folder: $($emailsToMove.Count)"

# Release Outlook COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doneFolder) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($inbox) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sharedMailbox) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

[GC]::Collect()
[GC]::WaitForPendingFinalizers()
