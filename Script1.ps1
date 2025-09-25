Add-Type -AssemblyName System.Windows.Forms

function Get-UniqueName {
    param (
        [string]$basePath,
        [string]$baseName,
        [string]$extension = ''
    )
    $i = 0
    $uniqueName = $baseName
    do {
        $suffix = if ($i -eq 0) { '' } else { " - $i" }
        $testPath = Join-Path $basePath ($baseName + $suffix + $extension)
        $i++
    } while (Test-Path $testPath)
    return $testPath
}

try {
    # Define base paths
    $sourcePath = "\\reiltys\iomgroot\DeptShare_DHSS_Nobles\Management\Director of Finance, Performance & Delivery\16. Manx Care\FAS DSC\Purchase Cards info\Purchase Cards Logs\To be sent to Treasury"
    $destinationPath = "\\reiltys\iomgroot\DeptShare_DHSS_Nobles\Management\Director of Finance, Performance & Delivery\16. Manx Care\FAS DSC\Purchase Cards info\Purchase Cards Logs\Sent to Treasury"
    $templatePath = "C:\Users\NADLUROB\Desktop\Email Templates\PurchaseCard_SendToTreasury.msg"

    # Collect all .xlsx logs
    $logs = Get-ChildItem -Path $sourcePath -Filter "*.xlsx" -File
    $logCount = $logs.Count
    if ($logCount -eq 0) {
        throw "No .xlsx logs found to process."
    }

    # Create folder name
    $datePrefix = Get-Date -Format 'dd MMMM yy'
    $baseFolderName = "$datePrefix - $logCount logs"
    $tempFolder = Join-Path $env:TEMP $baseFolderName
    $tempFolder = Get-UniqueName -basePath $env:TEMP -baseName "$datePrefix - $logCount logs"

    # Create temp folder
    New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null

    # Move logs to temp folder
    foreach ($log in $logs) {
        Move-Item -Path $log.FullName -Destination $tempFolder
    }

    # Zip the folder
    $zipFileTempPath = "$tempFolder.zip"
    Compress-Archive -Path "$tempFolder\*" -DestinationPath $zipFileTempPath -Force

    # Ensure unique destination folder and zip file names
    $finalFolderPath = Get-UniqueName -basePath $destinationPath -baseName "$datePrefix - $logCount logs"
    $finalZipPath = $finalFolderPath + ".zip"

    # Move both folder and zip
    Move-Item -Path $tempFolder -Destination $finalFolderPath
    Move-Item -Path $zipFileTempPath -Destination $finalZipPath

    # Open Outlook and load the template
    $outlook = New-Object -ComObject Outlook.Application
    $template = $outlook.Session.OpenSharedItem($templatePath)

    # Create new email (to get default signature)
    $mail = $outlook.CreateItem(0) # 0 = MailItem
    $mail.Subject = $template.Subject
    $mail.To = "Purchasecards.Treasury@gov.im"
    $mail.CC = $template.CC
    $mail.BCC = $template.BCC

    # Replace X in the body with actual log count
    $mail.Body = $template.Body -replace "\bX\b", "$logCount"

    # Attach zip file
    $mail.Attachments.Add($finalZipPath)

    # Send email
    $mail.Send()

    # Confirmation message
    [System.Windows.Forms.MessageBox]::Show("Successfully sent $logCount logs to Purchasecards.Treasury@gov.im", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
