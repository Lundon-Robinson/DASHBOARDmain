# Connect to Outlook
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# Get the default inbox
$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# Get or create "People" folder at same level as Inbox
$parentFolder = $inbox.Parent.Folders
$peopleFolder = $parentFolder | Where-Object { $_.Name -eq "People" }
if (-not $peopleFolder) {
    $peopleFolder = $parentFolder.Add("People")
}

# Cache folder names
$folderCache = @{}

# Clone inbox items to avoid collection change issues
$emails = @()
foreach ($item in $inbox.Items) {
    if ($item -is [Microsoft.Office.Interop.Outlook.MailItem]) {
        $emails += $item
    }
}

foreach ($mail in $emails) {
    $senderName = $mail.SenderName
    $senderEmail = $mail.SenderEmailAddress

    if ([string]::IsNullOrWhiteSpace($senderName)) { continue }

    # Normalize name: trim, remove multiple spaces
    $cleanName = ($senderName -replace '\s+', ' ').Trim()

    # Fallback: use part of email if name is too short
    $folderName = $cleanName
    if ($folderName.Split(" ").Count -lt 2 -and $senderEmail -match "(.+?)@") {
        $folderName = $Matches[1]
    }

    # Strip invalid folder name characters
    $folderName = $folderName -replace '[\\/:*?"<>|]', ''

    # Check cache
    if ($folderCache.ContainsKey($folderName)) {
        $targetFolder = $folderCache[$folderName]
    } else {
        $targetFolder = $peopleFolder.Folders | Where-Object { $_.Name -eq $folderName }
        if (-not $targetFolder) {
            try {
                $targetFolder = $peopleFolder.Folders.Add($folderName)
                Write-Host "Created folder: $folderName"
            } catch {
                Write-Warning "Could not create folder: $folderName - $_"
                continue
            }
        }
        $folderCache[$folderName] = $targetFolder
    }

    # Move the email
    try {
        $mail.Move($targetFolder) | Out-Null
        Write-Host "Moved email from $senderName to '$folderName'"
    } catch {
        Write-Warning "Failed to move email from $senderName - $_"
    }
}

# ------------------------------------
# 🔥 Delete all empty folders under "People"
# ------------------------------------
foreach ($subFolder in $peopleFolder.Folders) {
    try {
        if ($subFolder.Items.Count -eq 0 -and $subFolder.Folders.Count -eq 0) {
            Write-Host "Deleting empty folder: $($subFolder.Name)"
            $subFolder.Delete()
        }
    } catch {
        Write-Warning "Failed to delete folder: $($subFolder.Name) - $_"
    }
}
