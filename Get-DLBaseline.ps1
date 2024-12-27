# Distribution List Lifecycle Management - Version 1.4

# Load the necessary module for handling Excel files
Import-Module ImportExcel

# Connect Exchange Online
Connect-ExchangeOnline

# Define the path to the Excel file and the backup location
$excelFilePath = ".\DistributionLists.xlsx"
$BackupFiles = ".\Backup"

# Function to send email notifications
function Send-Notification {
    param (
        [string]$subject,
        [string]$body
    )
    $mailParams = @{
        SmtpServer = "smtp.yourserver.com"
        From       = "your-email@domain.com"
        To         = "recipient-email@domain.com"
        Subject    = $subject
        Body       = $body
    }
    Send-MailMessage @mailParams
}

# Function to get email addresses of managers
function Get-ManagerEmails {
    param (
        [array]$managedBy
    )
    $managedByEmails = @()
    foreach ($manager in $managedBy) {
        if ($manager -ne "Organization Management") {
            $recipient = Get-Recipient -Identity $manager.Trim()
            if ($recipient) {
                $managedByEmails += $recipient.PrimarySmtpAddress
            }
        }
    }
    return $managedByEmails -join "; "
}

# Function to handle backups and retain last 5 instances
function Manage-Backups {
    param (
        [string]$filePath,
        [string]$backupLocation
    )
    if (-Not (Test-Path $backupLocation)) {
        New-Item -Path $backupLocation -ItemType Directory
    }
    $backupFileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath) + "_" + (Get-Date -Format "yyyyMMddHHmmss") + ".xlsx"
    Copy-Item -Path $filePath -Destination (Join-Path -Path $backupLocation -ChildPath $backupFileName)

    $backupFiles = Get-ChildItem -Path $backupLocation -Filter "*.xlsx" | Sort-Object LastWriteTime -Descending
    if ($backupFiles.Count -gt 5) {
        $backupFiles | Select-Object -Skip 5 | Remove-Item
    }
}

# Start the main script logic
try {
    # Check if the Excel file exists
    if (-Not (Test-Path $excelFilePath)) {
        Write-Host "Excel file not found. Creating initial baseline..." -ForegroundColor Yellow
        
        $currentDistributionGroups = Get-DistributionGroup -ResultSize Unlimited -Filter 'RecipientType -ne "MailUniversalSecurityGroup"' | Select-Object -Property RecipientType, Alias, PrimarySMTPAddress, ManagedBy
        
        $baseline = @()
        foreach ($group in $currentDistributionGroups) {
            $MemberCount = ((Get-DistributionGroupMember $group.PrimarySmtpAddress -ResultSize Unlimited).Name).count
            $managedByEmailString = Get-ManagerEmails -managedBy $group.ManagedBy
            $baseline += [PSCustomObject]@{
                RecipientType      = $group.RecipientType
                Alias              = $group.Alias
                PrimarySMTPAddress = $group.PrimarySMTPAddress
                ManagedBy          = $managedByEmailString
                MemberCount        = $MemberCount
                IsDeleted          = $false
                LastEmailRecdDate  = $null
            }
        }
        
        $baseline | Export-Excel -Path $excelFilePath
        Write-Host "Initial baseline created and saved to $excelFilePath" -ForegroundColor Green

        Send-Notification -subject "Distribution List Baseline Created" -body "Initial baseline has been created and saved to $excelFilePath."
    }
    else {
        # Backup the existing Excel file
        Manage-Backups -filePath $excelFilePath -backupLocation $BackupFiles

        $baseline = Import-Excel -Path $excelFilePath
        
        $currentDistributionGroups = Get-DistributionGroup -ResultSize Unlimited -Filter 'RecipientType -ne "MailUniversalSecurityGroup"' | Select-Object -Property RecipientType, Alias, PrimarySMTPAddress, ManagedBy
        
        $currentGroupsHash = @{}
        foreach ($group in $currentDistributionGroups) {
            $currentGroupsHash[$group.PrimarySMTPAddress] = $group
        }
        
        $updatedEntries = 0
        $newEntries = 0
        $deletedEntries = 0
        
        foreach ($entry in $baseline) {
            if ($currentGroupsHash.ContainsKey($entry.PrimarySMTPAddress)) {
                $entry.IsDeleted = $false
                $entry.Alias = $currentGroupsHash[$entry.PrimarySMTPAddress].Alias
                $entry.ManagedBy = Get-ManagerEmails -managedBy $currentGroupsHash[$entry.PrimarySMTPAddress].ManagedBy
                $entry.MemberCount = ((Get-DistributionGroupMember $entry.PrimarySMTPAddress -ResultSize Unlimited).Name).count
                $updatedEntries++
            }
            else {
                $entry.IsDeleted = $true
                $deletedEntries++
            }
        }
        
        foreach ($group in $currentDistributionGroups) {
            if (-not ($baseline.PrimarySMTPAddress -contains $group.PrimarySMTPAddress)) {
                $MemberCount = ((Get-DistributionGroupMember $group.PrimarySmtpAddress -ResultSize Unlimited).Name).count
                $managedByEmailString = Get-ManagerEmails -managedBy $group.ManagedBy
                $newEntry = [PSCustomObject]@{
                    RecipientType      = $group.RecipientType
                    Alias              = $group.Alias
                    PrimarySMTPAddress = $group.PrimarySMTPAddress
                    ManagedBy          = $managedByEmailString
                    MemberCount        = $MemberCount
                    IsDeleted          = $false
                    LastEmailRecdDate  = $null
                }
                $baseline += $newEntry
                $newEntries++
            }
        }
        
        $baseline | Export-Excel -Path $excelFilePath

        Write-Host "### Distribution list baseline has been updated ###" -ForegroundColor Yellow
        Write-Host "Updated entries: $updatedEntries" -ForegroundColor Green
        Write-Host "New entries added: $newEntries" -ForegroundColor Magenta
        Write-Host "Entries marked as deleted: $deletedEntries" -ForegroundColor Red

        Send-Notification -subject "Distribution List Baseline Updated" -body "Distribution list baseline has been updated. Updated entries: $updatedEntries, New entries: $newEntries, Deleted entries: $deletedEntries."
    }
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
    Send-Notification -subject "Distribution List Script Error" -body "An error occurred while updating the distribution list baseline: $_"
}