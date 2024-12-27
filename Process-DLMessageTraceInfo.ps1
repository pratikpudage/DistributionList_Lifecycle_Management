# Import ImportExcel module if not already imported
Import-Module ImportExcel

# Path to Excel files
$distListFile = ".\DistributionLists.xlsx"
$msgTraceFile = ".\MsgTraceDetails.xlsx"

# Import data from DistributionLists.xlsx
$distributionLists = Import-Excel -Path $distListFile

# Import data from MsgTraceDetails.xlsx
$msgTraceDetails = Import-Excel -Path $msgTraceFile

# Define a hashtable to store merged data temporarily
$mergedData = @{}

# Merge the data based on PrimarySmtpAddress and RecipientAddress matching
foreach ($msgTrace in $msgTraceDetails) {
    $primarySmtp = $msgTrace.RecipientAddress

    foreach ($distList in $distributionLists) {
        if ($distList.PrimarySmtpAddress -eq $primarySmtp) {
            $mergedData[$distList.PrimarySmtpAddress] = [PSCustomObject]@{
                RecipientType = $distlist.RecipientType
                Alias = $distList.Alias
                PrimarySmtpAddress = $distList.PrimarySmtpAddress
                ManagedBy = $distList.ManagedBy
                IsDeleted = $distList.IsDeleted
                LastEmailRecdDate = $msgTrace.Received  # Received column from MsgTraceDetails
            }
        }
    }
}

# Update LastEmailRcdDate column in DistributionLists.xlsx
foreach ($item in $distributionLists) {
    $primarySmtpAddress = $item.PrimarySmtpAddress
    
    if ($mergedData.ContainsKey($primarySmtpAddress)) {
        $item.LastEmailRecdDate = $mergedData[$primarySmtpAddress].LastEmailRecdDate
    }
}

# Export updated data back to DistributionLists.xlsx
$distributionLists | Export-Excel -Path $distListFile -ClearSheet -AutoSize