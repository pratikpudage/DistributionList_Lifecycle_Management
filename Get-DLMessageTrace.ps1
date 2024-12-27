# Load the necessary module for handling Excel files
Import-Module ImportExcel

# Connect Exchange Online
Connect-ExchangeOnline

# Define the path to the Excel file
$excelFilePath = ".\MsgTraceDetails.xlsx"

# Check if the file exists, then delete it
if (Test-Path $excelFilePath) {
    Remove-Item $excelFilePath -Force
    Write-Output "Existing file $excelFilePath deleted."
}

# Define Variables
$StartDate = (Get-Date).AddDays(-9).ToString("MM/dd/yyyy")
$EndDate = (Get-Date).AddDays(+1).ToString("MM/dd/yyyy")

# Initialize an empty array to store message trace results
$AllMsgTraceResults = @()

# Initialize the page number
$Page = 1

do {
    # Retrieve message trace results for the current page
    $MsgTraceResult = Get-MessageTrace `
        -StartDate $StartDate `
        -EndDate $EndDate `
        -PageSize 5000 `
        -Page $Page `
        -Status Expanded `
    | Select RecipientAddress, Received

    # Add the current page results to the all results array
    $AllMsgTraceResults += $MsgTraceResult

    # Increment the page number
    $Page++

    # Check if the current page has less than 5000 results, indicating it's the last page
    $HasMoreResults = ($MsgTraceResult.Count -eq 5000)
} while ($HasMoreResults)

# Group by RecipientAddress and select the latest entry for each group
$LatestMsgTraceResult = $AllMsgTraceResults | Group-Object RecipientAddress | ForEach-Object {
    $_.Group | Sort-Object -Property Received -Descending | Select-Object -First 1
}

# Export the filtered message trace info to the Excel file
$LatestMsgTraceResult | Export-Excel -Path $excelFilePath