function JsonData() {
    param (
        [int]
        $id,
        [string]
        $invoice
    )
    $drummondSQLquery = @"
    SELECT JsonFact 
    FROM [BotAbc].[dbo].[tfact_ApiProcesos] 
    WHERE id = $($id)
"@
    $jsonFact = SQL_Query $drummondSQLquery
    $jsonFact = $jsonFact.JsonFact
    $JsonDataFile = "$PSScriptRoot\data\$($invoice).json"
    Set-Content -Path $JsonDataFile -Value $jsonFact
    $JsonDataFile = Get-Content $JsonDataFile | ConvertFrom-Json
    $ItemInformation = $JsonDataFile.ItemInformation
    foreach ($item in $ItemInformation) {
        $item.Reference
    }
    
    break
}

