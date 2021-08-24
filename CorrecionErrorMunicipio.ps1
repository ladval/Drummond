$JsonConfigFile = "$PSScriptRoot\settings\settings.json"
$Global:oJsonSettings = Get-Content -Path $JsonConfigFile | ConvertFrom-Json
. "$PSScriptRoot\sql.ps1"


$invoiceSQLquery = @"
    SELECT id,InvoiceNumber,JsonFact 
    FROM [BotAbc].[dbo].[tfact_ApiProcesos] 
    WHERE Status LIKE '%Error%'
"@
$errorData = SQL_Query $invoiceSQLquery
foreach ($data in $errorData) {
    $InvoiceNumber = $data.InvoiceNumber
    $idInvoice = $data.id
    $JsonDataFile = "$PSScriptRoot\data\$($data.$InvoiceNumber).json"
    Set-Content -Path $JsonDataFile -Value $data.JsonFact
    $JsonDataFileContent = Get-Content $JsonDataFile | ConvertFrom-Json
    $CityCode = $JsonDataFileContent.CustomerInformation.CityCode
    if ($CityCode.Length -ne 5) {
        Write-Host "$InvoiceNumber Error de municipio " -ForegroundColor DarkYellow
        $CityCode = "0$CityCode"
        $JsonDataFileContent.CustomerInformation.CityCode = $CityCode
        $JsonDataFileContent = $JsonDataFileContent | ConvertTo-Json -Depth 100

        $updateQuery = @"
    UPDATE [BotAbc].[dbo].[tfact_ApiProcesos]
    SET 
    Status = 'SinProcesar',
    JsonFact = '$JsonDataFileContent'
    WHERE id = $idInvoice
"@
        SQL_Query $updateQuery
    }
    else {
        Write-Host "$InvoiceNumber sin error de municipio " -ForegroundColor DarkGreen

    }
}
Remove-Item "$PSScriptRoot\data\*.json*"