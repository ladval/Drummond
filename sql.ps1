



function SQL_Query($Query) {
    if ([string]::IsNullOrEmpty($Query)) {
        Write-Output 'EmptyString'
        Return $False
    }
    $ServerInstance = $oJsonSettings.conn.instancia
    $Database = $oJsonSettings.conn.db
    $Username = $oJsonSettings.conn.user
    $Password = $oJsonSettings.conn.pass
    try {
        $oSqlData = Invoke-Sqlcmd -Query $Query `
            -ServerInstance $ServerInstance `
            -Database $Database `
            -Username $Username `
            -Password $Password `
            -maxcharlength ([int]::MaxValue)
        Return $oSqlData
    }
    catch {
        $_.Message.Exception
        Return "SQL Exception"
    }
}
