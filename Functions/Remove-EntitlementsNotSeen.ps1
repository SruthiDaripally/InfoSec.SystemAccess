function Remove-EntitlementsNotSeen {
    param(
        [Parameter(Mandatory=$true)]
        $System,

        [Parameter(Mandatory=$true)]
        $LastSeen,

        [Parameter(Mandatory=$true)]
        $Connection
    )

    $selectQuery = "SELECT ent_id FROM Acct_Ent WHERE system = '$System' AND last_seen < '$LastSeen'"
    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $selectQuery
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $existsTable = New-Object System.Data.DataTable
    $adapter.Fill($existsTable) | Out-Null

    foreach ($ent in $existsTable.Rows) {
        $acctEntQuery = "DELETE FROM Acct_Ent WHERE ent_id = '{0}' AND system = '{1}'" -f $ent["ent_id"],$System
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = $acctEntQuery
        $cmd.ExecuteNonQuery() | Out-Null
    }

    $entitlementQuery = "DELETE FROM Entitlements WHERE system = '$System' AND last_seen < '$LastSeen'"
    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $entitlementQuery
    $cmd.ExecuteNonQuery() | Out-Null
}