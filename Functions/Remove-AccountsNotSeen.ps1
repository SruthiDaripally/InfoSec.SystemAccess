function Remove-AccountsNotSeen {
    param(
        [Parameter(Mandatory=$true)]
        $System,

        [Parameter(Mandatory=$true)]
        $LastSeen,

        [Parameter(Mandatory=$true)]
        $Connection
    )

    $selectQuery = "SELECT sys_id FROM Accounts WHERE system = '$System' AND last_seen < '$LastSeen'"
    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $selectQuery
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $existsTable = New-Object System.Data.DataTable
    $adapter.Fill($existsTable) | Out-Null

    foreach ($acct in $existsTable.Rows) {
        $entQuery = "SELECT ent_id FROM Acct_Ent WHERE sys_id = '{0}' AND system = '{1}'" -f $acct["sys_id"],$System
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = $entQuery
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $entitlementsTable = New-Object System.Data.DataTable
        $adapter.Fill($entitlementsTable) | Out-Null

        $acctEntQuery = "DELETE FROM Acct_Ent WHERE sys_id = '{0}' AND system = '{1}'" -f $acct["sys_id"],$System
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = $acctEntQuery
        $cmd.ExecuteNonQuery() | Out-Null

        $deleteAccountsQuery = "DELETE FROM Accounts WHERE sys_id = '{0}' AND system = '{1}'" -f $acct["sys_id"],$System
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = $deleteAccountsQuery
        $cmd.ExecuteNonQuery() | Out-Null
    }
}