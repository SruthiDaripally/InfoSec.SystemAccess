function Save-Account {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Account,

        [Switch]
        $Passthru
    )

    BEGIN {
        $sqlCreds = @{}
        if ($Environment -eq "Prod") {
            $sqlCreds = "User ID=svtPapp_svc;Password=$(Get-SqlPassword)"
        } else {
            $sqlCreds = "Integrated Security=SSPI"
        }
        $sqlCstr = "Server=$($SqlInstance);Database=$($SqlDatabase);$sqlCreds"
        $conn = New-Object System.Data.SqlClient.SqlConnection($sqlCstr)
        $conn.Open()
    }
    PROCESS {
        $existsQuery = "SELECT COUNT(*) FROM Accounts WHERE (system = '$script:System' AND sys_id = '{0}') 
        OR ( system ='$script:System'  AND account_name = '{1}' AND sys_id ='{0}')" -f $Account.SystemId,$Account.AccountName
        $existsCommand = New-Object System.Data.SqlClient.SqlCommand $existsQuery, $conn
        $exists = $existsCommand.ExecuteScalar()

        if ($exists -gt 0) {
            $updateQuery = "UPDATE Accounts SET account_name = {0}, name = {1}, email = {2}, status = {3}, `
                last_seen = {4}, mailbox_location = {5} WHERE sys_id = {6} AND system = '$($script:System)'" `
                -f @(($Account.AccountName,$Account.Name,$Account.Email,$Account.Status,$script:LastSeen,$Account.MailboxLocation,$Account.SystemId) | Select-SqlValues)
            $updateCommand = New-Object System.Data.SqlClient.SqlCommand $updateQuery, $conn
            $updateCommand.ExecuteNonQuery() | Out-Null
        } else {
            $insertQuery = "INSERT INTO Accounts (sys_id, system, account_name, name, email, status, last_seen, mailbox_location) `
                VALUES ({0},{1},{2},{3},{4},{5},{6},{7})" -f @(($Account.SystemId,$script:System,$Account.AccountName,$Account.Name,$Account.Email,$Account.Status,$script:LastSeen,$Account.MailboxLocation) | Select-SqlValues)
            $insertCommand = New-Object System.Data.SqlClient.SqlCommand $insertQuery, $conn
            $insertCommand.ExecuteNonQuery() | Out-Null
        }

        if ($Passthru) { $Account }
    }
    END {
        Remove-AccountsNotSeen -System $script:System -LastSeen $script:LastSeen -Connection $conn
        Remove-EntitlementsNotSeen -System $script:System -LastSeen $script:LastSeen -Connection $conn

        $conn.Close()
    }
}