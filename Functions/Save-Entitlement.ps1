function Save-Entitlement {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Entitlement
    )

    BEGIN {
        $sqlCreds = @{}
        if ($Environment -eq "Prod") {
            $sqlCreds = "User ID=svtPapp_svc;Password=$(Get-SqlPassword)"
        } else {
            $sqlCreds = "Integrated Security=SSPI"
        }
        $sqlCstr = "Server=$($SqlInstance);Database=$($SqlDatabase);$sqlCreds"
        $entitlementConn = New-Object System.Data.SqlClient.SqlConnection($sqlCstr)
        $entitlementConn.Open()
    }
    PROCESS {
        $existsQuery = "SELECT COUNT(*) FROM Acct_Ent WHERE ent_id = '{0}' AND sys_id = '{1}' AND system = '$($script:System)'" -f $Entitlement.EntitlementId,$Entitlement.SystemId
        $existsCommand = New-Object System.Data.SqlClient.SqlCommand $existsQuery, $entitlementConn
        $exists = $existsCommand.ExecuteScalar()

        if ($exists -gt 0) {
            $updateQuery = "UPDATE Acct_Ent SET last_seen = '{0}' WHERE ent_id = '{1}' AND sys_id = '{2}' AND system = '$($script:System)'" -f $script:LastSeen,$Entitlement.EntitlementId,$Entitlement.SystemId
            $updateCommand = New-Object System.Data.SqlClient.SqlCommand $updateQuery, $entitlementConn
            $updateCommand.ExecuteNonQuery() | Out-Null
        } else {
            $insertQuery = "INSERT INTO Acct_Ent (ent_id, sys_id, system, last_seen) VALUES ({0}, {1}, {2}, {3})" -f @(($Entitlement.EntitlementId,$Entitlement.SystemId,$script:System,$script:LastSeen) | Select-SqlValues)
            $insertCommand = New-Object System.Data.SqlClient.SqlCommand $insertQuery, $entitlementConn
            $insertCommand.ExecuteNonQuery() | Out-Null
        }
    }
    END {
        $entitlementConn.Close()
    }
}