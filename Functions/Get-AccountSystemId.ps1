function Get-AccountSystemId {
    param(
        [Parameter(Mandatory=$true)]
        $AccountName        
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
        $existsQuery = "SELECT sys_id FROM Accounts WHERE  system ='$script:System'  AND account_name = '{0}'" -f $Account.AccountName
        $existsCommand = New-Object System.Data.SqlClient.SqlCommand $existsQuery, $conn
        $exists = $existsCommand.ExecuteScalar()

        if ($null -ne $exists) {
            $sysId = $exists
        }
        
         return $sysId
        
    }
    END {
        
        $conn.Close()
    }
}