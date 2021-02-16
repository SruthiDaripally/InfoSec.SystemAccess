function Save-EntitlementGroup {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Group,

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
        $existsQuery = "SELECT ent_id FROM entitlements WHERE (ent_id = {0} AND system = {1}) OR (name = {2} AND ent_type = {3} AND system = {1})" `
            -f (@($Group.EntitlementId,$Group.System,$Group.Name,$Group.EntitlementType) | Select-SqlValues)
        $existsCommand = New-Object System.Data.SqlClient.SqlCommand $existsQuery, $conn
        $ent_id = $existsCommand.ExecuteScalar()

        if ($null -eq $ent_id) {
            $insertQuery = "INSERT INTO Entitlements (ent_id, system, name, description, last_seen, ent_type) VALUES ({0},{1},{2},{3},{4},{5})" `
                -f (@($Group.EntitlementId,$Group.System,$Group.Name,$Group.Description,$Group.LastSeen,$Group.EntitlementType) | Select-SqlValues)
            $insertCommand = $conn.CreateCommand()
            $insertCommand.CommandText = $insertQuery
            $insertCommand.ExecuteNonQuery() | Out-Null
        } else {
            $updateQuery = "UPDATE Entitlements SET last_seen = '$($Group.LastSeen)' WHERE ent_id = {0} AND system = {1}" -f (@($ent_id,$Group.System) | Select-SqlValues)
            $updateCommand = $conn.CreateCommand()
            $updateCommand.CommandText = $updateQuery
            $updateCommand.ExecuteNonQuery() | Out-Null
            $Group.EntitlementId = $ent_id
        }

        if ($Passthru) { $Group }
    }
    END {
        $conn.Close()
    }
}
