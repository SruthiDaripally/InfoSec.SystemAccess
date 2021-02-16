function Get-Db2Password {
    param(
        [Parameter(Mandatory=$true)]
        $Db2Environment
    )
    $pwdFile = "$PSScriptRoot\db2-$Db2Environment.txt"
    $crypted = [Convert]::FromBase64String((Get-Content -LiteralPath $pwdFile))
    $clear = [System.Security.Cryptography.ProtectedData]::Unprotect($crypted, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
    $enc = [System.Text.Encoding]::Default
    $enc.GetString($clear)
}

function Select-Accounts {
    $db2Conn = New-Object System.Data.OleDb.OleDbConnection($script:cstr)
    $db2Conn.Open()
    $cmd = $db2Conn.CreateCommand()
    $cmd.CommandText = $usersQuery -f $script:Db2Prefix
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        $fn = ([string]$reader["CICL_FST_NM"]).Trim()
        $ln = ([string]$reader["CICL_LST_NM"]).Trim()
        $account = [pscustomobject]@{
            SystemId = ([string]$reader["SEC_USR_CLT_ID"]).Trim()
            System = $script:System
            AccountName = ([string]$reader["SEC_USR_ID"]).Trim()
            Name = "{1}, {0}" -f $fn,$ln
            Email = ([string]$reader["CIEM_EMAIL_ADR_TXT"]).Trim()
            Status = "Enabled"
            LastSeen = $script:LastSeen
            MailboxLocation = $null
        }
        $account
    }
    $reader.Close()
    $db2Conn.Close()
}

# Modified by Sruthi Daripally on 1/28/2020 to accomodate for 4 new entitlement types.
function Select-Entitlements {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Account
    )

    BEGIN {
        $groupLookup = @{}
        Select-EntitlementGroups | ForEach-Object { $_.LastSeen = $script:LastSeen; Save-EntitlementGroup $_ | Out-Null; $groupLookup.Add($_.Name, $_.EntitlementId) }
        $lossReserveLookup =@{}
        Select-EntitlementGroups-LossReserveAuthority | ForEach-Object {$_.LastSeen =$script:LastSeen; Save-EntitlementGroup $_| Out-Null; $lossReserveLookup.Add($_.Name, $_.EntitlementId)  }
        $lossPaymentLookup =@{}
        Select-EntitlementGroups-LossPaymentAuthority | ForEach-Object {$_.LastSeen =$script:LastSeen; Save-EntitlementGroup $_| Out-Null; $lossPaymentLookup.Add($_.Name, $_.EntitlementId)  }
        $expenseReserveLookup =@{}
        Select-EntitlementGroups-ExpenseReserveAuthority | ForEach-Object {$_.LastSeen =$script:LastSeen; Save-EntitlementGroup $_| Out-Null; $expenseReserveLookup.Add($_.Name, $_.EntitlementId) }
        $expensePaymentLookup =@{}
        Select-EntitlementGroups-ExpensePaymentAuthority | ForEach-Object {$_.LastSeen =$script:LastSeen; Save-EntitlementGroup $_| Out-Null; $expensePaymentLookup.Add($_.Name, $_.EntitlementId) }

        $conn = New-Object System.Data.OleDb.OleDbConnection($script:cstr)
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $script:AllUserGroupsQuery -f $script:Db2Prefix
        $groupsTable = New-Object System.Data.DataTable
        $groupsTable.Load($cmd.ExecuteReader())
        $conn.Close()
    }
    PROCESS {
        $entitlements = @()
        $groups = $groupsTable.Select("SEC_USR_ID = '" + $Account.AccountName + "'")
        foreach ($group in $groups) {
            $groupId = ([string]$group["SEC_GRP_ID"]).Trim()
            
            if ($null -ne $groupLookup[$groupId]) {
            $entitlements += [pscustomobject]@{SystemId = $Account.SystemId ; EntitlementId = $groupLookup[$groupId]}
            }
            if ($null -ne $lossReserveLookup[$groupId]) {
            $entitlements += [pscustomobject]@{SystemId = $Account.SystemId ; EntitlementId = $lossReserveLookup[$groupId]}   
            }
            if ($null -ne $lossPaymentLookup[$groupId]) {
            $entitlements += [pscustomobject]@{SystemId = $Account.SystemId ; EntitlementId = $lossPaymentLookup[$groupId]}
            }
            if ($null -ne $expenseReserveLookup[$groupId]) {
            $entitlements += [pscustomobject]@{SystemId = $Account.SystemId ; EntitlementId = $expenseReserveLookup[$groupId]}
            }
            if ($null -ne $expensePaymentLookup[$groupId]) {
            $entitlements += [pscustomobject]@{SystemId = $Account.SystemId ; EntitlementId = $expensePaymentLookup[$groupId]}
            }
            $entitlements
        }
    }
}

function Select-EntitlementGroups {
    $conn = New-Object System.Data.OleDb.OleDbConnection($script:cstr)
    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $groupsQuery -f $script:Db2Prefix
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        $groupNbr = $reader["SEC_USR_GRP_NBR"]
        $groupId = ([string]$reader["SEC_GRP_NM"]).Trim()
        $entitlementGroup = [pscustomobject]@{
            EntitlementId = $groupNbr
            System = $script:System
            Name = $groupId
            Description = $null
            LastSeen = ""
            EntitlementType = "Security Group"
        }
        $entitlementGroup
    }

    $reader.Close()
    $conn.Close()
}

# Added by Sruthi Daripally on 1/27/2020
function Select-EntitlementGroups-LossReserveAuthority {
    $conn = New-Object System.Data.OleDb.OleDbConnection($script:cstr)
    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $GroupsQuery_LossReserveAuthority -f $script:Db2Prefix
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        $groupId = ([string]$reader["CAJ_DIR_AUT_RES"]).Trim()
        $lossReserveAuthorityGroup = [pscustomobject]@{
            EntitlementId = [guid]::NewGuid().ToString()
            System = $script:System
            Name = $groupId
            Description = $null
            LastSeen = ""
            EntitlementType = "Loss Reserve Authority"
        }
        $lossReserveAuthorityGroup
    }
    
    $reader.Close()
    $conn.Close()
}

function Select-EntitlementGroups-LossPaymentAuthority {
    $conn = New-Object System.Data.OleDb.OleDbConnection($script:cstr)
    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $GroupsQuery_LossPaymentAuthority -f $script:Db2Prefix
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        $groupId = ([string]$reader["CAJ_DIR_AUT_PMT"]).Trim()
        $lossPaymentAuthorityGroup = [pscustomobject]@{
            EntitlementId = [guid]::NewGuid().ToString()
            System = $script:System
            Name = $groupId
            Description = $null
            LastSeen = ""
            EntitlementType = "Loss Payment Authority"
        }
        $lossPaymentAuthorityGroup
    }
    
    $reader.Close()
    $conn.Close()
}

function Select-EntitlementGroups-ExpenseReserveAuthority {
    $conn = New-Object System.Data.OleDb.OleDbConnection($script:cstr)
    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $GroupsQuery_ExpenseReserveAuthority -f $script:Db2Prefix
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        $groupId = ([string]$reader["CAJ_XPN_AUT_RES"]).Trim()
        $ExpenseReserveAuthorityGroup = [pscustomobject]@{
            EntitlementId = [guid]::NewGuid().ToString()
            System = $script:System
            Name = $groupId
            Description = $null
            LastSeen = ""
            EntitlementType = "Expense Reserve Authority"
        }
        $ExpenseReserveAuthorityGroup
    }
    
    $reader.Close()
    $conn.Close()
}

function Select-EntitlementGroups-ExpensePaymentAuthority {
    $conn = New-Object System.Data.OleDb.OleDbConnection($script:cstr)
    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $GroupsQuery_ExpensePaymentAuthority -f $script:Db2Prefix
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        $groupId = ([string]$reader["CAJ_XPN_AUT_PMT"]).Trim()
        $ExpensePaymentAuthorityGroup = [pscustomobject]@{
            EntitlementId = [guid]::NewGuid().ToString()
            System = $script:System
            Name = $groupId
            Description = $null
            LastSeen = ""
            EntitlementType = "Expense Payment Authority"
        }
        $ExpensePaymentAuthorityGroup
    }
    
    $reader.Close()
    $conn.Close()
}

function Start-Update {
    param(
        [Parameter(Mandatory=$true)]
        $Db2Environment
    )

    switch ($Db2Environment) {
        "Dev" {
            $script:Db2User = "DEXCEED"
            $script:Db2Port = "4004"
            $script:Db2Package = "DEV1"
            $script:Db2Address = "wbmdesl01"
            $script:Db2DataSource = "S390DEV"
            $script:Db2Prefix = "DBDEV"
            $script:System = "Exceed - Dev"
            break
        }
        "QA" {
            $script:Db2User = "QEXCEED"
            $script:Db2Port = "4004"
            $script:Db2Package = "QUA1"
            $script:Db2Address = "wbmdesl01"
            $script:Db2DataSource = "S390DEV"
            $script:Db2Prefix = "DBQA"
            $script:System = "Exceed - QA"
            break
        }
        "Prod" {
            $script:Db2User = "PEXCEED"
            $script:Db2Port = "4001"
            $script:Db2Package = "PRD1"
            $script:Db2Address = "wbmpesl01"
            $script:Db2DataSource = "S390PROD"
            $script:Db2Prefix = "DBPROD"
            $script:System = "Exceed"
            break
        }
    }

    $pass = Get-Db2Password -Db2Environment $Db2Environment
    $script:cstr = "User ID=$($script:Db2User);Password=$pass;Provider=DB2OLEDB;Persist Security Info=True;Initial Catalog=$($script:Db2DataSource);Network Transport Library=TCP;Host CCSID=37;PC Code Page=1252;Network Address=$($script:Db2Address);Network Port=$($script:Db2Port);Package Collection=$($script:Db2Package);Default Schema=WEBUSER;DBMS Platform=DB2/MVS;Process Binary as Character=False;Connection Pooling=False;Units of Work=RUW"
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}

Set-Variable -Name UsersQuery -Option Const -Value "SELECT u.SEC_USR_ID, u.SEC_USR_CLT_ID, e.CIEM_EMAIL_ADR_TXT, c.CICL_FST_NM, c.CICL_LST_NM FROM {0}.SEC_USRS u INNER JOIN {0}.CLIENT_EMAIL e ON e.CLIENT_ID = u.SEC_USR_CLT_ID INNER JOIN {0}.CLIENT_TAB c ON c.CLIENT_ID = u.SEC_USR_CLT_ID for read only with ur;"
Set-Variable -Name AllUserGroupsQuery -Option Const -Value "select * from {0}.SEC_USR_GRP_CTS for read only with ur;"
Set-Variable -Name GroupsQuery -Option Const -Value "select * from {0}.SEC_USR_GRP for read only with ur;"

# Added by SD on 1/28/2020
Set-Variable -Name GroupsQuery_LossReserveAuthority -Option Const -Value "select DISTINCT(CAJ_DIR_AUT_RES) from {0}.ADJUSTER_TAB for read only with ur;"
Set-Variable -Name GroupsQuery_LossPaymentAuthority -Option Const -Value "select DISTINCT(CAJ_DIR_AUT_PMT) from {0}.ADJUSTER_TAB for read only with ur;"
Set-Variable -Name GroupsQuery_ExpenseReserveAuthority -Option Const -Value "select DISTINCT(CAJ_XPN_AUT_RES) from {0}.ADJUSTER_TAB for read only with ur;"
Set-Variable -Name GroupsQuery_ExpensePaymentAuthority -Option Const -Value "select DISTINCT(CAJ_XPN_AUT_PMT) from {0}.ADJUSTER_TAB for read only with ur;"

Start-Update -Db2Environment "Dev"
Start-Update -Db2Environment "QA"
Start-Update -Db2Environment "Prod"