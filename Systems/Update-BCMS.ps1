function Import-MainframeGroupDefs {
    $mfgpath = "\\wbwpfil01\public$\IT General\Associate Setup\RACF Group Descriptions and Hard Coded Programs and GTAMS Descriptions.xlsx"
    if (-not(Test-Path -Path $mfgpath))
    {
        throw "$mfgpath wasn't found"
    }
    $lookup = Import-Excel -Path $mfgpath -WorksheetName "RACF Groups and Descriptions" | Group-Object -Property "Group" -AsHashTable
    $lookup
}

function Import-HardCodedProgramDefs {
    $progdefspath = "\\wbwpfil01\public$\IT General\Associate Setup\RACF Group Descriptions and Hard Coded Programs and GTAMS Descriptions.xlsx"
    if (-not(Test-Path -Path $progdefspath))
    {
        throw "$progdefspath wasn't found"
    }

    $x = Open-ExcelPackage -Path $progdefspath
    $sheet = $x.Workbook.Worksheets["Hard Coded Programs and GTAMS"]
    $lookup = @{}
    $start = $sheet.Dimension.Start
    $end = $sheet.Dimension.End
    for ($row = $start.Row; $row -le $end.Row; $row++) {
        $key = ($sheet.Cells[$row, 1]).Text.ToLower().Trim()
        if ($key) {
            $description = ($sheet.Cells[$row, 2]).Text
            if ($description) { $description = $description.Trim() }
            if ($lookup.ContainsKey($key)) { }
            else {
                $lookup.Add($key, $description)
            }
        }
    }
    $x | Close-ExcelPackage -NoSave
    Remove-Variable x
    $lookup
}

function Import-BcmsGtams {
    $bcmsGtamsPath = "\\wbmpfil01\public$\IT General\Associate Setup\GTAMS\Claims GTAMS\BCMS GTAMS\BCMS GTAMS.xlsx"
    if (-not(Test-Path $bcmsGtamsPath))
    {
        throw "$bcmsGtamsPath wasn't found"
    }

    $x = Open-ExcelPackage -Path $bcmsGtamsPath
    $lookup = @{}
    1..$x.Workbook.Worksheets.Count | ForEach-Object {
        $sheet = $x.Workbook.Worksheets.Item($_)
        if ($sheet.Name -match "Revision|OLD") { return } 
        $gtamDesc = $Script:HardCodedDefs[$sheet.Name]
        if ($gtamDesc -and $sheet.Name -notmatch "B350|B353") {
            $gtam = "{0} - {1}" -f $sheet.Name, $gtamDesc
        } else {
            $gtam = $sheet.Name
        }
        $end = $sheet.Dimension.End
        for ($row = 3; $row -le $end.Row; $row++) {
            $newrow = [ordered]@{
                Gtam = $gtam
                OpID = $sheet.Cells[$row,1].Text
                UserID = $sheet.Cells[$row,2].Text
                Name = $sheet.Cells[$row,3].Text
            }
            switch ($sheet.Name) {
                "B350" {
                    $newrow += [ordered]@{
                        BalanceByBatch = $sheet.Cells[$row,4].Text
                        DepositOption = $sheet.Cells[$row,5].Text
                        ApplyOption = $sheet.Cells[$row,6].Text
                    }
                    break
                }
                "B353" {
                    $newrow += [ordered]@{
                        AllocateCash = $sheet.Cells[$row,4].Text
                        WriteOffAuthority = $sheet.Cells[$row,5].Text
                        WriteOffLimit = $sheet.Cells[$row,6].Text
                        DisbursementAuthority = $sheet.Cells[$row,7].Text
                        DisbursementLimit = $sheet.Cells[$row,8].Text
                    }
                    break
                }
            }
            if ($lookup.ContainsKey($newrow.OpID)) {
                ($lookup[$newrow.OpID]).Add((New-Object PSCustomObject -Property $newrow)) | Out-Null
            } else {
                $lookup.Add($newrow.OpID, [System.Collections.ArrayList]@(New-Object PSCustomObject -Property $newrow))
            }
        }
    }
    $x | Close-ExcelPackage -NoSave
    Remove-Variable x
    $lookup
}

function Import-OpIds {
    $opidPath = "\\wbmpfil01\private$\Information Security\Manager Access Reviews\Mainframe\OPID"
    if (-not(Test-Path $opidPath))
    {
        throw "$opidPath wasn't found"
    }
    $csv = Import-Csv -Path $opidPath -Header OpID,UserId,Name
    $output = @{}
    $csv | ForEach-Object {
        $output.Add($_.UserId.Trim(), $_.OpID.Trim())
    }
    $output
}

function Select-Accounts {
    $users = Import-Csv -Path $groupsPath -Header Group,UserId,Name | Where-Object { $_.Group.StartsWith("BCMS") } | Group-Object -Property "UserId" | Select-Object @{n='UserId';e={$_.Name.TrimEnd()}}, @{n='Name';e={$_.Group[0].Name.TrimEnd()}}
    $users | ForEach-Object {
        $userid = $_.UserId
        if ($_.Name -match "\s") {
            $fn,$ln = $_.Name -split " "
            if ($fn -is [array]) { $fn = $fn[0] }
            $name = "{1}, {0}" -f $fn,$ln
            if ($ln -is [array]) { $name = $_.Name }
        } else {
            $name = $_.Name
        }
        $user = Get-ADUser -LDAPFilter "(anr=$($_.Name))" -Properties Mail | Where-Object { $_.Name -notmatch "\d\s|Admin|Test" } | Select-Object -First 1

        $account = [pscustomobject]@{
            SystemId = $userid
            System = $script:System
            AccountName = $userid
            Name = $name
            Email = $user.Mail
            Status = "Enabled"
            LastSeen = $script:LastSeen
            MailboxLocation = $null
        }
        $account
    }
}

function Select-Entitlements {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Account
    )

    BEGIN {
        $script:HardCodedDefs = Import-HardCodedProgramDefs
        $groupdefs =  Import-MainframeGroupDefs
        $opids = Import-OpIds
        $bcmsGtams = Import-BcmsGtams
        $entitlements = Import-Csv -Path $groupsPath -Header Group,UserId,Name | `
            Where-Object { $_.Group.StartsWith("BCMS") } | `
            Select-Object @{n='Group';e={$_.Group.TrimEnd()}}, @{n='UserId';e={$_.UserId.TrimEnd()}}, @{n='Name';e={$_.Name.TrimEnd()}} -OutVariable "g" | `
            Group-Object -Property "Group" | `
            Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={$script:System}}, @{n='Name';e={$_.Name}}, @{n='Description';e={$groupdefs[$_.Name]."Group Definition"}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Security Group"}} | `
            Save-EntitlementGroup -Passthru | `
            ForEach-Object -Begin {$e = @{}} -Process {$e.Add($_.Name, $_.EntitlementId)} -End {return $e}
        $userGroups = $g | Group-Object -Property "UserId" | ForEach-Object -Begin {$ug = @{}} -Process {$ug.Add($_.Name, $_.Group.Group)} -End {return $ug}
    }
    PROCESS {
        $acctEntitlements = [System.Collections.ArrayList]@()
        $opid = $opids[$Account.SystemId]
        if ($null -ne $opid) {
            if ($bcmsGtams.ContainsKey($opid)) {
                $gtamEntitlements = $bcmsGtams[$opid] | ForEach-Object -Begin {$g = @()} -Process {
                    if ($_.Gtam -match "B350") {
                        if ($_.BalanceByBatch -eq "Y") {
                            $g += [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System=$script:System; Name="B350 - Balance by Batch Total Amount"; Description=$null; LastSeen=$script:LastSeen; EntitlementType="Security Group"}
                        }
                        if ($_.DepositOption -eq "Y") {
                            $g += [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System=$script:System; Name="B350 - Deposit Option"; Description=$null; LastSeen=$script:LastSeen; EntitlementType="Security Group"}
                        }
                        if ($_.ApplyOption -eq "Y") {
                            $g += [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System=$script:System; Name="B350 - Apply Option"; Description=$null; LastSeen=$script:LastSeen; EntitlementType="Security Group"}
                        }
                    } elseif ($_.Gtam -match "B353") {
                        if ($_.AllocateCash -eq "Y") {
                            $g += [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System=$script:System; Name="B353 - Allocate Cash"; Description=$null; LastSeen=$script:LastSeen; EntitlementType="Security Group"}
                        }
                        if ($_.WriteOffAuthority -eq "Y") {
                            $g += [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System=$script:System; Name="B353 - Write-Off Authority"; Description=$null; LastSeen=$script:LastSeen; EntitlementType="Security Group"}
                        }
                        if ($_.WriteOffLimit -ne "") {
                            $g += [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System=$script:System; Name="B353 - Write-Off Limit = $($_.WriteOffLimit)"; Description=$null; LastSeen=$script:LastSeen; EntitlementType="Security Group"}
                        }
                        if ($_.DisbursementAuthority -eq "Y") {
                            $g += [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System=$script:System; Name="B353 - Disbursement Authority"; Description=$null; LastSeen=$script:LastSeen; EntitlementType="Security Group"}
                        }
                        if ($_.DisbursementLimit -ne "") {
                            $g += [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System=$script:System; Name="B353 - Disbursement Limit = $($_.DisbursementLimit)"; Description=$null; LastSeen=$script:LastSeen; EntitlementType="Security Group"}
                        }
                    } else {
                        $g += [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System=$script:System; Name=$_.Gtam; Description=$null; LastSeen=$script:LastSeen; EntitlementType="Security Group"}
                    }
                } -End { return $g }
                $acctEntitlements.AddRange(@($gtamEntitlements | Save-EntitlementGroup -Passthru | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$_.EntitlementId}}))
            }
        }

        $acctEntitlements.AddRange(@($userGroups[$Account.SystemId] | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$entitlements[$_]}}))
        $acctEntitlements
    }
}

function Start-Update {
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}

$script:System = "BCMS"
$groupsPath = "\\wbwpfil01\private`$\Information Security\Manager Access Reviews\Mainframe\GROUP"

Start-Update