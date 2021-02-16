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

function Import-ClaimsGtams {
    $claimsGtamsPath = "\\wbmpfil01\public$\IT General\Associate Setup\GTAMS\Claims GTAMS\CLAIMS GTAMS.xlsx"
    if (-not(Test-Path $claimsGtamsPath))
    {
        throw "$claimsGtamsPath wasn't found"
    }

    $x = Open-ExcelPackage -Path $claimsGtamsPath
    $lookup = @{}
    1..$x.Workbook.Worksheets.Count | ForEach-Object {
        $sheet = $x.Workbook.Worksheets.Item($_)
        if ($sheet.Name -eq "Revision History") { return } 
        $gtamDesc = $Script:HardCodedDefs[$sheet.Name]
        if ($gtamDesc) {
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
            if ($lookup.ContainsKey($newrow.OpID)) {
                ([System.Collections.ArrayList]$lookup[$newrow.OpID]).Add((New-Object PSCustomObject -Property $newrow)) | Out-Null
            } else {
                $lookup.Add($newrow.OpID, [System.Collections.ArrayList]@(New-Object PSCustomObject -Property $newrow))
            }
        }
    }
    $x | Close-ExcelPackage -NoSave
    Remove-Variable x
    $lookup
}

function Import-DataSets {
    $datasetPath = "\\wbwpfil01\private$\Information Security\Manager Access Reviews\Mainframe\DATASET"
    if (-not(Test-Path $datasetPath)) {
        throw "$datasetPath wasn't found"
    }
    $lookup = @{}
    $csv = Import-Csv -Path $datasetPath -Header UserId,Name,Level,Profile
    $profileLevels = $csv | Select-Object -Property @("Level", "Profile") -Unique
    $csv | ForEach-Object {
        if ($lookup.ContainsKey($_.UserId)) {
            ([System.Collections.ArrayList]$lookup[$_.UserId]).Add($_) | Out-Null
        } else {
            $lookup.Add($_.UserId, [System.Collections.ArrayList]@($_))
        }
    }
    [pscustomobject]@{
        ByUser = $lookup
        ProfileLevels = $profileLevels
    }
}

function Import-Genres {
    $genrePath = "\\wbwpfil01\private$\Information Security\Manager Access Reviews\Mainframe\GENRES"
    if (-not(Test-Path $genrePath)) {
        throw "$genrePath wasn't found"
    }
    $lookup = @{}
    $csv = Import-Csv -Path $genrePath -Header UserId,Name,Level,Class,Profile
    $csv | ForEach-Object {
        if ($lookup.ContainsKey($_.UserId)) {
            ([System.Collections.ArrayList]$lookup[$_.UserId]).Add($_) | Out-Null
        } else {
            $lookup.Add($_.UserId, [System.Collections.ArrayList]@($_))
        }
    }
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

function Import-HardCodedPrograms {
    $progspath = "\\wbwpfil01\public$\IT General\Associate Setup\Security Hard Coded Programs.xlsx"
    if (-not(Test-Path -Path $progspath))
    {
        throw "$progspath wasn't found"
    }
    $x = Open-ExcelPackage -Path $progspath
    $sheet = $x.Workbook.Worksheets["Hard Coded Security Info"]
    $lookup = @{}
    $start = $sheet.Dimension.Start
    $end = $sheet.Dimension.End
    for ($row = $start.Row + 1; $row -le $end.Row; $row++) {
        $opid = ($sheet.Cells[$row, 1]).Text.ToLower().Trim()
        if ($opid) {
            $groups = @()
            for ($col = 6; $col -le $sheet.Dimension.Columns; $col++) {
                $name = $sheet.Cells[1, $col].Text.Trim()
                $value = $sheet.Cells[$row, $col].Text.Trim()
                if ($value -eq "X") {
                    $groups += $name
                }
            }

            $user = [pscustomobject]@{
                OpId = $opid
                Groups = $groups
            }
            if ($lookup.ContainsKey($opid)) {
                ([System.Collections.ArrayList]$lookup[$opid]).Add($user) | Out-Null
            } else {
                $lookup.Add($opid, [System.Collections.ArrayList]@($user))
            }
        }
    }
    $x | Close-ExcelPackage -NoSave
    Remove-Variable x
    $lookup
}

function Select-Accounts {
    $users = Import-Csv -Path $groupsPath -Header Group,UserId,Name | Group-Object -Property "UserId" | Select-Object @{n='UserId';e={$_.Name.TrimEnd()}}, @{n='Name';e={$_.Group[0].Name.TrimEnd()}}
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
        $claimsGtams = Import-ClaimsGtams
        $datasets = Import-DataSets
        $genres = Import-Genres
        $progs = Import-HardCodedPrograms
        $entitlements = Import-Csv -Path $groupsPath -Header Group,UserId,Name | `
            Where-Object { -not($_.Group.StartsWith("BCMS")) } | `
            Select-Object @{n='Group';e={$_.Group.TrimEnd()}}, @{n='UserId';e={$_.UserId.TrimEnd()}}, @{n='Name';e={$_.Name.TrimEnd()}} -OutVariable "g" | `
            Group-Object -Property "Group" | `
            Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={$script:System}}, @{n='Name';e={$_.Name}}, @{n='Description';e={$groupdefs[$_.Name]."Group Definition"}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Security Group"}} | `
            Save-EntitlementGroup -Passthru | `
            ForEach-Object -Begin {$e = @{}} -Process {$e.Add($_.Name, $_.EntitlementId)} -End {return $e}
        $datasets.ProfileLevels | `
            Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={$script:System}}, @{n='Name';e={"$($_.Level) - $($_.Profile)"}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Security Group"}} | `
            Save-EntitlementGroup -Passthru | `
            ForEach-Object -Process {$entitlements.Add($_.Name, $_.EntitlementId)}
        $userGroups = $g | Group-Object -Property "UserId" | ForEach-Object -Begin {$ug = @{}} -Process {$ug.Add($_.Name, $_.Group.Group)} -End {return $ug}
    }
    PROCESS {
        $acctEntitlements = [System.Collections.ArrayList]@()
        $groups = [System.Collections.ArrayList]@()
        $opid = $opids[$Account.SystemId]
        if ($null -ne $opid) {
            if ($claimsGtams.ContainsKey($opid)) {
                $groups.AddRange(@($claimsGtams[$opid] | ForEach-Object {
                    [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System=$script:System; Name=$_.Gtam; Description=$null; LastSeen=$script:LastSeen; EntitlementType="Security Group"}
                }))
            }
            if ($progs.ContainsKey($opid)) {
                $groups.AddRange(@($progs[$opid] | ForEach-Object {
                    $_.Groups | ForEach-Object {
                        [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System=$script:System; Name=$_; Description=$null; LastSeen=$script:LastSeen; EntitlementType="Security Group"}
                    }
                }))
            }
        }

        $acctEntitlements.AddRange(@($datasets.ByUser[$Account.SystemId] | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$entitlements["$($_.Level) - $($_.Profile)"]}}))
        $groups.AddRange(@($genres[$Account.SystemId] | Where-Object { $null -ne $_.Level -and $null -ne $_.Profile } | ForEach-Object {
            [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System=$script:System; Name="$($_.Level) - $($_.Profile)"; Description=$null; LastSeen=$script:LastSeen; EntitlementType="Security Group"}
        }))
        $acctEntitlements.AddRange(@($groups | Save-EntitlementGroup -Passthru | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$_.EntitlementId}}))

        $acctEntitlements.AddRange(@($userGroups[$Account.SystemId] | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$entitlements[$_]}}))

        $acctEntitlements
    }
}

function Start-Update {
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}

$script:System = "PMS"
$groupsPath = "\\wbwpfil01\private`$\Information Security\Manager Access Reviews\Mainframe\GROUP"

Start-Update