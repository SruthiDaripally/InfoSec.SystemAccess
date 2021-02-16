function Select-Accounts {
    $usersPath = "$PSScriptRoot\..\Inputs\OneStream\OneStreamUserList.csv"
    $users = Get-Content -Path $usersPath | Select-Object -Skip 3 | ConvertFrom-Csv -Header @("Name","H1","Description","H2","H3","Enabled","UserName","Email","H4","H5","DateCreated")

    $users | Where-Object { ![string]::IsNullOrEmpty($_.Name) } | ForEach-Object {
        $account = [pscustomobject]@{
            SystemId = $_.Name
            System = $script:System
            AccountName = $_.UserName
            Name = $_.Name
            Email = $_.Email
            Status = switch ($_.Enabled) { "True" { "Enabled"; break; } "False" { "Disabled"; break; } }
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
        $entitlements = @{}
        $groupsPath = "$PSScriptRoot\..\Inputs\OneStream\OneStreamGroupList.csv"
        $lines = Get-Content -Path $groupsPath | Select-Object -Skip 3 | Select-Object -SkipLast 1

        # first pass to find the top-level groups
        :groupLoop foreach($line in $lines) {
            if ([regex]::IsMatch($line, "^,\w")) {
                $group = $line | ConvertFrom-Csv -Header @("H1","Name","H2","H3","Description")
                [pscustomobject]@{EntitlementId=$group.Name; System="$($script:System)"; Name=$group.Name; Description=$group.Description; LastSeen=$script:LastSeen; EntitlementType="Security Group"} | Save-EntitlementGroup | Out-Null
                $members = New-Object System.Collections.ArrayList
                $entitlements.Add($group.Name, $members) | Out-Null
                continue groupLoop
            }

            if ([regex]::IsMatch($line, "^,,(?!Group)\w")) {
                $child = $line | ConvertFrom-Csv -Header @("H1","H2","Name","Description")
                ([System.Collections.ArrayList]$entitlements[$group.Name]).Add("<<$($child.Name)") | Out-Null
                continue groupLoop
            }

            if ([regex]::IsMatch($line, "^,,,,,\w")) {
                $user = $line | ConvertFrom-Csv -Header @("H1","H2","H3","H4","H5","Name")
                ([System.Collections.ArrayList]$entitlements[$group.Name]).Add($user.Name) | Out-Null
            }
        }

        # seconds pass to replace child groups with members
        $copy = $entitlements.Clone()
        foreach ($entitlement in $copy.GetEnumerator()) {
            foreach ($member in $entitlement.Value) {
                if ($member.StartsWith("<<")) {
                    $childName = $member -replace "<<", ""
                    $entitlements[$entitlement.Key] = @(($entitlements[$entitlement.Key] | Where-Object { $_ -ne $member })) + $entitlements[$childName] | Get-Unique
                }
            }
        }
    }
    PROCESS {
        $entitlements.GetEnumerator() | Where-Object { $_.Value.Contains($Account.Name) } | ForEach-Object {
            $entitlement = [pscustomobject]@{
                SystemId = $Account.SystemId
                EntitlementId = $_.Key
            }
            $entitlement
        }
    }
}

function Start-Update {
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}

$script:System = "OneStream"
Start-Update