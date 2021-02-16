using namespace System.Text.RegularExpressions

enum Mode {
    User = 1
    UserType = 2
    UserName = 3
    NetworkAcct = 4
    Functions = 5
}

$lockedRegex = [regex]::new('Locked:\s+(\w+)', [RegexOptions]::Compiled)
$userIdRegex = [regex]::new('USER:\s+(\w+)', [RegexOptions]::Compiled)
$userNameRegex = [regex]::new('User Name:\s+(.*?)\s\s', [RegexOptions]::Compiled)
$networkAcctRegex = [regex]::new('Network Acct:\s+([\w\\]+)', [RegexOptions]::Compiled)
$wbmiRegex = [regex]::new('^wbmi\\(\w+)', [RegexOptions]::Compiled)
$entTypeRegex = [regex]::new('^\s+([^:$]+)$', [RegexOptions]::Compiled)
$entRegex = [regex]::new('^YES\s+([^$]+)', [RegexOptions]::Compiled)

function Select-CompiledRegexValue {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Value,

        [Parameter(Mandatory=$true)]
        [regex]$Pattern
    )

    $m = $Pattern.Matches($Value)
    $m.Groups | Where-Object { $_.Success } | Select-Object -Last 1 -ExpandProperty Value
}

function New-Account {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Line
    )

    $locked = $line | Select-CompiledRegexValue -Pattern $lockedRegex
    $account = [pscustomobject]@{
        SystemId = $null
        System = $script:System
        AccountName = $null
        Name = $null
        Email = $null
        Status = if ($locked -eq "Yes") { "Disabled" } else { "Enabled" }
        LastSeen = $script:LastSeen
        MailboxLocation = $null
        Entitlements = @{}
    }
    $account
}

function Select-Accounts {
    :inputLoop foreach ($operator in $operators) {
        $account = New-Account -Line $operator.Line
        $account.SystemId = $account.AccountName = $lines[$operator.LineNumber - 3] | Select-CompiledRegexValue -Pattern $userIdRegex
        $account.Name = $lines[$operator.LineNumber + 1] | Select-CompiledRegexValue -Pattern $userNameRegex
        $networkAcct = $lines[$operator.LineNumber + 11] | Select-CompiledRegexValue -Pattern $networkAcctRegex
        if ($null -ne $networkAcct -and $networkAcct -match $wbmiRegex) {
            try { $user = Get-ADUser $Matches[1] -Properties Mail } catch { }
            if ($user) { $account.Email = $user.Mail }
        }

        $funcStart = $operator.LineNumber + 18
        for ($f=$funcStart+1;$f -lt $lines.Count; $f++) {
            if ($lines[$f] -match "^Companies:") {
                $funcEnd = $f
                break
            }
        }

        for ($i=$funcStart;$i -lt $funcEnd;$i++) {
            if ($lines[$i] -match '\S') {
                switch -regex ($lines[$i]) {
                    $entTypeRegex {
                        $entType = $lines[$i] | Select-CompiledRegexValue -Pattern $entTypeRegex
                        if ($account -and !$account.Entitlements.ContainsKey($entType)) {
                            $account.Entitlements.Add($entType, (New-Object System.Collections.ArrayList)) | Out-Null
                        }
                        break
                    }
                    $entRegex {
                        if ($account) {
                            $entitlement = $lines[$i] | Select-CompiledRegexValue -Pattern $entRegex
                            ($account.Entitlements[$entType]).Add($entitlement) | Out-Null
                        }
                        break
                    }
                }
            }
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
        $e = @{}
        $script:Functions.GetEnumerator() | ForEach-Object {
            $type = $_.Key
            $_.Value | Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={$type}} | `
                Save-EntitlementGroup -Passthru | `
                Select-Object @("EntitlementId", "Name") | `
                ForEach-Object { $e.Add([System.Tuple]::Create($type, $_.Name), $_.EntitlementId) }
        }
    }
    PROCESS {
        $Account.Entitlements.GetEnumerator() | ForEach-Object {
            $type = $_.Key
            $_.Value | ForEach-Object {
                $t = [System.Tuple]::Create($type,$_)
                $_ | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$e[$t]}}
            }
        }
    }
}

function Start-Update {
    $accounts = Select-Accounts | Save-Account -Passthru
    $script:Functions = @{}
    $types = $accounts.Entitlements.Keys | Sort-Object -Unique
    foreach ($type in $types) {
        $functions = $accounts.Entitlements | ForEach-Object { $_[$type] } | Sort-Object -Unique
        $script:Functions.Add($type, $functions)
    }
    $accounts | Select-Entitlements | Save-Entitlement
}

$script:System = "EAS"
$path = "$PSScriptRoot\..\Inputs\EAS\ListUser.log.txt"
$prevHash = Get-Content -Path "$PSScriptRoot\EAS.txt" -ErrorAction SilentlyContinue
$currentHash = Get-FileHash -Path $path
if ($prevHash -ne $currentHash.Hash) {
    
    $lines = Get-Content $path
    # find the first operator
    $operators = $lines | Select-String -Pattern "User Type:    OPERATOR"

    Start-Update
    $currentHash.Hash | Out-File -FilePath "$PSScriptRoot\EAS.txt"
}

