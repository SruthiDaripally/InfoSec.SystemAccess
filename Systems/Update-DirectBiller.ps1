function Select-Accounts {
    $inputData | Group-Object -Property "Username" | ForEach-Object {
        $account = [pscustomobject]@{
            SystemId=$_.Name
            System=$script:System
            AccountName=$_.Name
            Name=$_.Name
            Email=$_.Name
            Status="Enabled"
            LastSeen=$script:LastSeen
            MailboxLocation=$null
            Role = $_.Group[0].Role
            Entitlements = $_.Group
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
        $roles = $inputData | `
            Group-Object -Property "Role" | `
            Select-Object @{n='EntitlementId';e={$_.Name}}, @{n='System'; e={$script:System}}, @{n='Name';e={$_.Name}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Role"}} | `
            Save-EntitlementGroup -Passthru | `
            Select-Object @("EntitlementId", "Name") | `
            ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}
        $groups = $inputData | `
            Group-Object -Property "Description" | `
            Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System'; e={$script:System}}, @{n='Name';e={$_.Name -replace "'",""}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Security Group"}} | `
            Save-EntitlementGroup -Passthru | `
            Select-Object @("EntitlementId", "Name") | `
            ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}
    }
    PROCESS {
        $entitlements = @()
        $entitlements += [pscustomobject]@{SystemId=$Account.SystemId; EntitlementId=$roles[$Account.Role]}
        $entitlements += $Account.Entitlements | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$groups[$_.Description -replace "'",""]}}
        $entitlements
    }
}

function Start-Update {
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}

$script:System = "DirectBiller"
$inputPath = "$PSScriptRoot\..\Inputs\DirectBiller\UserEntitlementReport_DirectBiller.csv"
$prevHash = Get-Content -Path "$PSScriptRoot\DirectBiller.txt" -ErrorAction SilentlyContinue
$currentHash = Get-FileHash -Path $inputPath
if ($prevHash -ne $currentHash.Hash) {
    $inputData = Import-Csv -Path $inputPath
    Start-Update
    $currentHash.Hash | Out-File -FilePath "$PSScriptRoot\DirectBiller.txt"
}