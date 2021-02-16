function Select-Accounts {
    $excelInput | Group-Object -Property "User.ID" | Where-Object { $null -ne $_.Group[0]."Login.Email Address" } | ForEach-Object {
        $fn = $_.Group[0]."Login.First Name"
        $ln = $_.Group[0]."Login.Last Name"
        $email = $_.Group[0]."Login.Email Address"
        $status = $_.Group[0]."User.Status"
        $account = [pscustomobject]@{
            SystemId = $_.Name
            System = $script:System
            AccountName = $email
            Name = "{1}, {0}" -f $fn,$ln
            Email = $email
            Status = $status
            LastSeen = $script:LastSeen
            MailboxLocation = $null
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
        $entitlements = $excelInput | `
            Where-Object { $null -ne $_."Security Group" } | `
            Group-Object -Property "Security Group" | `
            Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System'; e={$script:System}}, @{n='Name';e={$_.Name}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Security Group"}} | `
            Save-EntitlementGroup -Passthru | `
            Select-Object @("EntitlementId", "Name") | `
            ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}
    }
    PROCESS {
        $Account.Entitlements | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$entitlements[$_."Security Group"]}}
    }
}

function Start-Update {
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}

$script:System = "Equian"
$inputPath = "$PSScriptRoot\..\Inputs\Equian\West Bend User List.xlsx"
$prevHash = Get-Content -Path "$PSScriptRoot\Equian.txt" -ErrorAction SilentlyContinue
$currentHash = Get-FileHash -Path $inputPath
if ($prevHash -ne $currentHash.Hash) {
    $excelInput = Import-Excel -Path $inputPath -StartRow 6
    Start-Update
    $currentHash.Hash | Out-File -FilePath "$PSScriptRoot\Equian.txt"
}