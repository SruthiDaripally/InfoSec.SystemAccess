function Import-UserDetailReport {
    $pkg = Open-ExcelPackage -Path $inputPath
    $sheet = $pkg."User Detail Report"
    $start = $sheet.Dimension.Start
    $end = $sheet.Dimension.End

    for ($row = $start.Row + 1; $row -le $end.Row; $row++) {
        [pscustomobject]@{
            UserID = $sheet.Cells[$row,1].Text
            UserName = $sheet.Cells[$row,2].Text
            Email = $sheet.Cells[$row,3].Text
            Service = $sheet.Cells[$row,4].Text
            Entitlement = $sheet.Cells[$row,5].Text
        }        
    }
}

function Select-Accounts {
    $report | Group-Object -Property "UserID" | ForEach-Object {
        $login = $_.Name
        $name = $_.Group[0]."UserName"
        $email = $_.Group[0]."Email"

        $account = [pscustomobject]@{
            SystemId = $login
            System = $script:System
            AccountName = $login
            Name = $name
            Email = $email
            Status = "Enabled"
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
        $entitlements = $report | 
            Select-Object @("Service", "Entitlement") | `
            Sort-Object @("Service", "Entitlement") -Unique | `
            Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={$script:System}}, @{n='Name';e={$_.Service}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={$_.Entitlement}} | `
            Save-EntitlementGroup -Passthru | `
            ForEach-Object -Begin {$e = @{}} -Process {$key = [System.Tuple]::Create($_.Name,$_.EntitlementType); $e.Add($key, $_.EntitlementId)} -End {return $e}
    }
    PROCESS {
        $Account.Entitlements | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$key = [System.Tuple]::Create($_.Service,$_.Entitlement); $entitlements[$key]}} | Where-Object { $null -ne $_.EntitlementId }
    }
}

function Start-Update {
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}

$script:System = "BMO"
$inputPath = "$PSScriptRoot\..\Inputs\BMO\BMO User Access Report.xlsx"
$prevHash = Get-Content -Path "$PSScriptRoot\BMO.txt" -ErrorAction SilentlyContinue
$currentHash = Get-FileHash -Path $inputPath
if ($prevHash -ne $currentHash.Hash) {
    $report = Import-UserDetailReport
    Start-Update
    $currentHash.Hash | Out-File -FilePath "$PSScriptRoot\BMO.txt"
}