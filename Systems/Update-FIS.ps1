enum Mode {
    User = 1
    UserSettings = 2
    Companies = 3
    Processes = 4
    Ignored = 5
}

function Select-Mode {
    param(
        [Parameter(Mandatory=$true)]
        $Line
    )

    switch -regex ($Line) {
        'Permissions Report Run:' { return [Mode]::User }
        'User Settings' { return [Mode]::UserSettings }
        'Has Access To The Following Companies:' { return [Mode]::Companies }
        'Has Access To The Following Processes:' { return [Mode]::Processes }
        'Does NOT have Access To' { return [Mode]::Ignored }
    }
}

function Select-Accounts {
    $usersInput = Import-Excel -Path "$PSScriptRoot\..\Inputs\FIS\FIS_Investments_Users.xlsx"
    $currentMode = $null
    :inputLoop foreach ($line in $inputData) {
        if ($line -notmatch "\w") { continue :inputLoop }
        $mode = Select-Mode -Line $line
        if ($null -eq $mode) {
            switch ($currentMode) {
                ([Mode]::User) { 
                    if ($account) { $account }
                    $name = $line | Select-String -Pattern "\\(\w+)|(^\w+$)" | `
                        ForEach-Object { $_.Matches.Groups | Where-Object { $_.Success } | Select-Object -Last 1 -ExpandProperty Value }
                    $user = $usersInput | Where-Object { $_."User Name" -eq $name }
                    $account = [pscustomobject]@{
                        SystemId = $name
                        System = $script:System
                        AccountName = $name
                        Name = if ($null -ne $user) { $user."Display Name" } else { $name }
                        Email = if ($null -ne $user) { $user.Email } else { $null }
                        Status = $null
                        LastSeen = $script:LastSeen
                        MailboxLocation = $null
                        Entitlements = @{
                            Companies = New-Object System.Collections.ArrayList
                            Processes = New-Object System.Collections.ArrayList
                        }
                    }
                    break
                }
                ([Mode]::UserSettings) {
                    if ($line -match "Account Status:") {
                        if ($line -match "Enabled") { $account.Status = "Enabled" } else { $account.Status = "Disabled" }
                    }
                    break
                }
                ([Mode]::Companies) { $account.Entitlements.Companies.Add($line.Trim()) | Out-Null; break; }
                ([Mode]::Processes) { $account.Entitlements.Processes.Add($line.Trim()) | Out-Null; break; }
            }
        } else {
            $currentMode = $mode
        }
    }

    $account
}

function Select-Entitlements {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Account
    )

    BEGIN {
        $c = $script:Companies | Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Company"}} | `
            Save-EntitlementGroup -Passthru | `
            Select-Object @("EntitlementId", "Name") | `
            ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}
        $p = $script:Processes | Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Process"}} | `
            Save-EntitlementGroup -Passthru | `
            Select-Object @("EntitlementId", "Name") | `
            ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}
    }
    PROCESS {
        $Account.Entitlements.Companies | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$c[$_]}}
        $Account.Entitlements.Processes | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$p[$_]}}
    }
}

function Start-Update {
    $accounts = Select-Accounts | Save-Account -Passthru
    $script:Companies = $accounts.Entitlements.Companies | Select-Object -Unique
    $script:Processes = $accounts.Entitlements.Processes | Select-Object -Unique
    $accounts | Select-Entitlements | Save-Entitlement
}

$script:Companies = @()
$script:Processes = @()
$script:System = "FIS"
$inputData = Get-Content "$PSScriptRoot\..\Inputs\FIS\UserPerm.txt"
$prevHash = Get-Content -Path "$PSScriptRoot\FIS.txt" -ErrorAction SilentlyContinue
$currentHash = Get-FileHash -Path $inputData
if ($prevHash -ne $currentHash.Hash) {
    Start-Update
    $currentHash.Hash | Out-File -FilePath "$PSScriptRoot\FIS.txt"
}