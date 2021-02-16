function Select-Accounts {
    $excelInput | Group-Object -property "User ID"| ForEach-Object {
        $userId = $_.Group[0]."User ID"
        $email = $_.Group[0]."Email Address"
        $status =$_.Group[0]."Status"

        if ($null -ne $userId) {
            $account = [PSCustomObject]@{
                SystemId = [guid]::NewGuid().ToString()
                System = $script:System
                AccountName = $userId
                Name = $userId
                Email = $email
                Status =  switch ( $status ) {
                    "ACTIVE"    { "Enabled" }
                    "DISABLED"  { "Disabled" }
                  }
                LastSeen = $script:LastSeen
                MailboxLocation = $null
                Entitlements = $_.Group
            }
            $account
        }
    }
}

function Select-Entitlements {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Account
    )   

    BEGIN {
        # make sure entitlement groups exist
        $controlTotals = [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System="BMO - ACH Fraud"; Name="Control Totals"; Description=""; LastSeen=$script:LastSeen; EntitlementType="Security Group"}`
         | Save-EntitlementGroup -Passthru |`
         Select-Object @("EntitlementId", "Name") | `
         ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}

        $achPP = [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System="BMO - ACH Fraud"; Name="ACH PP"; Description=""; LastSeen=$script:LastSeen; EntitlementType="Security Group"}`
        | Save-EntitlementGroup -Passthru |`
        Select-Object @("EntitlementId", "Name") | `
        ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}

        $achPPApproval =[pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System="BMO - ACH Fraud"; Name="ACH PP Approval"; Description=""; LastSeen=$script:LastSeen; EntitlementType="Security Group"}`
        | Save-EntitlementGroup -Passthru |`
        Select-Object @("EntitlementId", "Name") | `
        ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}

        $adminRoles = [pscustomobject]@{EntitlementId=[guid]::NewGuid().ToString(); System="BMO - ACH Fraud"; Name="Admin Roles"; Description=""; LastSeen=$script:LastSeen; EntitlementType="Security Group"}`
        | Save-EntitlementGroup -Passthru |`
        Select-Object @("EntitlementId", "Name") | `
        ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}
       

    }
    PROCESS {
        $e = $Account.Entitlements

        $sysId = Get-AccountSystemId $Account.AccountName
        
        $a =$e."ACH PP "
        $b =$e."ACH PP Approval "
        $c =$e."Admin Roles"
        $d =$e."Control Totals"
        
        $e |  Where-Object {$null -ne $d -and  $d.ToUpper() -eq "X"} |  Select-Object @{n='SystemId';e={$sysId}}, @{n='EntitlementId';e={$controlTotals["Control Totals"]}}
        $e |  Where-Object {$null -ne $a -and  $a.ToUpper() -eq "X"} |  Select-Object @{n='SystemId';e={$sysId}}, @{n='EntitlementId';e={$achPP["ACH PP"]}}
        $e |  Where-Object {$null -ne $b -and  $b.ToUpper() -eq "X"} |Select-Object @{n='SystemId';e={$sysId}}, @{n='EntitlementId';e={$achPPApproval["ACH PP Approval"]}}
        $e |  Where-Object {$null -ne $c -and  $c.ToUpper() -eq "X" } |Select-Object @{n='SystemId';e={$sysId}}, @{n='EntitlementId';e={$adminRoles["Admin Roles"]}}
                     
     }
}

function Start-Update {
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}

$script:System = "BMO - ACH Fraud"
$inputPath = "$PSScriptRoot\..\Inputs\BMO - ACH Fraud\WBMI ACH PP.xlsx"
$prevHash = Get-Content "$PSScriptRoot\BMO-ACHFraud.txt" -ErrorAction SilentlyContinue
$currentHash = Get-FileHash -Path $inputPath
if ($prevHash -ne $currentHash.Hash) {
    $excelInput = Import-Excel -Path $inputPath
    $excelInput = $excelInput[0..($excelInput.Count-4)]
    Start-Update
    $currentHash.Hash | Out-File -FilePath "$PSScriptRoot\BMO-ACHFraud.txt"
}