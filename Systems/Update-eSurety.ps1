function Select-Accounts {
    $excelInput | ForEach-Object {
        $fn = $_."First Name"
        $ln = $_."Last Name"
        $uName = $_."User Name"
        $email = $_."Email Address"
        $status =$_."User Status"

        if ($null -ne $uName -and $null -ne $email) {
            $account = [PSCustomObject]@{
                SystemId = $uName
                System = $script:System
                AccountName = $uName
                Name = "{1}, {0}" -f $fn,$ln
                Email = $email
                Status =  switch ( $status ) {
                    "ACTIVE"    { "Enabled" }
                    "DISABLED"  { "Disabled" }
                  }
                LastSeen = $script:LastSeen
                MailboxLocation = $null
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
        #add roles
        $roles = $excelInput | Select-Object @("Name","Description") -Unique | Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_.Name}}, @{n='Description';e={$.Description}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Role"}} `
            | Save-EntitlementGroup -Passthru |`
            Select-Object @("EntitlementId","Name") | `
            ForEach-Object -Begin {$e =@{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}       
            
       

        #add secutity groups
       $CommericialAuthLevel =  $excelInput | Select-Object "Commercial Authority Level" -Unique | Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_."Commercial Authority Level"}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Commercial Authority Level"}} `
       | Save-EntitlementGroup -Passthru |`
       Select-Object @("EntitlementId","Name") | `
       ForEach-Object -Begin {$e =@{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}

        $ContractAuthLevel = $excelInput | Select-Object "Contract Authority Level" -Unique | Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_."Contract Authority Level"}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Contract Authority Level"}} `
        | Save-EntitlementGroup -Passthru |`
        Select-Object @("EntitlementId","Name") | `
        ForEach-Object -Begin {$e =@{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}

          $entitlements = $excelInput | Group-Object -Property "User Name" -AsHashTable 

    }
    PROCESS {
        $e = $entitlements[$Account.SystemId]
        $e | Select-Object @("Name","Description") -Unique | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$roles[$_."Name"]}}
        $e | Select-Object "Commercial Authority Level" -Unique | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$CommericialAuthLevel[$_."Commercial Authority Level"]}}
        $e | Select-Object "Contract Authority Level" -Unique | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$ContractAuthLevel[$_."Contract Authority Level"]}}
        
     }
}

function Start-Update {
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}

$script:System = "eSurety"
$inputPath = "$PSScriptRoot\..\Inputs\eSurety\eSurety_All_NonAgency_Users.xlsx"
$prevHash = Get-Content "$PSScriptRoot\eSurety.txt" -ErrorAction SilentlyContinue
$currentHash = Get-FileHash -Path $inputPath
if ($prevHash -ne $currentHash.Hash) {
    $excelInput = Import-Excel -Path $inputPath
    Start-Update
    $currentHash.Hash | Out-File -FilePath "$PSScriptRoot\eSurety.txt"
}