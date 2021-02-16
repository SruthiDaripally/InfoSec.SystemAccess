function Select-Accounts {
    $totalexcelInput | Group-Object -property "Login Name" | ForEach-Object {
        $uName =$_.Group[0]."Login Name"
        $fullName = $_.Group[0]."Full Name"
        $email = $_.Group[0]."Email Id"
        $status =$_.Group[0]."Account Disabled"

        if ($null -ne $uName) {
            $account = [PSCustomObject]@{
                SystemId = $uName
                System = $script:System
                AccountName = $uName
                Name = if ($fullName.Split(" ").count -gt 1 ) {
                    $fullName.Substring(0,$fullName.lastIndexOf(' '))
                } else {
                    $fullName
                }   
                Email = $email
                Status =  switch ( $status ) {
                    "NO"    { "Enabled" }
                    "YES"   { "Disabled" }
                  }
                LastSeen = $script:LastSeen
                MailboxLocation = $null
                Entitlements =$_.Group
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
        
      #add security group
      $entitlements =  $totalexcelInput | Select-Object "Security Group" -Unique | Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_."Security Group"}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Security Group"}}`
      | Save-EntitlementGroup -Passthru |`
      Select-Object @("EntitlementId", "Name") | `
      ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}

      #add privileges
     $privEntitlements =  $totalprivInput | Select-Object "Privilege Name" -Unique | Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_."Privilege Name"}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Privileges"}}`
      | Save-EntitlementGroup -Passthru |`
      Select-Object @("EntitlementId", "Name") | `
      ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}

      $privileges = $totalprivInput | Group-Object -Property "Security Group" -AsHashTable
    }

    PROCESS {
       
        $Account.Entitlements |  Select-Object "Security Group" -Unique |  Select-Object @{n ='SystemId'; e={$Account.SystemId}}, @{n ='EntitlementId';e={$entitlements[$_."Security Group"]}}
        $e = $privileges[$Account.Entitlements[0].'Security Group']
        $e  | Select-Object "Privilege Name" -Unique | Select-Object @{n='SystemId';e={$Account.SystemId}} , @{n='EntitlementId';e={$privEntitlements[$_."Privilege Name"]}}

    } 
    
}

function Start-Update {
    param (
        [Parameter(Mandatory =$true)]
        $envi
    )
   
    switch ($envi) {
        "Dev" {
            $script:System = "Informatica - Dev"
            $totalexcelInput = GetUserData -envi "Dev"
            $totalprivInput = GetPrivilegesData -envi "Dev"
            break
          }
        "QA" {
            $script:System = "Informatica - QA"
            $totalexcelInput = GetUserData -envi "QA"
            $totalprivInput = GetPrivilegesData -envi "QA"
            break
            }
        "PreProd"{
            $script:System = "Informatica - PreProd"
            $totalexcelInput = GetUserData -envi "PreProd"
            $totalprivInput = GetPrivilegesData -envi "PreProd"
            break
        }
        "Prod"{
            $script:System = "Informatica"
            $totalexcelInput = GetUserData -envi "Prod"
            $totalprivInput = GetPrivilegesData -envi "Prod"
            break
        }
        
    }
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}

function GetUserData {
    param (
        [Parameter(Mandatory =$true)]
        $envi
    )

    $inputPath ="$PSScriptRoot\..\Inputs\Informatica\$envi\"
    $files = Get-ChildItem -Path $inputPath -Depth 0 -Recurse -Filter "UsersPersonalInfo*"
    $totalexcelInput =@()
    if ($files.Count -gt 0) {
        foreach($file in $files)
        {            
           $excelInput = Import-Excel -path $file.FullName -StartRow 2
        #    if ($excelInput.Count -gt 0)
        #     {
                $sg = Import-Excel -path $file.FullName -StartRow 1 -EndRow 2 
                $securityGroup =[string]$sg
                $securityGroup = if ($securityGroup.Split(";").count -gt 0 ) {
                    $securityGroup.Split(";")[0].Remove(0,2)
                }

                foreach ($item in $excelInput) {
                    $item | Add-Member -MemberType NoteProperty  -Name "Security Group" -Value $securityGroup
                   }
            # }             
            $totalexcelInput += $excelInput
           
        }
    }
    return $totalexcelInput
        
}


function GetPrivilegesData {
    param (
        [Parameter(Mandatory =$true)]
        $envi
    )

    $inputPath ="$PSScriptRoot\..\Inputs\Informatica\$envi\"
    $privilegeFiles = Get-ChildItem -Path $inputPath -Depth 0 -Recurse -Filter "PrivilegesAssociation*"
    $totalprivInput = @()
    if ($privilegeFiles.Count -gt 0) {
        foreach ($file in $privilegeFiles) {
            $privInput = Import-Excel -Path $file.FullName -StartRow 2
            # if ($privInput.Count -gt 0)
            # {
                $sg = Import-Excel -path $file.FullName -StartRow 1 -EndRow 2 
                $securityGroup =[string]$sg
                $securityGroup = if ($securityGroup.Split(";").count -gt 0 ) {
                    $securityGroup.Split(";")[0].Remove(0,2)
                }

                foreach ($item in $privInput) {
                    $item | Add-Member -MemberType NoteProperty  -Name "Security Group" -Value $securityGroup
                   }
            # }             
            $totalprivInput += $privInput
        }
        
    }
    return $totalprivInput
    
}

Start-Update -envi "Dev"
Start-Update -envi "QA"
Start-Update -envi "PreProd"
Start-Update -envi "Prod"