function Select-Accounts {
    $userInput | Group-Object -property "User ID" | ForEach-Object {
        $fn = $_.Group[0]."First Name"
        $ln = $_.Group[0]."Last Name"
        $uId = $_.Group[0]."User ID"
        $status =$_.Group[0]."Status"
        # $lastAccessTime =$_.Group[0]."Last Access Time"

        if ($null -ne $uId) {
            $account = [PSCustomObject]@{
                SystemId = $uId
                System = $script:System
                AccountName = $uId
                Name = "{1}, {0}" -f $fn,$ln
                Email = $uId
                Status =  switch ( $status.ToUpper() ) {
                    "ACTIVE"    { "Enabled" }
                    "INACTIVE"  { "Disabled" }
                  }
                LastSeen =$script:LastSeen                
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
     #add security group
      $entitlements =  $totalPer | Select-Object "Security Group" -Unique | Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_."Security Group"}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Security Group"}}`
      | Save-EntitlementGroup -Passthru |`
      Select-Object @("EntitlementId", "Name") | `
      ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}
     
        
      #add permissions
      $privEntitlements =  $totalPer | Select-Object "Permission" -Unique | Select-Object @{n='EntitlementId';e={[guid]::NewGuid().ToString()}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_."Permission"}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Permissions"}}`
      | Save-EntitlementGroup -Passthru |`
      Select-Object @("EntitlementId", "Name") | `
      ForEach-Object -Begin {$e = @{}} -Process {$e[$_.Name] = $_.EntitlementId} -End {return $e}     

      $privileges = $totalPer | Group-Object -Property "Security Group" -AsHashTable
    }

    PROCESS {
      
        $Account.Entitlements |  Select-Object "Permission Group" -Unique |  Select-Object @{n ='SystemId'; e={$Account.SystemId}}, @{n ='EntitlementId';e={$entitlements[$_."Permission Group"]}}
        $e = $privileges[$Account.Entitlements[0].'Permission Group']
        $e  | Select-Object "Permission" -Unique | Select-Object @{n='SystemId';e={$Account.SystemId}} , @{n='EntitlementId';e={$privEntitlements[$_."Permission"]}}

    } 
    
}

function Get-Permissions {
    param (
        [Parameter(Mandatory =$true)]
        $groupFile
    )     
        $totalPer =@()   
        $perInput = Import-Csv -Path $groupFile.FullName -Header GroupID,Description,Status,EnabledPermission
        $perInput | Group-Object -Property "GroupID" | ForEach-Object {
            $group =$_.Group[0]."GroupID"
            $permissions =[string]$_.Group[0]."EnabledPermission" 
            if ($null -ne $group -and $null -ne $permissions) {           
                if ($permissions.Split("|").count -gt 0) {
                    $res = $permissions.Split("|")               
                        foreach ($item in $res) {                        
                            $totalPer += new-object psobject -Property @{
                                "Permission" = $item
                                "Security Group" = $group                            
                            }                      
                            
                        }                
                }
                
            }
        }
        
      $totalPer         
}

function Start-Update {    
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}



$script:System = "FIS - BillerDirect"
$fisPath ="\\wbwpfil05\infa_shared_prod\DataFeeds\FIS\Mailbox_FIS\inbox\"
$todaysDate= Get-Date -Format "yyyyMMdd"
$groupFile = Get-ChildItem -Path $fisPath  -Filter "WBM_Group_List_$todaysDate*" -Depth 0 

if ($groupFile.Count -gt 0) {     
$totalPer = Get-Permissions -groupFile $groupFile
}

$userFile = Get-ChildItem -Path $fisPath  -Filter "WBM_User_List_$todaysDate*" -Depth 0 
if ($userFile.Count -gt 0) {       
$userInput = Import-Csv -Path $userFile.FullName

Start-Update    

#backup the file to input folder
$groupFileName =$groupFile.Name
$groupDestination = "$PSScriptRoot\..\Inputs\FIS - BillerDirect\$groupFileName"

$userFileName =$userFile.Name
$userDestination = "$PSScriptRoot\..\Inputs\FIS - BillerDirect\$userFileName"


# Move both group and user files to input folder  after processing  
if (Test-Path $groupFile.FullName) {    
    Move-Item -Path $groupFile.FullName -Destination  $groupDestination
} 
else {
    throw "$groupFile was not found"
}   

if (Test-Path $userFile.FullName) {
    Move-Item -Path $userFile.FullName -Destination  $userDestination
}
else {
    throw "$userFile was not found"
}


}



