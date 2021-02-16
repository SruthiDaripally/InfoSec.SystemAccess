function Select-Accounts {
    $excelInput | Group-Object -Property "Employee ID"  | ForEach-Object {
        $userId =$_.Group[0]."Employee ID"
        $email = $_.Group[0]."Email Address"
        $fullName =$_.Group[0]."Employee Name"
        $status = $_.Group[0].Active

        if ($null -ne $userId) {
            $account = [pscustomobject]@{
                SystemId = $userId
                System = $script:System
                AccountName = $email
                Name = $fullName
                Email = $email
                Status = switch($status){
                    "Y"  {"Enabled"}
                    "N"  {"Disabled"}
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
        #add roles
        $excelInput | Select-Object "Role" -Unique | Select-Object @{n='EntitlementId';e={$_.Role}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_.Role}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Role"}} `
        | Save-EntitlementGroup | Out-Null

    #add Department - Code
     $excelInput | Select-Object "Department - Code" -Unique | Select-Object @{n='EntitlementId';e={$_."Department - Code"}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_."Department - Code"}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Department"}} `
        | Save-EntitlementGroup | Out-Null 

    #add Cost Center - Code
     $excelInput | Select-Object "Cost Center - Code" -Unique | Select-Object @{n='EntitlementId';e={$_."Cost Center - Code"}}, @{n='System';e={"$($script:System)"}}, @{n='Name'; e={$_."Cost Center - Code"}}, @{n='Description';e={$null}}, @{n='LastSeen';e={$script:LastSeen}}, @{n='EntitlementType';e={"Cost Center"}} `
        | Save-EntitlementGroup | Out-Null
    }
    PROCESS {

        $Account.Entitlements | Select-Object "Role" -Unique | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$_.Role}}
        $Account.Entitlements | Select-Object "Department - Code" -Unique | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$_."Department - Code"}}
        $Account.Entitlements | Select-Object "Cost Center - Code" -Unique | Select-Object @{n='SystemId';e={$Account.SystemId}}, @{n='EntitlementId';e={$_."Cost Center - Code"}}
    }
}

function Start-Update {
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}

function ReadMailFromConnectO365{

    $azureClientId ="f4235f05-bc87-4385-b06f-fbd85ead05f5"
    $azureTenantId ="48b0431c-82f6-4ad2-a023-ac96dbf5614e"  

    $sharedMailbox ="Saviynt@WBMI.com"   
    $subject ="Concur User Report"
    $emailsender="autonotification@concursolutions.com"

    # Set the path to your copy of EWS Managed API 
    $dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll" 

    # Load the Assembly 
    [void][Reflection.Assembly]::LoadFile($dllpath) 

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

    # Create a new Exchange service object 
    $service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService ([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)

    $Service.Url = 'https://outlook.office365.com/EWS/Exchange.asmx'
    $Service.UseDefaultCredentials = $false

    $pwdFile = "$PSScriptRoot\ews.txt"
    $crypted = [Convert]::FromBase64String((Get-Content -LiteralPath $pwdFile))
    $clear = [System.Security.Cryptography.ProtectedData]::Unprotect($crypted, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
    $enc = [System.Text.Encoding]::Default
    $secret = $enc.GetString($clear)

    $msalParams = @{
      ClientId = $azureClientId
      TenantId = $azureTenantId
      ClientSecret = (ConvertTo-SecureString $secret -AsPlainText -Force)
      Scopes   = "https://outlook.office.com/.default"
  }

    $token = Get-MsalToken @msalParams
    $Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$token.AccessToken

    $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList 2, $sharedMailbox
  
     
     $inboxfolderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$sharedMailbox)
     $inboxfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$inboxfolderid) 
     
    $sfsubject = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, $subject)
    $sfsender = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender, $emailsender)
   # $sfdateReceived= new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::d, $startdatetime)
     $sfattachment = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
     $sfcollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
     $sfcollection.add($sfattachment)
     $sfcollection.Add($sfsubject)
     $sfcollection.Add($sfsender)
         # $sfcollection.Add($sfdateReceived)

     $view = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1     

     $MailItems = $inboxfolder.FindItems($sfcollection, $view)
     $filePath = "$PSScriptRoot\..\Inputs\Concur\"

     foreach ($mailItem in $MailItems){        
        $mailItem.Load()
        # Process only todays file
        # if ($mailItem.DateTimeReceived.Date.Year -eq (Get-Date).Year -and $mailItem.DateTimeReceived.Date.Day -eq (Get-Date).Day) {
            foreach($Attachment in $MailItem.Attachments){
                $Attachment.Load()
                $File = new-object System.IO.FileStream(($filePath + $Attachment.Name.ToString()),[System.IO.FileMode]::Create)
                $File.Write($attachment.Content, 0, $attachment.Content.Length)
                $File.Close()
             }
             $mailItem.IsRead =$true
           #   The item or folder will be moved to the mailbox's Deleted Items folder.
             $mailItem.Delete(2)
        # }
                
     }    
       
}

$script:System = "Concur"
ReadMailFromConnectO365
$inputPath = "$PSScriptRoot\..\Inputs\Concur\Concur User Report - Operations (Daily).xlsx"
$prevHash = Get-Content -Path "$PSScriptRoot\Concur.txt" -ErrorAction SilentlyContinue
$currentHash = Get-FileHash -Path $inputPath
if ($prevHash -ne $currentHash.Hash) {
    $excelInput = Import-Excel -Path $inputPath -StartRow 2
    $excelInput = $excelInput[0..($excelInput.Count-2)]
    Start-Update
    $currentHash.Hash | Out-File -FilePath "$PSScriptRoot\Concur.txt"
}