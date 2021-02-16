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

   $msalParams = @{
     ClientId = $azureClientId
     TenantId = $azureTenantId
     ClientSecret = (ConvertTo-SecureString '3llZcy51t9W1AO_9_.Dn7~uEl67.CGZ6y6' -AsPlainText -Force)
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
       if ($mailItem.DateTimeReceived.Date.Year -eq (Get-Date).Year -and $mailItem.DateTimeReceived.Date.Day -eq (Get-Date).Day) {
           foreach($Attachment in $MailItem.Attachments){
               $Attachment.Load()
               $File = new-object System.IO.FileStream(($filePath + $Attachment.Name.ToString()),[System.IO.FileMode]::Create)
               $File.Write($attachment.Content, 0, $attachment.Content.Length)
               $File.Close()
            }
          #   The item or folder will be moved to the mailbox's Deleted Items folder.
            $mailItem.Delete(2)
       }
               
    }    
      
}