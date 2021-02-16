Add-Type -Assembly System.Security
Set-Variable -Name ConnectionUri -Option Const -Value "http://wiwpexc01.wbmi.com/powershell"

function Get-ExchangeCredential {
    $crypted = [Convert]::FromBase64String((Get-Content -LiteralPath "$PSScriptRoot\svc.txt"))
    $clear = [System.Security.Cryptography.ProtectedData]::Unprotect($crypted, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
    $enc = [System.Text.Encoding]::Default
    $pass = $enc.GetString($clear)
    $securePwd = ConvertTo-SecureString -AsPlainText -Force -String $pass
    $creds = New-Object System.Management.Automation.PSCredential("wbmi\svtPapp_svc",$securePwd)
    $creds
}

function Select-Recipients {
    $exchangeCredential = Get-ExchangeCredential
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Authentication Kerberos -Credential $exchangeCredential
    $recipients = Invoke-Command -Session $session -ScriptBlock {
        Get-Recipient -ResultSize Unlimited -RecipientType UserMailbox, MailUser | Select-Object SamAccountName,PrimarySmtpAddress,RecipientType
    }
    $output = @{}
    $recipients | ForEach-Object { 
        if (-not($output.ContainsKey($_.PrimarySmtpAddress.ToString().ToLower()))) {
            $output.Add($_.PrimarySmtpAddress.ToString().ToLower(), $_) 
        }
    }
    $output
}

function Select-Mailboxes {
    $exchangeCredential = Get-ExchangeCredential
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Authentication Kerberos -Credential $exchangeCredential
    $mailboxes = Invoke-Command -Session $session -ScriptBlock {
        Get-Recipient -ResultSize Unlimited -RecipientType UserMailbox | Get-Mailbox
    }
    $mailboxes
}

function Select-O365Mailboxes {
    Add-Type -AssemblyName System.Security
    Import-Module -Name ExchangeOnlineManagement -MinimumVersion 2.0.3
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

    if ($Environment -eq "Prod") {
        Connect-ExchangeOnline -CertificateThumbprint "86E9DFD491C906AAAD9C32FE3A670FAD54523170" -AppID "e1b8bf1d-2ebd-462e-8193-62b6abe7926e" -Organization "wbmi.onmicrosoft.com" -ShowBanner:$false
    } else {
        Connect-ExchangeOnline -ShowBanner:$false
    }

    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties IsMailboxEnabled -IncludeInactiveMailbox

    Disconnect-ExchangeOnline -Confirm:$false
    $mailboxes
}

function Update-SystemAccess {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )
    BEGIN {
        $recipients = Select-Recipients
    }
    PROCESS {
        $status = switch ($_.IsMailboxEnabled) { $true { "Enabled"; break } $false { "Disabled"; break } }
        $recipient = $recipients[$InputObject.PrimarySmtpAddress.Trim().ToLower()]
        if ($null -ne $recipient) {
            [pscustomobject]@{
                SystemId = $_.Guid
                System = "Exchange"
                AccountName = $recipient.SamAccountName
                Name = $_.Name -replace "'", "''"
                Email = $_.PrimarySmtpAddress
                Status = $status
                LastSeen = $script:LastSeen
                MailboxLocation = switch ($recipient.RecipientType.ToString()) { "UserMailbox" { "Exchange"; break; } "MailUser" { "Exchange Online"; break; } }
            }
        }
    }
    END {
    }
}

function Start-Update {
    ,@(Select-O365Mailboxes;Select-Mailboxes) | ForEach-Object { $_ | Update-SystemAccess | Save-Account }
}

$script:System = "Exchange"
Start-Update