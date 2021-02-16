function Select-Accounts {
    $excelInput | ForEach-Object {
        $login = $_."Login ID"
        $fn = $_."First Name"
        $ln = $_."Last Name"
        $email = $_."Email"

        if ($null -ne $login -and $null -ne $email) {
            $account = [pscustomobject]@{
                SystemId = $login
                System = $script:System
                AccountName = $login
                Name = "{1}, {0}" -f $fn,$ln
                Email = $email
                Status = "Enabled"
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
        # make sure entitlement groups exist
        [pscustomobject]@{EntitlementId="Appraiser"; System="CCC"; Name="Appraiser"; Description=""; LastSeen=$script:LastSeen; EntitlementType="Security Group"} | Save-EntitlementGroup | Out-Null
        [pscustomobject]@{EntitlementId="PortalAccess"; System="CCC"; Name="Portal Access"; Description=""; LastSeen=$script:LastSeen; EntitlementType="Security Group"} | Save-EntitlementGroup | Out-Null
        $appraisers = @()
        $excelInput | Where-Object { -not([string]::IsNullOrEmpty($_."Appraiser ID")) } | ForEach-Object { $appraisers += $_."Login ID" }
    }
    PROCESS {
        $isAppraiser = $appraisers.Contains($Account.AccountName)
        if ($isAppraiser) {
            $entitlement = [pscustomobject]@{
                SystemId = $Account.SystemId
                EntitlementId = "Appraiser"
            }
            $entitlement
        }

        $entitlement = [pscustomobject]@{
            SystemId = $Account.SystemId
            EntitlementId = "PortalAccess"
        }
        $entitlement
    }
}

function Start-Update {
    Select-Accounts | Save-Account -Passthru | Select-Entitlements | Save-Entitlement
}

$script:System = "CCC"
$inputPath = "$PSScriptRoot\..\Inputs\CCC\WestBend_All Active Users.xlsx"
$prevHash = Get-Content -Path "$PSScriptRoot\CCC.txt" -ErrorAction SilentlyContinue
$currentHash = Get-FileHash -Path $inputPath
if ($prevHash -ne $currentHash.Hash) {
    $excelInput = Import-Excel -Path $inputPath
    Start-Update
    $currentHash.Hash | Out-File -FilePath "$PSScriptRoot\CCC.txt"
}