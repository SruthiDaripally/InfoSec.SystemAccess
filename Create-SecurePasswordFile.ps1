$script:ScriptDir = &{
    $invocation = (Get-Variable MyInvocation -Scope 1).Value
    Split-Path $invocation.MyCommand.Path
}

Add-Type -Assembly System.Security

$password = Read-Host -Prompt "Enter password"
$pwdBytes = [System.Security.Cryptography.ProtectedData]::Protect($password.ToCharArray(), $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
Set-Content -LiteralPath "$script:ScriptDir\db2-PROD.txt" -Value ([Convert]::ToBase64String($pwdBytes))