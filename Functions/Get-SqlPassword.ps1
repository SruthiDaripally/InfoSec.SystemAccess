Add-Type -AssemblyName System.Security

function Get-SqlPassword {
    $crypted = [Convert]::FromBase64String((Get-Content -LiteralPath "$PSScriptRoot\sql.txt"))
    $clear = [System.Security.Cryptography.ProtectedData]::Unprotect($crypted, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
    $enc = [System.Text.Encoding]::Default
    $pass = $enc.GetString($clear)
    $pass
}