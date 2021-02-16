param(
    [Parameter(Mandatory=$false)]
    [string]$Filter = $null
)

$script:LastSeen = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

function Resolve-Error ($ErrorRecord=$Error[0])
{
    $ErrorRecord | Format-List * -Force
    $ErrorRecord.InvocationInfo | Format-List *
    $Exception = $ErrorRecord.Exception
    for ($i = 0; $Exception; $i++, ($Exception = $Exception.InnerException))
    {
        "$i" * 80
        $Exception | Format-List * -Force
    }
}

# source functions
$functionsPath = Join-Path -Path $PSScriptRoot -ChildPath "Functions"
if (Test-Path $functionsPath) {
	$functions = Get-ChildItem -Path $functionsPath -Filter "*.ps1"
	foreach($function in $functions) {
		. $($function.FullName)
	}
}

# source systems
$systemsPath = Join-Path -Path $PSScriptRoot -ChildPath "Systems"
if (Test-Path $systemsPath) {
	$systems = Get-ChildItem -Path $systemsPath -Filter "*.ps1"
	foreach($s in $systems) {
        if ($null -eq $Filter -or ($null -ne $Filter -and $s.FullName -match $Filter))
        {
            try {
                . $($s.FullName)
            } catch [System.Exception] {
                Resolve-Error $_ | Out-File -FilePath "$PSScriptRoot\update-log.txt" -Append
            }
        }
	}
}