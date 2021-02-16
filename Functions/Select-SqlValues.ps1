function Select-SqlValues {
    param(
        [Parameter(ValueFromPipeline=$true)]
        $InputObject
    )

    PROCESS {
        if ($null -eq $_) {
            "null"
        } else {
            "'{0}'" -f ($_ -replace "'", "''")
        }
    }
}