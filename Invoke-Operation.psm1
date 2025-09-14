function Invoke-Operation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock] $ScriptBlock,
        [string] $ErrorMessage = "Operation failed."
    )
    try {
        & $ScriptBlock
    } catch {
        throw "$ErrorMessage $($_.Exception.Message)"
    }
}
