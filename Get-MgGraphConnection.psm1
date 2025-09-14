# Invoke-Operation

function Get-MgGraphConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string[]] $Scopes,

        [Parameter(Mandatory = $false, Position = 1)]
        [string] $UserPrincipalName
    )

    Invoke-Operation -ScriptBlock {
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
            throw "Microsoft Graph module is not installed. Run: Install-Module Microsoft.Graph"
        }
    } -ErrorMessage "Module check failed."

    $context = Get-MgContext

    $canReuse = $false
    if ($context) {
        $isDelegated     = ($context.AuthType -eq 'Delegated')
        $accountMatches  = if ($UserPrincipalName) { $context.Account -eq $UserPrincipalName } else { $true }
        $currentScopes   = @($context.Scopes)
        $hasAllScopes    = (@($Scopes | Where-Object { $_ -in $currentScopes }).Count -eq $Scopes.Count)

        $canReuse = ($isDelegated -and $accountMatches -and $hasAllScopes)
    }

    if ($canReuse) { return }

    $requiredScopes = if ($context -and $context.Scopes) {
        @($context.Scopes + $Scopes | Select-Object -Unique)
    } else {
        $Scopes
    }

    if ($context) {
        Disconnect-MgGraph -Confirm:$false
    }
    Connect-MgGraph -NoWelcome -Scopes $requiredScopes
}
