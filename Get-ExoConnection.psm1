# Invoke-Operation

function Get-ExoConnection {
    [CmdletBinding()]
    param(
        [string] $UserPrincipalName,
        [switch] $ShowBanner,
        [switch] $UseRPSSession
    )

    Invoke-Operation -ScriptBlock {
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            throw "ExchangeOnlineManagement module is not installed. Run: Install-Module ExchangeOnlineManagement"
        }
    } -ErrorMessage "Module check failed."

    $active = Invoke-Operation -ScriptBlock {
        try {
            Get-ConnectionInformation 2>$null | Where-Object {
                $_.State -eq 'Connected' -and
                $_.TokenStatus -eq 'Active' -and
                -not $_.IsEopSession
            }
        } catch {
            @()
        }
    } -ErrorMessage "Failed to check connection state."

    if ($UserPrincipalName) { $active = $active | Where-Object { $_.UserPrincipalName -eq $UserPrincipalName } }
    if ($active.Count -gt 0) {return}

    $connectSplat = @{
        ShowBanner = [bool]$ShowBanner
    }
    if ($PSBoundParameters.ContainsKey('UserPrincipalName')) { $connectSplat.UserPrincipalName = $UserPrincipalName }
    if ($PSBoundParameters.ContainsKey('UseRPSSession') -and $UseRPSSession.IsPresent) { $connectSplat.UseRPSSession = $true }

    Disconnect-ExchangeOnline -Confirm:$false
    
    Invoke-Operation -ScriptBlock {
        Connect-ExchangeOnline @connectSplat
    } -ErrorMessage "Failed to connect to Exchange Online."
}
