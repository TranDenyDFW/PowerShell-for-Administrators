function Get-ExoMessageTrace {
    [CmdletBinding(DefaultParameterSetName='BySender')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='BySender')]
        [Parameter(Mandatory=$true, ParameterSetName='ByBoth')]
        [string] $SenderEmail,

        [Parameter(Mandatory=$true, ParameterSetName='ByRecipient')]
        [Parameter(Mandatory=$true, ParameterSetName='ByBoth')]
        [string] $RecipientEmail,

        [Parameter()]
        [datetime] $StartDate,

        [Parameter()]
        [datetime] $EndDate,

        [Parameter()]
        [ValidateRange(1, [int]::MaxValue)]
        [int] $First
    )

    Invoke-Operation -ScriptBlock { Get-ExoConnection }

    $now = Get-Date

    if (-not $PSBoundParameters.ContainsKey('StartDate')) { $StartDate = $now.AddDays(-10) }
    if (-not $PSBoundParameters.ContainsKey('EndDate'))   { $EndDate   = $now }

    if ($StartDate -gt $EndDate) {
        throw [System.ArgumentException]::new("StartDate ($StartDate) cannot be later than EndDate ($EndDate).")
    }

    $params = @{
        StartDate = $StartDate
        EndDate   = $EndDate
    }
    if ($PSBoundParameters.ContainsKey('SenderEmail'))    { $params.SenderAddress    = $SenderEmail }
    if ($PSBoundParameters.ContainsKey('RecipientEmail')) { $params.RecipientAddress = $RecipientEmail }

    $results = Get-MessageTraceV2 @params

    if ($PSBoundParameters.ContainsKey('First')) {
        $results | Select-Object -First $First -Property *
    } else {
        $results | Select-Object *
    }
}

# Get-ExoMessageTrace -SenderEmail 'email@outlook.com' -RecipientEmail 'email@domain.com'
# Get-ExoMessageTrace -SenderEmail 'email@outlook.com' -RecipientEmail 'email@domain.com' -StartDate '9/8/2025 1:35:21 AM' -EndDate '9/9/2025 1:40:21 AM' 
