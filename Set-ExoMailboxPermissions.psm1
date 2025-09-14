function Set-ExoMailboxPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$OwnersEmail,
        [Parameter(Mandatory)][string]$DelegatesEmail,
        [Parameter()][Switch]$Remove,

        [Parameter()][Switch]$GroupMailbox,
        [Parameter()][ValidateSet('None','All','Children','Descendents','SelfAndChildren')]
        [string]$InheritanceType = 'All',
        [Parameter()][bool]$AutoMapping = $true,
        [Parameter()][ValidateSet('ChangeOwner','ChangePermission','DeleteItem','ExternalAccount','FullAccess','ReadPermission')]
        [string[]]$AccessRights = @('FullAccess')
    )

    begin {
        Invoke-Operation -ScriptBlock { Get-ExoConnection } -ErrorMessage "Failed to connect to Exchange Online."
    }

    process {
        try {
            $owner = Invoke-Operation `
                -ScriptBlock { Get-User -Identity $OwnersEmail -ErrorAction Stop } `
                -ErrorMessage "User lookup failed for owner '$OwnersEmail'."

            $delegate = Invoke-Operation `
                -ScriptBlock { Get-User -Identity $DelegatesEmail -ErrorAction Stop } `
                -ErrorMessage "User lookup failed for delegate '$DelegatesEmail'."
        }
        catch {
            Write-Error $_
            return $false
        }

        $mailboxId    = $owner.UserPrincipalName
        $emailEscaped = [regex]::Escape($DelegatesEmail)

        $getPermsSplat = @{
            Identity    = $mailboxId
            ErrorAction = 'Stop'
        }
        if ($GroupMailbox) { $getPermsSplat['GroupMailbox'] = $true }

        try {
            $currentPerms = Invoke-Operation `
                -ScriptBlock { Get-MailboxPermission @getPermsSplat } `
                -ErrorMessage "Failed to get mailbox permissions for $mailboxId."
        }
        catch {
            Write-Error $_
            return $false
        }

        $entriesForUser = $currentPerms | Where-Object { $_.User -match $emailEscaped }

        if ($Remove) {
            $explicitEntries = $entriesForUser | Where-Object { -not $_.IsInherited }

            if (-not $explicitEntries) {
                if ($entriesForUser) {
                    Write-Host "ℹ️  All matching rights for $DelegatesEmail are inherited; nothing can be removed directly here."
                } else {
                    Write-Host "ℹ️  No permissions found for $DelegatesEmail; nothing to remove."
                }
            }
            else {
                Write-Host "ℹ️  Found $($explicitEntries.Count) permission entr$(if($explicitEntries.Count -eq 1){'y'}else{'ies'}) for $DelegatesEmail — removing all explicit rights..."
                foreach ($perm in $explicitEntries) {
                    $removeSplat = @{
                        Identity     = $mailboxId
                        User         = $DelegatesEmail
                        AccessRights = @($perm.AccessRights)
                        Confirm      = $false
                        ErrorAction  = 'Stop'
                    }
                    if ($GroupMailbox) { $removeSplat['GroupMailbox'] = $true }
                    if ($perm.Deny -eq $true) { $removeSplat['Deny'] = $true }

                    try {
                        Invoke-Operation `
                            -ScriptBlock { Remove-MailboxPermission @removeSplat } `
                            -ErrorMessage "❌ Failed to remove rights [$(@($perm.AccessRights) -join ', ')] for $DelegatesEmail —"
                        Write-Host "✅ Removed rights [$(@($perm.AccessRights) -join ', ')] for $DelegatesEmail."
                    }
                    catch {
                        Write-Error $_
                    }
                }
            }

            try {
                $currentPerms = Invoke-Operation `
                    -ScriptBlock { Get-MailboxPermission @getPermsSplat } `
                    -ErrorMessage "Failed to refresh mailbox permissions for $mailboxId."
            }
            catch {
                Write-Error $_
                return $false
            }
            return $currentPerms
        }
        else {
            $hasAllRequested = $entriesForUser | Where-Object {
                $entryRights = @($_.AccessRights)
                (@($AccessRights | Where-Object { $_ -in $entryRights }).Count -eq $AccessRights.Count)
            }

            if ($hasAllRequested) {
                Write-Host "✅ $DelegatesEmail already has [$($AccessRights -join ', ')] on $mailboxId."
                return $currentPerms
            }
            else {
                Write-Host "ℹ️  $DelegatesEmail does NOT have [$($AccessRights -join ', ')] — granting now..."

                $addSplat = @{
                    Identity        = $mailboxId
                    User            = $DelegatesEmail
                    AccessRights    = $AccessRights
                    InheritanceType = $InheritanceType
                    AutoMapping     = $AutoMapping
                    ErrorAction     = 'Stop'
                }
                if ($GroupMailbox) { $addSplat['GroupMailbox'] = $true }

                try {
                    Invoke-Operation `
                        -ScriptBlock { Add-MailboxPermission @addSplat } `
                        -ErrorMessage "❌ Failed to grant [$($AccessRights -join ', ')] to $DelegatesEmail —"
                    Write-Host "✅ Granted [$($AccessRights -join ', ')] to $DelegatesEmail (InheritanceType=$InheritanceType, AutoMapping=$AutoMapping)."
                }
                catch {
                    Write-Error $_
                }

                try {
                    return Invoke-Operation `
                        -ScriptBlock { Get-MailboxPermission @getPermsSplat } `
                        -ErrorMessage "Failed to refresh mailbox permissions for $mailboxId."
                }
                catch {
                    Write-Error $_
                    return $false
                }
            }
        }
    }

    end {
        # Disconnect-ExchangeOnline -Confirm:$false
    }
}

# Set-ExoMailboxPermissions -OwnersEmail 'email1@domain.com' -DelegatesEmail 'email2@domain.com' -Remove:$false
# Set-ExoMailboxPermissions -OwnersEmail 'email1@domain.com' -DelegatesEmail 'email2@domain.com' -Remove:$false -InheritanceType None -AutoMapping:$false
