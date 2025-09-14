function Export-GxMessageHeaders {
    param(
        [Parameter(Mandatory=$true)][string]$SenderEmail,
        [Parameter(Mandatory=$true)][string]$RecipientEmail,
        [Parameter(Mandatory=$true)][string]$DelegateEmail,
        [datetime]$StartDate = (Get-Date).AddDays(-10),
        [datetime]$EndDate = (Get-Date),
        [int]$First = 5
    )
    $path = 'C:\temp'
    $timestamp = [System.DateTime]::Now.ToString('yyyyMMddHHmmss', [System.Globalization.CultureInfo]::InvariantCulture)
    $newFolder = 'EmailHeaders', $timestamp, ($SenderEmail -split '@')[0],($RecipientEmail -split '@')[0] -join "_"
    $newFolderPath = [System.IO.Path]::Combine($path, $newFolder)
    Get-MgGraphConnection -Scopes @('Mail.Read')
    Invoke-Operation -ScriptBlock { [void][System.IO.Directory]::CreateDirectory($newFolderPath) } -ErrorMessage "CreateDirectory failed."
    $messageTrace = Invoke-Operation -ScriptBlock { Get-ExoMessageTrace -SenderEmail $SenderEmail -RecipientEmail $RecipientEmail -StartDate $StartDate -EndDate $EndDate -First $First } -ErrorMessage "Get-ExoMessageTrace failed."
    Invoke-Operation -ScriptBlock { Set-ExoMailboxPermissions -OwnersEmail $RecipientEmail -DelegatesEmail $DelegateEmail -Remove:$false } -ErrorMessage "Set-ExoMailboxPermissions failed."
    $traceQuarantined = $messageTrace | ? {$_.Status -eq 'Quarantined'}
    $traceDelivered = $messageTrace | ? {$_.Status -eq 'Delivered'}
    if ($traceQuarantined) {
        $traceQuarantined | ForEach-Object {
            $quarantinedMessage = Invoke-Operation -ScriptBlock { Get-QuarantineMessage -MessageId $_.MessageId -IncludeMessagesFromBlockedSenderAddress } -ErrorMessage "Get-QuarantineMessage failed."
            $quarantinedId = $quarantinedMessage.Identity.Split('\')[-1]
            $receivedTime = $quarantinedMessage.ReceivedTime.ToString('yyyyMMdd_HHmmss', [System.Globalization.CultureInfo]::InvariantCulture)
            $newFile = "Quarantine_$($receivedTime)_$($quarantinedId)"
            $newTxt = [System.IO.Path]::Combine($newFolderPath, "$newFile.headers.txt")
            $newEml = [System.IO.Path]::Combine($newFolderPath, "$newFile.eml")
            $qh = Invoke-Operation -ScriptBlock { (Get-QuarantineMessageHeader -Identity $quarantinedMessage.Identity).Header } -ErrorMessage "Get-QuarantineMessageHeader failed."
            $qh | Out-File $newTxt -Encoding ascii
            $export = Invoke-Operation -ScriptBlock { Export-QuarantineMessage -Identity $quarantinedMessage.Identity } -ErrorMessage "Export-QuarantineMessage failed."
            [IO.File]::WriteAllBytes($newEml, [Convert]::FromBase64String($export.eml))
        }
    }
    if ($traceDelivered) {
        $traceDelivered | ForEach-Object {
            $user = $_.RecipientAddress
            $internetId = ($_.MessageId -replace '\s+', '')
            $receivedDate = $_.Received.ToString('yyyyMMdd_HHmmss', [System.Globalization.CultureInfo]::InvariantCulture)
            $newFile = "Delivered_$($receivedDate)_$($internetId.Replace('<', '').Replace('>', ''))"
            if ($internetId -notmatch '^<.*>$') { $internetId = "<$internetId>" }
            $msg = Invoke-Operation -ScriptBlock { Get-MgUserMessage -UserId $user -Filter "internetMessageId eq '$($internetId.Replace("'", "''"))'" -Top 1 | Select * } -ErrorMessage "Get-MgUserMessage failed."
            if ($msg) {
                $newTxt = [System.IO.Path]::Combine($newFolderPath, "$newFile.headers.txt")
                $newEml = [System.IO.Path]::Combine($newFolderPath, "$newFile.eml")
                Invoke-Operation -ScriptBlock { Get-MgUserMessageContent -UserId $user -MessageId $msg.Id -OutFile $newEml } -ErrorMessage "Get-MgUserMessageContent failed."
                $headers = New-Object System.Collections.Generic.List[string]
                foreach ($line in [System.IO.File]::ReadLines($newEml)) { if ([string]::IsNullOrEmpty($line)) { break } $headers.Add($line) }
                $headers | Out-File $newTxt -Encoding ascii
            }
        }
    }
    explorer.exe $newFolderPath
}
