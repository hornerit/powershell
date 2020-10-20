#REQUIRES -Version 5 -Modules @{ModuleName="ExchangeOnlineManagement"; ModuleVersion="0.4368.1"}
<#
.SYNOPSIS
Performs a message trace for a spam message and searches for and purges it from recipient mailboxes

.DESCRIPTION
Connects to O365 - assuming you have Exchange Admin permissions - and performs a Get-MessageTrace.
From there, it will take all of the recipients, split them up separate Content Searches with the
action of purging all items.

.PARAMETER EmailDomain
REQUIRED Domain of the mailboxes affected by spam campaign (used for message trace search/filter). Defaults to
contoso.com and will prompt if you don't supply yours.
.PARAMETER NoMFA
OPTIONAL If the account you wish to use is enabled for basic auth and doesn't have expiring tokens, use this
switch to operate without MFA; otherwise, it will expect to use MFA and modern exchange.
.PARAMETER AccountLockdownScriptName
OPTIONAL If you wish to run a script to lockdown sender(s) that are part of the EmailDomain as a part of this
process (and your script has a $Users string parameter that can split based on commas), supply the full name of
the script (e.g. Secure-Account-Manually.ps1) and add the script to the same folder as this one
.PARAMETER Mailboxes2Exclude
OPTIONAL If there are certain mailbox address to exclude (e.g. an on-prem mailbox that cannot be managed by the
O365 Compliance and Security Center), supply them to this switch to ignore them in the attempts to fix things.

.NOTES
  Created by: Brendan Horner (www.hornerit.com)
  Notes: MUST BE RUN AS SCRIPT FILE, do NOT copy-paste into PS to run
  Version History:
  --2020-10-20-Converted to use Content Search, mirrored from new GUI version, and adjusted some code style
  --2019-07-16-Added feature: filters for email status
  --2019-07-15-Fixed bug for child windows again for MFA parameter
  --2019-06-27-Fixed bug for child windows due to changing MFA parameter to NoMFA and updated MFA Exchange Module
        to use the latest version
  --2019-06-19-Altered MFA parameter to be NoMFA so someone can force basic auth by setting that switch and
        adjusted MFA module to pull the latest version of the module on your machine
  --2019-05-28-Bug Fixes for MFA and errors in mailbox
  --2019-05-22-Added support for Exchange Online MFA Module
  --2019-05-21-Completed documentation and separate version of script that has a GUI, see other post for that
        version (https://www.hornerit.com/2019/05/o365-spam-remover-script-now-with-gui.html).
  --2019-05-16-Rewrote sections for dynamic window generation based on params, allows as many Exch Admin accts to
        assist as you can try...watch out for RAM usage
  --2019-05-02-Added better logic for throttling
  --2019-04-15-Initial public version

.EXAMPLE
.\O365-SPAM-REMOVER.ps1 -NoMFA
.\O365-SPAM-REMOVER.ps1 -NoMFA -AccountLockdownScriptName "Secure-Account.ps1"
#>
[CmdletBinding()]
param(
    [string]$EmailDomain = "contoso.com",
    [switch]$NoMFA,
    [string]$AccountLockdownScriptName,
    [string[]]$Mailboxes2Exclude
)

#Try to get the Exchange Online Powershell module that supports MFA
Import-Module ExchangeOnlineManagement
do {
    $Good = 0
    if ($NoMFA) {
        try {
            $Cred = Get-Credential -Message "Please enter exchange admin EMAIL ADDRESS...EMAIL" -ErrorAction Stop
            if ($Cred.Password.Length -eq 0) {
                throw
            }
        } catch {
            Write-Host "Error with your credential input, please try again or re-run without the -NoMFA switch"
            Continue
        }
        try {
            Connect-ExchangeOnline -Credential $Cred -ShowBanner:$False -ErrorAction Stop
        } catch {
            Read-Host "Error connecting to Exchange Online - $_. Press any key to exit."
            Exit
        }
        try {
            Get-OrganizationConfig | Select-Object Name
            $Good = 1
            $CredUPN = $Cred.UserName
        } catch {
            Write-Host "Supplied credential is not an Exchange Admin."
            Disconnect-ExchangeOnline -confirm:$false
        }
    } else {
        $CredUPN = Read-Host "Please enter an Exchange Admin email address"
        if ($CredUPN -match "^.+@.+\..+$") {
            try {
                Connect-ExchangeOnline -UserPrincipalName $CredUPN -ShowBanner:$False -ErrorAction Stop
            } catch {
                Read-Host "Error connecting to Exchange Online - $_. Press any key to exit."
                Exit
            }
            try {
                Get-OrganizationConfig | Select-Object Name
                $Good = 1
            } catch {
                Write-Host "Supplied credential is not an Exchange Admin."
                Disconnect-ExchangeOnline -confirm:$false
            }
        } else {
            Write-Host "Did not supply an email address, try again."
        }
    }
} until ($Good -eq 1)
$SendersHash = @{}
do {
    $Prompt = "[Required]Email address of evil sender. This prompt will repeat until you press enter with " +
        "no information. Do not enter quotes or empty, extra spaces"
    $BadSender = Read-Host $Prompt
    do {
        #Validate that each email address at least matches the typical email pattern
        $SendersGood = 1
        if ($BadSender.length -gt 0) {
            if (!($BadSender -match ".*@.*\..*")) {
                $SendersGood = 0
                $BadSender = Read-Host "Bad email address input, try again"
            } else {
                $SendersHash.add($BadSender,$null)
            }
        }
    } until ($SendersGood -eq 1)
} until ($BadSender.length -eq 0)
$Senders = $SendersHash.keys -join ","

#Create list of users to lock down if script is supplied and found
if ($AccountLockdownScriptName.Length -gt 0 -and
    (Test-Path $PSScriptRoot\$AccountLockdownScriptName) -and
    $Senders.IndexOf("@$EmailDomain") -gt 0) {
        $LockdownAddresses = (($SendersHash.Keys |
            Where-Object { $_ -match "^.+@$($EmailDomain.replace(".","\."))"}) -join ",")
        $Prompt = "Would you also like to run the lockdown script on these users in the " +
            $EmailDomain + ' domain?' + "`n$LockdownAddresses"
        if ((Read-Host $Prompt) -eq "y") {
            $Users2Lockdown = $LockdownAddresses
        }
}

#Get the start date and time of the range when that users should have received the spam
$Prompt = "[Required]Start Date AND TIME for range when users received message (e.g. 7/18/2018 12:20 AM)"
$StartDate = Read-Host $Prompt
do {
    $good = 0
    try {
        $StartDate = Get-Date($StartDate)
        $good = 1
    }
    catch {
        $StartDate = Read-Host "Start date invalid, try again"
    }
} until ($good -eq 1)
$SearchStartDate = (get-date $StartDate).ToUniversalTime()

#Get the end date and time of the range that users should have received the spam
$Prompt = "[Required]End Date AND TIME for range when users received message (e.g. 7/18/2018 3:59 PM)"
$EndDate = Read-Host $Prompt
do {
    $good = 0
    try {
        $EndDate = Get-Date($EndDate)
        if ($EndDate -gt $StartDate) { $good = 1 } else { throw }
    }
    catch {
        $EndDate = Read-Host "End date invalid or before start date, try again"
    }
} until ($good -eq 1)
$SearchEndDate = (get-date $EndDate).ToUniversalTime()

#Get the subject line(s) of the evil messages to filter
$SubjectLineFilter = @{}
$Prompt = "[Optional]Subject line to filter. This prompt will repeat until you press enter with no " +
    "information. Do not enter any quotes or backticks unless actually in the subject"
do {
    $SubjectLine = Read-Host $Prompt
    if ($SubjectLine.length -gt 0) {
        $SubjectLineFilter.add($SubjectLine,$null)
    }
} until ($SubjectLine.length -eq 0)
$SubjectLineStr = $null
if ($SubjectLineFilter.count -gt 0) {
    $SubjectLineStr = '(Subject:'
    foreach($subject in $SubjectLineFilter.keys) {
        $SubjectLineStr += '"'+$subject.Replace('"','""')+'" OR Subject:'
    }
    $SubjectLineStr = $SubjectLineStr.TrimEnd(" OR Subject:")
    $SubjectLineStr += ")"
}

#Get the desired Email status - Delivered, FilteredAsSpam, or both
$EmailStatusFilter = @("Pending")
$Prompt = "[Required]Status of emails being processed: `n[1]Delivered, [2]FilteredAsSpam, [3]Both`n" +
    "(Default is 3):"
do {
    try {
        $StatusInput = [int](Read-Host $Prompt -ErrorAction Stop)
        switch ($StatusInput) {
            1 { $EmailStatusFilter = "Pending","Delivered" }
            2 { $EmailStatusFilter = "Pending","FilteredAsSpam" }
            Default { $EmailStatusFilter = "Pending","FilteredAsSpam","Delivered" }
        }
        $Good = 1
    } catch {Write-Host "Invalid Response"}
} until ($Good -eq 1)

#Build Search Query from the SearchStartDate, SearchEndDate, subject line(s), and evil sender(s)

$SearchQuery = "(Received:`""+$SearchStartDate.toString()+".."+$SearchEndDate.toString()+"`") AND "
$SenderList = New-Object System.Collections.ArrayList
#If more than one sender, create filter string with a bunch of ORs for senders; otherwise, just set one
if ($Senders.IndexOf(",") -gt 0) {
    $SearchQuery += "("
    foreach($EvilSender in $Senders.Split(",")) {
        $addy = $EvilSender.Trim()
        $SearchQuery += "From:$addy OR "
        $SenderList.add($addy) | Out-Null
    }
    $SearchQuery = $SearchQuery.TrimEnd(" OR ")
    $SearchQuery += ")"
} else {
    $SearchQuery += "From:$Senders"
    $SenderList.add($Senders) | Out-Null
}
#If a subject line is specified, add it to the search query
if ($null -ne $SubjectLineStr) {
    $SearchQuery += " AND "+$SubjectLineStr
}

if ($Users2Lockdown.length -gt 0) {
    $Msg = "Starting script located at $PSScriptRoot\$AccountLockdownScriptName with the users $Users2Lockdown " +
        "and waiting until it is complete to continue."
    Write-Host $Msg
    $ArgString = "-file `"$("$PSScriptRoot\$AccountLockdownScriptName")`" -Users `"$Users2Lockdown`""
    Start-Process powershell -Passthru -Wait -ArgumentList $ArgString
}
$MyRecipients = @{}
$Page = 0
$SearchStartDateStr = $SearchStartDate.ToString()
$SearchEndDateStr = $SearchEndDate.ToString()
Write-Host "Getting recipients of evil message... from $SearchStartDateStr (UTC) to $SearchEndDateStr (UTC) from $SenderList"
#Since Message Traces cut off after 5k results, we use PageSize to limit it to 5k users and try another page of results till we run out
$MessageTraceArgs = @{
    SenderAddress = $SenderList
    StartDate = $SearchStartDateStr
    EndDate = $SearchEndDateStr
    Pagesize = 5000
    Status = $EmailStatusFilter
    ErrorAction = "Stop"
}
do {
    $Page++
    Write-Host "  Getting Page $Page of results, can take up to 5 minutes..."
    $a = (Get-MessageTrace @MessageTraceArgs -Page $Page | select-object recipientaddress).recipientaddress
    Write-Host "  Done."
    if ($null -ne $a) {
        <#For every person found in the trace, we look to make sure it is not already in the list and that it is
            an address in the supplied domain to which we can actually do something#>
        foreach($Recipient in $a) {
            if (!($MyRecipients.ContainsKey($Recipient)) -and
                $Recipient.IndexOf("@$EmailDomain") -gt 0 -and
                $Mailboxes2Exclude -notcontains $Recipient) {
                    $MyRecipients.Add($Recipient,$null)
            }
        }
    }
    #Just because these searches can be resource-intensive and occasionally freak, wait a second
    Start-Sleep -Seconds 1
} until ($null -eq $a)
Write-Host Done

try {
    if ($NoMFA) {
        Connect-IPPSSession -Credential $Cred -ErrorAction Stop
    } else {
        Connect-IPPSSession -UserPrincipalName $CredUPN -ErrorAction Stop
    }
} catch {
    $ErrMsg = "Error connecting to O365 Compliance and Security Center to create content search - $_. " +
        "Press any key to exit."
    Read-Host $ErrMsg
    Disconnect-ExchangeOnline -confirm:$false
    Exit
}

$ContentSearchCounter=0
$SearchNames=New-Object System.Collections.ArrayList
$CompletedSearches=0
$TotalRecipients2Process = $MyRecipients.Keys.Count
for($i=0;$i -lt $TotalRecipients2Process;$i+=50000) {
    $ContentSearchCounter++
    $SearchName = "SPAM-SearchAndHardPurge-$CredUPN-$(Get-Date -Format FileDateTime)"
    if ($MyRecipients.Keys.Count -gt 49999) {
        $SearchName+="-Part$ContentSearchCounter"
    }
    $SearchNames.Add($SearchName) | Out-Null
    $ComplianceArgs = @{
        Name = $SearchName
        Description = "Incident response, purging spam"
        ExchangeLocation = @($MyRecipients.Keys)[($i)..($i+49999)]
        AllowNotFoundExchangeLocationsEnabled = $true
        ContentMatchQuery = $SearchQuery
    }
    do {
        $Good = $false
        try {
            New-ComplianceSearch @ComplianceArgs -ErrorAction Stop
            $Good = $true
        } catch {
            $AmbiguousEntries = ([Regex]::new(".*: The location .* is ambiguous\..*")).Matches($_.exception.message)
            if($AmbiguousEntries.Count -gt 0){
                foreach($Entry in $AmbiguousEntries.Value.Trim()){
                    $EmailAddress = $Entry.substring(0,$Entry.IndexOf(":"))
                    try {
                        $ReplacementId = (Get-User -Identity $EmailAddress).ExternalDirectoryObjectId
                        Write-Host "Replacing $EmailAddress with $ReplacementId due to ambiguous O365 resolving"
                    } catch {
                        Write-Host "Removing $EmailAddress due to ambiguous error - $_"
                    } finally {
                        $MyRecipients.Remove($EmailAddress)
                        if($ReplacementId.length -gt 0 -and !($MyRecipients.ContainsKey($ReplacementId))){
                            $MyRecipients.Add($ReplacementId,$null)
                        }
                    }
                }
                $ComplianceArgs.ExchangeLocation = @($MyRecipients.Keys)[($i)..($i+49999)]
            } else {
                Read-Host "Error creating Compliance Search - $_. Exiting..."
                Disconnect-ExchangeOnline -Confirm:$false | Out-Null
                Exit
            }
        }
    } until ($Good -eq $true)
    try {
        Start-ComplianceSearch -Identity $ComplianceArgs.Name
    } catch {
        Read-Host "Compliance search $SearchName created but unable to start - $_. Exiting..."
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
        Exit
    }
    try {
        do {
            Start-Sleep -Seconds 5
            Write-Host "  $(Get-Date -Format u) - Checking to see if the search is completed and pausing for a few seconds if not."
        } until((Get-ComplianceSearch -Identity $SearchName).Status -eq "Completed")
    } catch {
        Read-Host "Error with the compliance search status - $_"
    }
    New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType HardDelete -Confirm:$false
}
$PauseInMinutes = 1
do {
    Write-Host "$(Get-Date -Format u) : Pausing for $PauseInMinutes minutes to allow job to progress."
    Start-Sleep -Seconds ($PauseInMinutes*60)
    foreach($SearchName in $SearchNames) {
        $ActionInfo = Get-ComplianceSearchAction -Identity ($SearchName+'_Purge') -Details
        $Progress = ($ActionInfo.Results -split ";" |
            Where-Object { $_ -like " Item count:*" -and $_ -notlike " Item count: 0"}).count
        if ($Progress -lt $TotalRecipients2Process -and $ActionInfo.Status -ne "Completed") {
            Write-Host ("  $(Get-Date -Format u) : Current Progress for $($SearchName)_Purge - " +
                "$Progress/$TotalRecipients2Process ($([int]($Progress/$TotalRecipients2Process*100))%) Complete")
        } else {
            Write-Host "  $(Get-Date -Format u) : Current Progress for $($SearchName)_Purge - 100% Complete"
            $CompletedSearches++
        }
    }
} until($CompletedSearches -eq $ContentSearchCounter)
#When done, we want to use Remove-PSSession to make sure that we properly close our powershell o365 sessions
Disconnect-ExchangeOnline -confirm:$false