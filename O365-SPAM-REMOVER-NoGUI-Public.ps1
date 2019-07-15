#REQUIRES -Version 5
<#
.SYNOPSIS
Performs a message trace for a spam message and searches for and purges it from recipient mailboxes with multiple windows
 
.DESCRIPTION
Connects to O365 - assuming you have Exchange Admin permissions for each credential you supply - and performs a Get-MessageTrace. From there, it will take all of the recipients, split them up into rounds of child windows, and perform Search-Mailbox -DeleteContent commands against each recipient.
 
.PARAMETER StartDate
REQUIRED The beginning of the search for the spam message - typically the day BEFORE users receive the message, must be within the last 7 days
.PARAMETER EndDate
REQUIRED The end of the search for the spam message - typically the day AFTER users receive the message
.PARAMETER Recipients
OPTIONAL Will be determined by message trace but can be supplied separately - expecting a comma-separated string of email addresses, max of 1000 for performance's sake
.PARAMETER CredU
OPTIONAL Username of the Exchange Admin credentials
.PARAMETER CredP
OPTIONAL Encrypted password secure string of the Exchange Admin credentials (password encrypted via convertfrom-securestring command)
.PARAMETER SearchQuery
OPTIONAL Uses the date when users are expected to have received the spam message and the evil senders to form a query for Search-Mailbox command
.PARAMETER EmailDomain
REQUIRED Domain of the mailboxes affected by spam campaign (used for message trace search/filter). Defaults to contoso.com and will prompt if you don't supply yours.
.PARAMETER WindowsPerCred
OPTIONAL This script defaults to 3 simultaneous powershell windows per Exchange Admin credential supplied due to typical tenant limits in O365, adjust at your own peril
.PARAMETER MailboxesPerWindow
OPTIONAL Number of Mailboxes to process per powershell window generated. Defaults to 500 and auto-adjusts for smaller jobs, adjust at your own peril. The total length of all email addresses plus the search terms must be less than 80k characters.
 
.NOTES
  Created by: Brendan Horner (www.hornerit.com)
  Notes: MUST BE RUN AS SCRIPT FILE, do NOT copy-paste into PS to run
  Version History:
  --2019-06-27-Fixed bug for child windows due to changing MFA parameter to NoMFA and updated MFA Exchange Module to use the latest version
  --2019-06-19-Altered MFA parameter to be NoMFA so someone can force basic auth by setting that switch and adjusted MFA module to pull the latest version of the module on your machine
  --2019-05-28-Bug Fixes for MFA and errors in mailbox
  --2019-05-22-Added support for Exchange Online MFA Module
  --2019-05-21-Completed documentation and separate version of script that has a GUI, see other post for that version (https://www.hornerit.com/2019/05/o365-spam-remover-script-now-with-gui.html).
  --2019-05-16-Rewrote sections for dynamic window generation based on params, allows as many Exch Admin accts to assist as you can try...watch out for RAM usage
  --2019-05-02-Added better logic for throttling
  --2019-04-15-Initial public version
 
.EXAMPLE
.\O365-SPAM-REMOVER.ps1
.\O365-SPAM-REMOVER.ps1 -Recipients "someone@CONTOSO.COM,someoneelse@CONTOSO.COM" -SearchQuery "FROM:bob@something.com AND Received:04/19/2018"
#>
param(
[string]
$StartDate,
[string]
$EndDate,
[string]
$Recipients,
[string]
$CredU,
[string]
$CredP,
[string]
$SearchQuery,
[string]
$EmailDomain = "contoso.com",
[int]
$WindowsPerCred = 3,
[int]
$MailboxesPerWindow = 500,
[switch]
$NoMFA
)

if(!($NoMFA)){
    #Try to get the Exchange Online Powershell module that supports MFA
    try{
        $getChildItemSplat = @{
            Path = "$Env:LOCALAPPDATA\Apps\2.0\*\CreateExoPSSession.ps1"
            Recurse = $true
            ErrorAction = 'Stop'
            Verbose = $false
        }
        $MFAExchangeModule = ((Get-ChildItem @getChildItemSplat | Sort-Object LastWriteTime -Descending | where-object {(Test-Path "$($_.PSParentPath)\Microsoft.Exchange.Management.ExoPowershellModule.dll") -eq $true} | Select-Object -First 1 | Select-Object -ExpandProperty fullname).Replace("\CreateExoPSSession.ps1", ""))
        . "$MFAExchangeModule\CreateExoPSSession.ps1" 3>$null
        Write-Host "MFA Module found and imported"
    } catch {
        $NoMFA = $true
        Write-Host "MFA Module not found. If legacy auth is disabled for your tenant, this script will most likely fail. To install the latest module, go to https://aka.ms/exopspreview"
    }
}
 
#If supplied, create the Credential object used to log into O365 session. If we are using MFA, the credential token cache should hopefully still be working so just connecting without creds will work
if($null -ne $CredU -and $CredU.Length -gt 0){
    write-host "Username supplied, attempting to connect to O365"
    if(!($NoMFA)){
        Get-PSSession | Remove-PSSession
        Connect-EXOPSSession -UserPrincipalName $CredU 3>$null
        $Session = Get-PSSession
        try {
            Invoke-Command -Session $Session -ScriptBlock { Get-OrganizationConfig | Select-Object Name } -ErrorAction Stop
        } catch {
            Read-Host "The account supplied does not appear to be an Exchange Admin account, please try again. Press any key to close..."
            exit
        }
    } else {
        $Cred = New-Object PSCredential($CredU,(ConvertTo-SecureString $CredP))
        try{
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Authentication Basic -AllowRedirection -Credential $Cred -ErrorAction Stop
        } catch {
            Write-Host "Unable to establish a connection to O365. You will need to re-run the script using this command to simply try again or you can change something:"
            Write-Host "$($MyInvocation.MyCommand.Definition) -CredU $CredU -CredP $CredP -SearchQuery `"$SearchQuery`" -Recipients $Recipients -MFA:$NoMFA"
            do{
                $ReadyToClose = Read-Host "Press 'Y' key to close this script (make sure you copy the above command to paste and run in another powershell window to try again without running the whole giant script again)"
            } until ($ReadyToClose -eq "y")
            exit
        }
    }
} else {
    $Creds = New-Object -TypeName System.Collections.ArrayList
}
 
#Change window appearance if this is a child window so that it is smaller
if($Recipients.Length -gt 0){
    $title = "SPAM REMOVER - Processing "+$Recipients.Substring(0,$Recipients.IndexOf("@"))+" thru "+$Recipients.Substring($Recipients.LastIndexOf(",")+1,$Recipients.LastIndexOf("@")-$Recipients.LastIndexOf(",")-1)
    $Host.ui.RawUI.WindowTitle = $title
    $newSize = $Host.UI.RawUI.WindowSize
    $newSize.Height = 30
    $newSize.Width = 75
    $Host.UI.RawUI.WindowSize = $newSize
    $newBuffer = $Host.UI.RawUI.BufferSize
    $newBuffer.Height = 3000
    $newBuffer.Width = 75
    $Host.UI.RawUI.BufferSize = $newBuffer
}
 
#If SearchQuery has not been supplied, get the evil sender(s) then the received date and, finally, the subject line(s) of the evil emails
if($SearchQuery.Length -eq 0){
    #In case someone forgot to change the email domain, ask for it
    if($EmailDomain -eq "contoso.com"){
        do{
            $EmailDomain = Read-Host "You have not updated the script with your email domain for your recipients. Please enter the domain (the part after the @ symbol, e.g. mycompany.com)"
            if($EmailDomain -notmatch "^[a-zA-Z0-9]+\..*$"){
                Write-Host "Invalid entry"
                $EmailDomain = $null
            }
        } until ($null -ne $EmailDomain)
    }
    do {
        if($Creds.Count -eq 0){
            $CredEntry = Read-Host "[Required]Email address of Exchange Admin account to use for this script. This prompt will repeat until you press enter with no information."
        } else {
            $CredEntry = Read-Host "[Optional]Email address of another Exchange Admin account to speed up the process. This prompt will repeat until you press enter with no information."
        }
        if($CredEntry.Length -gt 0){
            if($CredEntry -match "^.+@.+\..+$" -and $null -eq ($Creds.Username | Where-Object { $_ -eq $CredEntry })){
                if(!($NoMFA)){
                    Connect-EXOPSSession -UserPrincipalName $CredEntry 3>$null
                    $TestPSSession = Get-PSSession
                } else {
                    $CredEntry = Get-Credential -UserName $CredEntry -Message "Please enter password for $CredEntry"
                    Write-Host "  Attempting to connect to O365 and verify this is an Exchange admin"
                    $TestPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $CredEntry -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                }
                try {
                    Invoke-Command -Session $TestPSSession -ScriptBlock { Get-OrganizationConfig | Select-Object Name } -ErrorAction Stop
                    Remove-PSSession $TestPSSession
                    if(!($NoMFA)){
                        $Creds.Add((New-Object PSCredential($CredEntry,(ConvertTo-SecureString " " -AsPlainText -Force)))) | Out-Null
                    } else {
                        $Creds.Add($CredEntry) | Out-Null
                    }
                    Write-Host "  Successful, prompting for another one."
                } catch {
                    Write-Host "There was an error connecting to O365: Not an admin, account cannot use basic auth, bad password, or bad email"
                }
            } else {
                Write-Host "That was not a valid entry, try again"
            }
        }
    } until ($CredEntry.Length -eq 0 -and $Creds.Count -gt 0)
    $SendersHash = @{}
    do {
        $BadSender = read-host "[Required]Email address of evil sender. This prompt will repeat until you press enter with no information. Do not enter quotes or empty, extra spaces"
        do {
            #Validate that each email address at least matches the typical email pattern
            $SendersGood = 1
            if($BadSender.length -gt 0){
                if(!($BadSender -match ".*@.*\..*")){
                    $SendersGood = 0
                    $BadSender = read-host "Bad email address input, try again"
                } else {
                    $SendersHash.add($BadSender,$null)
                }
            }
        } until ($SendersGood -eq 1)
    } until ($BadSender.length -eq 0)
    $Senders = $SendersHash.keys -join ","
 
    #Get the start date and time of the range when that users should have received the spam
    $StartDate = read-host "[Required]Start Date AND TIME for range when users received message (e.g. 7/18/2018 12:20 AM)"
    do {
        $good = 0
        try {
            $StartDate = Get-Date($StartDate)
            $good = 1
        }
        catch {
            $StartDate = read-host "Start date invalid, try again"
        }
    } until ($good -eq 1)
    $SearchStartDate = (get-date $StartDate).ToUniversalTime()
 
    #Get the end date and time of the range that users should have received the spam
    $EndDate = read-host "[Required]End Date AND TIME for range when users received message (e.g. 7/18/2018 3:59 PM)"
    do {
        $good = 0
        try {
            $EndDate = Get-Date($EndDate)
            if($EndDate -gt $StartDate){ $good = 1 } else { throw }
        }
        catch {
            $EndDate = read-host "End date invalid or before start date, try again"
        }
    } until ($good -eq 1)
    $SearchEndDate = (get-date $EndDate).ToUniversalTime()
 
    #Get the subject line(s) of the evil messages to filter
    $SubjectLineFilter = @{}
    do {
        $SubjectLine = read-host "[Optional]Subject line to filter. This prompt will repeat until you press enter with no information. Do not enter any quotes or backticks unless actually in the subject"
        if($SubjectLine.length -gt 0){
            $SubjectLineFilter.add($SubjectLine,$null)
        }
    } until ($SubjectLine.length -eq 0)
    $SubjectLineStr = $null
    if($SubjectLineFilter.count -gt 0){
        $SubjectLineStr = '(Subject:'
        foreach($subject in $SubjectLineFilter.keys){
            $SubjectLineStr += '"'+$subject.Replace('"','""')+'" OR Subject:'
        }
        $SubjectLineStr = $SubjectLineStr.TrimEnd(" OR Subject:")
        $SubjectLineStr += ")"
    }
 
    #Build Search Query from the SearchStartDate, SearchEndDate, subject line(s), and evil sender(s)
 
    $SearchQuery = "(Received:`""+$SearchStartDate.toString()+".."+$SearchEndDate.toString()+"`") AND "
    $SenderList = New-Object System.Collections.ArrayList
    #If more than one sender, create filter string with a bunch of ORs for senders; otherwise, just set one
    if($Senders.IndexOf(",") -gt 0){
        $SearchQuery += "("
        foreach($EvilSender in $Senders.Split(",")){
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
    if($null -ne $SubjectLineStr) {
        $SearchQuery += " AND "+$SubjectLineStr
    }
}
 
#If a string of spam recipients has not been supplied, we perform a message trace to get them; otherwise, they were supplied - probably by this script
if($Recipients.length -eq 0){
    $MyRecipients = @{}
    $Page = 0
    $SearchStartDateStr = $SearchStartDate.ToString()
    $SearchEndDateStr = $SearchEndDate.ToString()
    Write-Host "Connecting to Exchange Online..."
    try {
        if(!($NoMFA)){
            Connect-EXOPSSession -UserPrincipalName ($Creds[0].UserName) 3>$null
            $Session = Get-PSSession
        } else {
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid" -Authentication Basic -AllowRedirection -Credential ($Creds[0])
        }
    } catch {
        Read-Host "There was an error connecting to Exchange Online. Press any key to close."
        Exit
    }
    Write-Host "Done."
    Write-Host "Getting recipients of evil message... from $SearchStartDateStr (UTC) to $SearchEndDateStr (UTC) from $SenderList"
    #Since Message Traces cut off after 5k results, we use PageSize to limit it to 5k users and try another page of results till we run out
    do {
        $Page++
        Write-Host "  Getting Page $Page of results, can take up to 5 minutes..."
        $a = (Invoke-Command -Session $Session -ScriptBlock { Get-MessageTrace -SenderAddress $Using:SenderList -StartDate $Using:SearchStartDateStr -EndDate $Using:SearchEndDateStr -Pagesize 5000 -Page $Using:Page -Status "Pending","Delivered" -ErrorAction Stop | select-object recipientaddress} -HideComputerName).recipientaddress
        Write-Host "  Done."
        if($null -ne $a){
            #For every person found in the trace, we look to make sure it is not already in the list and that it is an LU address to which we can actually do something
            foreach($Recipient in $a){
                if(!($MyRecipients.ContainsKey($Recipient)) -and $Recipient.IndexOf("@$EmailDomain") -gt 0){
                    $MyRecipients.Add($Recipient,$null)
                }
            }
        }
        #Just because these searches can be resource-intensive and occasionally freak out, wait a second before trying again
        Start-Sleep -Seconds 1
    } until ($null -eq $a)
    Write-Host Done
    #When done, we want to use Remove-PSSession to make sure that we properly close our powershell o365 sessions
    Remove-PSSession $Session
 
    #Test that you are running this as a script (needed to spawn child windows)
    Write-Host Testing that you ran this as a script and did not copy paste it...
    try{
        $ScriptPath = $MyInvocation.MyCommand.Definition
        Resolve-Path $ScriptPath | out-null
    } catch {
        Write-Host STOPPING - You are not running this as a script. Press any key to close...
        Exit;
    }
    Write-host Done
 
    #Just for clarity to the person running the script
    Write-Host "Spawning child powershell windows with these parameters:"
    Write-Host "  Path to script file - $ScriptPath"
    Write-Host "  Received Local Date/Time -"$SearchStartDate.ToLocalTime().ToString()"to"$SearchEndDate.ToLocalTime().ToString()
    Write-Host "  Search Query (UTC DateTime) - $SearchQuery"
    Write-Host "  Total Mailboxes being processed -"$MyRecipients.Count
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
 
    #When sending stuff with double quotes to the child powershell windows, double quotes get lost due to how powershell works. This adds a backslash to escape them so it works correctly.
    $SearchQuery = $SearchQuery.replace('"','\"')
 
    #Here is where we sort the list of recipients for later chunking into smaller groups
    #$MaxMailboxesProcessedPerRound tells how many mailboxes will be processed for each round of open windows. Currently, with 3 windows per credential supplied, 1500 is expected because we want 500 per window per cred
    $Mailboxes = @($MyRecipients.keys | Sort-Object | foreach-object { $_.toString() })
 
    #If the total number of mailboxes to process is smaller than the number of windows per session * number of credentials provided, remove some creds because the extra are pointless
    if($Mailboxes.Count -lt ($Creds.Count*$WindowsPerCred)){
        do {
            if($Creds.Count -gt 1){
                $Creds.RemoveAt(($Creds.Count)-1)
            }
        } until ($Mailboxes.Count -ge ($Creds.Count*$WindowsPerCred) -or $Creds.Count -eq 1)
    }
     
    #Figure out how many mailboxes will be processed per credential supplied.
    $MailboxesPerCred = if($Creds.Count -gt 1){[math]::Ceiling($Mailboxes.Count/$Creds.Count)} else { $Mailboxes.Count }
 
    #Adjust the number of mailboxes processed in each window if we have such a small number that the (number of mailboxes per window * the windows per credential) is too big
    if($MailboxesPerCred -lt ($MailboxesPerWindow*$WindowsPerCred)){ $MailboxesPerWindow = [math]::Ceiling($MailboxesPerCred/$WindowsPerCred) }
    $MaxMailboxesProcessedPerRound = $MailboxesPerWindow * $WindowsPerCred * ($Creds.Count)
    if($MaxMailboxesProcessedPerRound -gt $MyRecipients.Count){ $MaxMailboxesProcessedPerRound = $MyRecipients.Count }
 
    #Tell the script-runner the results of our calculations on creds, windows, and mailboxes
    Write-Host "  Total number of accounts being used - $($Creds.Count)"
    Write-Host "  Number of mailboxes per child window - $MailboxesPerWindow"
    Write-Host "  Number of mailboxes per round of child windows - $MaxMailboxesProcessedPerRound"
 
    #Begin the process of spawning child windows. Number of windows will be the WindowsPerCred * number of creds you provided and waits for all windows in each round to complete before attempting another
    Write-Host "Child Window Data:"
    for($m=0;$m -lt $Mailboxes.count;$m+=$MaxMailboxesProcessedPerRound){
        $RoundMinimum = $m
        $ChildWindows = $(
            for($c=0;$c -lt $Creds.Count;$c++){
                for($w=0;$w -lt $WindowsPerCred;$w++){
                    $min=$RoundMinimum
                    $max=$RoundMinimum+$MailboxesPerWindow-1
                    if($null -eq $Mailboxes[$max]){ $max = $Mailboxes.Count-1}
                    if($null -ne $Mailboxes[$min]){
                        Write-Host "  Window$c$w will be $($Mailboxes[$min]) to $($Mailboxes[$max])"
                        $u = $Creds[$c].UserName
                        $p = ConvertFrom-SecureString $Creds[$c].Password
                        $Recipients = $Mailboxes[($min)..($max)] -join ","
                        #This is the freaking magic that opens another powershell window and supplies all the values that are set as parameters up at the top of this script
                        #!!!NOTICE THE BACKTICK CHARACTERS FOR FILE AT BEGINNING AND SEARCH QUERY AT THE END! IF YOU TAKE THEM AWAY, THE SEARCH QUERY FOR THESE SPAWNED PROCESSES IS INCOMPLETE AND DELETES LOTS MORE EMAILS!!!
                        if(!($NoMFA)){
                            Start-Process powershell -Passthru -ArgumentList "-file `"$ScriptPath`" -Recipients $Recipients -CredU $u -CredP $p -SearchQuery `"$SearchQuery`""
                        } else {
                            Start-Process powershell -Passthru -ArgumentList "-file `"$ScriptPath`" -Recipients $Recipients -CredU $u -CredP $p -SearchQuery `"$SearchQuery`" -NoMFA"
                        }
                    }
                    $RoundMinimum+=$MailboxesPerWindow
                }
            }
        )
        #Wait for all of the child windows spawned up there to complete and close before opening a new set.
        $ChildWindows | Wait-Process
    }
    $timer.Stop()
    read-host "Done, the runtime for this entire process was"($timer.Elapsed.TotalMinutes)"minutes. Press any key to complete script"
    exit
} else {
    #Go ahead and process the recipients supplied to the script already
    Write-Host "Processing recipients using this query: $SearchQuery"
    $total = ([regex]::Matches($Recipients,"@")).count
    $counter = 0
    foreach ($Recipient in ($Recipients.split(",") | Sort-Object)){
        $counter++
        $errorCounter = 0
        do{
            try{
                $SearchResults = Invoke-Command -Session $Session -Scriptblock { Search-Mailbox -Identity $Using:Recipient -SearchQuery $Using:SearchQuery -deletecontent -Force -ErrorAction Stop -WarningAction SilentlyContinue } -HideComputerName -WarningVariable SearchWarning -ErrorAction Stop
                if($SearchWarning.Message -like "*exceeded*" -or $SearchWarning.Message -like "*throttl*" -or $SearchWarning.Message -like "*frequent*" -or $SearchWarning.Message -like "*The I/O Operation*"){
                    throw $SearchWarning.Message
                }
            } catch {
                $errorCounter++
                if($_.exception.message -like "*exceeded*" -or $_.exception.message -like "*throttl*" -or $_.exception.message -like "*frequent*" -or $_.exception.message -like "*The I/O Operation*"){
                    Write-Host "O365 throttling limit hit. Pausing for 5 minutes starting now $(Get-Date)"
                    Start-Sleep -Seconds 300
                    Remove-PSSession -Session $Session
                }
                if($Session.State -ne "Opened"){
                    Get-PSSession | Remove-PSSession
                    if(!($NoMFA)){
                        Connect-EXOPSSession -UserPrincipalName $CredU 3>$null
                        $Session = Get-PSSession
                    } else {
                        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    }
                }
            }
        } until ( $null -ne $SearchResults.ResultItemsCount -or $errorCounter -eq 2)
        if($errorCounter -eq 2){
            Write-Host "ERROR - $Recipient - $counter of $total boxes"
        } else {
            Write-Host "$($SearchResults.ResultItemsCount) item removed from $Recipient - $counter of $total boxes"
        }
    }
 
    #When done, we want to use Remove-PSSession to make sure that we properly close our powershell o365 sessions
    Get-PSSession | Remove-PSSession
}
