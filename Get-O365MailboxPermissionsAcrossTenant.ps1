#Requires -Module ExchangeOnlineManagement,ActiveDirectory
<#
.NOTES
    Created By: Brendan Horner (hornerit.com)
    Purpose: Get all custom permissions entries across the entire tenant and store in csv files
    Version History:
    --2020-03-05-Fixed bug in resumes if it is resuming a full download where it would delete the final file since there is no merging of data. Also added many line breaks to make code more readable.
    --2020-02-24-Added extra filters for certain mailbox folder value returns like SIDs with a NT: prefix or ExchangePublishedUser prefix. Changed Tee-Object to use a file instead of memory to make more memory efficient.
    --2020-02-20-Added a pause before merging the final contents of files together so you can verify. Fixed error for new runs. Fixed Mailbox only runs to get objectid correctly.
    --2020-02-19-Reduced the time for processing new mailboxes so that it doesn't reach so far back past the previous run. Fixed check for folderpath for resumes.
    --2020-02-17-Minor tweak for removing temp file only if exists, fixed creation datetime values for final exports, fixed bug in full download for sorting
    --2020-02-13-Added fix for forward slashes in mailbox folder names as it becomes [char]63743 or a question mark inside a box
    --2020-02-12-Updated folderpath to show full path, added escaping of single quotes for resolving folder permissions, added sorting for folder perms
    --2020-02-11-Added documentation and a regex check for email domain to be the proper format
    --2020-02-10-Added check for Sid function to throw error if no results return, fixed mailbox identity for folder permissions to remove ":\", several folder bugs
    --2020-02-06-Initial public version. Performance with new EXO modules appears to be 8-10 hours per 100k mailboxes for mailbox perms; 12 hours for 50k mailbox folders

.SYNOPSIS
Obtains all custom mailbox permissions for the entire tenant with some options to reduce the scope as necessary
.DESCRIPTION
Obtains all mailboxes, only recently changed mailboxes, or resumes previous efforts
.PARAMETER NoMFA
[OPTIONAL] If you know for certain that MFA is not enabled on your tenant or for the account under which you wish to run this, use this switch. No value is expected.
.PARAMETER CustomPermsCSVPath
[REQUIRED] CSV output path and name where the output needs to go (and temp and other files will be generated within the same folder). Defaults to the current script folder\CustomPermEntriesForMailboxes.csv"
.PARAMETER GetMailboxResultSize
[OPTIONAL] If you are testing, use a number here to reduce overall query size. Defaults to "Unlimited" for entire tenant
.PARAMETER ProcessNewlyCreatedOrChangedMailboxesOnly
[OPTIONAL] To reduce the overall download, use this switch to check the full download file and download only items changed 24 hours earlier than either the created or modified date of the full download CSV, whichever is earliest
.PARAMETER IncludeFolders
[OPTIONAL] If mailbox folder permissions wish to be downloaded, set this value to $true. This will retrieve all Inbox folders, Calendar, and custom folders created at the root of the mailbox.
.PARAMETER Resume
[OPTIONAL] If this script was interrupted or errored previously (preferably very recently ONLY), use this switch to resume the download and be sure to add -IncludeFolders $true if you intended to run this for folders and the Modified switch if you remember setting that. If the interruption was more than 1 week, do not use this switch - just download the full dataset.
.PARAMETER UseModifiedDateForPermsDelta
[OPTIONAL] If using the ProcessNewlyCreatedOrChangedMailboxesOnly switch, this switch forces the use of the modified date instead of the created date for determining how far back to retrieve permissions. This is sometimes required due to overwriting files keeping old created date.
#>
[CmdletBinding()]
param(
    [switch]$NoMFA,
    [string]$CustomPermsCSVPath = "$PSScriptRoot\CustomPermEntriesForMailboxes.csv",
    [string]$GetMailboxResultSize = "Unlimited",
    [string]$EmailDomain = "@contoso.com",
    [switch]$ProcessNewlyCreatedOrChangedMailboxesOnly,
    [bool]$IncludeFolders,
    [switch]$Resume,
    [switch]$UseModifiedDateForPermsDelta
)
#Import the modules used for this script - ExchangeOnlineManagement, which is the modern Exchange V2 powershell module, and ActiveDirectory, which can be obtained within RSAT
Import-Module ExchangeOnlineManagement
Import-Module ActiveDirectory

#This Test-ObjectId function is needed because sometimes a permissions entry on a mailbox or folder is not a userprincipalname (e.g. Sid, DisplayName) so we try to resolve it
function Test-ObjectId{
    param([string]$ObjectId)
    try {
        #In Exchange, 'Default' is typically used on calendars to allow AvailabilityOnly to see Free/Busy info, it should not otherwise be used as it is all authenticated users. Anonymous is, well, anyone.
        if($ObjectId -eq "Default" -or $ObjectId -eq "Anonymous"){ return $ObjectId }
        $ObjectId = $ObjectId.Replace("'","''")
        #If the object id starts with S-1-5-21, we know it is a Sid so we can ask for an ADObject filtering on ObjectSid; if it looks like an email - search by proxyAddresses, otherwise use displayname
        if($ObjectId -match "S-1-5-21"){
            #Sometimes the permissions entry looks like this - NT:S-1-5-21-blahblah. So it's a SID but it begins with 'NT:' so we try to only grab the SID to resolve
            $ADO = Get-ADObject -Filter "objectSid -eq '$($ObjectId.Substring($ObjectId.IndexOf("S-1")))'" -Properties UserPrincipalName -ErrorAction Stop
        } elseif($ObjectId -match "$EmailDomain"){
            #Some entries look like ExchangePublishedUser.someone@something.com so we grab the email afterward and try to find that
            if($ObjectId -like "ExchangePublishedUser*"){
                $ObjectId = $ObjectId.Substring($ObjectId.IndexOf(".")+1)
            }
            $ADO = Get-ADObject -Filter "proxyAddresses -eq 'smtp:$ObjectId'" -Properties UserPrincipalName
        } else {
            $ADO = Get-ADObject -Filter "displayName -eq '$ObjectId' -or samaccountname -eq '$ObjectId'" -Properties UserPrincipalName -ErrorAction Stop
        }
        #After retrieving, an empty resultset does not throw an error so we force one since no results were returned
        if($ADO.distinguishedName.Length -eq 0){ throw }
        #If multiple results return (e.g. two people with same displayname and mailbox folder only shows displaynames), don't force it, just log the info
        if($ADO.count -gt 1){
            "ERROR: Ambiguous: $ObjectId"
        } else {
            if($ADO.UserPrincipalName.Length -gt 0){
                $ADO.UserPrincipalName
            } else {
                $ADO.Name
            }
        }
    } catch {
        "ERROR: Unresolvable: $ObjectId"
    }
}

#Just in case you forgot to change the default or supply it when running the script, this forces you to supply an email domain
if($EmailDomain -eq "@contoso.com"){
    do{
        $EmailDomain = Read-Host "Please enter your email domain as @yourdomain.com (e.g. @myschool.edu)"
    } until ($EmailDomain.Length -gt 0 -and $EmailDomain -ne "@contoso.com" -and $EmailDomain -match "^@.+\..+")
}

#Prompt for credential to use and use some validation to make sure it is an Exchange Admin. If using MFA, we don't need a password because the MFA window will ask for it.
do{
    $CredEntry = if(!($NoMFA)){
        New-Object PSCredential((Read-Host "[Required]Email address of Exchange Admin account to use for this script. This prompt will repeat until you supply a valid entry."),(ConvertTo-SecureString " " -AsPlainText -Force))
    } else {
        Get-Credential -UserName (Read-Host "[Required]Email address of Exchange Admin account to use for this script. This prompt will repeat until you supply a valid entry.") -Message "Please enter password"
    }
    #Verify they entered something
    if($CredEntry.UserName.Length -gt 0){
        #Verify they entered an email address using a regex match
        if($CredEntry.UserName -match "^.+@.+\..+$"){
            #If you are using MFA, we expect that this is an interactive session and you have to supply the MFA approval. Note that if you have session expiration in Azure AD conditional access policies then it will cause the script to stop/fail without reauth
            if(!($NoMFA)){
                Write-Verbose "Attempting to MFA connect $CredEntry"
                Connect-ExchangeOnline -UserPrincipalName $CredEntry.UserName
            } else {
                Write-Verbose "Attempting to connect to O365 and verify this is an Exchange admin"
                Connect-ExchangeOnline -Credential $CredEntry
            }
            try {
                #This command will only work if you are an Exchange Admin
                Get-OrganizationConfig -ErrorAction Stop | Select-Object Name
            } catch {
                Write-Host "There was an error connecting to O365: Not an admin, account cannot use basic auth, bad password, or bad email"
                Get-PSSession | Remove-PSSession
                $CredEntry = $null
            }
        } else {
            Write-Host "That was not a valid entry, try again"
            $CredEntry = $null
        }
    }
} until ($CredEntry.UserName.Length -gt 0)

#Setup and begin logging and start the timer to track the total time to complete the script
$LogPath = "$PSScriptRoot\TranscriptLog-GetTenantMailboxPerms.txt"
Start-Transcript -Path $LogPath -Append
$Timer = [System.Diagnostics.Stopwatch]::StartNew()
$StartTime = (Get-Date)
$TotalMailboxesProcessed = 0

#Begin processing
Write-Host "Retrieving Mailbox Permissions...current time is $(Get-Date -format g)"
if($ProcessNewlyCreatedOrChangedMailboxesOnly -or $Resume){
    #If we are only trying to handle recent changes, we need the oldest date of the custom permissions entries between its creation and recent modifications (in case you removed an entry or something)
    if((Get-Item -Path $CustomPermsCSVPath).LastWriteTimeUtc -le ((Get-Item -Path $CustomPermsCSVPath).CreationTimeUtc) -or $UseModifiedDateForPermsDelta){
        $DateForDelta = Get-Date (Get-Date (Get-Item -Path $CustomPermsCSVPath).LastWriteTimeUtc).AddHours(-4) -Format u
    } else {
        $DateForDelta = Get-Date (Get-Date (Get-Item -Path $CustomPermsCSVPath).CreationTimeUtc).AddHours(-4) -Format u
    }
}

#Build a message to the user know the time when beginning to request mailboxes so that expectations can be managed and it's obvious how long this has run
$Message = $(Get-Date -format filedatetime)+" Retrieving mailboxes"

#Fundamental logic problem here - if a mailbox has changed recently to have permissions REMOVED, then they simply won't be in the retrieval since they may have no custom permissions and there's no way to use the new commands and know enough to remove them from existing permissions entries in the CSV. Not a horrible problem because it gives extra investigation but it means data is not perfectly accurate until you run a full/fresh download.
try {
    if($Resume){
        #We need to know if we are resuming a full download or a recently-changed download
        $CustomPermsCSVPathNew = ($CustomPermsCSVPath.Substring(0,$CustomPermsCSVPath.LastIndexOf("."))+'-NEW.csv')
        $ResumeCSVPath = ""
        $FoldersFoundInFile = $false
        if(Test-Path $CustomPermsCSVPathNew){
            #Reaching here means we are resuming an attempt to only process new or recently-changed mailboxes
            $LastEntry = import-csv $CustomPermsCSVPathNew | Select-Object -Last 1
            $LastMailbox = $LastEntry.Mailbox
            #If there is a folder somewhere in the file, we will treat this as if $IncludeFolders was set to True and continue retrieving folder permissions for this resume
            if($LastEntry.FolderPath.Length -gt 1 -or (import-csv $CustomPermsCSVPathNew | Where-Object { $_.FolderPath.Length -gt 1}).Count -gt 0){
                $FoldersFoundInFile = $true
            }
            if($FoldersFoundInFile){
                $Message += "`nResuming last download...downloading mailbox folder permissions whose mailbox name attribute is after $LastMailbox and whenChangedUTC attribute after $DateForDelta from $CustomPermsCSVPathNew"
            } else {
                $Message += "`nResuming last download...downloading mailboxes whose name attribute is after $LastMailbox and whenChangedUTC attribute after $DateForDelta from $CustomPermsCSVPathNew.`nPlease note: The previous attempt did not yet have any folder permissions so it is unknown if that was desired. If you did not specify the `$IncludeFolder = `$true then you will need to run this again for folder changes."
            }
            $ResumeCSVPath = $CustomPermsCSVPathNew
        } else {
            #Reaching here means that we are resuming an attempt to download all mailboxes and work fresh
            $LastEntry = import-csv $CustomPermsCSVPath | Select-Object -Last 1
            $LastMailbox = $LastEntry.Mailbox
            #Since we are retrieving all mailboxes, we don't filter like we would for recently-changed mailboxes...to keep the command the same, we just set a really old date here so we get all mailboxes.
            $DateForDelta = Get-Date "1/1/2000" -Format u
            #If there is a folder somewhere in the file, we will treat this as if $IncludeFolders was set to True and continue retrieving folder permissions for this resume
            if($LastEntry.FolderPath.Length -gt 1 -or (import-csv $CustomPermsCSVPath | Where-Object { $_.FolderPath.Length -gt 1}).Count -gt 0){
                $FoldersFoundInFile = $true
            }
            if($FoldersFoundInFile){
                $Message += "`nResuming last download...downloading mailbox folder permissions whose mailbox name attribute is after $LastMailbox and whenChangedUTC attribute after $DateForDelta from $CustomPermsCSVPath"
            } else {
                $Message += "`nResuming last download...downloading mailboxes whose name attribute is after $LastMailbox and whenChangedUTC attribute after $DateForDelta from $CustomPermsCSVPath.`nPlease note: The previous attempt did not yet have any folder permissions so it is unknown if that was desired. If you did not specify the `$IncludeFolder = `$true then you will need to run this again for folder changes."
            }
            $ResumeCSVPath = $CustomPermsCSVPath
        }
        #Tell the user that we are starting now and what we are doing
        Write-Host $Message
        #Actually go get the mailbox and folder data as necessary, the first is if folders were already found in the file so it just picks up with retrieving folder permissions from mailboxes
        if($FoldersFoundInFile){
            try {
                #The filters for the FolderStatistics try to focus on normal mail/calendar folders or ones created by the user or the Top of Information Store
                #The filters for the FolderPermission ignore when someone is blocked and make sure it is not the mailbox owner where possible and excludes the normal permissions a calendar has: Default users get AvailabilityOnly
                #The Select-Object calculations are there in the middle to force the appropriate format needed to send to the FolderPermission command, the second Select-Object formats for file output
                #Replacing [char]63743 is used to fix forward slashes that get interpreted as a box with a question mark in it. This will be fixed once we can use folderId instead of a path for folders
                Get-EXOMailbox -Filter "name -ge '$LastMailbox' -and whenChangedUTC -gt '$DateForDelta'" -ResultSize $GetMailboxResultSize -Properties ExternalDirectoryObjectId |
                    Tee-Object -FilePath "$PSScriptRoot\Temp-Resume-FoldersFoundInFile-Mailboxes.txt" |
                    Get-EXOMailboxFolderStatistics |
                    Where-Object {
                        $_.SearchFolder -eq $false -and
                        @("Root","Calendar","Inbox","User Created") -contains $_.FolderType -and
                        (@("IPF.Note","IPF.Appointment",$null) -contains $_.ContainerClass -or $_.Name -eq "Top of Information Store")} |
                    Select-Object @{Label="Identity";Expression={ 
                        if($_.Name -eq "Top of Information Store"){
                            $_.Identity.Substring(0,$_.Identity.IndexOf("\"))
                        } else {
                            $_.Identity.Substring(0,$_.Identity.IndexOf("\"))+':'+$_.Identity.Substring($_.Identity.IndexOf("\")).Replace([char]63743,"/")
                        }}} |
                    Get-EXOMailboxFolderPermission |
                    Where-Object {
                        $_.AccessRights -ne "None" -and
                        @("NT AUTHORITY\SELF") -notcontains $_.User -and
                        $_.User -ne ($_.Identity.Substring(0,$_.Identity.IndexOf(":\"))) -and
                        !($_.User -eq "Default" -and $_.AccessRights -eq "AvailabilityOnly")} |
                    Select-Object @{Label="Mailbox";Expression={$_.Identity.Substring(0,$_.Identity.IndexOf(":\"))}},
                        @{Label="FolderPath";Expression={
                            if($_.FolderName -eq "Top of Information Store"){
                                "Top of Information Store"
                            } else {
                                $_.Identity.Substring($_.Identity.IndexOf(":\")+2)
                            }
                        }},
                        @{Label="UserGivenAccess";Expression={Test-ObjectId -ObjectId $_.User}},
                        @{Label="AccessRights";Expression={$_.AccessRights -join ","}} |
                    Where-Object { $_.UserGivenAccess -ne ($_.Mailbox+"$EmailDomain")} |
                    Export-CSV -Path $ResumeCSVPath -Append
                #Grab the text file generated from Tee-Object and count the IDs so we can count how many mailboxes were processed, then remove the file
                $TotalMailboxesProcessed += (Get-Content "$PSScriptRoot\Temp-Resume-FoldersFoundInFile-Mailboxes.txt" |
                    Where-Object { $_ -match "^ExternalDirectoryObjectId :"}).Count
                Remove-Item "$PSScriptRoot\Temp-Resume-FoldersFoundInFile-Mailboxes.txt" -Force -Confirm:$false
            } catch {
                throw
            }
        } else {
            #Reaching this means no folders were found in the file so we just need to work on mailboxes and we'll check for the IncludeFolders boolean afterward
            try {
                #The filters for MailboxPermission try to ignore inherited permissions, users who are blocked, and the mailbox owner to find truly other users with permissions
                #The Select-Object is to format the file output
                Get-EXOMailbox -Filter "name -ge '$LastMailbox' -and whenChangedUTC -gt '$DateForDelta'" -ResultSize Unlimited -Properties ExternalDirectoryObjectId |
                    Tee-Object -FilePath "$PSScriptRoot\Temp-Resume-NoFoldersFoundInFile-Mailboxes.txt" |
                    Get-EXOMailboxPermission -ExternalDirectoryObjectId $_.ExternalDirectoryObjectId -ResultSize Unlimited |
                    Where-Object {
                        $_.IsInherited -eq $false -and
                        $_.Deny -eq $false -and
                        @("NT AUTHORITY\SELF") -notcontains $_.User -and
                        $_.User -ne ($_.Identity+"$EmailDomain")} |
                    Select-Object @{Label="Mailbox";Expression={$_.Identity}},
                        @{Label="FolderPath";Expression={''}},
                        @{Label="UserGivenAccess";Expression={Test-ObjectId -ObjectId $_.User}},
                        @{Label="AccessRights";Expression={$_.AccessRights -join ","}} |
                    Export-CSV -Path $ResumeCSVPath -Append
                #Grab the text file generated from Tee-Object and count the IDs so we can count how many mailboxes were processed, then remove the file
                $TotalMailboxesProcessed += (Get-Content "$PSScriptRoot\Temp-Resume-NoFoldersFoundInFile-Mailboxes.txt" |
                    Where-Object { $_ -match "^ExternalDirectoryObjectId :"}).Count
                Remove-Item "$PSScriptRoot\Temp-Resume-NoFoldersFoundInFile-Mailboxes.txt" -Force -Confirm:$false
            } catch {
                throw
            }
            #If the script is told to include folder permissions
            if($IncludeFolders){
                try {
                    #The filters for the FolderStatistics try to focus on normal mail/calendar folders or ones created by the user or the Top of Information Store
                    #The filters for the FolderPermission ignore when someone is blocked and make sure it is not the mailbox owner where possible and excludes the normal permissions a calendar has: Default users get AvailabilityOnly
                    #The Select-Object calculations are there in the middle to force the appropriate format needed to send to the FolderPermission command, the second Select-Object formats for file output
                    #Replacing [char]63743 is used to fix forward slashes that get interpreted as a box with a question mark in it. This will be fixed once we can use folderId instead of a path for folders
                    Get-EXOMailbox -Filter "whenChangedUTC -gt '$DateForDelta'" -ResultSize Unlimited -Properties ExternalDirectoryObjectId |
                        Tee-Object -FilePath "$PSScriptRoot\Temp-Resume-NoFoldersFoundInFile-IncludeFolders-Mailboxes.txt" |
                        Get-EXOMailboxFolderStatistics |
                        Where-Object {
                            $_.SearchFolder -eq $false -and
                            @("Root","Calendar","Inbox","User Created") -contains $_.FolderType -and
                            (@("IPF.Note","IPF.Appointment",$null) -contains $_.ContainerClass -or $_.Name -eq "Top of Information Store")} |
                        Select-Object @{Label="Identity";Expression={
                            if($_.Name -eq "Top of Information Store"){
                                $_.Identity.Substring(0,$_.Identity.IndexOf("\"))
                            } else {
                                $_.Identity.Substring(0,$_.Identity.IndexOf("\"))+':'+$_.Identity.Substring($_.Identity.IndexOf("\")).Replace([char]63743,"/")
                            }}} |
                        Get-EXOMailboxFolderPermission |
                        Where-Object {
                            $_.AccessRights -ne "None" -and
                            @("NT AUTHORITY\SELF") -notcontains $_.User -and
                            $_.User -ne ($_.Identity.Substring(0,$_.Identity.IndexOf(":\"))) -and
                            !($_.User -eq "Default" -and $_.AccessRights -eq "AvailabilityOnly")} |
                        Select-Object @{Label="Mailbox";Expression={$_.Identity.Substring(0,$_.Identity.IndexOf(":\"))}},
                            @{Label="FolderPath";Expression={
                                if($_.FolderName -eq "Top of Information Store"){
                                    "Top of Information Store"
                                } else {
                                    $_.Identity.Substring($_.Identity.IndexOf(":\")+2)
                                }
                            }},
                            @{Label="UserGivenAccess";Expression={Test-ObjectId -ObjectId $_.User}},
                            @{Label="AccessRights";Expression={$_.AccessRights -join ","}} |
                        Where-Object { $_.UserGivenAccess -ne ($_.Mailbox+"$EmailDomain")} |
                        Export-CSV -Path $ResumeCSVPath -Append    
                    #Grab the text file generated from Tee-Object and count the IDs so we can count how many mailboxes were processed, then remove the file
                    $TotalMailboxesProcessed += (Get-Content "$PSScriptRoot\Temp-Resume-NoFoldersFoundInFile-IncludeFolders-Mailboxes.txt" |
                        Where-Object { $_ -match "^ExternalDirectoryObjectId :"}).Count
                    Remove-Item "$PSScriptRoot\Temp-Resume-NoFoldersFoundInFile-IncludeFolders-Mailboxes.txt" -Force -Confirm:$false
                } catch {
                    throw
                }
            }
        }

        #Since we are resuming, there were some changes recently and so we need to deduplicate the results and create a final product, make sure that the Creation date/time shows the most recent creation data, then remove the temp stuff
        if((Read-Host "Run has completed. Please review file at $ResumeCSVPath before entries are merged/sorted. Ready to merge/sort ? Type 'y' and press 'enter'.") -eq "y"){
            $NewEntries = (Import-CSV $ResumeCSVPath).Mailbox | Select-Object -Unique
            if($IncludeFolders -or $FoldersFoundInFile){
                @(Import-CSV $ResumeCSVPath) +
                    @(if(Test-Path $CustomPermsCSVPathNew){
                        Import-CSV $CustomPermsCSVPath | Where-Object { $NewEntries -notcontains $_.Mailbox}
                    }) |
                    Sort-Object -Property Mailbox,FolderPath,UserGivenAccess |
                    Export-CSV $CustomPermsCSVPath -Force
            } else {
                @(Import-CSV $ResumeCSVPath) +
                    @(if(Test-Path $CustomPermsCSVPathNew){
                        Import-CSV $CustomPermsCSVPath | Where-Object { $NewEntries -notcontains $_.Mailbox}
                    }) +
                    @(if(Test-Path $CustomPermsCSVPathNew){
                        Import-CSV $CustomPermsCSVPath | Where-Object { $_.FolderPath.Length -gt 0}
                    }) |
                    Sort-Object -Property Mailbox,FolderPath,UserGivenAccess |
                    Export-CSV $CustomPermsCSVPath -Force
            }
            if($ResumeCSVPath -ne $CustomPermsCSVPath){
                $CustomPermsCSVFile = Get-Item $CustomPermsCSVPath
                $CustomPermsCSVFile.CreationTime = (Get-Item $ResumeCSVPath).CreationTime
                Remove-Item $ResumeCSVPath -Force -Confirm:$false
            }
        }
    } elseif($ProcessNewlyCreatedOrChangedMailboxesOnly){
        #Tell the user that we are starting now and what we are doing
        $Message += " Since $DateForDelta (UTC)"
        Write-Host $Message
        #Create a file called NEW to indicate this is only recently changed data. Once we get the NEW data, we deduplicate and merge into the main list of perms entries and then remove the NEW file since it was temporary
        $CustomPermsCSVPathNew = ($CustomPermsCSVPath.Substring(0,$CustomPermsCSVPath.LastIndexOf("."))+'-NEW.csv')
        try {
            #The filters for MailboxPermission try to ignore inherited permissions, users who are blocked, and the mailbox owner to find truly other users with permissions
            #The Select-Object is to format the file output
            Get-EXOMailbox -Filter "whenChangedUTC -gt '$DateForDelta'" -ResultSize $GetMailboxResultSize -Properties ExternalDirectoryObjectId |
                Tee-Object -FilePath "$PSScriptRoot\TEMP-ProcessNewOnly-Mailboxes.txt" |
                Get-EXOMailboxPermission -ExternalDirectoryObjectId $_.ExternalDirectoryObjectId -ResultSize Unlimited |
                Where-Object {
                    $_.IsInherited -eq $false -and
                    $_.Deny -eq $false -and
                    @("NT AUTHORITY\SELF") -notcontains $_.User -and
                    $_.User -ne ($_.Identity+"$EmailDomain")} |
                Select-Object @{Label="Mailbox";Expression={$_.Identity}},
                    @{Label="FolderPath";Expression={''}},
                    @{Label="UserGivenAccess";Expression={Test-ObjectId -ObjectId $_.User}},
                    @{Label="AccessRights";Expression={$_.AccessRights -join ","}} |
                Export-CSV -Path $CustomPermsCSVPathNew -Append
            #Grab the text file generated from Tee-Object and keep the mailbox IDs initially for counting and then use if we need to process folders from these mailboxes
            $arrMailboxes = Get-Content "$PSScriptRoot\TEMP-ProcessNewOnly-Mailboxes.txt" |
                Where-Object { $_ -match "^ExternalDirectoryObjectId :"} |
                Foreach-Object { [pscustomobject]@{"ExternalDirectoryObjectId"=($_ -split ":")[1].Trim()} }
            $TotalMailboxesProcessed += $arrMailboxes.Count
        } catch {
            throw
        }
        if($IncludeFolders){
            try{  
                #The filters for the FolderStatistics try to focus on normal mail/calendar folders or ones created by the user or the Top of Information Store
                #The filters for the FolderPermission ignore when someone is blocked and make sure it is not the mailbox owner where possible and excludes the normal permissions a calendar has: Default users get AvailabilityOnly
                #The Select-Object calculations are there in the middle to force the appropriate format needed to send to the FolderPermission command, the second Select-Object formats for file output
                #Replacing [char]63743 is used to fix forward slashes that get interpreted as a box with a question mark in it. This will be fixed once we can use folderId instead of a path for folders
                $arrMailboxes |
                    Get-EXOMailboxFolderStatistics |
                    Where-Object {
                        $_.SearchFolder -eq $false -and
                        @("Root","Calendar","Inbox","User Created") -contains $_.FolderType -and
                        (@("IPF.Note","IPF.Appointment",$null) -contains $_.ContainerClass -or $_.Name -eq "Top of Information Store")} |
                    Select-Object @{Label="Identity";Expression={
                        if($_.Name -eq "Top of Information Store"){
                            $_.Identity.Substring(0,$_.Identity.IndexOf("\"))
                        } else {
                            $_.Identity.Substring(0,$_.Identity.IndexOf("\"))+':'+$_.Identity.Substring($_.Identity.IndexOf("\")).Replace([char]63743,"/")
                        }
                    }} |
                    Get-EXOMailboxFolderPermission |
                    Where-Object {
                        $_.AccessRights -ne "None" -and
                        @("NT AUTHORITY\SELF") -notcontains $_.User -and
                        $_.User -ne ($_.Identity.Substring(0,$_.Identity.IndexOf(":\"))) -and
                        !($_.User -eq "Default" -and $_.AccessRights -eq "AvailabilityOnly")} |
                    Select-Object @{Label="Mailbox";Expression={$_.Identity.Substring(0,$_.Identity.IndexOf(":\"))}},
                        @{Label="FolderPath";Expression={
                            if($_.FolderName -eq "Top of Information Store"){
                                "Top of Information Store"
                            } else {
                                $_.Identity.Substring($_.Identity.IndexOf(":\")+2)
                            }
                        }},
                        @{Label="UserGivenAccess";Expression={Test-ObjectId -ObjectId $_.User}},
                        @{Label="AccessRights";Expression={$_.AccessRights -join ","}} |
                    Where-Object { $_.UserGivenAccess -ne ($_.Mailbox+"$EmailDomain")} |
                    Export-CSV -Path $CustomPermsCSVPathNew -Append
                $TotalMailboxesProcessed += $arrMailboxes.Count
            } catch {
                throw
            }
        }
        Remove-Item "$PSScriptRoot\TEMP-ProcessNewOnly-Mailboxes.txt" -Force -Confirm:$false
        Remove-Variable -Name "arrMailboxes"
        #Since we are processing only recent changes, we need to deduplicate the results and create a final product, then remove the temp stuff
        if((Read-Host "Run has completed. Please review file at $CustomPermsCSVPathNew before entries are merged. Ready to merge back into one file? Type 'y' and press 'enter'.") -eq "y"){
            $NewEntries = (Import-CSV $CustomPermsCSVPathNew).Mailbox | Select-Object -Unique
            if($IncludeFolders){
                @(Import-CSV $CustomPermsCSVPathNew) +
                    @(Import-Csv $CustomPermsCSVPath | Where-Object { $NewEntries -notcontains $_.Mailbox }) |
                    Sort-Object -Property Mailbox,FolderPath,UserGivenAccess |
                    Export-CSV $CustomPermsCSVPath -Force
            } else {
                @(Import-CSV $CustomPermsCSVPathNew) +
                    @(Import-Csv $CustomPermsCSVPath | Where-Object { $NewEntries -notcontains $_.Mailbox }) +
                    @(Import-Csv $CustomPermsCSVPath | Where-Object { $_.FolderPath.Length -gt 0 }) |
                    Sort-Object -Property Mailbox,FolderPath,UserGivenAccess |
                    Export-CSV $CustomPermsCSVPath -Force
            }
            $CustomPermsCSVFile = Get-Item $CustomPermsCSVPath
            $CustomPermsCSVFile.CreationTime = (Get-Item $CustomPermsCSVPathNew).CreationTime
            Remove-Item $CustomPermsCSVPathNew -Force
        }
    } else {
        #We aren't resuming and we aren't processing only recently created/changed mailboxes, so this is a fresh/full download of data for the entire tenant
        Write-Host $Message
        try {
            #The filters for MailboxPermission try to ignore inherited permissions, users who are blocked, and the mailbox owner to find truly other users with permissions
            #The Select-Object is to format the file output
            Get-EXOMailbox -ResultSize $GetMailboxResultSize -Properties ExternalDirectoryObjectId |
                Tee-Object -FilePath "$PSScriptRoot\TEMP-FullDownload-Mailboxes.txt" |
                Get-EXOMailboxPermission -ExternalDirectoryObjectId $_.ExternalDirectoryObjectId -ResultSize Unlimited |
                Where-Object {
                    $_.IsInherited -eq $false -and
                    $_.Deny -eq $false -and
                    @("NT AUTHORITY\SELF") -notcontains $_.User -and
                    $_.User -ne ($_.Identity+"$EmailDomain")} |
                Select-Object @{Label="Mailbox";Expression={$_.Identity}},
                    @{Label="FolderPath";Expression={''}},
                    @{Label="UserGivenAccess";Expression={Test-ObjectId -ObjectId $_.User}},
                    @{Label="AccessRights";Expression={$_.AccessRights -join ","}} |
                Export-CSV -Path $CustomPermsCSVPath -Force
            #Grab the text file generated from Tee-Object and keep the mailbox IDs initially for counting and then use if we need to process folders from these mailboxes
            $arrMailboxes = Get-Content "$PSScriptRoot\TEMP-FullDownload-Mailboxes.txt" |
                Where-Object { $_ -match "^ExternalDirectoryObjectId :"} |
                Foreach-Object { [pscustomobject]@{"ExternalDirectoryObjectId"=($_ -split ":")[1].Trim()} }
            $TotalMailboxesProcessed += $arrMailboxes.Count
        } catch {
            throw
        }
        if($IncludeFolders){
            try{
                #The filters for the FolderStatistics try to focus on normal mail/calendar folders or ones created by the user or the Top of Information Store
                #The filters for the FolderPermission ignore when someone is blocked and make sure it is not the mailbox owner where possible and excludes the normal permissions a calendar has: Default users get AvailabilityOnly
                #The Select-Object calculations are there in the middle to force the appropriate format needed to send to the FolderPermission command, the second Select-Object formats for file output
                #Replacing [char]63743 is used to fix forward slashes that get interpreted as a box with a question mark in it. This will be fixed once we can use folderId instead of a path for folders
                $arrMailboxes |
                    Get-EXOMailboxFolderStatistics |
                    Where-Object {
                        $_.SearchFolder -eq $false -and
                        @("Root","Calendar","Inbox","User Created") -contains $_.FolderType -and
                        (@("IPF.Note","IPF.Appointment",$null) -contains $_.ContainerClass -or $_.Name -eq "Top of Information Store")} |
                    Select-Object @{Label="Identity";Expression={
                        if($_.Name -eq "Top of Information Store"){
                            $_.Identity.Substring(0,$_.Identity.IndexOf("\"))
                        } else {
                            $_.Identity.Substring(0,$_.Identity.IndexOf("\"))+':'+$_.Identity.Substring($_.Identity.IndexOf("\")).Replace([char]63743,"/")
                        }
                    }} |
                    Get-EXOMailboxFolderPermission |
                    Where-Object {
                        $_.AccessRights -ne "None" -and
                        @("NT AUTHORITY\SELF") -notcontains $_.User -and
                        $_.User -ne ($_.Identity.Substring(0,$_.Identity.IndexOf(":\"))) -and
                        !($_.User -eq "Default" -and $_.AccessRights -eq "AvailabilityOnly")
                    } |
                    Select-Object @{Label="Mailbox";Expression={$_.Identity.Substring(0,$_.Identity.IndexOf(":\"))}},
                        @{Label="FolderPath";Expression={
                            if($_.FolderName -eq "Top of Information Store"){
                                "Top of Information Store"
                            } else {
                                $_.Identity.Substring($_.Identity.IndexOf(":\")+2)
                            }
                        }},
                        @{Label="UserGivenAccess";Expression={Test-ObjectId -ObjectId $_.User}},
                        @{Label="AccessRights";Expression={$_.AccessRights -join ","}} |
                    Where-Object { $_.UserGivenAccess -ne ($_.Mailbox+"$EmailDomain")} |
                    Export-CSV -Path $CustomPermsCSVPath -Append
                $TotalMailboxesProcessed += $arrMailboxes.Count
                #Since this was a fresh attempt and the Folders were added to the entries, we want to sort everything so the folders appear alongside the mailbox permissions
                $CustomPermsCSV = Import-Csv -Path $CustomPermsCSVPath
                $CustomPermsCSV | Sort-Object Mailbox,FolderPath,UserGivenAccess | Export-CSV -Path $CustomPermsCSVPath -Force
            } catch {
                throw
            }
        }
        Remove-Item "$PSScriptRoot\TEMP-FullDownload-Mailboxes.txt" -Force -Confirm:$false
        Remove-Variable -Name "arrMailboxes"
        $CustomPermsCSVFile = Get-Item $CustomPermsCSVPath
        $CustomPermsCSVFile.CreationTime = $StartTime
    }
} catch {
    Write-Host "Error - $_"
    Read-Host "$(Get-Date -format filedatetime) - The command to get mailboxes or permissions stopped due to error. Sorry about that"
    Get-PSSession | Remove-PSSession
    exit
}
Get-PSSession | Remove-PSSession
$Timer.Stop()
Write-host "Done, the runtime for this entire process was"($timer.Elapsed.TotalMinutes)"minutes. Total Mailboxes processed (whether mailboxes, folders, or both): $TotalMailboxesProcessed. Total Errors $($Error.Count)"
Stop-Transcript
Read-Host "Press any key to exit"
exit