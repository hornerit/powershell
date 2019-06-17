<#
.SYNOPSIS
Large Group AD Rollup Group Sync to Azure, other group sync for misc reasons: Synchronizes AD groups with a mirror rollup group, possibly creating subgroups of a specified size; optionally, it can pull users from nested groups inside the source AD group
 
.DESCRIPTION
Uses hash tables to improve comparison performance and manages groups. Note that the subgroups must be in the format of GroupName_# and members of the parent group if they are created before running this script.
 
.PARAMETER SrcGrp
REQUIRED Specifies an AD group username (not diplay name) that contains the MASTER list of users - does NOT expand groups within this group by default
.PARAMETER DestGrp
REQUIRED Specifies the username of the parent group whose members should roll up to match SrcGrp members; it will create if group does not exist; if a user is not a member of the SrcGrp but is found in DestGrp, they are removed from DestGrp
.PARAMETER MaxGroupSize
REQUIRED Specify the maximum number of users for your receiving system. For Azure, I have been using 45000.
.PARAMETER ExpandNestedGroups
OPTIONAL Add this switch to have nested groups inside the source AD group expanded to find all of the possible members
.PARAMETER IgnoreIfInGroup
OPTIONAL Use this to remove a user from the process if they are already in another specific AD group. This is being used for groups related to licensing but can be useful for ordinal group membership (e.g. if VIP, put in VIP group DONE; if Engineer but already in VIP, ignore, otherwise, add to Engineer group)
 
.NOTES
  Created by: Brendan Horner (hornerit)
  Credit: Learned comparing hashtables for performance from CompareObjects2 function created by Ashley McGlone
  Version History:
  --2019-04-11-Initial Public version
   
.EXAMPLE
Giant-Group is too big to sync to Azure and has groups inside of it, create a separate set of groups that mirrors Giant-Group and all of its nested members so that I can use those groups in Azure
SplitAndOrSyncHugeADGroups -SrcGrp 'Giant-Group' -DestGrp "Azure-Licensing-Group" -ExpandNestedGroups -MaxGroupSize 45000
 
.EXAMPLE
An Application we use is limited to AD groups of 5000 or less and our 'MyCompany' AD Group is too large and contains nested groups, create a set of groups that mirror 'MyCompany' but only have a max of 5000 users. Additionally, this app is not supposed to be used by VIPs, so exclude them from the process and make sure they aren't in the created mirror groups.
SplitAndOrSyncHugeADGroups -SrcGrp 'MyCompany' -DestGrp "Azure-MyCompany" -ExpandNestedGroups -IgnoreIfInGroup VIP -MaxGroupSize 5000
 
.EXAMPLE
An application we use only allows users and does not support nested subgroups but does not care how many users are in the group. Create a flattened version of the source group so that we can use the flattened version for our application.
SplitAndOrSyncHugeADGroups -SrcGrp 'MyGroup' -DestGrp 'MyGroup-Flattened' -ExpandNestedGroups -MaxGroupSize 999999999
#>
 
#VARIABLES YOU NEED TO SET IN THIS SCRIPT
$NotificationTo = "YourEmailAddress@YourDomain.com"
$NotificationFrom = "AnotherEmailAddress@YourDomain.com"
$NotificationSmtpServer = 0.0.0.0
$TranscriptLogFolderPath = "C:\yourFolder" #(this folder must already exist)
You will specify the running of this script function at the bottom
#DONE WITH VARIABLES YOU NEED TO SET
  
function SplitAndOrSyncHugeADGroups {
    param(
    [Parameter(Mandatory=$true)]
    [string]
    $SrcGrp,
    [Parameter(Mandatory=$true)]
    [string]
    $DestGrp,
    [Parameter(Mandatory=$true)]
    [int]
    $MaxGroupSize,
    [switch]
    $ExpandNestedGroups,
    [string[]]
    $IgnoreIfInGroup
    )
    #Get the SrcGrp AD group and its list of members, hold on to the name of it for later use
    write-host "**********Beginning SplitAndSync for $SrcGrp to $DestGrp"
    write-host "Getting source group $SrcGrp..." -NoNewline
    $SourceADGroup = Get-ADGroup -Identity $SrcGrp -Properties members -ErrorAction Stop
    $SrcName = $SourceADGroup.Name
    write-host "Done"
   
    write-host "Checking for $DestGrp..." -NoNewline
    #Check to see if DestGrp exists; if not, create it
    try {
        Get-ADGroup -Identity $DestGrp -Properties members -ErrorAction Stop | Out-Null
    }
    catch {
        write-host "`n  Rollup group does not exist, creating..." -NoNewline
        $Path = ($SourceADGroup.DistinguishedName -split ",",2)[1]
        $Description = "Rollup group containing subgroups to match $SrcName"
        New-ADGroup -Name $DestGrp -GroupScope 2 -Description $Description -GroupCategory 1 -Path $Path -DisplayName $DestGrp -SamAccountName $DestGrp -ErrorVariable errNewGroup
        if($errNewGroup[0].Message.length -gt 0){
            write-host "`n  Quitting script, there was an error creating the main destination group..."
            $ScriptName = split-path $MyInvocation.PSCommandPath -Leaf
            Send-MailMessage `
            -To $NotificationTo `
            -From $NotificationFrom `
            -Subject "Error creating group from script" `
            -Body "There was an error trying to create group titled $DestGrp via $ScriptName script on $env:COMPUTERNAME" `
            -SmtpServer $NotificationSmtpServer
            Stop-Transcript
            exit
        }
        write-host "Done"
    }
    $RollupADGroup = Get-ADGroup -Identity $DestGrp -Properties members -ErrorAction Stop
    write-host "Done"
   
    write-host "Loading all Source group members into memory..." -NoNewline
    #Create hash table for all users in the master group, then remove the group object to save memory
    $RefHash = @{}
    $GrpHash = @{}
    foreach($dn in $SourceADGroup.members){
        $RefHash.Add($dn,$SourceADGroup.Name)
    }
    if($ExpandNestedGroups){
        #Use LDAP Query language to get all GROUP descendants of the source group
        $LDAPQry = "(&(objectClass=group)(memberOf:1.2.840.113556.1.4.1941:="+$SourceADGroup.DistinguishedName+"))"
        $SrcChildGrps = @(Get-ADGroup -LDAPFilter $LDAPQry -Properties members)
        #If there are any group descendants: add them to a list of groups, get their members and add to the list of child group members
        if($SrcChildGrps.count -gt 0 -or $SrcChildGrps.gettype().Name -eq "ADGroup"){
            Write-Host "`n  Expanding child groups to load into memory..."
            foreach($grp in $SrcChildGrps){
                if(!($GrpHash.ContainsKey($grp.DistinguishedName))){
                    $GrpHash.Add($grp.DistinguishedName,$null)
                }
            }
            write-host " "$GrpHash.Count"descendant groups were found"
            for($i=0;$i -lt $SrcChildGrps.count;++$i){
                $DN = $SrcChildGrps[$i].DistinguishedName
                $subCounter = 0
                foreach($member in (Get-ADGroup -Identity $DN -Properties members).members){
                    if((!($RefHash.ContainsKey($member))) -and (!($GrpHash.ContainsKey($member)))){
                        $RefHash.Add($member,$SrcChildGrps[$i].Name)
                        $subCounter++
                    }
                }
                write-host "    Found $subCounter not-previously-found users from"$SrcChildGrps[$i].Name
                $RefHash.Remove($DN)
            }
            Write-Host "  Done"
        }
    }
    Remove-Variable -Name SourceADGroup
    #In the process of handling the destination groups, we need to know if we even need subgroups based on the max group size
    $NeedsSubgroups = $false
    if($RefHash.Keys.Count -gt $MaxGroupSize){
        $NeedsSubgroups = $true
    }
    write-host "Done,"$RefHash.count"users added from $SrcGrp" -NoNewline
   
    write-host "Loading rollup group members into memory..." -NoNewline
    $DifHash = @{}
    $RollupGrpHash = @{}
   
    #Get all subgroups within the destination or rollup group
    $RollupLDAPQry = "(&(objectClass=group)(memberOf:1.2.840.113556.1.4.1941:="+$RollupADGroup.DistinguishedName+"))"
    $RollupChildGrps = @(Get-ADGroup -LDAPFilter $RollupLDAPQry)
    foreach($grp in $RollupChildGrps.DistinguishedName){
        if(!($RollupGrpHash.ContainsKey($grp))){
            $RollupGrpHash.add($grp,$null)
        }
    }
   
    #Loop thru the sub groups of the rollup AD group, get the users inside them, then shove them all into a hash table
    foreach($member in $RollupADGroup.members){
        #If the member found in the rollup group is an actual group found in the previous step, get their non-group members
        if($RollupGrpHash.ContainsKey($member)){
            $DestGrpSub = Get-ADGroup -Identity $member -Properties members -ErrorAction Stop
            foreach($dn in $DestGrpSub.members){
                try {
                    if(!($RollupGrpHash.ContainsKey($dn))){
                        $DifHash.add($dn,$DestGrpSub.Name)
                    }
                } catch {
                    $username = $dn.substring(3,$dn.indexOf(",")-3)
                    write-host "`n  $username in multiple groups, removing..." -NoNewline
                    Remove-ADGroupMember -Identity $DestGrpSub.DistinguishedName -Members $dn -Confirm:$false
                    write-host "Done" -NoNewline
                }
            }
        } else {
            #The entity we tried to load as an AD group is most likely a user that was in the rollup group, go ahead and add them to the hash as being part of the parent group
            try {
                if($NeedsSubgroups){
                    throw
                } else{
                    $DifHash.add($member,$RollupADGroup.Name)
                }
            } catch {
                $username = $member.substring(3,$member.indexOf(",")-3)
                #Depending on whether our rollup group actually has enough members to need more than one group, we remove the user from the parent group or the child group
                if($NeedsSubgroups){
                    write-host "`n  $username in parent group, removing from parent..." -NoNewline
                    Remove-ADGroupMember -Identity $RollupADGroup.DistinguishedName -Members $member -Confirm:$false
                    write-host "Done" -NoNewline
                } else {
                    write-host "`n  $username in child and parent groups but child groups are not necessary, removing from child group..." -NoNewline
                    Remove-ADGroupMember -Identity ($DifHash.$member) -Members $member -Confirm:$false
                    write-host "Done" -NoNewline
                }
            }
        }
    }
    write-host "Done,"$DifHash.count"users added from $DestGrp" -NoNewline
   
    #If the user is in an Ignore group, remove them from the process of pushing...this will cause them to be removed from existing groups if they are already there, DOES NOT EXPAND NESTED
    foreach($grp in $IgnoreIfInGroup){
        write-host "Calculating users to ignore or skip from source group who are immediate children in $grp..." -NoNewline
        try{
            $IgnoreGroup = Get-ADGroup -Identity $grp -Properties members -ErrorAction Stop
        } catch {
            write-host "`n  The AD group supplied in $grp to be ignored is invalid"
            continue
        }
        foreach($dn in $IgnoreGroup.members){
            if($RefHash.ContainsKey($dn)){
                $RefHash.Remove($dn)
            }
        }
        write-host "Done" -NoNewline
    }
   
    write-host "Removing users already synchronized from this process..." -NoNewline
    #If the user is in both the source and destination groups, remove from this process (we are only looking for unmatched users). Must use Clone function because you can't remove an item while you are looking directly at it
    $totalAlreadySyncd = 0
    $RefHash.keys.Clone() | ForEach-Object {
        If ($DifHash.ContainsKey($_)) {
            $DifHash.Remove($_)
            $RefHash.Remove($_)
            $totalAlreadySyncd++
        }
    }
    write-host "Done, $totalAlreadySyncd were already synchronized" -NoNewline
   
    write-host "Removing"$DifHash.count"users from group or groups as they should no longer be there..." -NoNewline
    #If the user is ONLY in the child group, they aren't on the master/source list and should be removed from their group, freeing slots for others
    $UsersRemoved = "`nUsers removed from groups:"
    $DifHash.GetEnumerator() | Foreach-Object {
        $Usr = Get-ADUser -Identity $($_.key)
        try {
            Remove-ADGroupMember -Identity $($_.value) -Members $Usr -Confirm:$false -ErrorAction Stop
            $UsersRemoved+=($usr.SamAccountName + ",")
        } catch {
            write-host "  ERROR REMOVING"$usr.SamAccountName
        }
    }
    if($UsersRemoved.length -gt 29){
        $UsersRemoved=$UsersRemoved.substring(0,$UsersRemoved.length)+"`n"
        write-host $UsersRemoved
    }
    write-host "Done"
   
    #Check to see if there are any users needing to be pushed into a child group; if so, do so.
    if($RefHash.Count -gt 0){
        write-host "Determining available slots..." -NoNewline
        #Since removing unneeded users, find out how many slots are available per child group and store in a hash table
        $AvailableSlots = @{}
        if($NeedsSubgroups){
            foreach($grp in $RollupADGroup.members){
                $DestGrpSub = Get-ADGroup -Identity $grp -Properties members
                if($DestGrpSub.members.count -lt $MaxGroupSize){
                    $AvailableSlots.add($DestGrpSub.Name,$MaxGroupSize-$DestGrpSub.members.count)
                }else{
                    if($DestGrpSub.members.count -gt $MaxGroupSize){
                        write-host "`n "$DestGrpSub.Name"is too big, shrinking..." -NoNewline
                        $min = $MaxGroupSize
                        $max = $DestGrpSub.members.count-1
                        $total = $max-$min
                        foreach($member in $DestGrpSub.members[$min..$max]){
                            $RefHash.Add($member,$DestGrpSub.Name)
                        }
                        Remove-ADGroupMember -Identity $DestGrpSub -Members $DestGrpSub.members[$min..$max] -Confirm:$false
                        write-host "Done, $total users removed to be sent elsewhere"
                    }
                }
            }
        } else {
            $AvailableSlots.add($RollupADGroup.Name,$MaxGroupSize-$RollupADGroup.members.count)
        }
        write-host "Done"
   
        write-host "Filling"$RefHash.Count"users into available slots in rollup groups..."
        #Look at the slots available in each row in the hash table (sorted alphabetically) and then fill them from users, then remove those users from the list of users to add
        foreach($slot in ($AvailableSlots.GetEnumerator() | Sort-Object -Property Name)){
            if($RefHash.Count -gt 0){
                $UsrsToAdd = New-Object System.Collections.ArrayList
                $TotalUsers = 0
                write-host "  Creating list of users to add to"$slot.key
                $RefHash.GetEnumerator() | Get-Random -Count $slot.value | ForEach-Object {
                    [void]$UsrsToAdd.add("$($_.key)")
                    $RefHash.Remove($($_.key))
                    $TotalUsers++
                }
                for($i=$UsrsToAdd.count-1;$i -ge 0;$i-=5000){
                    if($i-4999 -lt 0){ $min = 0 } else { $min = $i-4999 }
                    $GrpName = $slot.key
                    write-host "  Adding"($min+1)"thru"($i+1)"users to $GrpName..." -NoNewline
                    Add-ADGroupMember -Identity $GrpName -Members $UsrsToAdd[$min..$i]
                    write-host "Done"
                }
            }
        }
        write-host "Done"
   
        #If there are leftover users, create new groups and populate them
        if($RefHash.count -gt 0 -and $NeedsSubgroups){
            write-host "There are still"$RefHash.Count"users who need a group but no slots are available, creating subgroups..."
            #If we had to create the rollup group earlier, there are no child groups - create the first one
            if($RollupADGroup.members.count -eq 0){
                $Path = ($RollupADGroup.DistinguishedName -split ",",2)[1]
                $RollupName = $RollupADGroup.Name
                $SubGroupName = $RollupADGroup.Name + "_1"
                $Description = "Child group of $RollupName, related to $SrcName"
                write-host "  Creating $SubGroupName and adding to $RollupName..." -NoNewline
                $NewGrp = New-ADGroup -Name $SubGroupName -GroupScope 2 -Description $Description -GroupCategory 1 -Path $Path -DisplayName $SubGroupName -SamAccountName $SubGroupName -PassThru -ErrorVariable errNewGroup
                if($errNewGroup[0].Message.length -gt 0){
                    write-host "`n Quitting script, there was an error creating the $SubGroupName..."
                    $ScriptName = split-path $MyInvocation.PSCommandPath -Leaf
                    Send-MailMessage `
                    -To $NotificationTo `
                    -From $NotificationFrom `
                    -Subject "Error creating group from script" `
                    -Body "There was an error trying to create group titled $SubGroupName via $ScriptName script on $env:COMPUTERNAME" `
                    -SmtpServer $NotificationSmtpServer
                    Stop-Transcript
                    exit
                }
                Add-ADGroupMember -Identity $RollupADGroup -Members $NewGrp
                $AvailableSlots.add("$SubGroupName",$MaxGroupSize)
                write-host "Done"
                $MaxGrp = 1
                #Create additional groups if needed
                for($i=1; $i -le ([math]::Truncate($RefHash.count/$MaxGroupSize));$i++){
                    $GrpNum = $MaxGrp + $i
                    $GrpName = $RollupName + "_" + $GrpNum
                    write-host "  Creating $GrpName and adding to $RollupName..." -NoNewline
                    $NewGrp = New-ADGroup -Name $GrpName -GroupScope 2 -Description "Child group of $RollupName, related to $SrcName" -DisplayName $GrpName -GroupCategory 1 -Path $Path -SamAccountName $GrpName -PassThru -ErrorVariable errNewGroup
                    if($errNewGroup[0].Message.length -gt 0){
                        write-host "`n Quitting script, there was an error creating the $GrpName..."
                        $ScriptName = split-path $MyInvocation.PSCommandPath -Leaf
                        Send-MailMessage `
                        -To $NotificationTo `
                        -From $NotificationFrom `
                        -Subject "Error creating group from script" `
                        -Body "There was an error trying to create group titled $GrpName via $ScriptName script on $env:COMPUTERNAME" `
                        -SmtpServer $NotificationSmtpServer
                        Stop-Transcript
                        exit
                    }
                    Add-ADGroupMember -Identity $RollupADGroup -Members $NewGrp
                    $AvailableSlots.add("$GrpName",$MaxGroupSize)
                    write-host "Done"
                }
            } else {
                #Get the subgroup with the largest number on the end by getting all DestGrp members, extracting numbers, throwing them into hashtable, sorting descending, then grab top number
                $MaxGrpInt = @{}
                $GrpCounter = 0
                foreach($Grp in $RollupADGroup.members){
                    $GrpCounter++
                    $GrpName = $Grp.substring(0,$Grp.indexOf(","))
                    $GrpNumber = [int]$GrpName.substring($GrpName.lastindexof("_")+1)
                    $MaxGrpInt.add($GrpCounter,$GrpNumber)
                }
                $MaxGrp = ($MaxGrpInt.GetEnumerator() | Sort-Object value -Descending | Select-Object -First 1).value
                #Empty hash table to hold the new groups and how many slots they have
                $AvailableSlots = @{}
                $Path = ($RollupADGroup.DistinguishedName -split ",",2)[1]
                $RollupName = $RollupADGroup.Name
   
                #Create groups
                for($i=1; $i -le ([math]::Truncate($RefHash.count/$MaxGroupSize)+1);$i++){
                    $GrpNum = $MaxGrp + $i
                    $GrpName = $RollupName + "_" + $GrpNum
                    Write-Host "  Creating $GrpName and adding to $RollupName..." -NoNewline
                    $NewGrp = New-ADGroup -Name $GrpName -GroupScope 2 -Description "Child group of $RollupName, related to $SrcName" -DisplayName $GrpName -GroupCategory 1 -Path $Path -SamAccountName $GrpName -PassThru -ErrorVariable errNewGroup
                    if($errNewGroup[0].Message.length -gt 0){
                        write-host "`n Quitting script, there was an error creating the $GrpName..."
                        $ScriptName = split-path $MyInvocation.PSCommandPath -Leaf
                        Send-MailMessage `
                        -To $NotificationTo `
                        -From $NotificationFrom `
                        -Subject "Error creating group from script" `
                        -Body "There was an error trying to create group titled $GrpName via $ScriptName script on $env:COMPUTERNAME" `
                        -SmtpServer $NotificationSmtpServer
                        Stop-Transcript
                        exit
                    }
                    Add-ADGroupMember -Identity $RollupADGroup -Members $NewGrp
                    $AvailableSlots.add("$GrpName",$MaxGroupSize)
                    Write-Host "Done"
                }
            write-host "  Sending message to IT about the new groups so that they can do anything needed further like Azure licensing adjustments or policies that are group-bound" -NoNewline
            Send-MailMessage `
            -To $NotificationTo `
            -From $NotificationFrom `
            -Subject "New AD Group Created from $env:COMPUTERNAME" `
            -Body "$SrcGrp being split and/or syncrhonized by script on $env:COMPUTERNAME had to create a new subgroup in $DestGrp to handle more users. If you use Azure group licensing or a similar product that does not support nested groups, please update these systems accordingly to handle this new group in addition to the old ones: $GrpName" `
            -SmtpServer $NotificationSmtpServer
            }
            write-host "Done"
   
            write-host "Populating newly created subgroups..."
            #Fill the groups, same logic as above, 5k at a time due to typical AD web services constraints
            foreach($slot in ($AvailableSlots.GetEnumerator() | Sort-Object -Property Name)){
                $UsrsToAdd = New-Object System.Collections.ArrayList
                $TotalUsers = 0
                write-host "  Creating list of users to add to"$slot.key"..." -NoNewline
                $RefHash.GetEnumerator() | Get-Random -Count $slot.value | ForEach-Object {
                    [void]$UsrsToAdd.add("$($_.key)")
                    $RefHash.Remove($($_.key))
                    $TotalUsers++
                }
                write-host "Done"
                for($i=$UsrsToAdd.count-1;$i -ge 0;$i-=5000){
                    if($i-4999 -lt 0){ $min = 0 } else { $min = $i-4999 }
                    $GrpName = $slot.key
                    write-host "  Adding"($min+1)"thru"($i+1)"users to $GrpName" -NoNewline
                    Add-ADGroupMember -Identity $GrpName -Members $UsrsToAdd[$min..$i]
                    write-host "Done" -NoNewline
                }
            }
            write-host "Done"
        }
    } else {
        write-host "No users need to be added to subgroups at this time"
    }
}
Start-Transcript "$TranscriptLogFolderPath\RunningLog.txt" -Force -Append
#EXAMPLE: SplitAndOrSyncHugeADGroups -SrcGrp 'All-Staff' -DestGrp "Azure-Licensing-Staff" -MaxGroupSize 48000
Stop-Transcript