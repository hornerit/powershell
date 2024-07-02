<#
.SYNOPSIS
Obtain all DNS records for a given IPV4 address and allows for updating them in bulk.

.DESCRIPTION
This script will query the closest DNS server for all zones that it has, then it will query for all
"A" and "PTR" records within each zone, then will combine them all into a single hashtable based on IPv4 address.
Does NOT support IPv6. Once queried, presents a menu to allow one to update all A or PTR or both records for their
TimeToLive values (possibly others as time progresses). Assumes a single-forest environment and AD integrated
zones, mostly expecting Forest replication for all zones.

.NOTES
    Created by: Brendan Horner
    Version History:
    2024-07-02-Added an escape option when querying to return to menu
    2021-02-03-Couple of bug fixes, added Zone to the CSV exports, added timestamp to limited csv export
    2021-01-21-Made many updates to try to get the full feature set working over the last several days.
    2020-12-07-Initial version
    <#SAMPLE A RECORDS
        DistinguishedName : DC=MY-COMPUTER.subdomain.domain.com,DC=domain.com,cn=MicrosoftDNS,DC=ForestDnsZones
                            ,DC=domain,DC=com
        HostName          : MY-COMPUTER.subdomain.domain.com
        RecordType        : A
        Type              : 1
        RecordClass       : IN
        TimeToLive        : 00:20:00
        Timestamp         : 12/5/2020 8:00:00 AM
        RecordData        : 10.0.150.6
        ---------------------------------------------------
        DistinguishedName : DC=MY-COMPUTER.subdomain,DC=domain.com,cn=MicrosoftDNS,DC=ForestDnsZones,DC=domain,
                            DC=com
        HostName          : MY-COMPUTER.subdomain
        RecordType        : A
        Type              : 1
        RecordClass       : IN
        TimeToLive        : 00:20:00
        Timestamp         : 12/5/2020 8:00:00 AM
        RecordData        : 10.0.150.6
        ---------------------------------------------------
        The RecordData is misleading because it is actually a .NET IPv4Address object so it looks like this:
            IPv4Address       : 10.0.150.6
            PSComputerName    :
        And the IPv4Address is also an object consisting of something like this:
            Address           : 123456789 <--this is a calculated thing that makes IP Addresses sortable!!
            AddressFamily     : InterNetwork
            ScopeId           :
            IsIPv6Multicast   : False
            IsIPv6LinkLocal   : False
            IsIPv6SiteLocal   : False
            IsIPv6Teredo      : False
            IsIPv4MappedToIPv6: False
            IPAddressToString : 10.0.150.6 <--this is the thing you often want for scripting
    #>
    <#SAMPLE PTR:
        DistinguishedName : DC=6,DC=150.0.10.in-addr.arpa,cn=MicrosoftDNS,DC=ForestDnsZones,DC=domain,DC=com
        HostName          : 6
        RecordType        : PTR
        Type              : 12
        RecordClass       : IN
        TimeToLive        : 00:20:00
        Timestamp         : 8/8/2008 08:00:00 PM
        RecordData        : something.domain.com.
        ---------------------------------------------
    The RecordData is misleading in that it actually is an object that looks like this:
        PtrDomainName     : something.domain.com.
        PSComputerName    :
    Also note that PTR records for the host to which they are attached always ends in a period due to FQDN reqs
    #>
#>
#REQUIRES -Modules DNSServer
function Get-DNSRecordsFromTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$entry,
        [Parameter(Mandatory=$true)][hashtable]$allDnsRecords
    )
    if ($entry -match "\d+\.\d+\.\d+\.\d") { $entryType = "IP"} else { $entryType = "HOSTNAME" }
    #The DNS records are stored with the IP as the key so IP queries are efficient and simple
    if ($entryType -eq "IP") {
        if ($allDnsRecords.$entry.count -gt 0) {
            [System.Collections.DictionaryEntry]@{
                Name = $entry
                Key = $entry
                Value = $allDnsRecords.$entry
            }
        } else { $null }
    } else {
        #For hostname queries, we have to do some digging to find all the related records.
        $entryEscaped = [System.Text.RegularExpressions.Regex]::Escape($entry)
        ($allDnsRecords.GetEnumerator() |
            Foreach-Object {
                $Record = $_
                if ($Record.Value.count -eq 1) {
                    if ($Record.Value.Hostname -eq $entry -or
                    $Record.Value.RecordData.PtrDomainName -eq $entry -or
                    $Record.Value.Hostname -match "$entryEscaped\..*" -or
                    $Record.Value.RecordData.PtrDomainName -match "$entryEscaped\..*") {
                        $Record
                    }        
                } else {
                    $IncludeRecord = $false
                    foreach ($Value in $Record.Value) {
                        if ($Value.Hostname -eq $entry -or
                        $Value.RecordData.PtrDomainName -eq $entry -or
                        $Value.Hostname -match "$entryEscaped\..*" -or
                        $Value.RecordData.PtrDomainName -match "$entryEscaped\..*") {
                            $IncludeRecord = $true
                        }
                    }
                    if ($IncludeRecord) { $Record }
                } 
            }
        )
    }
}
function Get-AllDnsRecords {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$DC
    )
    $ptrIpRegex = "DC=(?<Part4>\d+)(\.|,DC=)(?<Part3>\d+)(\.|,DC=)(?<Part2>\d+)(\.|,DC=)(?<Part1>\d+)\.in-addr.*"
    Write-Verbose "FUNCTION Get-AllDnsRecords: $(Get-Date -format u) - Obtaining all 'A' Records..."
    $aRecords = Get-DnsServerZone -ComputerName $DC |
        Where-Object { $_.IsReverseLookupZone -eq $false -and
            $_.ZoneType -eq "Primary" -and
            $_.IsDsIntegrated -eq $true } |
        Tee-Object -Variable "aZones" |
        ForEach-Object {
            #Write-Verbose "FUNCTION Get-AllDnsRecords: Obtaining A for $($_.ZoneName)"
            Get-DnsServerResourceRecord -ComputerName $DC -ZoneName "$($_.ZoneName)" -RRType A
        }
    Write-Verbose "FUNCTION Get-AllDnsRecords: $(Get-Date -format u) - Done, $($aRecords.Count) found"
    Write-Verbose "FUNCTION Get-AllDnsRecords: $(Get-Date -format u) - Obtaining all 'PTR' Records..."
    $ptrRecords = Get-DNSServerZone -ComputerName $DC |
        Where-Object { $_.IsReverseLookupZone -eq $true -and
            $_.ZoneType -eq "Primary" -and
            $_.IsDsIntegrated -eq $true } |
        Tee-Object -Variable "ptrZones" |
        ForEach-Object {
            #Write-Verbose "FUNCTION Get-AllDnsRecords: Obtaining PTR for $($_.ZoneName)"
            Get-DnsServerResourceRecord -ComputerName $DC -ZoneName "$($_.ZoneName)" -RRType PTR
        }
    Write-Verbose -Message "FUNCTION Get-AllDnsRecords: $(Get-Date -format u) - Done, $($ptrRecords.Count) found"
    Write-Verbose -Message "FUNCTION Get-AllDnsRecords: $(Get-Date -format u) - Combining Records..."
    $allRecords = @{}
    Write-Verbose -Message "FUNCTION Get-AllDnsRecords: $(Get-Date -format u) -   Processing 'A' Records"
    foreach ($record in $aRecords) {
        $key = $record.RecordData.IPv4Address.IPAddressToString
        if (!($allRecords.ContainsKey($key))) {
            $allRecords.Add($key,$record)
        } else {
            $allRecords.$key = @($allRecords.$key)+@($record)
        }
    }
    Write-Verbose -Message "FUNCTION Get-AllDnsRecords: $(Get-Date -format u) -   Processing 'PTR' Records"
    foreach ($record in $ptrRecords) {
        if ($record.DistinguishedName -match $ptrIpRegex) {
            $key = $matches.Part1 + '.' + $matches.Part2 + '.' + $matches.Part3 + '.' + $matches.Part4
        } else {
            Write-Error ("PTR Record does not match $ptrIpRegex - $($record.DistinguishedName) - " +
                "$($record.RecordData.PtrDomainName)")
            continue
        }
        if (!($allRecords.ContainsKey($key))) {
            $allRecords.Add($key,$record)
        } else {
            $allRecords.$key = @($allRecords.$key)+@($record)
        }
    }
    Write-Verbose "FUNCTION Get-AllDnsRecords: $(Get-Date -format u) - Done"
    return $allRecords,@($aZones+$ptrZones)
}

function Update-DNSTable {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)][hashtable]$recordChanges,
        [parameter(Mandatory=$true)][hashtable]$allDnsRecords
    )
    #This is intended to update the local table / variable of the DNS records with the changes that are successful
    #We want to minimize re-querying the server for a full update since that can take several minutes.
    $PtrIPFromDNRegex = "^DC=(?<ptrHostname>.*),DC=(?<zoneName>.*)\.in-addr.arpa,.*$"
    Write-Verbose -Message "Beginning update of the DNS table using recent updates"
    $Removals = $recordChanges.GetEnumerator() | Where-Object {
        ($_.Value.OldInputObject.DistinguishedName.Length -gt 0 -and
        $null -eq $_.Value.NewInputObject) }
    $AddsChanges = $recordChanges.GetEnumerator() | Where-Object {
        $null -ne $_.Value.NewInputObject }
    foreach ($Removal in $Removals.Value.OldInputObject) {
        Write-Verbose "  Processing Record"
        $oldDN = $Removal.DistinguishedName
        Write-Verbose -Message "    Old record DN beginning - $($oldDN.substring(0,$oldDN.IndexOf(",cn=")))"
        $oldTTL = $Removal.TimeToLive
        switch ($Removal.RecordType) {
            "A" {
                $key2Remove = $Removal.RecordData.IPv4Address.IPAddressToString
                $oldHostname = $Removal.Hostname
                $type2Remove = "A"
            }
            "PTR" {
                if ($oldDN -match $PtrIPFromDNRegex) {
                    $key2Remove = ("$($Matches.ptrHostname).$($Matches.zoneName)" -split "\.")[-1..-4] -join "."
                    $oldHostname = $Removal.RecordData.PtrDomainName
                    $type2Remove = "PTR"
                } else {
                    $key2Remove = $null
                    Write-Verbose -Message "    OLD PTR record supplied does not match normal DN - $oldDN"
                }
            }
            default { $null }
        }
        if ($null -ne $key2Remove) {
            Write-Verbose -Message "    Old $type2Remove record - hostname $oldHostname : TTL $oldTTL"
            Write-Verbose -Message "    Attempting to remove from $key2Remove in table"
            if ($allDnsRecords.ContainsKey($key2Remove)){
                #Here is similar to the Change/Add only if it's empty after removing, remove the key.
                $numOfRecordsBeforeRemoval = $allDnsRecords.$key2Remove.count
                $allDnsRecords.$key2Remove = @($allDnsRecords.$key2Remove | Where-Object {
                    !($_.DistinguishedName -like "$oldDN*" -and
                    $_.RecordType -eq $type2Remove -and
                    ($null -ne $oldTTL -or $_.TimeToLive -eq $oldTTL) -and
                    ($_.RecordData.PtrDomainName -eq $oldHostname -or $_.Hostname -eq $oldHostname))})
                $numOfRecordsAfterRemoval = $allDnsRecords.$key2Remove.count
                if ($null -eq $allDnsRecords.$key2Remove -or
                $numOfRecordsAfterRemoval -eq 0 -or
                ($numOfRecordsAfterRemoval -eq 1 -and $null -eq $allRecords.$key2Remove.Value)) {
                    $allDnsRecords.Remove($key2Remove)
                    Write-Verbose -Message "    No records remained for $key2Remove, removed from table"
                } else {
                    if ($numOfRecordsBeforeRemoval -ne $numOfRecordsAfterRemoval) {
                        Write-Verbose -Message "    Removed this particular '$type2Remove' record from $key2Remove."
                    } else {
                        Write-Verbose -Message "    Exact match not found in $key2Remove or already removed"
                    }
                }
            } else {
                Write-Verbose -Message "    $key2Remove does not exist or was already deleted in local table"
            }
            $key2Remove = $null
        }
    }
    foreach ($recordChange in $AddsChanges.Value) {
        Write-Verbose -Message "  Processing Record"
        $oldRecord = $recordChange.OldInputObject
        $newRecord = $recordChange.NewInputObject
        $oldDN = $oldRecord.DistinguishedName
        $newDN = $newRecord.DistinguishedName
        if ($oldDN.length -gt 0) {
            Write-Verbose -Message "    Old record DN beginning - $($oldDN.substring(0,$oldDN.IndexOf(",cn=")))"
        }
        if ($newDN.length -gt 0) {
            Write-Verbose -Message "    New record DN beginning - $($newDN.substring(0,$newDN.IndexOf(",cn=")))"
        }
        $oldKey = switch ($oldRecord.RecordType) {
            "A" {
                $oldRecord.RecordData.IPv4Address.IPAddressToString
            }
            "PTR" {
                if ($oldDN -match $PtrIPFromDNRegex) {
                    ("$($Matches.ptrHostname).$($Matches.zoneName)" -split "\.")[-1..-4] -join "."
                } else {
                    $null
                    Write-Verbose -Message "    OLD PTR record supplied does not match normal DN - $newDN"
                }
            }
            default { $null }
        }
        if ($null -ne $oldKey) {
            Write-Verbose -Message "    Old record has IP of $oldKey"
        }
        $newKey = switch ($newRecord.RecordType) {
            "A" {
                $newRecord.RecordData.IPv4Address.IPAddressToString
            }
            "PTR" {
                if ($newDN -match $PtrIPFromDNRegex) {
                    ("$($Matches.ptrHostname).$($Matches.zoneName)" -split "\.")[-1..-4] -join "."
                } else {
                    $null
                    Write-Verbose -Message "    NEW PTR record supplied does not match normal DN - $newDN"
                }
            }
            default { $null }
        }
        if ($null -ne $newKey) {
            Write-Verbose -Message "    New record has IP of $newKey"
        }
        if ($null -eq $oldKey){
            #This means we need to just add new records to our local table but don't know if it exists
            $keys2ChangeOrAdd = $newKey
            Write-Verbose -Message "    Requested to add a record"
        } else {
            #This means records existed that were updated so we need to update them locally
            $keys2ChangeOrAdd = $oldKey
            Write-Verbose -Message "    Requested to update a record"
        }
        foreach ($chgKey in $keys2ChangeOrAdd) {
            if (!($allDnsRecords.ContainsKey($chgKey))){
                #No entry existed, so we're finally adding it to the local table
                Write-Verbose -Message ("    New entry being added for record dn beginning with " +
                    "$($newDN.substring(0,$newDN.IndexOf(",cn=")))")
                $allDNSRecords.add($chgKey,@($newRecord))
            } else {
                if ($null -eq $oldKey) {
                    #So the request didn't include the already-existent value to check, so just exclude any dups
                    $allDnsRecords.$chgKey = @($allDnsRecords.$chgKey |
                        Where-Object { $_ -ne $newRecord}) +
                        @($newRecord)
                    Write-Verbose -Message "    Successfully updated $chgKey and excluded or removed duplicates"
                } elseif ($oldKey -ne $newKey) {
                    #The request included an old record to alter but the IP changed
                    if ($allDnsRecords.ContainsKey($oldKey)) {
                        $allDnsRecords.$oldKey =
                            @($allDnsRecords.$oldKey | Where-Object { $_ -ne $oldRecord })
                        if ($null -eq $allDnsRecords.$oldKey) {
                            $allDnsRecords.Remove($oldKey)
                            Write-Verbose -Message "    No records remained for $oldKey, removed from table"
                        }
                    }
                    $allDnsRecords.$newKey = @($allDnsRecords.$newKey) + @($newRecord)
                } else {
                    #The request included an old record to alter and the IP did not change
                    $allDnsRecords.$chgKey = 
                        @($allDnsRecords.$chgKey | Where-Object {
                            $_ -ne $oldRecord -and
                            $_ -ne $newRecord }) + 
                        @($newRecord)
                    Write-Verbose -Message "    Successfully updated $chgKey and removed the old data"
                }
            }
        }
        $keys2ChangeOrAdd = $null
    }
    Write-Verbose -Message "Completed updating DNS table"
    return $allDnsRecords
}

function Get-PtrZoneAndNameFromIP {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][object]$zones,
        [Parameter(Mandatory=$true)][string]$ip
    )
    #This section uses the IP, splits it into octets, reverses them, then looks for a matching zone
        #starting with the most narrow to the broadest. E.g. if the new IP is 10.1.2.3 then the reverse is
        #3.2.1.10 and it looks for a zone for 2.1.10.in-addr.arpa, then 1.10, and finally 10. If no zone
        #exists, it will return a null value and something can be done from there if so desired.
    $ptrZones = $zones | Where-Object { $_.IsReverseLookupZone -eq $true }
    $ipArray = $ip -split "\."
    $ipIndex = -2
    #Using negative index values, we iterate the IP address backwards to get the PTR zone info
    do {
        $zoneCheck = ($ipArray[$ipIndex..-4] -join ".")
        $zoneGood = $false
        if ($ptrZones.ZoneName -match "$zoneCheck\.in-addr\.arpa") {
            $zoneGood = $true
        } else {
            $ipIndex--
        }
    } until ($zoneGood -eq $true -or $ipIndex -eq -5)
    if ($zoneGood) {
        return @{
            Zone = $($ptrZones | Where-Object { $_.ZoneName -match "$zoneCheck\.in-addr\.arpa" })
            Name = $($ipArray[-1..-4] -join ".").substring(0,$($ipArray[-1..-4] -join ".").IndexOf($zoneCheck)-1)
        }
    } else {
        return $null
    }
}

function Update-DNSRecord {
    [CmdletBinding()]
    param(
        [string]$newIP,
        [string]$newHostname,
        [timespan]$newTTL,
        [Parameter(Mandatory=$true)][object[]]$existingRecords,
        [Parameter(Mandatory=$true)][object]$zones,
        [Parameter(Mandatory=$true)][string]$DC
    )
    #We will tabulate all changes occurring and send them to another function to update the local table
    $changedRecords = @{}
    #The below regex uses named matches so that when a -match is used then the $Matches gives those pieces names
    $zoneRegex = "^DC=.*,DC=(?<zoneName>.*),cn=MicrosoftDNS,DC=(ForestDnsZones|DomainDnsZones).*"
    $recordsUpdated = 0
    #Out of the records submitted, separate processes guide PTR and A record updates
    $existingARecords = @($existingRecords | Where-Object { $_.RecordType -eq "A"})
    $existingPtrRecords = @($existingRecords | Where-Object { $_.RecordType -eq "PTR"})
    if ($existingARecords.count -gt 0) {
        foreach ($existingRecord in $existingARecords) {
            $existingDN = $existingRecord.DistinguishedName
            #We always start with cloning the existing record and its data structure + data
            Write-Verbose -Message ("Processing existing 'A' record - " +
                "$($existingDN.substring(0,$existingDN.IndexOf(",cn=")))")
            $newRecord = $existingRecord.Clone()
            if ($existingDN -match $zoneRegex) {
                $zoneName = $Matches.zoneName
                Write-Verbose -Message "  ZoneName determined to be $zoneName"
            } else {
                Write-Error -Message "Something went wrong analyzing 'A' record for ZoneName for $existingDN"
                continue
            }
            if ($newTTL.Hours -gt 0 -or $newTTL.Minutes -gt 0 -or $newTTL.Seconds -gt 0) {
                $aTTL = $newTTL
                #$newRecord.TimeToLive = $newTTL
            } else {
                $aTTL = $existingRecord.TimeToLive
            }
            if ($newIP.Length -gt 0) {
                $aIP = [System.Net.IPAddress]::parse($newIP)
            } else {
                $aIP = $existingRecord.RecordData.IPv4Address
            }
            if ($newHostname.Length -gt 0){
                #NEED TO create a new record
                #If this is a subdomain "A" record, try to create like that
                if ($existingRecord.Hostname -like "*.*") {
                    $newHostname += $existingRecord.Hostname.substring($existingRecord.Hostname.IndexOf("."))
                }
                #Before we actually try to add the 'A' record, we need to verify that the zone for its PTR exists
                $newAip = $aIP.IPAddressToString
                try {
                    $newAPtrZoneAndName = Get-PtrZoneAndNameFromIP -ip $newAip -zones $zones
                    if ($null -eq $newAPtrZoneAndName) {
                        throw "No Zone Found"
                    }
                } catch {
                    #Need to create a minimal zone on the server to house this record. Minimal = first 2 octets.
                        #In general, we would expect APIPA or private IPs in a business to have a replicated zone -
                        #10.x, 172.x, and 192.x, even though the latter 2 have IPs in that range that are public.
                    $newIPArray = $newAip -split "\."
                    $addPtrZoneArgs = @{
                        NetworkId = "$($newIPArray[0..1] -join ".").0.0/16"
                        ComputerName = $DC
                        ReplicationScope = "Forest"
                        DynamicUpdate = "Secure"
                        PassThru = $true
                        ErrorAction = "Stop"
                    }
                    Write-Verbose -Message "  New PTR zone needed, creating one for $($addPtrZoneArgs.NetworkId)"
                    try {
                        $newZoneName = (Add-DnsServerPrimaryZone @addPtrZoneArgs).ZoneName
                        Write-Verbose -Message "  Done creating zone, moving on"
                    } catch {
                        Write-Error -Message ("Unable to create PTR zone for $($addPtrZoneArgs.NetworkId) on $DC " +
                            "for this new 'A' record. Proceeding but you will need to create the zone and edit "+
                            "the existing 'A' record in this tool to have the PTR generate. - $_")
                    }
                }
                $addAArgs = @{
                    CreatePtr = $true
                    ComputerName = $DC
                    TimeToLive = $aTTL
                    IPv4Address = $aIP
                    ZoneName = $zoneName
                    PassThru = $true
                    Name = $newHostname
                    ErrorAction = "STOP"
                }
                try {
                    $newRecord = Add-DnsServerResourceRecordA @addAArgs
                    $recordsUpdated++
                    $ChangedRecord = @{
                        NewInputObject = $newRecord
                        ZoneName = $zoneName
                    }
                    $changedRecords.add("CreateA$recordsUpdated",$ChangedRecord)
                    Write-Verbose -Message "  new 'A' and PTR record successfully added to $DC."
                    Write-Verbose -Message "  'A' record submitted to be added to local table"
                } catch {
                    if ($_.exception.message -like ("*Failed to create PTR record. Resource record * in zone * " +
                    "on server * is created successfully, but corresponding PTR record could not be created*")) {
                        Write-Verbose -Message "  'A' Record was successfully created but not the PTR."
                        try {
                            $newRecord = Get-DnsServerResourceRecord -RRType "A" -ZoneName $zoneName -Name $newHostname
                            if ($newRecord.Count -gt 1) {
                                $newRecord = $newRecord | Where-Object {
                                    $_.RecordData.IPv4Address -eq $aIP -and $_.TimeToLive -eq $aTTL
                                }
                            }
                            $recordsUpdated++
                            $ChangedRecord = @{
                                NewInputObject = $newRecord
                                ZoneName = $zoneName
                            }
                            $changedRecords.add("CreateA$recordsUpdated",$ChangedRecord)
                        } catch {
                            Write-Error -Message "Unable to add new 'A' record, skipping removal of the old. - $_"
                            continue
                        }
                    } else {
                        Write-Error -Message "Unable to add new 'A' record, skipping removal of the old. - $_"
                        continue
                    }
                }
                #If the addition of the 'a' record worked, then the addition of the new PTR worked, update locally
                try {
                    $newAPtrZoneAndName = Get-PtrZoneAndNameFromIP -ip $newAip -zones $zones
                    $newAPtrArgs = @{
                        Name = $newAPtrZoneAndName.name
                        ZoneName = $newAPtrZoneAndName.Zone.ZoneName
                        RRType = "PTR"
                    }
                    $newAptr = Get-DnsServerResourceRecord @newAPtrArgs -ErrorAction Stop
                    if ($null -ne $newAptr.Count -and $newAptr.Count -gt 1) {
                        $newAptr = $newAptr |
                            Where-Object { $_.RecordData.PtrDomainName -match $newRecord.Hostname }
                    }
                } catch {
                    Write-Error -Message "  Unable to find corresponding PTR"
                }
                #After verifying it existed by getting it from server, check the TTL - it defaults to 1 hour
                if ($null -ne $newAptr -and $newAptr.TimeToLive -ne $aTTL) {
                    $newAptrUpdated = $newAptr.Clone()
                    try {
                        $newAptrUpdated.TimeToLive = $aTTL
                        $newAPtrArgs.Remove("Name")
                        $newAPtrArgs.Remove("RRType")
                        $newAPtrArgs.OldInputObject = $newAptr
                        $newAPtrArgs.NewInputObject = $newAptrUpdated
                        $newAPtrArgs.ErrorAction = "Stop"
                        $newAptrFixed = Set-DnsServerResourceRecord @newAPtrArgs -PassThru
                        Write-Verbose -Message "  Newly-created corresponding PTR had wrong TTL, fixed"
                    } catch {
                        Write-Error -Message "  New PTR generated had wrong TTL but fixing failed - $_"
                    }
                    if ($null -ne $newAptrFixed) { $newAptr = $newAptrFixed }
                }
                Write-Verbose -Message ("  Retrieved newly-created corresponding PTR info:`n           DN " +
                    "starts $($newAptr.DistinguishedName.substring(0,$newAptr.DistinguishedName.IndexOf(",cn=")))" +
                    "`n           Hostname $($newAptr.RecordData.PtrDomainName)" +
                    "`n           TTL $($newAptr.TimeToLive)")
                $ChangedRecord = @{
                    NewInputObject = $newAptr
                    ZoneName = $zoneName
                }
                $changedRecords.add("PtrFromA$recordsUpdated",$ChangedRecord)
                Write-Verbose -Message "  new PTR submitted to be added to local table"
                #Try to remove the old record from the server
                try {
                    $RemRecArgs = @{
                        InputObject = $existingRecord
                        ZoneName = $zoneName
                        Force = $true
                        PassThru = $true
                    }
                    $RemRec = Remove-DnsServerResourceRecord @RemRecArgs
                    Write-Verbose -Message ("  Successfully removed old 'A' record since hostname changes require " +
                        "new records be made")
                    $recordsUpdated++
                    $ChangedRecord = @{
                        OldInputObject = $RemRec
                        ZoneName = $zoneName
                    }
                    $changedRecords.add("RemoveA$recordsUpdated",$ChangedRecord)
                } catch {
                    #If the record wasn't found on the server, then submit to be removed from local table
                    if (!($_.exception.message -like "*Failed to get * record in * zone on * server*")) {
                        Write-Error "Unable to remove old 'A' record for $existingDN - $_"
                    } else {
                        #Either the object doesn't exist or the remove failed. We call that a win.
                        $recordsUpdated++
                        $ChangedRecord = @{
                            OldInputObject = $existingRecord
                            ZoneName = $zoneName
                        }
                        $changedRecords.add("A$recordsUpdated",$ChangedRecord)
                    }
                }
            } else {
                #We are solely updating the "A" record here.
                $newRecord.TimeToLive = $aTTL
                $newRecord.RecordData.IPv4Address = $aIP
                $SetRecordArgs = @{
                    OldInputObject = $existingRecord
                    NewInputObject = $newRecord
                    ZoneName = $zoneName
                    ComputerName = $DC
                    ErrorAction = "STOP"
                }
                try {
                    Set-DNSServerResourceRecord @SetRecordArgs
                    $recordsUpdated++
                    Write-Verbose -Message "  Successfully updated 'A' record"
                    $SetRecordArgs.Remove("ErrorAction")
                    $SetRecordArgs.Remove("ComputerName")
                    #It appears possible to change multiple "A" records to the same IP, so append as needed
                    $changedRecords.add("a$recordsUpdated",$SetRecordArgs)
                } catch {
                    #When you try to add a record, there sometimes exists a duplicate due to the nature of
                        #subdomain records - which gives you 2 "A" records but somehow it's really only one
                        #or else any changes to one are automatically reflected to the other. If that happens
                        #we just need to treat it like it was successful so we can update the local table
                    if (!($_ -like "*Resource record in OldInputObject not found in*")){
                        Write-Error "Unable to update 'A' record for $existingDN) - $_"
                    } else {
                        Write-Verbose -Message ("  Technical error recorded where old A record not found, " +
                            "probably subdomain duplicate which is not really an error. Submitting to update " +
                            "local table as if it was properly updated.")
                        $recordsUpdated++
                        $SetRecordArgs.Remove("ErrorAction")
                        $SetRecordArgs.Remove("ComputerName")
                        $changedRecords.add("a$recordsUpdated",$SetRecordArgs)
                    }
                }
                #Now that we've updated the "A" record, we need to figure out if there's a corresponding PTR record
                #Often, changes to the "A" record will auto delete the corresponding PTR on the server
                Write-Verbose -Message "  Verifying and creating corresponding PTR record for this 'A' record"
                $GetPtrArgs = @{
                    RRType = "PTR"
                    ComputerName = $DC
                    ErrorAction = "STOP"
                }
                #Calculate what the existing record's corresponding PTR should have been using the IP, zone, and
                    #hostname and request that record directly from the server
                $oldAip = $existingRecord.RecordData.IPv4Address.IPAddressToString
                try {
                    $oldAPtrZoneAndName = Get-PtrZoneAndNameFromIP -ip $oldAip -zones $zones
                    if ($null -eq $oldAPtrZoneAndName) {
                        throw "No Zone Found"
                    }
                } catch {
                    Write-Error -Message "No zone was found that could contain PTR records for $oldAip"
                }
                $oldAPtrZoneName = $oldAPtrZoneAndName.Zone.ZoneName
                $oldAPtrName = $oldAPtrZoneAndName.Name
                $oldAPtrHostname = $existingRecord.Hostname + "." + $zoneName + "."
                Write-Verbose -Message ("    Checking for PTR for old IP $oldAip for hostname " +
                    "'$oldAPtrHostname' in zone $oldAptrZoneName")
                try {
                    $oldAptr = @(Get-DnsServerResourceRecord @GetPtrArgs -ZoneName $oldAPtrZoneName -Name $oldAPtrName)
                    if ($null -eq $oldAptr -or
                    ($oldAPtr.Count -gt 1 -and
                    $null -eq ($oldAptr.RecordData.PtrDomainName | Where-Object { $_ -eq $oldAPtrHostname })) -or
                    ($oldAptr.Count -eq 1 -and $oldAptr.RecordData.PtrDomainName -ne $oldAPtrHostname)) {
                        throw
                    }
                    Write-Verbose -Message ("    $($oldAptr.Count) PTR for old IP found:`n" +
                        "             DN = $($oldAptr.DistinguishedName)`n" +
                        "             PtrDomainName = $($oldAptr.RecordData.PtrDomainName)`n" +
                        "             TTL = $($oldAptr.TimeToLive)")
                } catch {
                    Write-Verbose -Message "    No corresponding PTR was found"
                    #We now know that the PTR for this A record doesn't exist. Always try to remove from local tbl
                    #But we only want to attempt this if there's not an existing PTR with the same hostname and DN
                        #being processed in the PTR section once these 'a' records are done
                    $existingPtrDNCheck = "DC=$oldAPtrName,DC=$oldAptrZoneName,cn=MicrosoftDNS*"
                    $existingPtrDupCount = ($existingPtrRecords |
                        Where-Object {
                            $_.DistinguishedName -like $existingPtrDNCheck -and
                            $_.RecordData.PtrDomainName -eq "oldAPtrHostname" }).Count
                    if ($existingPtrDupCount -eq 0) {
                        $record2Remove = @{
                            DistinguishedName = "DC=$oldAPtrName,DC=$oldAptrZoneName,cn=MicrosoftDNS"
                            RecordType = "PTR"
                            RecordData = @{
                                PtrDomainName = $oldAPtrHostname
                            }
                        }
                        $ChangedRecord = @{
                            OldInputObject = $record2Remove
                            ZoneName = $zoneName
                        }
                        $changedRecords.add("ptrRemoval$recordsUpdated",$ChangedRecord)
                        Write-Verbose -Message ("    Submitted request to remove old record (if exists) from " +
                            "local table")
                    } else {
                        Write-Verbose -Message "    It appears the PTR for the old ip will be processed, skipping"
                    }
                }
                #Now we get to creating a corresponding PTR for the new record
                $newAip = $newRecord.RecordData.IPv4Address.IPAddressToString
                try {
                    $newAPtrZoneAndName = Get-PtrZoneAndNameFromIP -ip $newAip -zones $zones
                    if ($null -eq $newAPtrZoneAndName) {
                        throw "No Zone Found"
                    }
                } catch {
                    Write-Error -Message "No zone was found that could contain PTR records for $newAip"
                }
                $newAPtrZoneName = $newAPtrZoneAndName.Zone.ZoneName
                $newAPtrName = $newAPtrZoneAndName.name
                $newAPtrHostname = $newRecord.Hostname + "." + $zoneName + "."
                Write-Verbose -Message ("    Checking for PTR for new IP $newAip for hostname " +
                    "'$newAPtrHostname' in zone $newAptrZoneName")
                try {
                    $newAptr = @(Get-DnsServerResourceRecord @GetPtrArgs -ZoneName $newAPtrZoneName -Name $newAPtrName)
                    if ($null -eq $newAptr -or
                    ($newAptr.Count -gt 1 -and
                    $null -eq ($newAptr.RecordData.PtrDomainName | Where-Object { $_ -eq $newAPtrHostname })) -or
                    ($newAptr.Count -eq 1 -and $newAptr.RecordData.PtrDomainName -ne $newAPtrHostname)) {
                        throw
                    }
                    Write-Verbose -Message ("    $($newAptr.Count) PTR for new IP found:`n" +
                        "             DN = $($newAptr.DistinguishedName)`n" +
                        "             PtrDomainName = $($newAptr.RecordData.PtrDomainName)`n" +
                        "             TTL = $($newAptr.TimeToLive)")
                } catch {
                    Write-Verbose -Message "    No corresponding PTR was found"
                    $newPtrDNCheck = "DC=$newAPtrName,DC=$newAptrZoneName,cn=MicrosoftDNS*"
                    $newPtrDupCount = ($existingPtrRecords |
                        Where-Object {
                            $_.DistinguishedName -like $newPtrDNCheck -and
                            $_.RecordData.PtrDomainName -eq $newAPtrHostname }).Count
                    if ($newPtrDupCount -eq 0) {
                        #SINCE NO PTR WAS FOUND FOR THE NEW INFO, WE NEED TO CREATE ONE
                        $addPTRArgs = @{
                            Name = $newAPtrName
                            ZoneName = $newAPtrZoneName
                            ComputerName = $DC
                            PtrDomainName = $newAPtrHostname
                            TimeToLive = $newRecord.TimeToLive
                            ErrorAction = "STOP"
                        }
                        try {
                            $newRecord = Add-DnsServerResourceRecordPtr @addPTRArgs -PassThru
                            Write-Verbose -Message "    new PTR record successfully added to $DC."
                            $ChangedRecord = @{
                                NewInputObject = $newRecord
                                ZoneName = $zoneName
                            }
                            $changedRecords.add("ptrFromA$recordsUpdated",$ChangedRecord)
                        } catch {
                            Write-Error -Message "Unable to update/replace corresponding PTR record for $newAip - $_"
                        }
                    } else {
                        Write-Verbose -Message "    It appears the PTR for the new ip will be processed, skipping"
                    }
                }
            }
        }
        Write-Information -InformationAction "Continue" -MessageData ("Successfully created/updated/removed " +
            "$recordsUpdated 'A' records")
    }
    if ($existingPtrRecords.count -gt 0) {
        $recordsUpdated = 0
        foreach ($existingRecord in $existingPtrRecords) {
            $existingDN = $existingRecord.DistinguishedName
            Write-Verbose -Message ("Processing existing PTR record - " +
                "$($existingDN.substring(0,$existingDN.IndexOf(",cn=")))")
            if ($existingDN -match $zoneRegex) {
                $zoneName = $matches.zoneName
                Write-Verbose -Message "  ZoneName determined to be $zoneName"
            } else {
                Write-Error -Message "Something went wrong analyzing 'PTR' record for ZoneName for $existingDN"
                continue
            }
            if ($newTTL.Hours -gt 0 -or $newTTL.Minutes -gt 0 -or $newTTL.Seconds -gt 0) {
                $ptrTTL = $newTTL
                Write-Verbose "  new TTL supplied - $($newTTL.Hours):$($newTTL.Minutes):$($newTTL.Seconds)"
            } else {
                #$ptrTTL = New-TimeSpan -Hours 1
                $ptrTTL = $existingRecord.TimeToLive
                Write-Verbose -Message "  no new TTL supplied, using from existing"
            }
            if ($newHostName.Length -gt 0) {
                $domainSuffix = $existingRecord.RecordData.PtrDomainName
                $domainSuffix = ($domainSuffix.substring($domainSuffix.indexof(".")+1)).TrimEnd(".")
                $hostName = "$newHostname.$domainSuffix"
                Write-Verbose "  New Hostname domain suffix calculated to be $domainSuffix"
            } else {
                $hostName = ($existingRecord.RecordData.PtrDomainName).TrimEnd(".")
                Write-Verbose -Message "  old PtrDomainName of $hostName being used"
            }
            if ($newIP.Length -gt 0) {
                Write-Verbose -Message "  NewIP requested, beginning calculations"
                #PTR records are stored based on the split between IP and zone, so a new IP = delete old/make new
                Write-Verbose -Message "  Calculating new PTR hostname and Zone name"
                try {
                    $newPtrZoneAndName = Get-PtrZoneAndNameFromIP -zones $zones -ip $newIP
                    $newZoneName = $newPtrZoneAndName.Zone.ZoneName
                    $ptrName = $newPtrZoneAndName.name
                    if ($null -eq $newPtrZoneAndName) {
                        throw
                    }
                } catch {
                    #Need to create a minimal zone on the server to house this record. Minimal = first 2 octets.
                        #In general, we would expect APIPA or private IPs in a business to have a replicated zone -
                        #10.x, 172.x, and 192.x, even though the latter 2 have IPs in that range that are public.
                    $newIPArray = $newIP -split "\."
                    $addPtrZoneArgs = @{
                        NetworkId = "$($newIPArray[0..1] -join ".").0.0/16"
                        ComputerName = $DC
                        ReplicationScope = "Forest"
                        DynamicUpdate = "Secure"
                        PassThru = $true
                        ErrorAction = "Stop"
                    }
                    Write-Verbose -Message "  New PTR zone needed, creating one for $($addPtrZoneArgs.NetworkId)"
                    try {
                        $newZoneName = (Add-DnsServerPrimaryZone @addPtrZoneArgs).ZoneName
                        $ptrName = $($newIPArray[-1..-3] -join ".")
                        Write-Verbose -Message "  Done creating zone, moving on"
                    } catch {
                        Write-Error -Message ("Unable to create PTR zone for $($addPtrZoneArgs.NetworkId) on $DC " +
                            "for record using the IP of $newIP - $_")
                        continue
                    }
                }
                #At this point, we don't want to create the PTR 2x if it was made already for an "A" change
                $alreadyChanged = @($changedRecords.GetEnumerator() |
                    Where-Object {
                        $_.Value.NewInputObject.DistinguishedName -like "DC=$ptrName,DC=$newZoneName,*" -and
                        $_.Value.NewInputObject.HostName -eq $ptrName -and
                        $_.Value.NewInputObject.TimeToLive -eq $ptrTTL -and
                        $_.Value.NewInputObject.RecordData.PtrDomainName -eq "$hostName." -and
                        $_.Value.NewInputObject.RecordType -eq "PTR"
                    }).count
                if ($alreadyChanged -eq 0) {
                    Write-Verbose -Message ("  new PTR record being created:`n           Hostname = $ptrName`n" +
                    "           ZoneName = $zoneName`n           PtrDomainName = $hostName`n" +
                    "           TTL = $ptrTTL")
                    $addPTRArgs = @{
                        Name = $ptrName
                        ZoneName = $newZoneName
                        ComputerName = $DC
                        PtrDomainName = $hostName
                        TimeToLive = $ptrTTL
                    }
                    try {
                        $newRecord = Add-DnsServerResourceRecordPtr @addPTRArgs -PassThru
                        $recordsUpdated++
                        Write-Verbose -Message "  new PTR record successfully added to $DC."
                        $ChangedRecord = @{
                            NewInputObject = $newRecord
                            ZoneName = $zoneName
                        }
                        $changedRecords.add("ptr$recordsUpdated",$ChangedRecord)
                    } catch {
                        Write-Error -Message "Unable to update/replace PTR record - $_"
                    }
                } else {
                    Write-Verbose -Message "  new PTR record appears to have been made via 'A' record work"
                    $recordsUpdated++
                    $ChangedRecord = @{
                        OldInputObject = $existingRecord
                        ZoneName = $zoneName
                    }
                    $changedRecords.add("ptr$recordsUpdated",$ChangedRecord)
                    Write-Verbose -Message "  submitted request to remove old record (if exists) from local table"
                }
                #Now begins the attempted removal since the addition of the new one was hopefully successful
                try {
                    Remove-DnsServerResourceRecord -InputObject $existingRecord -ZoneName $zoneName -Force
                    Write-Verbose -Message "  Successfully removed old PTR record since IP changes require new " +
                        "records be made"
                    $recordsUpdated++
                    $ChangedRecord = @{
                        OldInputObject = $existingRecord
                        ZoneName = $zoneName
                    }
                    $changedRecords.add("ptr$recordsUpdated",$ChangedRecord)
                } catch {
                    #If the record wasn't found on the server, then submit to be removed from local table
                    if (!($_.exception.message -like "*Failed to get * record in *.in-addr.arpa zone on * server*")) {
                        Write-Error "Unable to remove old PTR record for $existingDN - $_"
                    } else {
                        $recordsUpdated++
                        $ChangedRecord = @{
                            OldInputObject = $existingRecord
                            ZoneName = $zoneName
                        }
                        $changedRecords.add("ptr$recordsUpdated",$ChangedRecord)
                    }
                }
            } else {
                Write-Verbose -Message "  No new IP detected, proceeding with updates"
                $newRecord = $existingRecord.Clone()
                $newRecord.RecordData.PtrDomainName = "$hostName."
                $newRecord.TimeToLive = $ptrTTL
                $SetRecordArgs = @{
                    OldInputObject = $existingRecord
                    NewInputObject = $newRecord
                    ZoneName = $zoneName
                    ComputerName = $DC
                    ErrorAction = "STOP"
                }
                #At this point, we don't want to create the PTR 2x if it was made already for an "A" change
                $alreadyChanged = @($changedRecords.GetEnumerator() |
                    Where-Object {
                        $_.Value.NewInputObject.DistinguishedName -eq $newRecord.DistinguishedName -and
                        $_.Value.NewInputObject.HostName -eq $newRecord.Hostname -and
                        $_.Value.NewInputObject.TimeToLive -eq $newRecord.TimeToLive -and
                        $_.Value.NewInputObject.RecordData.PtrDomainName -eq $newRecord.RecordData.PtrDomainName
                    }).count
                if ($alreadyChanged -eq 0) {
                    try {
                        Set-DNSServerResourceRecord @SetRecordArgs
                        $recordsUpdated++
                        Write-Verbose -Message "  Successfully updated record on server"
                        $SetRecordArgs.Remove("ErrorAction")
                        $SetRecordArgs.Remove("ComputerName")
                        $changedRecords.add("ptr$recordsUpdated",$SetRecordArgs)
                    } catch {
                        #When you attempt a change for a record, it will say the Old was not found if it's been
                            #modified already on the server in any way
                        if (!($_ -like "*Resource record in OldInputObject not found in*")){
                            Write-Error "Unable to update 'PTR' record - $_"
                        } else {
                            #Let's try to add it, just to be safe...if that fails, then it's a duplicate
                            try {
                                $newRecord = Add-DnsServerResourceRecord -InputObject $newRecord -ZoneName $zoneName -PassThru
                                $recordsUpdated++
                                Write-Verbose -Message "  new PTR record successfully added to $DC."
                                $ChangedRecord = @{
                                    OldInputObject = $existingRecord
                                    NewInputObject = $newRecord
                                    ZoneName = $zoneName
                                }
                                $changedRecords.add("ptr$recordsUpdated",$ChangedRecord)
                            } catch {
                                $recordsUpdated++
                                Write-Verbose -Message "  PTR record appears to have been a duplicate or already altered"
                                $SetRecordArgs.Remove("ErrorAction")
                                $SetRecordArgs.Remove("ComputerName")
                                $changedRecords.add("ptr$recordsUpdated",$SetRecordArgs)
                            }
                        }
                    }
                } else {
                    Write-Verbose -Message ("  Skipped updating record because it was submitted with an 'A' " +
                        "record or already done with previous PTR record update.")
                }
            }
        }
    Write-Information -InformationAction "Continue" -MessageData ("Successfully created/updated/altered " +
        "$recordsUpdated 'PTR' records")
    }
    return $changedRecords
}

function Get-NewIP {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][hashtable]$allDnsRecords
    )
    #Can't take credit for this regex, found on the internet, makes sure it's max 255.255.255.255
    $ipRegex = "^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]).){3}" +
        "([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$"
    do {
        $ipGood = $false
        $existingGood = $false
        $newIP = Read-Host "`nWhat IP would you like to assign"
        if ($newIP -match $ipRegex) {
            $confirmIP = Read-Host "Please confirm the new IP"
            if ($confirmIP -eq $NewIP) {
                $ipGood = $true
            } else {
                Write-Information -InformationAction "Continue" -MessageData "IPs do not match, starting over"
                continue
            }
        } else {
            Write-Information -InformationAction "Continue" -MessageData "Not a valid IP, try again"
            continue
        }
        #We do try to prevent errors from submitting a change to an already-existing IP, unless you're sure
        $existingRecords = Get-DNSRecordsFromTable -entry $newIP -allDnsRecords $allDnsRecords
        if ($existingRecords.Count -gt 0) {
            $confirmExisting = Read-Host ("$newIP already exists and has $($existingRecords.Count) DNS " +
                "records...press 'y' to map to this IP anyway")
            if ($confirmExisting -eq "y") {
                $existingGood = $true
            } else {
                Write-Information -InformationAction "Continue" -MessageData "Starting over"
                continue
            }
        } else {
            $existingGood = $true
        }
    } until ($ipGood -and $existingGood)
    return $NewIP
}

function Get-NewTTL {
    [CmdletBinding()]
    param()
    $Prompt = "`nPlease enter a valid timespan in the format HH:MM:SS and press Enter"
    do {
        try {
            $Good = $false
            $newTimeSpan = Read-Host $Prompt
            if (!($newTimeSpan -match "^\d\d:\d\d:\d\d$")) { throw }
            $newTimeSpan = $newTimeSpan -split ":"
            $Good = $true
        } catch {
            Write-Information -InformationAction "Continue" -MessageData ("Invalid timespan, try again " +
                "following the format exactly")
        }
    } until ($Good)
    return (New-TimeSpan -Hours $newTimeSpan[0] -Minutes $newTimeSpan[1] -Seconds $newTimeSpan[2])
}

function Get-NewHostname {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][hashtable]$allDnsRecords
    )
    $Prompt = "`nPlease enter a new Hostname using only letters, numbers, dashes, and underscores"
    do {
        try {
            $Good = $false
            $newHostname = Read-Host $Prompt
            if (!($newHostname -match "^[0-9a-zA-Z\-_]+$")) { throw }
            #Check AD and DNS for any machines that already have that name as that could break things
            try {
                $ADObjects = @(Get-ADComputer -Identity $newHostname)
            } catch {}
            $ExistingRecords = Get-DNSRecordsFromTable -entry $newHostname -allDnsRecords $allDnsRecords
            if ($ADObjects.count -gt 0 -or $ExistingRecords.count -gt 0) {
                $Confirm = Read-Host ("$newHostname already exists and has $($existingRecords.Count) DNS " +
                "records...press 'y' to remap to this hostname anyway")
                if ($Confirm -eq "y") {
                    $Good = $true
                } else {
                    continue
                }
            } else {
                $confirmHostname = Read-Host "Please confirm hostname"
                if ($newHostName -eq $confirmHostname) {
                    $Good = $true
                } else {
                    Write-Information -InformationAction "Continue" -MessageData "Hostnames no matchie, try again"
                }
            }
        } catch {
            Write-Information -InformationAction "Continue" -MessageData "Invalid hostname, try again"
        }
        
    } until ($Good)
    return $newHostname
}

function Initialize-UpdateActionFromPrompt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][hashtable]$allDnsRecords,
        [Parameter(Mandatory=$true)][object[]]$existingRecords,
        [Parameter(Mandatory=$true)][object]$zones,
        [Parameter(Mandatory=$true)][string]$dc
    )
    #This functions as a repeating menu prompt for all update requests since they're the same
    do {
        $Prompt = "`nWhat would you like to update?`n" +
            "(1)New IP`n" +
            "(2)New TTL`n" +
            "(3)New Hostname`n" +
            "(4)New IP and TTL`n" +
            "(5)New TTL and Hostname`n" +
            "(6)New IP and Hostname`n" +
            "(7)New IP, TTL, and Hostname (seriously!?)`n" +
            "(9)Return to results menu`n" +
            "Please choose"
        $DoneUpdatingRecords = $false
        $UpdateArgs = @{
            ExistingRecords = $existingRecords
            DC = $dc
            zones = $zones
            Verbose = $true
        }
        $UpdateTableArgs = @{
            allDnsRecords = $allDnsRecords
        }
        switch ([int](Read-Host $Prompt)) {
            1 {
                $UpdateArgs.newIP = Get-NewIP -allDnsRecords $allDnsRecords
            }
            2 {
                $UpdateArgs.newTTL = Get-NewTTL
            }
            3 {
                $UpdateArgs.newHostName = Get-NewHostname -allDnsRecords $allDnsRecords
            }
            4 {
                $UpdateArgs.newIP = Get-NewIP -allDnsRecords $allDnsRecords
                $UpdateArgs.newTTL = Get-NewTTL
            }
            5 {
                $UpdateArgs.newTTL = Get-NewTTL
                $UpdateArgs.newHostName = Get-NewHostname -allDnsRecords $allDnsRecords
            }
            6 {
                $UpdateArgs.newIP = Get-NewIP -allDnsRecords $allDnsRecords
                $UpdateArgs.newHostName = Get-NewHostname -allDnsRecords $allDnsRecords
            }
            7 {
                $UpdateArgs.newIP = Get-NewIP -allDnsRecords $allDnsRecords
                $UpdateArgs.newTTL = Get-NewTTL
                $UpdateArgs.newHostName = Get-NewHostname -allDnsRecords $allDnsRecords
            }
            9 { $DoneUpdatingRecords = $true }
            Default { "Invalid Entry" }
        }
        if (!($DoneUpdatingRecords)) {
            $UpdateTableArgs.recordChanges = Update-DNSRecord @UpdateArgs
            $allDnsRecords = Update-DNSTable @UpdateTableArgs -Verbose
            $DoneUpdatingRecords = $true
        }
    } until ($DoneUpdatingRecords -eq $true)
    return $allDnsRecords
}

###MAIN###
#REQUIRES #Requires -Module DnsServer,ActiveDirectory
$DC = [string](Get-ADDomainController -Discover).Hostname
Write-Information -InformationAction "Continue" -MessageData ("DNS Record retrieval requires obtaining all " +
    "records from DNS server (will use appx 2Gb of RAM)")
Write-Information -InformationAction "Continue" -MessageData "This will take around 7 minutes."
Read-Host "Press Enter to continue or press CTRL + C to cancel..."
Write-Information -InformationAction "Continue" -MessageData "Beginning Retrieval - $(Get-Date -Format u)"
try {
    $AllRecords,$AllZones = Get-AllDnsRecords -DC $DC -Verbose
} catch {
    Read-Host ("$(Get-Date -Format u) - Unable to retrieve all records successfully. Press any key to quit." +
        "If the error is related to credentials, please run PowerShell as a DNS or Domain Admin - $_")
    exit
}
Write-Information -InformationAction "Continue" -MessageData "Done - $(Get-Date -Format u)"
do {
    $DoneWithScript=$false
    $Prompt = ("`nWhat would you like to do?`n" +
        "(1) Query DNS by Hostname or IP`n" +
        "(2) Export all retrieved DNS records to CSV`n" +
        "(3) Update local table of DNS records by redownloading all records`n" +
        "(9) Exit`n" +
        "Type your choice and press enter")
    switch ([int](Read-Host $Prompt)) {
        1 {
            $DoneWithActions = $false
            $Entry = $null
            do {
                if ($Entry.Length -eq 0) {
                    $Entry = Read-Host "`nWhat IP or Hostname would you like to query (enter x! to return to menu)?"
                }
                if ($Entry -eq "x!") {
                    $DoneWithActions = $true
                    continue
                }
                $EntryRecords = Get-DNSRecordsFromTable -allDnsRecords $AllRecords -entry $Entry
                $ARecords = @($EntryRecords.Value | Where-Object { $_.RecordType -eq "A" })
                $PtrRecords = @($EntryRecords.Value | Where-Object { $_.RecordType -eq "PTR" })
                if ($ARecords.Count -eq 0 -and $PtrRecords.Count -eq 0) {
                    "$Entry not found, try again."
                    $Entry = $null
                    continue
                }
                $Prompt = ("Found $($ARecords.Count + $PtrRecords.Count) total records. " +
                    "What would you like to do?`n" +
                    "(1) Show list of A records`n" +
                    "(2) Show list of PTR records`n" +
                    "(3) Show list of A and PTR records`n" +
                    "(4) Update all PTR Records found`n" +
                    "(5) Update all A and PTR Records found`n" +
                    "(6) Export all A and PTR records found to CSV`n" +
                    "(9) Return to main menu`n" +
                    "Please type your choice and press enter")
                $UpdatePromptArgs = @{
                    allDnsRecords = $AllRecords
                    zones = $AllZones
                    dc = $DC
                    Verbose = $true
                }
                switch ([int](Read-Host $Prompt)) {
                    1 {
                        if ($ARecords.Count -gt 0) {
                            $ARecords | Select-Object -ExcludeProperty Type,RecordType |
                                Sort-Object @{ Expression = { $_.RecordData.IPv4Address.Address } } |
                                Format-Table -AutoSize
                        }
                    }
                    2 {
                        if ($PtrRecords.Count -gt 0) {
                            $PtrRecords | Select-Object -ExcludeProperty Type,RecordType |
                                Sort-Object @{ Expression = { $_.RecordData.PtrDomainName } } |
                                Format-Table -AutoSize
                            Write-Information -InformationAction "Continue" -MessageData ("**Please note that a" +
                                " PTR record hostname is the portion of the IP address not used for the zone " +
                                "name, often the last or last two octets of the IP`n")
                        }
                    }
                    3 {
                        $EntryRecords.Value |
                            Sort-Object -Property RecordType,@{Expression = {
                                if ($_.RecordType -eq "A") {
                                    $_.RecordData.IPv4Address.Address
                                } else {
                                    $_.RecordData.PtrDomainName
                                }
                            }} |
                            Select-Object Hostname,RecordType,Timestamp,TimeToLive,
                            @{Name="RecordData";Expression={
                                if ($_.RecordType -eq "A") {
                                    $_.RecordData.IPv4Address.IPAddressToString
                                } else {
                                    $_.RecordData.PtrDomainName.trimEnd(".")
                                }
                            }} |
                            Format-Table
                        Write-Information -InformationAction "Continue" -MessageData ("**Please note that a" +
                            " PTR record hostname is the portion of the IP address not used for the zone " +
                            "name, often the last or last two octets of the IP`n")
                    }
                    4 {
                        if ($PtrRecords.Count -gt 0) {
                            $UpdatePromptArgs.existingRecords = $PtrRecords
                            $AllRecords = Initialize-UpdateActionFromPrompt @UpdatePromptArgs
                        } else {
                            Write-Information -InformationAction "Continue" -MessageData "No PTR records supplied"
                        }
                    }
                    5 {
                        if ($PtrRecords.Count -gt 0 -or $ARecords.Count -gt 0) {
                            $UpdatePromptArgs.existingRecords = $ARecords + $PtrRecords
                            $AllRecords = Initialize-UpdateActionFromPrompt @UpdatePromptArgs
                        } else {
                            Write-Information -InformationAction "Continue" -MessageData "No records supplied"
                        }
                    }
                    6 { 
                        $CSVPath = "$PSScriptRoot\$entry-DNSRecords-$(Get-Date -format FileDateTime).csv"
                        $EntryRecords.Value | foreach-object {
                            [pscustomobject]@{
                                DistinguishedName = $_.DistinguishedName
                                Hostname = $_.Hostname
                                RecordType = $_.RecordType
                                Timestamp = $_.Timestamp
                                TimeToLive = $_.TimeToLive
                                RecordData = if ($_.RecordType -eq "PTR") {
                                    $_.RecordData.PtrDomainName.TrimEnd(".")
                                } else {
                                    $_.RecordData.IPv4Address.IPAddressToString
                                }
                                Zone = if ($_.DistinguishedName.Length -gt 0) {
                                    $Short = $_.DistinguishedName.substring($_.DistinguishedName.IndexOf(",DC=")+4)
                                    $Short.substring(0,$Short.IndexOf(",cn="))
                                }
                            }
                        } | Export-CSV -Path $CSVPath -NoTypeInformation
                        Write-Information -InformationAction "Continue" -MessageData "Saved to $CSVPath"
                    }
                    9 { $DoneWithActions = $true }
                    Default { Write-Information -InformationAction "Continue" -MessageData "Invalid entry" }
                }
            } until ($DoneWithActions -eq $true)
        }
        2 {
            $CSVPath = "$PSScriptRoot\AllDnsRecords-$(Get-Date -format FileDateTime).csv"
            Write-Information -InformationAction "Continue" -MessageData "Beginning export to $CSVPath..."
            $AllRecords.GetEnumerator() |
                Select-Object -ExpandProperty Value |
                Select-Object DistinguishedName,
                    Hostname,
                    RecordType,
                    Timestamp,
                    TimeToLive,
                    @{ Name = "RecordData"; Expression = {
                        if ($_.RecordType -eq "PTR") {
                            $_.RecordData.PtrDomainName.TrimEnd(".")
                        } else {
                            $_.RecordData.IPv4Address.IPAddressToString
                        }
                    }},
                    @{ Name = "Zone"; Expression = {
                        $Short = $_.DistinguishedName.substring($_.DistinguishedName.IndexOf(",DC=")+4)
                        $Short.substring(0,$Short.IndexOf(",cn="))
                    }} |
                Export-CSV -Path $CSVPath -NoTypeInformation
            Write-Information -InformationAction "Continue" -MessageData "Done."
        }
        3 {
            Write-Information -InformationAction "Continue" -MessageData ("Beginning Retrieval - " +
                "$(Get-Date -Format u)")
            try {
                $AllRecords,$AllZones = Get-AllDnsRecords -DC $DC -Verbose
            } catch {
                Read-Host ("$(Get-Date -Format u) - Unable to retrieve all records. Press any key to quit." +
                    "If the error is related to credentials, please run PowerShell as a DNS or Domain Admin - $_")
                exit
            }
            Write-Information -InformationAction "Continue" -MessageData "Done - $(Get-Date -Format u)"
        }
        9 { $DoneWithScript = $true }
        Default { "Invalid Entry";continue }
    }
    
} until ($DoneWithScript -eq $true)