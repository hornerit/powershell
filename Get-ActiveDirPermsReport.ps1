<#
.SYNOPSIS
Get all Active Directory ACLs and optionally filter for a specific user or security group for the domain of the
    computer on which this script is run.

.PARAMETER Credential
OPTIONAL - When provided, uses a PSCredential object to connect to AD. This may be needed to see some objects
.PARAMETER UserOrGroup
OPTIONAL - Use to filter the final output to focus only on certain users(s) and group(s) samaccountnames
.PARAMETER ReportOutputFolderPath
OPTIONAL - Defaults to C:\temp, path to the folder where you want the output CSV to be placed.
.PARAMETER IncludeGroupMemberships
OPTIONAL - Switch to use if you want to show full capabilities of a supplied UserOrGroup by including their group
    memberships in the output report. For instance, if a user is part of GroupA and GroupA is given access to
    something in ActiveDirectory, they would not be included by default but this switch would include them.
.PARAMETER ObjectsToScan
OPTIONAL - Defaults to "OUs", supply which Active Directory objects that should be scanned for an ACL: OUs, Groups,
    Computers, or All. We are unable to scan user objects for ACLs because it takes too long. OUs will scan
    both OrganizationalUnits and Containers.
.PARAMETER OutputFilteredReportOnly
OPTIONAL - Default is to output the entire ACL report for AD in addition to any UserOrGroup filters you supply.
    This will only output the filtered report instead.
.PARAMETER ExistingPermsReportPath
OPTIONAL - If you have already generated a full output, provide the path to the CSV and this will bypass querying
    Active Directory and just leverage the existing CSV to output a filtered report

.NOTES
    Created By: Brendan Horner (MIT)
    Credit: https://devblogs.microsoft.com/powershell-community/understanding-get-acl-and-ad-drive-output/
    Version History
    --2025-06-04-Some code refactor
    --2025-05-19-Added option to use the local ADPerms-ALL output CSV to avoid re-querying all of AD
    --2025-04-16-Attempting to simply log those OUs with backslashes that break things instead of failing
    --2024-12-05-Fixed bug when filtering for user or group that would filter out valid entries
    --2024-04-16-Ready for production use

.EXAMPLE
.\Get-ActiveDirPermsReport.ps1
.\Get-ActiveDirPermsReport.ps1 -ReportOutputFolderPath "C:\MyFolder\MySubFolder" -UserOrGroup "Group1"
.\Get-ActiveDirPermsReport.ps1 -UserOrGroup User1 -IncludeGroupMemberships -ObjectsToScan All
.\Get-ActiveDirPermsReport.ps1 -UserOrGroup User1 -IncludeGroupMemberships -ObjectsToScan All -OutputFilteredReportOnly $true
#>
[CmdletBinding()]
param(
    [pscredential]$Credential,
    [string[]]$UserOrGroup,
    [string]$ReportOutputFolderPath = "C:\temp",
    [bool]$IncludeGroupMemberships,
    [bool]$OutputFilteredReportOnly,
    [ValidateSet("OUs","Groups","Computers","All")][string]$ObjectsToScan = "OUs",
    [string]$ExistingPermsReportPath
)
#REQUIRES -Modules ActiveDirectory
###FUNCTIONS###
function Get-AllGroupMemberships {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$Identity
    )

    #Initialize an empty array to store group names
    $allGroups = New-Object -TypeName System.Collections.ArrayList
    #Get all domains in the current Forest of the current user using current AD PSDrive
    try {
        $domains = Get-ADForest |
            Select-Object -ExpandProperty Domains |
            ForEach-Object {
                Get-ADDomain -Identity $_ | Select-Object NetBIOSName,DNSRoot,DistinguishedName,Name,Forest
            }
    } catch {
        Write-Error -Message "Error retrieving forest or domains from Active Directory - $_"
        throw
    }
    $getArgs = @{
        ErrorAction = "Stop"
    }
    $adoArgs = @{
        Server = "$($domains[0].DNSRoot):3268"
        Properties = @("samaccountname","canonicalname")
        ErrorAction = "Stop"
    }
    #Set the Identity or Filter of the Get-ADObject call based on supplied formatting
    if ($Identity -like "CN=*") {
        $adoArgs.Identity = $Identity
    } elseif ($Identity -like "*@*") {
        $adoArgs.Filter = "userPrincipalName -eq '$Identity'"
    } else {
        $adoArgs.Filter = "samaccountname -eq '$Identity'"
    }
    #Get the AD Object for the identity supplied
    try {
        $ado = Get-ADObject @adoArgs
        if ($null -eq $ado) { throw "Nothing found in AD using current args"}
        if ($ado.count -gt 1) { throw "More than one entity with the identity '$Identity'" }
    } catch {
        Write-Error -Message "Error retrieving identity '$Identity' from AD: $_"
        throw
    }
    #Need the initial domain where the AD Object is actually found so we can use that for future Server parameters
    $adoDomain = $domains | Where-object { $ado.canonicalname -match "^$($_.DNSRoot)/.+`$"}
    if ($adoDomain.Count -gt 1) {
        Write-Error -Message "more than 1 comain matches the CN of the identity '$Identity', which should be impossible"
        throw
    }
    if ($null -eq $adoDomain) {
        Write-Error -Message ("No domain found that matches identity '$Identity's canonical name of '" +
            "$($ado.canonicalname)', which should be impossible")
        throw
    }
    $adoDomDNS = $adoDomain.DNSRoot

    # Retrieve the initial group memberships for the specified identity
    try {
        $initArgs = @{
            Identity = $ado.DistinguishedName
            Server = $adoDomDNS
        }
        #Get the group membership for the domain where the identity was found THEN all other domains
        #Once we have the groups, we calculate the domain of the groups for future ACL output comparison
        $initialGroups = (@(Get-ADPrincipalGroupMembership -ResourceContextServer $adoDomDNS @initArgs @getArgs |
            Where-Object { $_.GroupCategory -eq "Security" } |
            Select-Object -Property DistinguishedName,Name,@{N="Domain";E={if ($adoDomain.NetBIOSName) {
                $adoDomain.NetBIOSName
            } else {
                $adoDomDNS
            }}}) +
            @(foreach ($domain in $domains | Where-Object { $_.DNSRoot -ne $adoDomDNS }) {
                $domDNS = $domain.DNSRoot
                $domName = if ($domain.NetBIOSName) { $domain.NetBIOSName } else { $domDNS }
                Get-ADPrincipalGroupMembership -ResourceContextServer $domDNS @initArgs @getArgs |
                    Select-Object -Property DistinguishedName,Name,@{N="Domain";E={$domName}}
            })) |
            Sort-Object -Property Domain,Name
    } catch {
        Write-Error -Message "Error getting initial groups - $_"
        throw
    }
    #Add initial batch of groups to the ArrayList
    $allGroups.AddRange($($initialGroups | ForEach-Object {"$($_.Domain)\$($_.Name)"}))
    # Process each group recursively
    foreach ($group in $initialGroups) {
        #Get the domain that matches the group's calculated domain from above
        $grpDomain = $domains | Where-Object { $_.DNSRoot -eq $group.Domain -or $_.NetBIOSName -eq $group.Domain }
        if ($grpDomain.Count -gt 1) {
            throw "More than 1 domain matched dnsroot or netbiosname, which should be impossible"
        }
        if ($null -eq $grpDomain) {
            throw "No domain matched the initial group '$($group.DistinguishedName)'"
        }
        $grpDNS = $grpDomain.DNSRoot

        #Recursively get group memberships
        try {
            $nestArgs = @{
                Identity = $group.DistinguishedName
                Server = $grpDNS
            }
            #Repeats the same logic from above, gathering groups from source domain then others, calcs Domain, sorts
            $nestedGroups = (@(Get-ADPrincipalGroupMembership -ResourceContextServer $grpDNS @nestArgs @getArgs |
                Where-Object { $_.GroupCategory -eq "Security" } |
                Select-Object -Property DistinguishedName,Name,@{N="Domain";E={if ($grpDomain.NetBIOSName) {
                    $grpDomain.NetBIOSName
                } else {
                    $grpDNS
                }}} |
                Where-Object { $allGroups -notcontains "$($_.Domain)\$($_.Name)" }) +
                @(foreach ($domain in $domains | Where-Object { $_.DNSRoot -ne $grpDNS }) {
                    $domDNS = $domain.DNSRoot
                    $domName = if ($domain.NetBIOSName) { $domain.NetBIOSName } else { $domDNS }
                    Get-ADPrincipalGroupMembership -ResourceContextServer $domDNS @nestArgs @getArgs |
                        Where-Object { $_.GroupCategory -eq "Security" } |
                        Select-Object -Property DistinguishedName,Name,@{N="Domain";E={$domName}} |
                        Where-Object { $allGroups -notcontains "$($_.Domain)\$($_.Name)" }
                })) |
                Sort-Object -Property Domain,Name
        } catch {
            Write-Error -Message "Error getting first set of nested groups - $_"
            throw
        }
        #If any groups are found that show that the current group is a member of another, keep repeating the steps
        while ($nestedGroups) {
            #Add all found groups to the arraylist
            $allGroups.AddRange(@($($nestedGroups | ForEach-Object { "$($_.Domain)\$($_.Name)" })))
            #Same as above, get from group's source domain, then others, calc Domain, then sort
            try {
                $nestedGroups = @($nestedGroups | ForEach-Object {
                    $nestDom = $nestedGrp = $domName = $nestedArgs = $null
                    $nestedGrp = $_
                    $nestDom = $domains | Where-Object {
                        $_.DNSRoot -eq $nestedGrp.Domain -or $_.NetBIOSName -eq $nestedGrp.Domain
                    }
                    $nestDNS = $nestDom.DNSRoot
                    $nestedArgs = @{
                        Identity = $nestedGrp.DistinguishedName
                        Server = $nestDNS
                    }
                    if ($nestDom.count -gt 1) {
                        throw "More than 1 domain matched dnsroot or netbiosname, which should be impossible"
                    }
                    if ($null -eq $nestDom) {
                        throw ("No domain matched the nested group '$($nestedGrp.DistinguishedName)', which " +
                            "should be impossible")
                    }
                    try {
                        @(Get-ADPrincipalGroupMembership -ResourceContextServer $nestDNS @nestedArgs @getArgs |
                            Where-Object { $_.GroupCategory -eq "Security" } |
                            Select-Object -Property DistinguishedName,Name,@{N="Domain";E={if ($nestDom.NetBIOSName) {
                                $nestDom.NetBIOSName
                            } else {
                                $nestDNS
                            }}} |
                            Where-Object { $allGroups -notcontains "$($_.Domain)\$($_.Name)" }) +
                            @(foreach ($domain in $domains | Where-Object { $_.DNSRoot -ne $nestDNS }) {
                                $domDNS = $domain.DNSRoot
                                $domName = if ($domain.NetBIOSName) { $domain.NetBIOSName } else { $domDNS }
                                @(Get-ADPrincipalGroupMembership -ResourceContextServer $domDNS @nestedArgs @getArgs |
                                    Where-Object { $_.GroupCategory -eq "Security" } | 
                                    Select-Object -Property DistinguishedName,Name,@{N="Domain";E={$domName}}) |
                                    Where-Object { $allGroups -notcontains "$($_.Domain)\$($_.Name)" }
                            })
                    } catch {
                        throw "Error with group '$($nestedGrp.Domain)\$($nestedGrp.Name)' - $($_.Exception.Message)"
                    }
                }) |
                Sort-Object -Property Domain,Name
            } catch {
                Write-Error -Message "Error getting a batch in the while loop nested groups - $_"
                throw
            }
        }
    }

    # Return the unique list of group names in the format DOMAIN\NAME, DOMAIN could be netbiosname or DNSRoot/FQDN
    $allGroups | Select-Object -Unique | Sort-Object
}

function Get-LUACL {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$dn,
        [Parameter(Mandatory=$true)][string]$domName,
        [Parameter(Mandatory=$true)][hashtable]$systemSidsHash,
        [Parameter(Mandatory=$true)][hashtable]$customSidsHash,
        [Parameter(Mandatory=$true)][object]$driveArgs,
        [Parameter(Mandatory=$true)][hashtable]$ObjTypeGUIDHash,
        [switch]$DoNotRetry
    )
    if ($dn -like "*\*" -and $dn -notlike "Microsoft.ActiveDirectory.Management.dll*") {
        $dn = "Microsoft.ActiveDirectory.Management.dll\ActiveDirectory:://RootDSE/$dn"
        $dn2 = "Microsoft.ActiveDirectory.Management\ActiveDirectory:://RootDSE/$dn"
    }
    try {
        Get-Acl $dn -ErrorAction Stop |
            Select-Object -ExpandProperty Access |
            Where-Object { $_.IsInherited -eq $false} |
            Select-Object -Property @{N="DistinguishedName";E={$dn}},
                @{N="IdentityUsername";E={
                    if ($systemSidsHash.ContainsKey($_.IdentityReference.Value)) {
                        "$domName\$($systemSidsHash."$($_.IdentityReference.Value)")"
                    } elseif ($customSidsHash.ContainsKey($_.IdentityReference.Value)) {
                        $customSidsHash."$($_.IdentityReference.Value)"
                    } elseif ($_.IdentityReference.Value -notlike "S-1-5-*" -and 
                    $_.IdentityReference -notlike "S-1-5-*") {
                        $_.IdentityReference
                    } else {
                        try {
                            $adoArgs = @{
                                Filter = "objectSid -eq '$($_.IdentityReference)'"
                                Properties = @("samaccountname","name")
                                Server = $driveArgs.Server
                            }
                            $ado = Get-ADObject @adoArgs @ADArgs
                            if ($null -eq $ado) {
                                $ado = @{SID = $_.IdentityReference; samaccountname = "UNKNOWN USER" }
                            }
                        } catch {
                            Write-Error "Error getting ADObject using sid - $_"
                            throw
                        }
                        $ado.samaccountname
                        $customSidsHash.Add($ado.SID.Value,$ado.samaccountname)
                    }
                }},
                AccessControlType,
                ActiveDirectoryRights,
                @{N="RightsToObject";E={$ObjTypeGUIDHash.[GUID]"$($_.ObjectType)"}},
                @{N="PropagatesToChildren";E={
                    if ($_.InheritanceFlags -eq "None") { "FALSE" } else { "TRUE" }
                }},
                @{N="AppliesTo";E={
                    if ($_.InheritanceType -eq "None") {
                        "This Object Only"
                    } elseif ($_.InheritanceType -eq "Children") {
                        "Immediate Children Only and NOT This Object"
                    } elseif ($_.InheritanceType -eq "SelfAndChildren") {
                        "This Object and Immediate Children Only"
                    } else { $_.InheritanceType }
                }},
                PropagationFlags
    } catch {
        if ($_.Exception.Message -like "*Cannot find a provider with the name *" -and $dn -like "*.dll*") {
            $ADWeirdArgs = @{
                dn = $dn2
                domName = $domName
                systemSidsHash = $systemSidsHash
                customSidsHash = $customSidsHash
                driveArgs = $driveArgs
                ObjTypeGUIDHash = $ObjTypeGUIDHash
            }
            Get-LUACL @ADWeirdArgs
        } else {
            Write-Host -ForegroundColor Red -Object "Error retrieving ACL for DN '$dn' - $_"
            if(!$DoNotRetry){
                $arrRetry.Add($dn) | Out-Null
            }
        }
    }
}
###MAIN###
$ADArgs = @{
    ErrorAction = "Stop"
}
$outputFileName = ""
if ($ReportOutputFolderPath -like "*\") {
    $ReportOutputFolderPath = $ReportOutputFolderPath.Substring(0,$ReportOutputFolderPath.Length)
}
#This is a list of commonly-known sids that exist in Active Directory in each Domain for faster lookup.
$systemSidsHash = @{
    'S-1-5-32-544' = 'Administrators'
    'S-1-5-32-545' = 'Users'
    'S-1-5-32-546' = 'Guests'
    'S-1-5-32-547' = 'Power Users'
    'S-1-5-32-548' = 'Account Operators'
    'S-1-5-32-549' = 'Server Operators'
    'S-1-5-32-550' = 'Print Operators'
    'S-1-5-32-551' = 'Backup Operators'
    'S-1-5-32-552' = 'Replicators'
    'S-1-5-32-554' = 'Builtin\Pre-Windows 2000 Compatible Access'
    'S-1-5-32-555' = 'Builtin\Remote Desktop Users'
    'S-1-5-32-556' = 'Builtin\Network Configuration Operators'
    'S-1-5-32-557' = 'Builtin\Incoming Forest Trust Builders'
    'S-1-5-32-558' = 'Builtin\Performance Monitor Users'
    'S-1-5-32-559' = 'Builtin\Performance Log Users'
    'S-1-5-32-560' = 'Builtin\Windows Authorization Access Group'
    'S-1-5-32-561' = 'Builtin\Terminal Server License Servers'
    'S-1-5-32-562' = 'Builtin\Distributed COM Users'
    'S-1-5-32-568' = 'Builtin\IIS_IUSRS'
    'S-1-5-32-569' = 'Builtin\Cryptographic Operators'
    'S-1-5-32-573' = 'Builtin\Event Log Readers'
    'S-1-5-32-574' = 'Builtin\Certificate Service DCOM Access'
    'S-1-5-32-575' = 'Builtin\RDS Remote Access Servers'
    'S-1-5-32-576' = 'Builtin\RDS Endpoint Servers'
    'S-1-5-32-577' = 'Builtin\RDS Management Servers'
    'S-1-5-32-578' = 'Builtin\Hyper-V Administrators'
    'S-1-5-32-579' = 'Builtin\Access Control Assistance Operators'
    'S-1-5-32-580' = 'Builtin\Remote Management Users'
}
#As we resolve sids to a user-friendly dataset, put them here so we can pull from this faster than AD query again
$customSidsHash = @{}
#Variable to hold all the permissions to be output later
if ($ExistingPermsReportPath -and (Test-Path $ExistingPermsReportPath) -and $ExistingPermsReportPath -like "*.csv") {
    Write-Verbose -Message "$(Get-date -format u) -- Importing perms CSV"
    $arrPerms = Import-CSV $ExistingPermsReportPath
    Write-Verbose -Message "$(Get-date -format u) -- Done"
    $OutputFilteredReportOnly = $true
} else {
    $arrPerms = New-Object -TypeName System.Collections.ArrayList
}
$arrPermsFiltered = New-Object -TypeName System.Collections.ArrayList
if ($arrPerms.Count -eq 0) {
    $arrRetry = New-Object -TypeName System.Collections.ArrayList
    $OURegex = "^CN=(?<dnName>.+?),(?<parentOU>(?:OU|CN|DC)=.+$)"
    #Get the current location so we can set the path back to that when done
    $startingPath = (Get-Location).Path
    #Getting list of all properties and objects permissions can be granted to
    $ObjTypeGUIDHash = @{}
    $GetADObjectParameter=@{
        SearchBase=(Get-ADRootDSE @ADArgs).SchemaNamingContext
        LDAPFilter='(SchemaIDGUID=*)'
        Properties=@("Name", "SchemaIDGUID")
    }
    $SchGUID = Get-ADObject @GetADObjectParameter
    foreach ($schemaItem in $SchGUID){
        $ObjTypeGUIDHash.Add([GUID]$schemaItem.SchemaIDGUID,$schemaItem.Name)
    }
    $ADObjExtPar = @{
        SearchBase = "CN=Extended-Rights,$((Get-ADRootDSE @ADArgs).ConfigurationNamingContext)"
        LDAPFilter = '(ObjectClass=ControlAccessRight)'
        Properties = @("Name", "RightsGUID")
    }
    $SchExtGUID = Get-ADObject @ADObjExtPar @ADArgs
    foreach ($schExtItem in $SchExtGUID) {
        $key = [GUID]$schExtItem.RightsGUID
        if (!($ObjTypeGUIDHash.ContainsKey($key))) {
            $ObjTypeGUIDHash.Add($key,$schExtItem.Name)
        }
    }
    #Adding an entry for when an ACL has an inherited object type of all zeroes, which means all properties
    $ObjTypeGUIDHash.Add([GUID]"00000000-0000-0000-0000-000000000000","AllProperties")

    #Using Get-ADRootDSE above should trigger the loading of the AD: drive, now we can set to that for Get-ACL
    $start = Get-date
    $startStr = $start.ToString("yyyy-MM-ddTHH-mm-ss")
    Write-Host "$startStr -- BEGINNING"
    #Get the AD Forest, focus on each domain, build a PSDrive so AD calls work in that domain, use it, get all DNs
        #of objects chosen to scan, and use the Get-ACL command to pull all ACLs on an DN. We filter for those ACLs
        #that are not inherited, and finally format output.
    Get-ADForest | Select-Object -ExpandProperty Domains | ForEach-Object {
        $dom = Get-ADDomain -Identity $_ | Select-Object NetBIOSName,DNSRoot,DistinguishedName,Name
        $domDNS = $dom.DNSRoot
        $driveArgs = @{
            Name = "AD$(if ($dom.NetBIOSName) { $dom.NetBIOSName } else { $dom.Name })"
            Scope = "Global"
            root = "//RootDSE/"
            PSProvider = "ActiveDirectory"
            Server = $domDNS
        }
        if ($Credential) {
            $driveArgs.Credential = $Credential
        }
        try {
            New-PSDrive @driveArgs @ADArgs | Out-Null
        } catch {
            Write-Error "Error creating PSDrive on server '$($driveArgs.Server)' with name '$($driveArgs.Name)' - $_"
            break
        }
        try {
            Set-Location "$($driveArgs.Name):" -ErrorAction Stop
            Write-Host "Processing Domain '$($driveArgs.Name)'"
        } catch {
            Write-Error "Error setting current location to PSDrive - $_"
            break
        }
        $scanArgs = @{
            SearchScope = "Subtree"
            Server = $domDNS
        }
        $scannedObjects = ($(if ($ObjectsToScan -contains "Groups" -or $ObjectsToScan -contains "All") {
                @((Get-ADGroup -Filter "*" @scanArgs @ADArgs).DistinguishedName)
            }) +
            $(if ($ObjectsToScan -contains "OUs" -or $ObjectsToScan -contains "All") {
                $ouFilter = 'ObjectClass -eq "organizationalunit" -or ObjectClass -eq "container"'
                @((Get-ADObject -Filter $ouFilter @scanArgs @ADArgs).DistinguishedName) +
                @($dom.DistinguishedName)
            }) +
            $(if ($ObjectsToScan -contains "Computers" -or $ObjectsToScan -contains "All") {
                @((Get-ADComputer -Filter "*" @scanArgs @ADArgs).DistinguishedName)
            })) |
            Where-Object { $null -ne $_ -and $_.length -gt 1 }
        $expectedMinutes = [Math]::Round($scannedObjects.Count/438)
        Write-Warning -Message "BE AWARE THAT THIS PART WILL LIKELY TAKE $expectedMinutes minutes to complete!!"
        Write-Verbose -Message ("$(Get-date -format yyyy-MM-ddTHH-mm-ss) -- Found $($scannedObjects.Count) objects " +
            "in scan. Proceeding to obtain permissions.")
        for ($i=0;$i -lt $scannedObjects.Count;$i+=1000) {
            $batch = $scannedObjects[($i)..($i+999)]
            Write-Verbose -Message "$(Get-Date -format u) -- Processing batch of $($batch.Count) objects"
            $perms = @($batch | ForEach-Object { 
                $dn = $_
                $domName = if ($dom.NetBIOSName) { $dom.NetBIOSName } else { $domDNS }
                $GetLUACLArgs = @{
                    dn = $dn
                    domName = $domName
                    systemSidsHash = $systemSidsHash
                    customSidsHash = $customSidsHash
                    driveArgs = $driveArgs
                    ObjTypeGUIDHash = $ObjTypeGUIDHash
                }
                try {
                    Get-LUACL @GetLUACLArgs -ErrorAction Stop
                } catch {
                    $arrRetry.Add($dn) | Out-Null
                }
            })
            if ($perms.count -gt 0) {
                $arrPerms.AddRange($perms)
            }
        }
        if ($arrRetry.Count -gt 0) {
            Write-Warning -Message "Processing $($arrRetry.Count) items that failed the first attempt."
            if ($VerbosePreference -eq "Continue") {
                Write-Warning -Message "First $(($arrRetry[0..4]).Count) items:`n$($arrRetry[0..4] -join "`n`t")"
            }
            $retryArgs = @{
                Server = $domDNS
            }
            for ($i=0;$i -lt $arrRetry.Count;$i+=1000) {
                $batch = $arrRetry[($i)..($i+999)]
                Write-Verbose -Message "$(Get-Date -format u) -- Processing another batch of $($batch.Count) objects"
                $perms = @($batch | ForEach-Object {
                    $dnFilter = $dn = $domName = $null
                    $dn = $_
                    $domName = if ($dom.NetBIOSName) { $dom.NetBIOSName } else { $domDNS }
                    #Try to remove the AD DLL reference from PSDrive stuff
                    if ($dn -like "Microsoft.ActiveDirectory.Management.dll*") {
                        $dn = $dn.substring(68)
                    } elseif ($dn -like "Microsoft.ActiveDirectory.Management*") {
                        $dn = $dn.substring(64)
                    }
                    if ($dn -match $OURegex) {
                        try {
                            $dnFilter = "name -eq '$($Matches.dnName)'"
                            $dn = (Get-ADObject -Filter $dnFilter @retryArgs @ADArgs).DistinguishedName
                        } catch {
                            Write-Host -ForegroundColor Red -Object "Still unable to find object dn of '$dn'"
                            continue
                        }
                    } else {
                        Write-Host -ForegroundColor Red -Object "DN doesn't match regex. DN supplied is '$dn'"
                        continue
                    }
                    $GetLUACLArgs = @{
                        dn = $dn
                        domName = $domName
                        systemSidsHash = $systemSidsHash
                        customSidsHash = $customSidsHash
                        driveArgs = $driveArgs
                        ObjTypeGUIDHash = $ObjTypeGUIDHash
                    }
                    Get-LUACL @GetLUACLArgs -DoNotRetry
                })
                if ($Perms.count -gt 0) {
                    $arrPerms.AddRange($Perms)
                }
            }
        }
        Set-Location $startingPath
        #The AD PSDrives we make should be temporary, so remove them now that we are done querying that domain
        Remove-PSDrive $driveArgs.Name
    }
} else {
    $start = Get-Date
    if ($ExistingPermsReportPath -match ".+-(?<startStr>[0-9T-]{19})\.csv$") {
        $startStr = $Matches.startStr
    } else {
        Write-Warning "Invalid existing report filename submitted. Output Filename will be inaccurate."
        $startStr = $start.ToString("yyyy-MM-ddTHH-mm-ss")
    }
    Write-Host "$($start.ToString("yyyy-MM-dd HH-mm-ss")) -- BEGINNING USING EXISTING CSV"
}
#This section will reduce the output if a user or group is supplied where we want to filter
if ($UserOrGroup) {
    #Use an arrayList to build the list of potential entries to use to filter output
    $fullListOfFilters = New-Object -TypeName System.Collections.ArrayList
    $allGCs = (Get-ADForest).GlobalCatalogs
    $domDNS = $allGCs | Get-Random -Count 1
    #Most ACL output will be a username or DOMAIN\username, so get that (samaccountname)
    Write-Verbose -Message "$(Get-Date -format u) -- Beginning to build list of entries to filter"
    foreach ($entry in $UserOrGroup) {
        try {
            $adoArgs = @{
                Filter = "samaccountname -eq '$entry'"
                Server = "$($domDNS):3268"
                Properties = @("samaccountname")
            }
            $ado = Get-ADObject @adoArgs @ADArgs
            if ($null -eq $ado -or $ado.Count -gt 1) {
                if ($null -eq $ado -or $ado.samaccountname.length -eq 0) {
                    throw "Directory object not found or zero-length samaccountname"
                } else {
                    throw "Multiple accounts for same samaccountname '$entry' found. Should be impossible"
                }
            }
        } catch {
            throw "Unable to find '$entry' when retrieving ADObject - $_"
        }
        $fullListOfFilters.Add($ado.samaccountname) | Out-Null
        #If we want to also show what powers an identity has based on their group memberships, this will include them
        if ($IncludeGroupMemberships) {
            Write-Verbose -Message "$(Get-Date -format u) -- Adding groups of $($ado.samaccountname) to filter"
            if ($ado.objectClass -eq "User") {
                $ado = Get-ADUser $ado.DistinguishedName -Server "$($domDNS):3268" @ADArgs
            } elseif ($ado.objectClass -eq "Group") {
                $ado = Get-ADGroup $ado.DistinguishedName -Server "$($domDNS):3268" @ADArgs
            }
            try {
                $allGrps = @(Get-AllGroupMemberships -Identity $ado.DistinguishedName -Verbose)
            } catch {
                Write-Error -Message "Error retrieving all memberships for $entry - $_"
                throw
            }
            if ($allGrps.count -gt 0) {
                Write-Verbose -Message "$(Get-Date -format u) -- Found $($allGrps.Count) groups to add to filter"
                $fullListOfFilters.AddRange($allGrps)
            }
        }
    }
    Write-Verbose -Message "$(Get-Date -format u) -- Found a total of $($fullListOfFilters.Count) entries to filter"
    if ($fullListOfFilters.Count -gt 1) {
        $fullListOfFilters = $fullListOfFilters | Select-Object -Unique
    }
    Write-Verbose -Message "$(Get-Date -format u) -- Found a total of $($fullListOfFilters.Count) entries to filter after removing duplicates"
    #The actual filtering of the permissions output
    Write-Host "$(Get-Date -format u) -- Extracting filtered perms from full set of perms"
    $arrPermsFiltered = @(foreach ($entry in $fullListOfFilters) {
        $arrPerms | Where-Object { $_.IdentityUsername -like "*\$entry" -or $_.IdentityUsername -eq $entry }
    })
    Write-Host -Message "$(Get-Date -format u) -- Completed, beginning output to file"
    #Change the output filename to indicate Filtered, Filtered including groups, otherwise defaults to ALL
    if ($IncludeGroupMemberships) {
        $outputFileName = "ADPerms-Fltrd-$UserOrGroup-$($ObjectsToScan -join "-") Objects-InclMbrshp-$startStr.csv"
    } else {
        $outputFileName = "ADPerms-Fltrd-$UserOrGroup-$($ObjectsToScan -join "-") Objects-$startStr.csv"
    }
}
#Actually output the permissions to a CSV at the supplied path and calculated output filename
if ($arrPerms.count -gt 0 -or $arrPermsFiltered.Count -gt 0) {
    if ($UserOrGroup -and $arrPermsFiltered.Count -gt 0) {
        $arrPermsFiltered | Export-CSV "$ReportOutputFolderPath\$outputFileName" -NoTypeInformation -Append
        Write-Host "$(Get-Date -format u) -- CSV Exported to '$ReportOutputFolderPath\$outputFileName'"
    }
    if (!$OutputFilteredReportOnly) {
        $outputFullFileName = "ADPerms-ALL-$($ObjectsToScan -join "-") Objects-$startStr.csv"
        $arrPerms | Export-CSV "$ReportOutputFolderPath\$outputFullFileName" -NoTypeInformation -Append
        Write-Host "$(Get-Date -format u) -- CSV Exported to '$ReportOutputFolderPath\$outputFullFileName'"
    }
} else {
    Write-Host "NO PERMISSIONS FOUND USING CURRENT PARAMETERS"
}
$end = Get-Date
$totalElapsed = [math]::Round($(($end - $start).TotalMinutes))
Write-Host "$(Get-Date -format u) -- DONE. Took a total of $totalElapsed minutes"