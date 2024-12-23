<#
.SYNOPSIS
From a SharePoint 2010 server, extracts InfoPath XML files and attachments embedded within those files to local
	folders from a SharePoint library.

.DESCRIPTION
From a SharePoint server, attempts to find all XML files inside of a SharePoint library specified, downloads the
	raw XML files locally unless they already exist locally, and reads through each one for any fields that have
	another file attachment embedded and will extract all of them into organized folders

.PARAMETER SiteURL
REQUIRED URL to the SharePoint site where the library exists
.PARAMETER LibraryName
REQUIRED Name of the library where the InfoPath forms exist
.PARAMETER FolderStructureNodes
REQUIRED In InfoPath, each 'Data Source' is an XML element with a name. Which 'Data Source' or XML element
	contains unique information that will be used to create a folder structure for the attachments. E.G. a
	username, unique ID, LastName + a FirstName (simply supply them as separate fields like "LastName","FirstName"
	or separate by comma when prompted)
.PARAMETER CutOffDate
OPTIONAL If you desire to only download/process files up to a certain date then use this parameter to specify that
	date and the script will process through 11:59:59 PM for that date using the Server's date and time
.PARAMETER LocalDownloadPath
REQUIRED Need an existing folder parent path to which the script can download or process but should NOT be the
	name of the library (the script will create/find a subfolder for the library being processed). If you already
	have XML files downloaded to process, place them in a subfolder titled '_DownloadedRawFiles' inside the
	subfolder for the library name (e.g. C:\temp\LibraryName\_DownloadedRawFiles)
.PARAMETER SkipDownload
OPTIONAL If you have previously run this script for the library or already have the files in a folder
.PARAMETER SkipExtraction
OPTIONAL If you wish to download the XML files and do something separately with them, use this switch
.PARAMETER UseLastModifiedInsteadOfCreatedDate
OPTIONAL By default, the date filters are based on when the infopath file was created. Use this switch to use
	LastModified attribute of the file instead of CreatedDate. This could be useful if you are trying to include
	forms that normally would have been created before your start date but were modified during the range of time
	that you wish to preserve.
.PARAMETER Credential
OPTIONAL If downloading, you will need credentials for connecting to SharePoint. You can supply it here or via prompt
.PARAMETER DataToExtract
REQUIRED If you are extracting data from the forms, you can generate a "CSV" of XML data, "ATTACHMENTS", or "BOTH"
.PARAMETER DateFilterIndexed
OPTIONAL If the Created field is an indexed column (and indexing is complete), this will use the indexed column

.NOTES
  Created by: Brendan Horner (www.hornerit.com)
  Notes: MUST BE RUN AS SCRIPT FILE, do NOT copy-paste into PS to run. MUST BE RUN ON SHAREPOINT SERVER WITH
  	SHAREPOINT SNAP-IN
  Credit: The following urls helped me compile enough to create this
  --https://sharepoint.stackexchange.com/questions/222076/reading-and-writing-xml-in-form-library-with-powershell
  --http://chrissyblanco.blogspot.ie/2006/07/infopath-2007-file-attachment-control.html
  --https://stackoverflow.com/questions/14905396/using-powershell-to-read-modify-rewrite-sharepoint-xml-document
  --https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/
  		encode-and-decode-attachment-using-visual-c-in-infopath-2010
  Version History:
  --2024-11-18-Added logic for some infopath forms that use FormFields instead of myFields for its nodes
  --2023-12-19-Added error handling for XML files that fail to be converted to XML in the engine with err output
  --2022-06-01-Bugfix for CSVs with forward slashes in the library name
  --2022-05-12-Adjusted how the final CSV output is generated so that it properly merges all fields found in all
		xml files retrieved in a library before exporting instead of using the first file as the master template
		for all fields being exported as some infopath form templates changed over time without relinking.
  --2022-04-19-Adjusted filename calculation evaluate the first 4 bytes of the base64 string. According to MS,
  		the 'signature' of an InfoPath form attachment is the first 4 bytes equaling 199,73,70,65. If that exists,
		then the file is a proper attachment with a header and filename. If not, the raw string is likely an image
		that was pasted directly or built into the infopath form template. Additionally, there is a double null
		at the end as a terminator between the header + filename and the content of a proper attachment, hence the
		need to trim it.
  --2022-03-10-Added switch to filter query if the columns are indexed (avoids threshold issue), performance boost
  --2022-03-07-Added feature for very basic extraction of data in the files to CSV
  --2022-02-16-Re-arranged SiteUrl and added char replace for xml files with nodes with invalid char
  --2021-10-26-Added CmdletBinding() and Credential parameter
  --2021-01-26-Updated styling and documentation to fit better for more narrow screens.
  --2020-02-04-Fixed bug in file content logic that was breaking attachment extraction
  --2019-07-29-Adjusted filename logic to use appropriate encoding and added some more documentation
  --2019-07-24-Added switch for using modified date instead of created date for differenct scenarios
  --2019-07-23-Updated wording for local download path, added StartDate filter and clarified language
  --2019-06-17-Initial public version in GitHub, adjusted some settings from previous private work

.EXAMPLE
.\Get-SharePointInfoPathFilesAndAttachments.ps1
#>
[CmdletBinding()]
param(
	[string]$SiteUrl,
	[string]$LibraryName = (Read-Host ("What is the LIBRARY name? If skipping download, what is the folder name " +
		"that contains the '_DownloadedRawFiles' folder?")),
	[string[]]$FolderStructureNodes,
	[string]$StartDate = (Read-Host ("Please enter the start date for files you wish to archive. Leave empty " +
		"for the earliest files in the library. It will assume midnight of the start date you supply. E.g. " +
		"1/1/2010")),
	[string]$CutOffDate = (Read-Host ("Please enter the last date you wish to archive, leave empty for all " +
		"dates - we will process till Modified date/time is 11:59:59 PM of cutoff day")),
	[string]$LocalDownloadPath = (Read-Host ("Please type a path to a folder in which this script will work. " +
		"This script will create a subfolder inside this path for the library and subfolders within that.")),
	[switch]$SkipDownload,
	[switch]$SkipExtraction,
	[switch]$UseLastModifiedInsteadOfCreatedDate,
	[System.Management.Automation.PSCredential]$Credential,
	[Parameter(Mandatory=$true)][ValidateSet("CSV","ATTACHMENTS","BOTH")][string]$DataToExtract,
	[switch]$DateFilterIndexed
)
if (!$SkipDownload -and $SiteUrl.Length -eq 0) {
	$SiteUrl = Read-Host "What is the url to the SharePoint SITE in question?"
}
if (!$SkipExtraction -and $FolderStructureNodes.Count -eq 0) {
	$FolderStructureNodes = @((Read-Host ("If InfoPath forms found, what data source (or sources, " +
		"comma-separated) should be used to create folders? If the data source is actually an attribute of a " +
		"parent data source, type the data source then add a period and type the attribute name: e.g. " +
		"mydatasourcewithattributes.attribute5")) -split ",")
}
$LibraryName = $LibraryName.trim().trimend("/\")

#Try to create a folder locally for the file downloading
$FilePath1 = "$LocalDownloadPath\$LibraryName\".replace("/","")
$FilePath2 = $FilePath1+"_DownloadedRawFiles\"
if (!(test-path $FilePath2 -PathType Container)) {
	New-Item -ItemType Directory -Force -Path $FilePath2 | out-null
} else {
	if (!($SkipDownload) -and (Read-Host ("A folder already exists for these files, type skip to skip " +
	"downloading and just process these files")) -eq "skip") {
		$SkipDownload = $true
	}
}

if (!($SkipDownload)) {
	$SiteUrl = $SiteUrl.trim().trimend("/\")
	$Cred = if ($Credential) { $Credential } else { Get-credential }
	if ($null -eq (Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue)) {
		try {
			Add-PSSnapin "Microsoft.SharePoint.PowerShell"
		} catch {
			Write-Host "This script requires the use of the SharePoint PowerShell snap-in"
			Read-Host "Press enter to exit script..."
			exit
		}
	}

	#Create an internet browser object for downloading and set the authentication information for it
	$WebClient = New-Object System.Net.WebClient
	$WebClient.Credentials = $Cred

	#Check if the cutoff date is specified
	if ($StartDate.Length -gt 0) {
		$DtStartDate = (Get-date -date "$StartDate 12:00:00 AM")
	} else {
		$DtStartDate = ""
	}

	#Check if the cutoff date is specified
	if ($CutOffDate.Length -gt 0) {
		$DtCutOffDate = (Get-date -date "$CutOffDate 11:59:59 PM")
	} else {
		$DtCutOffDate = ""
	}

	#Get SharePoint web aka website, then get the library in question
	try {
		$Web = Get-SPWeb -Identity $SiteUrl -ErrorAction SilentlyContinue
		$SiteUrl2 = $SiteUrl -replace "%20", ' '
		$Web2 = Get-SPWeb -Identity $SiteUrl2 -ErrorAction SilentlyContinue
		if ($null -eq $Web -or $Web -eq "") { 
			if ($null -eq $Web2 -or $Web2 -eq "") {
				throw
			} else {
				$Web = $Web2
			}
		}
	} catch {
		Write-Error "Failed to obtain the website, possibly bad url or bad credentials" -ErrorAction Stop
	}
	try {
		$List = $Web.lists[$LibraryName]
		$LibraryName2 = $LibraryName -replace "%20", ' '
		$List2 = $Web.lists[$LibraryName2]
		if ($null -eq $List) { 
			if ($null -eq $List2) {
				throw
			} else {
				$List = $List2
			}
		}
	} catch {
		Write-Error "Failed to obtain list, possibly bad Name or URL" -ErrorAction Stop
	}

	Write-Host "All files will be downloaded to $FilePath2 and processed from there"
	#Start a stopwatch to find out just how long the download part takes
	$Timer = [System.Diagnostics.Stopwatch]::StartNew()

	#Create the query for only 1000 records from the list with only 3 fields of data to keep the query small
	$Query = New-Object Microsoft.SharePoint.SPQuery
	$Query.ViewAttributes = "Scope='Recursive'"
	$Query.RowLimit = 1000
	$Query.ViewFields = "<FieldRef Name='ID'/><FieldRef Name='LinkFilenameNoMenu'/>" +
		"<FieldRef Name='Last_x0020_Modified'/><FieldRef Name='Created_x0020_Date'/>"
	$Query.ViewFieldsOnly = $true
	#If the Created or Modified date is indexed, this will set a WHERE clause to have SP filter the results before
		#sending them to us, so we don't have to evaluate them - they are already within the date range. We can
		#only do this because indexing allows us to bypass the list view threshold limit for those columns.
	if ($DateFilterIndexed) {
		Write-Host "Attempting to use CAML query to speed up processing"
		if ($UseLastModifiedInsteadOfCreatedDate) { $FilterField = "Modified" } else { $FilterField = "Created" }
		$Query.Query = '<Where>' +
			$(if ($DtStartDate -ne "" -and $DtCutOffDate -ne "") {
				'<And>'
			}) +
			$(if ($DtStartDate -ne "") {
				'<Geq>' +
					'<FieldRef Name="' + $FilterField + '" />' +
					'<Value IncludeTimeValue="TRUE" Type="DateTime">' +
						$(Get-Date $DtStartDate -Format "yyyy-MM-ddTHH:mm:ssZ") +
					'</Value>' +
				'</Geq>'
			}) +
			$(if ($DtCutOffDate -ne "") {
				'<Leq>' +
					'<FieldRef Name="' + $FilterField + '" />' +
					'<Value IncludeTimeValue="TRUE" Type="DateTime">' +
						$(Get-Date $DtCutOffDate -Format "yyyy-MM-ddTHH:mm:ssZ") +
					'</Value>' +
				'</Leq>'
			}) +
			$(if ($DtStartDate -ne "" -and $DtCutOffDate -ne "") {
				'</And>'
			}) +
		'</Where>'
	}
	
	#Looping logic - approximating the size of the library because a giant library would use all of your RAM
		#before you could process it
	$LoopCounter = 0
	$LoopTotal = $List.itemcount
	$Interval = [math]::Round($LoopTotal/20)
	$downloadedCounter = 0
	if ($Interval -lt 1) {
		$Interval = 1
	}
	
	#Execute the query to get the list items, get the position in case there are more than 1000 items, loop
		#through the files, show our progress, download file
	$PrgAct = "Step 1 of 2: Downloading Files"
	if ($StartDate -or $CutOffDate) {
		$PrgAct += " (Ignoring files found outside of $StartDate-$CutOffDate)"
	}
	$PrgStat = "Working on $LoopCounter of up to $LoopTotal (Updates every $Interval files reviewed)"
	$PrgPrcnt = ($LoopCounter/$LoopTotal*100)
	Write-Progress -id 1 -activity $PrgAct -status $PrgStat -percentComplete $PrgPrcnt
	do {
		$myFiles = $List.GetItems($Query)
		$Query.ListItemCollectionPosition = $myFiles.ListItemCollectionPosition
		foreach ($file in $myFiles) {
			$LoopCounter++
			if (($LoopCounter % $Interval) -eq 0) {
				$PrgStat = "Working on $LoopCounter of appx $LoopTotal (Updates every $Interval files reviewed)"
				$PrgPrcnt = ($LoopCounter/$LoopTotal*100)
				Write-Progress -id 1 -activity $PrgAct -status $PrgStat -percentComplete $PrgPrcnt
			}
			#If a startDate or cutoffDate is supplied, we can filter with the data comparison. Since many libraries
				#don't have the Created and/or Modified columns as indexed columns, we cannot expect to be able to
				#query them directly in this manner - we have to query for raw entries of the entire library and
				#then evaluate the data being returned and THEN download the xml file
			if (!$DateFilterIndexed) {
				if (!($UseLastModifiedInsteadOfCreatedDate)) {
					$comparisonDate = (Get-Date -Date ($file["Created_x0020_Date"]))
				} else {
					$comparisonDate = (Get-Date -Date ($file["Last_x0020_Modified"]))
				}
				if ($DtStartDate -ne "" -and $comparisonDate -lt $DtStartDate) {
					continue
				}
				if ($DtCutOffDate -ne "" -and $comparisonDate -gt $DtCutOffDate) {
					continue
				}
			}
			$WebClient.DownloadFile($SiteUrl + "/" + $file.Url + "?NoRedirect=true",$FilePath2+$file.Name)
			$downloadedCounter++
		}
	} while ($null -ne $Query.ListItemCollectionPosition);
	Write-Progress -id 1 -activity $PrgAct -status "Completed" -Completed
	
	#Clean up the web object to prevent memory leak
	$Web.dispose()
	$Timer.Stop()
	Write-Output -InputObject "Download Stats:"
	Write-Output -InputObject "Total Source files found in library: $LoopTotal"
	Write-Output -InputObject "Total files downloaded: $downloadedCounter"
	Write-Output -InputObject "Total time to download files: $($Timer.Elapsed.TotalSeconds) seconds"
}

if (!($SkipExtraction)) {
	if (@("BOTH","CSV") -contains $DataToExtract) {
		#Create a couple of functions and variables to use to take the raw XML data and create a CSV of the values
			#CSV creation is in beta because testing needs to occur for repeating infopath fields and currently
			#has no way of setting data types - everything is output as if it was text.
			$outputCSVPath = "$filePath1\$LibraryName.csv".Replace("/","")
		function Get-ChildNodes {
			[CmdletBinding()]
			param(
				[Parameter(Mandatory=$true)]$Node,
				[Parameter(Mandatory=$true)][object]$BreadCrumbs,
				[Parameter(Mandatory=$true)][hashtable]$FlattenedData
			)
			$BreadCrumbs += @("[$($Node.LocalName)]")
			foreach ($child in $Node.ChildNodes) {
				if ($child.HasChildNodes -and $null -eq $child.'#text') {
					$getArgs = @{
						FlattenedData = $flattenedData
						BreadCrumbs = $breadCrumbs
					} 
					$FlattenedData,$BreadCrumbs = Get-ChildNodes -Node $child @getArgs
				} else {
					$key = "$($BreadCrumbs -join '')$($child.LocalName)"
					$FlattenedData[$key] = (Get-StringOrEmpty -Node $child)
				}
			}
			# remove this node from the bread crumbs
			$BreadCrumbs = $BreadCrumbs[0..$($BreadCrumbs.Length -2)]
			return $FlattenedData,$BreadCrumbs
		}
		function Get-StringOrEmpty {
			[CmdletBinding()]
			param(
				[Parameter(Mandatory=$true)]$NodeText
			)
			if ($null -eq $NodeText.'#text') {
				return ''
			}
			$text = $NodeText.'#text'.Trim()
			if ($text.Length -gt 100 -and
			$text.IndexOf(" ") -eq -1 -and
			($text.length % 4) -eq 0 -and
			$text -notlike "http://*" -and
			$text -notlike "https://*") {
				return "See attachment"
			}
			return $text
		}
		function ConvertTo-Object ($hashTable) {
			return New-Object PSObject -property $hashTable
		}
		function Invoke-NormalizeHashTables {
			[CmdletBinding()]
			param(
				[ref]$hashTables
			)
			$hashSet = New-Object -TypeName 'System.Collections.Generic.HashSet[string]'
			foreach ($ht in $hashTables.Value) {
				foreach ($key in $ht.Keys) {
					if (!($hashSet.Contains($key))) {
						$hashSet.Add($key) | Out-Null
					}
				}
			}
			foreach ($ht in $hashTables.Value) {
				foreach ($key in $hashSet.GetEnumerator()) {
					if (!($ht.ContainsKey($key))) {
						$ht.Add($key, "")
					}
				}
			}
		}
	}
	#Start a timer to see how long the extraction process takes
	$Timer = [System.Diagnostics.Stopwatch]::StartNew()
	Write-Host "All attachments will be extracted to subfolders in $FilePath1"

	#Grab all xml (InfoPath) files in the download to process for embedded attachments;
		#if there aren't any, we are done; if there are, find out how many and set loop info
	$MyFiles = Get-ChildItem -Path "$FilePath2\*" -Include "*.xml" -Recurse
	if ($MyFiles.Count -eq "" -or $null -eq $MyFiles) {
		return
	}
	if (@("BOTH","CSV") -contains $DataToExtract) {
		if ((Test-Path $outputCSVPath)) {
			$OverallCsvData = Import-CSV -Path $outputCSVPath
		} else {
			$OverallCsvData = @()
		}
	}
	$LoopCounter = 0
	$ErrorCounter = 0
	$FileErrorTotal = 0
	$FilesExtracted = 0
	$LoopTotal = $MyFiles.count
	$Interval = [math]::Round($LoopTotal/20)
	if ($Interval -lt 0) {
		$Interval = 1
	}
	$InvalidCharsRegex = '[^\p{L}\p{Nd}/(/_/)/./@/,/-]'
	$XmlPrefix = "//*[translate(local-name(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='"
	$XmlSuffix = "']"

	#Progress variables
	$PrgAct = "Step 2 of 2: Extracting Data and/or Attachments"
	$PrgStat = "Working on $LoopCounter of appx $LoopTotal (Updates every $Interval files processed)"
	$PrgPrcnt = 0

	#Begin processing files
	Write-Progress -id 1 -activity $PrgAct -status $PrgStat -percentComplete $PrgPrcnt
	foreach ($file in $MyFiles) {
		$fileErrorCounter = 0
		$LoopCounter++
		if (($LoopCounter % $Interval) -eq 0) {
			$PrgStat = "Working on $LoopCounter of appx $LoopTotal (Updates every $Interval files processed)"
			$PrgPrcnt = ($LoopCounter/$LoopTotal*100)
			Write-Progress -id 1 -activity $PrgAct -status $PrgStat -percentComplete $PrgPrcnt
		}
		#Thanks to Steve Smith for finding that some XML nodes may contain a §, which is an invalid char.
		#Expanded this so we can add more replacement options in the future
		try {
			[xml]$xml = Get-Content $file
		} catch {
			try {
				[xml]$xml = (Get-Content $file).Replace("§","")
			} catch {
				Write-Error -Message ("$(Get-Date -format u) - Error converting file at path '$file' to XML" +
					" - $($_.exception.message)")
				$ErrorCounter++
				$FileErrorTotal++
				continue
			}
		}
		if (@("BOTH","CSV") -contains $DataToExtract) {
			$flattenedData = @{ SrcFileName = $file.Name }
			$breadCrumbs = @()
			$getArgs = @{
				FlattenedData = $flattenedData
				BreadCrumbs = $breadCrumbs
			}
			#Thanks to Michael Adams for finding out that some InfoPath forms use FormFields not myFields.
				#Might be ones built in 2k3, might be ones deployed as solutions, unsure.
			if ($xml.myFields) { $getArgs.Node = $xml.myFields } else { $getArgs.Node = $xml.FormFields }
			try {
				$flattenedData,$breadCrumbs = Get-ChildNodes @getArgs
			} catch {
				Write-Error -Message ("$(Get-Date -format u) - Error extracting child nodes for " +
					"$($flattenedData.SrcFileName) - $($_.exception.message)")
				$ErrorCounter++
				$FileErrorTotal++
				continue
			}
			try {
				$OverallCsvData += @($flattenedData)
			} catch {
				Write-Error -Message "Unable to merge new CSV data with old in memory - $_"
				$ErrorCounter++
				$FileErrorTotal++
				continue
			}
		}
		if (@("BOTH","ATTACHMENTS") -contains $DataToExtract) {
			$myNodes = $xml.SelectNodes("//*")
			$foldername = ""
			if ($FolderStructureNodes.count -gt 0) { 
				for ($i=0;$i -lt $FolderStructureNodes.count;$i++) {
					$folderNode = $FolderStructureNodes[$i].ToLower()
					if ($folderNode.IndexOf(".") -gt 0) {
						$tmpFolderNodeName = ($FolderNode -split "\.")[0]
						$tmpFolderNodeAttr = ($FolderNode -split "\.")[1]
						$nodeSearch = "$XmlPrefix$tmpFolderNodeName$XmlSuffix"
					} else {
						$tmpFolderNodeAttr = "innertext"
						$nodeSearch = "$XmlPrefix$folderNode$XmlSuffix"
					}
					$folderXml = $xml.SelectSingleNode($nodeSearch).$tmpFolderNodeAttr
					$folderName += ($folderXml -replace $InvalidCharsRegex,'')
					$folderName += "-"
				}
				$folderName = $folderName.TrimEnd("-")
				$folderName = $folderName -replace '`n',''
			}
			if ($folderName -eq "" -or $null -eq $folderName -or $FolderStructureNodes.count -eq 0) {
				$folderName = $file.BaseName
			}
			$createFolder = 0
			if (!(test-path $FilePath1$folderName -PathType Container)) {
				$createFolder = 1
			}
			$fileNamePrepend = $file.BaseName
			for ($j=0;$j -lt $myNodes.Count;$j++) {
				$b64 = $myNodes.Item($j) | select-object -ExpandProperty "#text" -ErrorAction SilentlyContinue
				if ($b64.length -gt 100 -and 
				$b64.indexOf(" ") -eq -1 -and
				($b64.length % 4) -eq 0 -and
				$b64 -notlike "http://*" -and
				$b64 -notlike "https://*") {
					$b64name = $myNodes.Item($j) | select-object -ExpandProperty "name"
					$b64name = $b64name.Substring(3)
					$bytes = [Convert]::FromBase64String($b64)
					if ($bytes.length -gt 0) {
						#BYTE WORK
						#When the attachment is broken into byte strings, the 20th byte tells you how many bytes are
							#used for the filename. Multiply by 2 for Unicode encoding
						$fileNameByteLen = $bytes[20]*2
						#Test whether the file has an infopath file attachment header. If not, all base64 = content
						if ($bytes[0] -ne 199 -or $bytes[1] -ne 73 -or $bytes[2] -ne 70 -or $bytes[3] -ne 65) {
							$fileByteHeader = 0
							$arrFileNameBytes = $null
							$fileName = "uploadedImage.jpg"
							$arrFileContentBytes = $bytes
						} else {
							#The Header is 24 bytes long for InfoPath attachments
							$fileByteHeader=24
							#Extract the bytes for the filename
							$arrFileNameBytes = for ($i=0;$i -lt $fileNameByteLen;$i++) {
								$bytes[$fileByteHeader+$i]
							}
							try {
								$fileName = [System.Text.Encoding]::Unicode.GetString($arrFileNameBytes)
							} catch {
								Write-Host ("Error extracting attachment. Attempted data: FieldName = $b64name. " +
									"Source File = $fileNamePrepend")
								$fileErrorCounter++
								continue
							}
							$fileName = $fileName.substring(0,$fileName.length -1)
							#Determine content length by Total - Header - Filename
							$fileContentByteLen = $bytes.length-$fileByteHeader-$fileNameByteLen
							$fileContentBytesStart = $fileByteHeader+$fileNameByteLen
							$fileContentBytesEnd = $fileContentBytesStart+$fileContentByteLen
							#Create new array by cloning the content bytes into new array
							$arrFileContentBytes = $bytes[($fileContentBytesStart)..($fileContentBytesEnd)]
						}

						#PROCESSING BYTE WORK RESULTS
						#Clean up filename to get rid of spaces and illegal characters and files with too short a name
						$fileName = $fileName.trim()
						$fileName = $fileName -replace $InvalidCharsRegex,''
						if ($fileName.length -lt 6) {
							$fileName = "---"+$fileName
						}
						if (($fileName.indexOf(".",$fileName.length - 5)) -eq -1 -or
						(($fileName.indexOf(".") -eq -1) -and $fileName.length -lt 5)) {
							$fileName = "$fileName.pdf"
						}
						$fileName = $fileNamePrepend+$b64name+"-"+$fileName
						if ($createFolder -eq 1) {
							New-Item -ItemType Directory -Force -Path $FilePath1$folderName | out-null 
							$createFolder = 0
						}
						$folderName += "\"
						#If, for some reason, the file path is longer than 260 (max for older OS's for windows), we need
							#to truncate the filename and re-attach the extension on the end. Adjusted to 255 since
							#browser url lengths can also be affected and their original limit was 255.
						if ("$FilePath1$foldername$fileName".Length -gt 255) {
							$currentPathLength = "$FilePath1$folderName$fileName".Length
							#Get File extension length itself, e.g. xlsx = 4. Then add 1 for the period
							$fileExtension = ($fileName.substring($fileName.length-5).split("."))[1]
							#Remove a few extra characters in case the file already exists and we need to append numbers
							$length2Remove = ($currentPathLength - 255) + $fileExtension.Length + 1 + 4
							$fileName = "$($fileName.substring(0,($fileName.Length - $length2Remove))).$fileExtension"
						}
						#If the filename already exists, don't overwrite - just add a number to the end
						if (test-path $FilePath1$folderName$fileName) {
							$myLoop = 1
							$lenMin5 = $fileName.length-5
							#This is a weird calc where we get close to the end and figure out where the . is
							$fileNamePre = $fileName.substring(0,$lenMin5+($fileName.substring($lenMin5).indexOf(".")))
							$fileNamePost = $fileName.trimStart($fileNamePre)
							while (test-path $FilePath1$folderName$fileName) {
								$fileName = $fileNamePre+"("+$myLoop+")"+$fileNamePost
								$myLoop++
							}
						}
						#Final step - save the document to the local computer
						try {
							[IO.File]::WriteAllBytes($FilePath1+$folderName+$fileName,$arrFileContentBytes)
							$FilesExtracted++
						} catch {
							Write-Host ("Error saving file. Attempted data: Foldername = $foldername. Filename = " +
								"$filename. Source File = $fileNamePrepend")
							$fileErrorCounter++
						}
					}
				}
			}
			if ($fileErrorCounter -gt 0) { $ErrorCounter += $fileErrorCounter; $FileErrorTotal++ }
		}
	}
	if (@("CSV","BOTH") -contains $DataToExtract) {
		try {
			Invoke-NormalizeHashTables -hashTables ([ref]$OverallCsvData)
			$OverallCsvData | foreach-object { ConvertTo-Object -hashTable $_ } |
				Export-CSV $outputCSVPath -Force -NoTypeInformation
		} catch {
			Write-Error -Message "Unable to create or overwrite CSV at $outputCSVPath - $_"
		}
	}
	Write-Progress -id 1 -activity $PrgAct -status "Completed" -Completed
	Write-Output "Error stats: $ErrorCounter attachments failed to be extracted from $FileErrorTotal files"
	Write-Output "Extraction stats:"
	if (@("BOTH","ATTACHMENTS") -contains $DataToExtract) {
		Write-Output "Total attachments extracted: $FilesExtracted (from appx $LoopTotal InfoPath source files)"
	}
	Write-Output "Total time to extract data/attachments: $($Timer.Elapsed.TotalSeconds) seconds"
}
Read-Host "Please press enter to close"