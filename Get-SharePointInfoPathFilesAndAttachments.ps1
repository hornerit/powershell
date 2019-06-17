<#
.SYNOPSIS
Extracts InfoPath XML files and attachments embedded within those files to local folders from a SharePoint library.

.DESCRIPTION
Attempts to find all XML files inside of a SharePoint library specified, downloads the raw XML files locally unless they already exist locally, and reads through each one for any fields that have another file attachment embedded and will extract all of them into organized folders

.PARAMETER SiteURL
REQUIRED URL to the SharePoint site where the library exists
.PARAMETER LibraryName
REQUIRED Name of the library where the InfoPath forms exist
.PARAMETER FolderStructureNodes
REQUIRED In InfoPath, each 'Data Source' is an XML element with a name. Which 'Data Source' or XML element contains unique information that will be used to create a folder structure for the attachments. E.G. a username, unique ID, LastName + a FirstName (simply supply them as separate fields like "LastName","FirstName" or separate by comma when prompted)
.PARAMETER CutOffDate
OPTIONAL If you desire to only download/process files up to a certain date then use this parameter to specify that date and the script will process through 11:59:59 PM for that date using the Server's date and time
.PARAMETER LocalDownloadPath
REQUIRED Need an existing folder parent path to which the script can download or process but should NOT be the name of the library (the script will create/find a subfolder for the library being processed). If you already have XML files downloaded to process, place them in a subfolder titled '_DownloadedRawFiles' inside the subfolder for the library name (e.g. C:\temp\LibraryName\_DownloadedRawFiles)
.PARAMETER SkipDownload
OPTIONAL If you have previously run this script for the library or already have the files in a folder

.NOTES
  Created by: Brendan Horner (www.hornerit.com)
  Notes: MUST BE RUN AS SCRIPT FILE, do NOT copy-paste into PS to run
  Credit: The following urls helped me compile enough to create this
  --https://sharepoint.stackexchange.com/questions/222076/reading-and-writing-xml-in-form-library-with-powershell
  --http://chrissyblanco.blogspot.ie/2006/07/infopath-2007-file-attachment-control.html
  --https://stackoverflow.com/questions/14905396/using-powershell-to-read-modify-rewrite-sharepoint-xml-document
  Version History:
  --2019-06-17-Initial public version in GitHub, adjusted some settings from previous private work

.EXAMPLE
.\Get-SharePointInfoPathFilesAndAttachments.ps1
#>
param(
	[string]$siteurl = (Read-Host "What is the url to the site in question?"),
	[string]$libraryname = (Read-Host "What is the library name?"),
	[string[]]$folderStructureNodes = ((Read-Host "If InfoPath forms found, what data source (or sources, comma-separated) should be used to create folders? If the data source is actually an attribute of a parent data source, type the data source then add a period and type the attribute name: e.g. mydatasourcewithattributes.attribute5") -split ","),
	[string]$CutOffDate = (Read-Host "Please enter the last date you wish to archive, leave empty for all dates - we will process till Modified date/time is 11:59:59 PM of cutoff day"),
	[string]$LocalDownloadPath = (Read-Host "Please type a path to a folder in which this script will store downloaded files and begin processing them into subfolders"),
	[switch]$SkipDownload
)
$siteurl = $siteurl.trim().trimend("/\")
$libraryname = $libraryname.trim().trimend("/\")
$cred = get-credential
if ($null -eq (Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue)){
    try{
		Add-PSSnapin "Microsoft.SharePoint.PowerShell"
	} catch {
		Write-Host "This script requires the use of the SharePoint PowerShell snap-in"
	}
}

#Create an internet browser object for downloading and set the authentication information for it
$webclient = New-Object System.Net.WebClient
$webclient.Credentials = $cred

#Check if the cutoff date is specified or a folder structure node to create folders for embedded attachments within XML files
if($CutOffDate.Length -gt 0){
	$dtCutOffDate = (get-date -date "$CutOffDate 11:59:59 PM")
} else {
	$dtCutOffDate = ""
}

#Create local paths for storing the downloaded files and their attachments, change filepath1 to your liking but leave the $libraryname as an easy folder name
$filepath1 = "$LocalDownloadPath\$libraryname\".replace("/","")
$filepath2 = $filepath1+"_DownloadedRawFiles\"

#Get SharePoint web aka website, then get the library in question
try{
	$web = Get-SPWeb -Identity $siteurl -ErrorAction SilentlyContinue
	$siteurl2 = $siteurl -replace "%20", ' '
	$web2 = Get-SPWeb -Identity $siteurl2 -ErrorAction SilentlyContinue
	if($null -eq $web -or $web -eq ""){ 
		if($null -eq $web2 -or $web2 -eq ""){
			throw
		} else {
			$web = $web2
		}
	}
} catch {
	Write-Error "Failed to obtain the website, possibly bad url or bad credentials" -ErrorAction Stop
}
try{
	$list = $web.lists[$libraryname]
	$libraryname2 = $libraryname -replace "%20", ' '
	$list2 = $web.lists[$libraryname2]
	if($null -eq $list){ 
		if($null -eq $list2){
			throw
		} else {
			$list = $list2
		}
	}
} catch {
	Write-Error "Failed to obtain list, possibly bad Name or URL or list is not actually in that site" -ErrorAction Stop
}

#Assuming everything went well, try to create a folder locally for the file downloading
if(!(test-path $filepath2 -PathType Container)){
	New-Item -ItemType Directory -Force -Path $filepath2 | out-null
} else {
	if(!($SkipDownload) -and (Read-Host "A folder already exists for these files, type skip to skip downloading and just process these files") -eq "skip"){
		$SkipDownload = $true
	}
}

if(!($SkipDownload)){
	Write-Host "All files will be downloaded to $filepath2 and processed from there"
	#Start a stopwatch to find out just how long the download part takes
	$timer = [System.Diagnostics.Stopwatch]::StartNew()

	#Create the query for only 1000 records from the list with only 3 fields of data to keep the query small
	$query = New-Object Microsoft.SharePoint.SPQuery
	$query.ViewAttributes = "Scope='Recursive'"
	$query.RowLimit = 1000
	$query.ViewFields = "<FieldRef Name='ID'/><FieldRef Name='LinkFilenameNoMenu'/><FieldRef Name='Last_x0020_Modified'/>"
	$query.ViewFieldsOnly = $true
	
	#Looping logic - approximating the size of the library because a giant library would use all of your RAM before you could process it
	$loopCounter = 0
	$loopTotal = $list.itemcount
	$interval = [math]::Round($loopTotal/20)
	if($interval -lt 0){
		$interval = 1
	}
	
	#Execute the query to get the list items, get the position in case there are more than 1000 items, loop through the files, show our progress, download file
	Write-Progress -id 1 -activity "Step 1 of 2: Downloading Files" -status "Working on $loopCounter of appx $loopTotal" -percentComplete ($loopCounter/$loopTotal*100)
	do{
		$myFiles = $list.GetItems($query)
		$query.ListItemCollectionPosition = $myFiles.ListItemCollectionPosition
		foreach($file in $myFiles){
			$loopCounter++
			if(($loopCounter % $interval) -eq 0){
				Write-Progress -id 1 -activity "Step 1 of 2: Downloading Files" -status "Working on $loopCounter of appx $loopTotal" -percentComplete ($loopCounter/$loopTotal*100)
			}
			if($dtCutOffDate -ne "" -and (Get-Date -date $file["Last_x0020_Modified"]) -gt $dtCutOffDate){
				continue
			}
			$webclient.DownloadFile($siteurl + "/" + $file.Url + "?NoRedirect=true",$filepath2+$file.Name)
		}
		Write-Progress -id 1 -activity "Step 1 of 2: Downloading Files" -status "Completed" -Completed
	} while($null -ne $query.ListItemCollectionPosition);
	
	#Clean up the web object to prevent memory leak
	$web.dispose()
	$timer.Stop()
	Write-Output "Part 1 Stats:"
	Write-Output "Total Source files: $loopTotal"
	Write-Output "Total time to download source files: $($timer.Elapsed.TotalSeconds) seconds"
}

#Start a timer to see how long the extraction process takes
$timer = [System.Diagnostics.Stopwatch]::StartNew()
Write-Host "All attachments will be extracted to subfolders in $filepath1"

#Grab all xml (InfoPath) files in the download to process for embedded attachments; if there aren't any, we are done; if there are, find out how many and set loop info
$myFiles = Get-ChildItem -Path "$filepath2\*" -Include "*.xml" -Recurse
if ($myFiles.Count -eq "" -or $null -eq $myFiles){
	return
}
$loopCounter = 0
$errorCounter = 0
$fileErrorTotal = 0
$filesExtracted = 0
$loopTotal = $myFiles.count
$interval = [math]::Round($loopTotal/20)
if($interval -lt 0){
	$interval = 1
}

#Begin processing files
Write-Progress -id 1 -activity "Step 2 of 2: Extracting Attachments" -status "Working on $loopCounter of $loopTotal" -percentComplete 0
foreach($file in $myFiles){
	$fileErrorCounter = 0
	$loopCounter++
	if(($loopCounter % $interval) -eq 0){
		Write-Progress -id 1 -activity "Step 2 of 2: Extracting Attachments" -status "Working on $loopCounter of $loopTotal" -percentComplete ($loopCounter/$loopTotal*100)
	}
	[xml]$xml = Get-Content $file
	$myNodes = $xml.SelectNodes("//*")
	$foldername = ""
	if($folderStructureNodes.count -gt 0){ 
		for($i=0;$i -lt $folderStructureNodes.count;$i++){
			$folderNode = $folderStructureNodes[$i].ToLower()
			if($folderNode.IndexOf(".") -gt 0){
				$tempFolderNodeName = ($FolderNode -split "\.")[0]
				$tempFolderNodeAttribute = ($FolderNode -split "\.")[1]
				$folderName += ($xml.SelectSingleNode("//*[translate(local-name(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='$tempFolderNodeName']").$tempFolderNodeAttribute) -replace '[^\p{L}\p{Nd}/(/_/)/./@/,/-]',''
			} else {
				$folderName += ($xml.SelectSingleNode("//*[translate(local-name(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='$folderNode']").innertext) -replace '[^\p{L}\p{Nd}/(/_/)/./@/,/-]',''
			}
			$folderName += "-"
		}
		$folderName = $folderName.TrimEnd("-")
		$folderName = $folderName -replace '`n',''
	}
	if($folderName -eq "" -or $null -eq $folderName -or $folderStructureNodes.count -eq 0){
		$folderName = $file.BaseName
	}
	$createFolder = 0
	if(!(test-path $filepath1$folderName -PathType Container)){
		$createFolder = 1
	}
	$fileNamePrepend = $file.BaseName
	for($j=0;$j -lt $myNodes.Count;$j++){
		$b64 = $myNodes.Item($j) | select-object -ExpandProperty "#text" -ErrorAction SilentlyContinue
		if($b64.length -gt 2000 -and $b64.indexOf(" ") -eq -1){
			$b64name = $myNodes.Item($j) | select-object -ExpandProperty "name"
			$b64name = $b64name.Substring(3)
			$bytes = [Convert]::FromBase64String($b64)
			if($bytes.length -gt 0){
				$arrFileNameBytes = @()
				#When the attachment is broken into byte strings, the 20th byte tells you how many bytes are used for the filename. Multiply by 2 for ASCII encoding
				$fileNameByteLen = $bytes[20]*2
				#Recreate the filename using every other char from filename bytes
				for($i=0;$i -lt $fileNameByteLen-2;$i+=2){
					$arrFileNameBytes+=$bytes[24+$i]
				}
				$arrFileContentBytes = @()
				#Determine content length by Total - Header - Filename
				$fileContentByteLen = $bytes.length-(24+$fileNameByteLen)
				#Create new array by cloning the content bytes into new array
				$arrFileContentBytes = $bytes[(24+$fileNameByteLen)..($fileContentByteLen+24+$fileNameByteLen)]
				$fileName = [System.Text.Encoding]::ASCII.GetString($arrFileNameBytes)
				#Clean up filename to get rid of spaces and illegal characters and files with too short a name
				$fileName = $fileName.trim()
				$fileName = $fileName -replace '[^\p{L}\p{Nd}/(/_/)/./-]',''
				if($fileName.length -lt 6){
					$fileName = "---"+$fileName
				}
				if(($fileName.indexOf(".",$fileName.length - 5)) -eq -1 -or (($fileName.indexOf(".") -eq -1) -and $fileName.length -lt 5)){
					$fileName = "$fileName.pdf"
				}
				$fileName = $fileNamePrepend+$b64name+"-"+$fileName
				if($createFolder -eq 1) {
					New-Item -ItemType Directory -Force -Path $filepath1$folderName | out-null 
					$createFolder = 0
				}
				$folderName += "\"
				#If the filename already exists, don't overwrite - just add a number to the end
				if(test-path $filepath1$folderName$fileName){
					$myLoop = 1
					$fileNamePre = $fileName.substring(0,$fileName.length-5+($fileName.substring($fileName.length-5).indexOf(".")))
					$fileNamePost = $fileName.trimStart($fileNamePre)
					while(test-path $filepath1$folderName$fileName){
						$fileName = $fileNamePre+"("+$myLoop+")"+$fileNamePost
						$myLoop++
					}
				}
				#Final step - save the document to the local computer
				try{
					[IO.File]::WriteAllBytes($filepath1+$folderName+$fileName,$arrFileContentBytes)
					$filesExtracted++
				} catch {
					Write-Host "Error saving file. Attempted data: Foldername = $foldername. Filename = $filename. Source File = $fileNamePrepend"
					$fileErrorCounter++
				}
			}
		}
	}
	if($fileErrorCounter -gt 0){ $errorCounter++; $fileErrorTotal++ }
}
Write-Progress -id 1 -activity "Step 2 of 2: Extracting Attachments" -status "Completed" -Completed
Write-Output "Error stats: $errorCounter attachments failed to be extracted from $fileErrorTotal files"
Write-Output "Part 2 Stats:"
Write-Output "Total attachments extracted: $filesExtracted (from appx $loopTotal InfoPath source files)"
Write-Output "Total time to extract attachments: $($timer.Elapsed.TotalSeconds) seconds"
Read-Host "Please press enter to close"