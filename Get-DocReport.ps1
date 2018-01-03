<#

.SYNOPSIS
Retrieves information from one or more remote computers and creates a report in multiple formats.


.DESCRIPTION
The getDocReport function uses WMI to create reports from one or more remote computers. The function gathers basic system inforamtion, fixed local volumes, shared folders,local users, server roles and features, installed software, services and scheduled tasks.

The report can be output to one or more of screen, plain text and Sharepoint Wiki formatted html.

.PARAMETER Computer
A computer name, or comma-separated list of computer names to report on. If not specified, the report will run on the local computer.

.PARAMETER HostList
Path to a file containing a list of computer names. This overrides the -Computer parameter.

.PARAMETER OutputScreen
Output results to the users' screen.

.PARAMETER OutputText
Output results to a text file. The file will be named "<SERVERNAME>_Text_<yyyy>-<MM>-<dd>.txt"

.PARAMETER OutputWiki
Output results to a Sharepoint Wiki HTML file. The file will be named "<SERVERNAME>_Wiki_<yyyy>-<MM>-<dd>.aspx"

.PARAMETER OutputPath
Path to a directory where the output files will be saved. Defaults to current directory.

.PARAMETER WikiHeaderFile
Specifies the location of the file containing the content of the wiki header. The default file name is '_template_wikiHeader'.

.PARAMETER WikiFooterFile
Specifies the location of the file containing the content of the wiki footer. The default file name is '_template_wikiFooter'.

.PARAMETER UploadWiki
Uploads the generated Wiki files to a SharePoint site.

.PARAMETER SharePointFilePrefix
String of text to add before the filename when uploading to SharePoint.

.PARAMETER SharePointFileSuffix
String of text to add before the filename when uploading to SharePoint. The default value is '(Computer)'.

.PARAMETER SharePointSiteUrl
Sets the location of the Sharepoint site. The default site is 'https://southhunsley.sharepoint.com/teams/itsupport/documentation'.

.PARAMETER SharePointLibraryName
Sets the SharePoint library. This is the nake of the Wiki we are uploading to. The default library name is 'IT Support Wiki'.

.PARAMETER SharePointUserName
The username to use to connect to the SharePoint site

.PARAMETER SharePointPassword
The password to use to connect to the SharePoint site. If you do not specify a password, you will be prompted for one if required. It is recommened that you do not use this parameter as it will expose your password.

.PARAMETER SharePointAuthMethod
Which authentication method to use to connect to SharePoint. Accepted values are CurrentUser, Credentials, Office365.
	CurrentUser : Attempt to sign in with the account of the user running the script.
	Credentials : Submit a username and password.
	Office365   : Submit Office 365 login details.

#>
Param (
	[Parameter(Position=1)]
	[string]
	$Computer = "$env:computername",
	
	[string]
	$HostList,
	
	[switch]
	$ExcludeUsers,
	
	[switch]	
	$IncludeFeatures,

	[switch]
	$IncludeNetwork,
	
	[switch]
	$IncludeProcesses,

	[switch]
	$IncludeRoles,

	[switch]
	$IncludeServices,

	[switch]
	$IncludeShares,

	[switch]
	$IncludeSoftware,

	[switch]
	$IncludeStorage,

	[switch]
	$IncludeSystem,

    [switch]
    $IncludeSystemChecks,

	[switch]
	$IncludeTasks,

	[switch]
	$IncludeUsers,
	
	[switch]
	$OutputScreen,
	
	[switch]
	$OutputText,
	
	[switch]
	$OutputWiki,
	
	[switch]
	$OutputWikiIndex,
	
	[string]
	$OutputPath = ".\",
	
	[ValidateSet("CurrentUser","Credentials","Office365")]
	[string]
	$SharePointAuthMethod,
	
	[string]
	$SharePointFilePrefix,
	
	[string]
	$SharePointFileSuffix = "(Computer)",
		
	[string]
	$SharePointLibraryName = "Wiki Library Name",
	
	[string]
	$SharePointPassword,
	
	[string]
	$SharePointSite = "https://<SUBDOMAIN>.sharepoint.com",
	
	[string]
	$SharePointSubSite = "path/to/subsite",
		
	[string]
	$SharePointUsername,
	
	[switch]
	$UploadWiki,
	
	[string]
	$WikiHeaderFile = "_template_wikiHeader",
	
	[string]
	$WikiFooterFile = "_template_wikiFooter",
	
	[string]
	$WikiIndexFile = "Computers"
)

if (!($OutputText) -and !($OutputWiki)) {$OutputScreen = $TRUE}

$Validation = 0

if ($UploadWiki) {
	if (!($OutputWiki) -and !($OutputWikiIndex)) { write-host "To upload the Wiki, you must specify -OutputWiki or -OutputWikiIndex"; $Validation++ 
	} else {	
		if (!($SharePointAuthMethod -eq "CurrentUser")){
			if (!($SharePointUserName)) {
				write-host "When uploading to Wiki you must specify a username if not using -SharePointAuthMethod CurrentUser"
				$Validation++ 
			}
			# if(!($SharePointPassword)){
				# $SharePointPassword = Read-Host 'Please enter your SharePoint password' -AsSecureString
				# $SharePointPasswordIsSecure = $TRUE
			# }
		}
		
		#Load SharePoint CSOM Assemblies
		[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

		#Setup Credentials to connect
		Switch ($SharePointAuthMethod) {
			"CurrentUser"{
				#Current User Credentials
				$Credentials = [System.Net.CredentialCache]::DefaultCredentials
			}
			"Credentials"{
				#connect using user account/password
				If (!($SharePointPasswordIsSecure)) {
					$Credentials = New-Object System.Net.NetworkCredential($UserName, (ConvertTo-SecureString $SharePointPassword -AsPlainText -Force))
				} else {
					$Credentials = New-Object System.Net.NetworkCredential($UserName, $(Read-Host 'Please enter your SharePoint password' -AsSecureString))
				}
			}
			"Office365"{
				#For Office 365, Use:
				if ($SharePointPassword) {
					$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($SharePointUserName, (ConvertTo-SecureString $SharePointPassword -AsPlainText -Force))
				} else {
					$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($SharePointUserName, $(Read-Host 'Please enter your SharePoint password' -AsSecureString))
				}
			}
		}
	}
	$SharePointSiteUrl = "${SharePointSite}/${SharePointSubSite}"
}

If (!($IncludeFeatures) -and !($IncludeNetwork) -and !($IncludeRoles) -and !($IncludeServices) -and !($IncludeShares) `
        -and !($IncludeSoftware) -and !($IncludeStorage) -and !($IncludeSystem) -and !($IncludeTasks) -and !($IncludeUsers) `
        -and !($IncludeSystemChecks) -and !($IncludeProcesses)) {
	$IncludeFeatures = $TRUE; $IncludeNetwork = $TRUE; $IncludeRoles = $TRUE; $IncludeServices = $TRUE; $IncludeShares = $TRUE
	$IncludeSoftware = $TRUE; $IncludeStorage = $TRUE; $IncludeSystem = $TRUE ;$IncludeTasks = $TRUE ;$IncludeUsers = $TRUE
    $IncludeSystemChecks = $TRUE; $IncludeProcesses =  $TRUE
	}
	
if ($ExcludeUsers) { $IncludeUsers = $false }

if ($Validation -gt 0) {
	write-host ""
	exit $Validation
}

if ($hostlist -and (Test-path "${hostlist}")) {
	$Computers = Get-Content "${hostlist}"
} else {
	$Computers = $Computer.split(",")
}

if ($outputpath -match '^\.\\') { $outputpath = $outputpath -replace "^\.", (Get-Item -path ".\").fullname }

Function humanSize ($s) {
	$loops = 0
	$fin = $false
	while ($s -ge 1) {
		$s = $s / 1024
		$loops = $loops + 1
		#write-host Loop $loops. $s
	}
	switch ($loops){
		1{$u="bytes"}
		2{$u="KB"}
		3{$u="MB"}
		4{$u="GB"}
		5{$u="TB"}
	}
	$s = $s * 1024
	$s = "{0:N2}" -f $s
	$output = "$s$u"
	return $output
}

Function computerHeading ($title) {
	$loops = $title.length + 6
	for ($i=0; $i -lt $loops; $i++) {
		$topbottom = "${topbottom}="
	}
	$caption = "   ${title}   "
	
	if ($OutputScreen) {
		write-host ""
		write-host $caption
		write-host $topbottom -foregroundcolor red
	}
	if ($OutputText) {
		$textstream.writeline($caption)
		$textstream.writeline($topbottom)
	}
	if ($OutputWiki) {
		$wikistream.writeline("<H1>${title}</H1>")
	}
}

Function sectionHeading ($title) {
	$loops = $title.length + 4
	for ($i=0; $i -lt $loops; $i++) {
		$topbottom = "${topbottom}#"
	}
	$space = "#"
	$loops = $title.length + 2
	for ($i=0; $i -lt $loops; $i++) {
		$space = "${space} "
	}
	$space = "${space}#"
	$caption = "${title}"
	
	if ($OutputScreen) {
		write-host ""
		write-host $topbottom -foregroundcolor green
		write-host $space -foregroundcolor green
		write-host "# " -nonewline -foregroundcolor green
		write-host $caption -nonewline -foregroundcolor white
		write-host " #" -foregroundcolor green
		write-host $space -foregroundcolor green
		write-host $topbottom -foregroundcolor green
		write-host ""
	}
	if ($OutputText) {
		$textstream.writeline()
		$textstream.writeline($topbottom)
		$textstream.writeline($space)
		$textstream.write("# ")
		$textstream.write($caption)
		$textstream.writeline(" #")
		$textstream.writeline($space)
		$textstream.writeline($topbottom)
		$textstream.writeline()
	}
	if ($OutputWiki) {
		$wikistream.writeline("<H2>${Title}</H2>")
		$wikistream.writeline("<blockquote>")
	}
}

function itemStart () {
	if ($OutputWiki){
		$wikistream.writeline("<table width=""100%"" class=""ms-rteTable-2"" cellspacing=""0""><tbody>")
		$oddrow = $True
	}
}

function itemHeading ($title) {
	$underline = ""
	for ($u=0; $u -lt $title.length; $u++) {
		$underline = "${underline}~"
	}
	
	if ($outputscreen) { 
		write-host $title -foregroundcolor white
		write-host $underline -foregroundcolor cyan
	}
	if ($OutputText) {
		$textstream.writeline($title)
		$textstream.writeline($underline)
	}
	if ($OutputWiki){
		$wikistream.writeline("<tr class=""ms-rteTableHeaderRow-2""><th class=""ms-rteTableHeaderEvenCol-2"" rowspan=""1"" colspan=""2"" style=""width: 25%;"">${title}</th></tr>")
	}
}

function itemEnd () {
	if ($outputscreen) {
		write-host ""
	}
	if ($OutputText) {
		$textstream.writeline()
	}
	if ($OutputWiki){
		$wikistream.writeline("</tbody></table>")
		$wikistream.writeline("<br />")
	}
}

function sectionEnd () {
	if ($OutputWiki) {
		$wikistream.writeline("</blockquote>")
	}
}

function speak ($text, $screenonly = $TRUE) {
	write-host "$text"
	if (!($ScreenOnly)) {
		if ($OutputText) {
			$textstream.writeline("$text")
		}
		if ($OutputWiki) {
			$wikistream.write("$text")
			$wikistream.writeline("<br/>")
		}
	}
}

function dataout ($field, $value, $xlength = 0) {
	$linelength = 16 + $xlength
	$fieldlength = $field.length
	$linediff = $linelength - $fieldlength
	$caption = "$field"
	for ($i=0; $i -lt $linediff; $i++) {
		$caption = "${caption} "
	}
	
	if ($outputscreen) { 
		write-host $($caption + ": " + $value)
	}
	if ($OutputText) {
		$textstream.writeline($($caption + ": " + $value))
	}
	if ($OutputWiki) {
		if ($OddRow) {
			$wikistream.writeline("<tr class=""ms-rteTableOddRow-2"">")
			$OddRow = $False
		} else {
			$wikistream.writeline("<tr class=""ms-rteTableEvenRow-2"">")
			$OddRow = $True
		}
		$wikistream.writeline("<td class=""ms-rteTableEvenCol-2"" style=""width: 25%;"">${field}</td>")
		$wikistream.writeline("<td class=""ms-rteTableOddCol-2"">${value}</td>")
		$wikistream.writeline("</tr>")
	}
}

function newline () {
	if ($outputscreen) { 
		write-host ""
	}
	if ($OutputText) {
		$textstream.writeline()
	}
}

function checkresult ($status, $text) {
	if ($OutputWiki) {
		$wikistream.WriteLine("<table width=""100%"" class=""ms-rteTable-2"" cellspacing=""0""><tbody>")
		If ($OddRow) {
			$wikistream.WriteLine("<tr class=""ms-rteTableOddRow-2"">")
			$OddRow = $False
		} else {
			$wikistream.WriteLine("<tr class=""ms-rteTableEvenRow-2"">")
			$OddRow = $True
		}
	}
	switch ($status) {
		0{
			if ($OutputScreen) { write-host "OK      " -foregroundcolor green -nonewline }
			if ($OutputText) {$textstream.Write("OK      ")}
			if ($OutputWiki) {$wikistream.Write("<td class=""ms-rteTableEvenCol-2"" style=""width: 25%;"">OK</td>")}
		}
		1{
			if ($OutputScreen) { write-host "WARNING " -foregroundcolor yellow -nonewline }
			if ($OutputText){$textstream.Write("WARNING ")}
			if ($OutputWiki) {$wikistream.Write("<td class=""ms-rteTableEvenCol-2"" style=""width: 25%;"">WARNING</td>")}
		}
		2{
			if ($OutputScreen) { write-host "ERROR   " -foregroundcolor red -nonewline }
			if ($OutputText){$textstream.Write("ERROR   ")}
			if ($OutputWiki) {$wikistream.Write("<td class=""ms-rteTableEvenCol-2"" style=""width: 25%;"">ERROR</td>")}
		}
		default{
			$defaulted = $TRUE 
			if ($OutputScreen) { write-host "UNKNOWN " -foregroundcolor yellow -nonewline}
			if ($OutputText){$textstream.Write("UNKNOWN ")}
			if ($OutputWiki) {$wikistream.Write("<td class=""ms-rteTableEvenCol-2"" style=""width: 25%;"">UNKNOWN</td>")}
		}
	}
	If ($defaulted) { write-host "Unknown result. Probably a scripting error. " -foregroundcolor yellow -nonewline} 
	if ($defaulted -or $OutputScreen) { write-host "$text" }
	if ($OutputText) {$textstream.WriteLine("$text")}
	if ($OutputWiki) {
		$wikistream.Write("<td class=""ms-rteTableOddCol-2"">${text}</td></tr>")
		$wikistream.WriteLine("</tbody></table>")
	}
	$defaulted = ""
}

Function uploadWiki ($SourceFile) {
	if ($OutputWikiIndex) {
		speak "Uploading ${wikifile} to ${SharePointSiteUrl}/${SharePointLibraryName}/${WikiIndexFile}.aspx as ${SharePointUserName}."
	} else {
		speak "Uploading ${wikifile} to ${SharePointSiteUrl}/${SharePointLibraryName}/${SharePointFilePrefix}${ComputerName}${SharePointFileSuffix}.aspx as ${SharePointUserName}."
	}

	#Set up the context
	$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SharePointSiteUrl) 
	$Context.Credentials = $credentials
	$web = $Context.Web

	#Get the Library
	$List = $web.Lists.GetByTitle($SharePointLibraryName)
	$Context.Load($List)
	$Context.ExecuteQuery()

	#Get File Name from source file path
	$SourceFileName = Split-path $SourceFile -leaf
	
	#Set destination file name
	if ($OutputWikiIndex) {
		$DestinationFileName = "${WikiIndexFile}.aspx"
	} else {
		$DestinationFileName = "${SharePointFilePrefix}${ComputerName}${SharePointFileSuffix}.aspx"
	}

	#Get Source file contents
	$FileStream = ([System.IO.FileInfo] (Get-Item $SourceFile)).OpenRead()

	#Upload to SharePoint
	$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
	$FileCreationInfo.Overwrite = $true
	$FileCreationInfo.ContentStream = $FileStream
	$FileCreationInfo.URL = $DestinationFileName
	$FileUploaded = $List.RootFolder.Files.Add($FileCreationInfo)
	$Context.Load($FileUploaded)
	$Context.ExecuteQuery()

	$FileStream.Close()
}

If ($OutputWikiIndex) {
	If (!(Test-path "${OutputPath}" -PathType Container)){
		try {
			New-Item -path "${OutputPath}\" -ItemType Directory -ErrorAction Stop | out-null
		} catch {
			write-host "Unable to create directory ${OutputPath}"
			exit 1
		}
	}
	$wikifile = "${OutputPath}\Index_wiki_" + (get-date -f "yyyy-MM-dd") + ".aspx"
	$wikistream = [System.IO.StreamWriter] "$wikifile"
	speak "Saving wiki index to ${wikifile}."
	$wikiHeader = $(Get-Content "${WikiHeaderFile}") -replace "\!\!COMPUTERNAME\!\!", "$OutputWikiIndex"
	foreach ($line in $wikiheader) {
		$wikistream.WriteLine($line)
	}
	$wikistream.writeline("<H1>${WikiIndexFile}</h1>")
	$wikistream.writeline("<ul>")
	#$SharePointLibraryName = $SharePointLibraryName -replace " ", "%20"
	$AId = 0
	foreach ($ComputerName in $Computers) {
		$ComputerName = ($ComputerName.ToLower())
		$ComputerName = ([Regex]::Replace($ComputerName, '\b(\w)', { param($m) $m.Value.ToUpper() }))
		$wikistream.writeline("<li dir=""ltr""><h2>")
		$wikistream.write("<a id=""${Aid}&#58;&#58;${SharePointFilePrefix}${ComputerName}${SharePointFileSuffix}|${ComputerName}"" class=""ms-wikilink"" href=""/${SharePointSubSite}/")
		$wikistream.write($SharePointLibraryName -replace " ", "%20")
		$wikistream.writeline("/${SharePointFilePrefix}${ComputerName}${SharePointFileSuffix}.aspx"">${ComputerName}</a>")
		$wikistream.writeline("</h2></li>")
		$AId++
	}
	$wikifooter = $(Get-Content "${WikiFooterFile}")
	foreach ($line in $wikifooter) {
		$wikistream.writeline($line)
	}
	$wikistream.close()
	if ($UploadWiki) { uploadwiki "${wikifile}" }
} else {
	foreach ($ComputerName in $Computers) {
        #flush out any existing information
        $cpuinfo = ""
		$computerinfo = ""
		$osinfo = ""
        $NetworkAdapters = ""
		$StorageVolumes = ""
		$sharedfolders = ""
		$ComputerADSI = ""
		$userlist = ""
		$Roles = ""
		$Features = ""
		$installedsoftware = ""
		$Services = ""
		$Processes = ""
		$ScheduledTasks = ""

		$ComputerName = ($ComputerName.ToLower())
		$ComputerName = ([Regex]::Replace($ComputerName, '\b(\w)', { param($m) $m.Value.ToUpper() }))
		speak "Running report for ${ComputerName}..."
		speak "...Checking system is online..."
		if (!(Test-Connection -ComputerName $ComputerName -count 1 -quiet )) {
			speak "${ComputerName} is offline."
		} else {
			if ($OutputText -or $OutputWiki) {
				If (!(Test-path "${OutputPath}" -PathType Container)){
					try {
						New-Item -path "${OutputPath}\" -ItemType Directory -ErrorAction Stop | out-null
					} catch {
						write-host "Unable to create directory ${OutputPath}"
						exit 1
					}
				}
			}
			
			# System Information
			if ($IncludeSystem) {
				speak "...Gathering basic system information..." $false
				$computerInfo = get-wmiobject win32_computersystem -computername $ComputerName | select name, domain, manufacturer, model, description, dnshostname, partofdomain, numberofprocessors, workgroup
				$cpuInfo = get-wmiobject win32_processor -computername $ComputerName | select name, description, datawidth
				$osInfo = get-wmiobject win32_operatingsystem -computername $ComputerName | select caption, OSArchitecture
				$osType = $osinfo.OSArchitecture
				$physicalMemory = Get-WmiObject win32_physicalmemory -computername $ComputerName | Measure-Object capacity -sum | select @{N="memory"; E={[math]::round(($_.Sum / 1GB),2)}}
				$physicalMemory = $physicalMemory.memory
			}
			
			# Networking information
			if ($IncludeNetwork) {
				speak "...Gathering networking information..."
				$NetworkAdapters = Get-WmiObject win32_networkadapterconfiguration -ComputerName $ComputerName | 
					where {$_.ipaddress} | 
					select Description, DhcpEnabled, DnsServerSearchOrder, IpAddress, DefaultIpGateway, IpSubnet, MacAddress
			}
			
			# Find all fixed local volumes and optical drives
			If ($IncludeStorage) {
				speak "...Gathering local storage volumes..." $false
				$StorageVolumes = Get-WMIOBJECT win32_volume -computername $ComputerName | where { $_.DriveType -eq 3 -or $_.DriveType -eq 5 -and $_.Label -ne "System Reserved"} | sort-object { $_.Name }
			}
			
			# Shared folders
			if ($IncludeShares) {
			speak "...Gathering shared folders..." $false
			$sharedfolders = get-WmiObject -class Win32_Share -computer $ComputerName |
				where {$_.name -inotmatch '^[A-Z]\$$' `
				-and $_.name -inotmatch '^ADMIN\$$' `
				-and $_.name -inotmatch '^IPC\$$' `
				} | select Name, Path, Description
			}
			
			# Local users
			If ($IncludeUsers) {
				speak "...Gathering local users..." $false
				$ComputerADSI = [ADSI]"WinNT://$ComputerName,computer" 
				$userlist = $ComputerADSI.psbase.Children | Where-Object { $_.psbase.schemaclassname -eq 'user' }
				$ComputerADSI = ""
			}
			
			# Roles and features
			if (get-wmiobject win32_operatingsystem -ComputerName $ComputerName | where {$_.Name -match "Server"}){
				if ($IncludeRoles){
					Import-Module ServerManager
					speak "...Gathering server roles..." $false
					$Roles = get-windowsfeature -computername $ComputerName | where-object {$_.Installed -match "True" -and $_.FeatureType -match "Role" `
						-and $_.Name -notmatch "FileAndStorage-Services" `
						-and $_.Name -notmatch "Storage-Services" `
						-and $_.name -notmatch "File-Services" `
					} | select name, displayname, description, systemservice | sort-object {$_.name}
				}
				if ($IncludeFeatures) {
					speak "...Gathering server features..." $false 
					$Features = get-WindowsFeature -computername $ComputerName |
						where-object {$_.Installed -match "True" `
						-and $_.FeatureType -match "Feature" `
						-and $_.Name -notmatch "RDC" `
						-and $_.Name -notmatch "RSAT" `
						-and $_.Name -notmatch "RSAT-Role-Tools" `
						-and $_.Name -notmatch "FS-FileServer" `
						-and $_.Name -notmatch "FS-SMB1" `
						-and $_.Name -notmatch "User-Interfaces-Infra" `
						-and $_.Name -notmatch "Server-Gui-Mgmt-Infra" `
						-and $_.Name -notmatch "Server-Gui-Shell" `
						-and $_.Name -notmatch "PowerShellRoot" `
						-and $_.Name -notmatch "PowerShell" `
						-and $_.Name -notmatch "PowerShell-V2" `
						-and $_.Name -notmatch "PowerShell-ISE" `
						-and $_.Name -notmatch "WoW64-Support" `
					} | select name, displayname, description | sort-object {$_.name}
						# Not sure if to exclude .net 4 stuff as it is installed as default on our server image
						# -and $_.Name -notmatch "NET-Framework-45-Features" `
						# -and $_.Name -notmatch "NET-Framework-45-Core" `
						# -and $_.Name -notmatch "NET-Framework-45-ASPNET" `
						# -and $_.Name -notmatch "NET-WCF-Services45" `
						# -and $_.Name -notmatch "NET-WCF-TCP-PortSharing45" `
				}
			}

			# Installed software
			if ($IncludeSoftware) {
				# Installed software ( old version. doesn't work with remote computers )
				<#
				$InstalledSoftware = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ -system $ComputerName | Get-ItemProperty
				if (Test-path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ -computername $ComputerName ) {
				   $InstalledSoftware += Get-ChildItem HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ -computername $ComputerName | Get-ItemProperty
				}
				$InstalledSoftware = ($InstalledSoftware |
					where {$_.DisplayName -ne $Null `
					-AND $_.SystemComponent -ne "1" `
					-AND $_.ParentKeyName -eq $Null} |
					Select DisplayName, DisplayVersion, Publisher, InstallDate).GetEnumerator() | 
					Sort-Object {"$_"}
				#>	
				
				# Installed software (new version, works with remote computers, but is a bit clunky )
				speak "...Gathering installed software..." $false 
				$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
				$regKey = $Reg.openSubKey("SOFTWARE\\Microsoft\Windows\CurrentVersion\Uninstall")
				$subkeys = $regkey.getsubkeynames()
				$regresults = """DisplayName"",""DisplayVersion"",""Publisher"",""InstallDate"",""Architecture"",""ParentKeyName"",""SystemComponent""`n"
				foreach ($key in $subkeys) {
					$regkey = $Reg.opensubkey($("SOFTWARE\\Microsoft\Windows\CurrentVersion\Uninstall\" + "$key"))
					$regresults += $("""" + $RegKey.GetValue("Displayname") + """,""" `
						+ $RegKey.GetValue("DisplayVersion") + """,""" `
						+ $RegKey.GetValue("Publisher") + """,""" `
						+ $RegKey.GetValue("InstallDate") + """,""" `
						+ "x64"",""" `
						+ $RegKey.GetValue("ParentKeyName") + """,""" `
						+ $RegKey.GetValue("SystemComponent") + """`n")
				}
				$regKey = $Reg.openSubKey("SOFTWARE\\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\")
				$subkeys = $regkey.getsubkeynames()
				foreach ($key in $subkeys) {
					$regkey = $Reg.opensubkey($("SOFTWARE\\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" + "$key"))
					$regresults += $("""" + $RegKey.GetValue("Displayname") + """,""" `
						+ $RegKey.GetValue("DisplayVersion") + """,""" `
						+ $RegKey.GetValue("Publisher") + """,""" `
						+ $RegKey.GetValue("InstallDate") + """,""" `
						+ "x86"",""" `
						+ $RegKey.GetValue("ParentKeyName") + """,""" `
						+ $RegKey.GetValue("SystemComponent") + """`n")
				}
				$installedsoftware = $regresults | ConvertFrom-Csv
				$InstalledSoftware = ($InstalledSoftware |
					Where {$_.DisplayName -ne "" `
					-AND $_.SystemComponent -ne "1" `
					-AND $_.ParentKeyName -eq ""} |
					Select DisplayName, DisplayVersion, Publisher, InstallDate).GetEnumerator() |
					Sort-Object {"$_"}
				$regresults = ""
			}
			
			# Services
			If ($IncludeServices) {
				speak "...Gathering services..."
				$Services = Get-wmiobject win32_service -computername $ComputerName | 
					where { ($_.Caption -notmatch "Windows" `
					-and $_.PathName -notmatch "Windows" `
					-and $_.PathName -notmatch "policyhost.exe" `
					-and $_.Name -ne "LSM" `
					-and $_.PathName -notmatch "OSE.EXE" `
					-and $_.PathName -notmatch "OSPPSVC.EXE" `
					-and $_.PathName -notmatch "Microsoft Security Client") `
					-or $_.Caption -match "DFS Namespace" `
					-or $_.Caption -match "DFS Replication" `
					-or $_.Caption -match "DHCP Server" `
					-or ($_.Displayname -match "Hyper-V" -and $_.State -match "Running")`
					-or $Caption -match "IIS Admin Service" `
					-or $_.Caption -match "Network Policy Server" `
					-or $_.Caption -match "Windows Internal Database" `
					-or $_.Caption -match "World Wide Web Publishing Service" `
					} | sort-object {$_.displayname}
			}
			
			# Processes
			If ($IncludeProcesses){
				speak "...Gathering running processes..."
				$Processes = Get-Process -ComputerName $ComputerName | 
					where { $_.Name -notmatch "conhost" `
					-and $_.Name -notmatch "csrss" `
					-and $_.Name -notmatch "dwm" `
					-and $_.name -notmatch "msseces" `
					-and $_.name -notmatch "explorer" `
					-and $_.name -notmatch "idle" `
					-and $_.name -notmatch "logonui" `
					-and $_.name -notmatch "lsass" `
					-and $_.name -notmatch "msdtc" `
					-and $_.name -notmatch "rdpclip" `
					-and $_.name -notmatch "services" `
					-and $_.Name -notmatch "smss" `
					-and $_.name -notmatch "spoolsv" `
					-and $_.Name -notmatch "svchost" `
					-and $_.name -notmatch "system" `
					-and $_.name -notmatch "taskeng" `
					-and $_.name -notmatch "taskhostex" `
					-and $_.name -notmatch "taskmgr" `
					-and $_.name -notmatch "wininit" `
					-and $_.name -notmatch "winlogon" `
					-and $_.Name -notmatch "wmiprvse"`
					} | select Name, Id
					
				if (!($IncludeServices)) {
					$Services = Get-wmiobject win32_service -computername $ComputerName | 
						where { ($_.Caption -notmatch "Windows" `
						-and $_.PathName -notmatch "Windows" `
						-and $_.PathName -notmatch "policyhost.exe" `
						-and $_.Name -ne "LSM" `
						-and $_.PathName -notmatch "OSE.EXE" `
						-and $_.PathName -notmatch "OSPPSVC.EXE" `
						-and $_.PathName -notmatch "Microsoft Security Client") `
						-or ($_.Displayname -match "Hyper-V" -and $_.State -match "Running")`
						-or $_.Caption -match "Windows Internal Database" `
						} | select Name, PathName, ProcessID
				}
				
				$NonDuplicateProcesses = """Name"",""Id""`n"
				foreach ($Process in $Processes) {
					$ProcessName = $Process.name
					$Duplicate = $False
					foreach ($Service in $Services) {
						If ($Service.Pathname -match "${ProcessName}" -or $Service.ProcessID -eq $Process.Id) { $Duplicate = $True }
					}
					If (!($Duplicate)) {$NonDuplicateProcesses += """" + $Process.name + """,""" + $Process.id + """`n"}
				}
				If (!($IncludeServices)) { $Services = ""}	
				$Processes = $NonDuplicateProcesses | ConvertFrom-Csv
				$NonDuplicateProcesses = ""
			}
			
			if ($IncludeTasks) {
				# Scheduled tasks
				speak "...Gathering scheduled tasks..." $false 
				#Get all scheduled tasks on the system in a Csv format
				$Tasks = schtasks /query /s $ComputerName /v /fo csv | ConvertFrom-Csv
				#Filtering out all Windows tasks for Windows 2k3 and 2k12 (and 2k8?)
				$ScheduledTasks = $Tasks | Where-Object { $_.HostName -eq $ComputerName `
					-and $_.Author -ne "N/A" `
					-and $_.'Next Run Time' -ne "N/A" `
					-and $_.Author -notmatch "Microsoft" `
					-and $_.TaskName -notmatch "User_Feed_Synchronization" `
					} | sort-object {$_.name}
				# -and $_.'Scheduled Task State' -ne "Disabled" # <-- Add this to the where-object selection in the line above to filter out disabled tasks as well
				$Tasks = ""
			}
			speak "...Finished gathering."
			speak ""
			speak "Generating report..."

			
			if ($OutputText) {
				$textfile = "${OutputPath}\${computername}_text_" + (get-date -f "yyyy-MM-dd") + ".txt"
				$textstream = [System.IO.StreamWriter] "$textfile" 
				speak "Saving text report to ${textfile}." 
			}
			
			if ($OutputWiki) {
				$wikifile = "${OutputPath}\${computername}_wiki_" + (get-date -f "yyyy-MM-dd") + ".aspx"
				$wikistream = [System.IO.StreamWriter] "$wikifile"
				speak "Saving wiki report to ${wikifile}."
				$wikiHeader = $(Get-Content "${WikiHeaderFile}") -replace "\!\!COMPUTERNAME\!\!", $ComputerName
				foreach ($line in $wikiheader) {
					$wikistream.WriteLine($line)
				}			
			}

			
			## Display system information ##
			computerHeading "$computername"

			if ($IncludeSystem) {
				sectionHeading "System Information"
				switch($cpuInfo.DataWidth){
					"64"{$arch = "x64-based processor"}
					"32"{$arch = "x86-based processor"}
				}
				itemStart
				itemHeading "Report Generated $(Get-Date)"
				dataout "Manufacturer" $computerinfo.manufacturer 5
				dataout "Model" $computerInfo.model 5
				dataout "Computer Description" $computerInfo.description 5
				dataout "CPU" $cpuInfo.Name 5
				dataout "CPU Description" $cpuInfo.description 5
				dataout "Number of CPUs" $computerInfo.numberofprocessors 5
				dataout "Memory" "${physicalmemory}GB" 5
				dataout "Operating System" $osinfo.caption 5
				dataout "System Type" "${ostype} operating system, ${arch}" 5
				if ($computerinfo.partofdomain -eq "True"){
					dataout "Full Computer Name" $($computerInfo.dnshostname + "." + $computerinfo.domain) 5
					dataout "Domain" $computerInfo.domain 5
				} else {
					dataout "Workgroup" $computerInfo.workgroup 5
				}
				itemEnd
				sectionEnd
			}
			## Display network adapters ##
			
			if ($NetworkAdapters) {
				sectionheading "Network Adapters"
				foreach ($adapter in $NetworkAdapters) {
					itemstart
					$adapterdescription = $adapter.description
					$adaptermac = $adapter.macaddress
					itemheading "$adapterdescription - $adaptermac"
					dataout "IP Address" $adapter.IpAddress[0]
					dataout "Subnet Mask" $adapter.IpSubnet[0]
					dataout "Default Gateway" $adapter.DefaultIpGateway
					dataout "DHCP Enabled" $adapter.DhcpEnabled
					dataout "DNS Servers" $adapter.DnsServerSearchOrder
					itemend
				}
				sectionend
			}
			
			## Display storage volumes ##

			if ($StorageVolumes) {
				sectionheading "Storage Volumes"
				foreach ($volume in $StorageVolumes) {
					$Name = $volume.Name
					itemStart
					#dataout "Name" $Name
					itemHeading $Name
					if ($volume.DriveType -eq 3) {
						$HSize = humanSize $volume.Capacity
						$Size = $volume.Capacity
						dataout "Label" $volume.Label
						dataout "DriveType" "Local Disk"
						dataout "Size" "${HSize} (${Size} bytes)"
					} elseif ($volume.DriveType -eq 5) {
						dataout "DriveType" "Optical Drive"
					}
					itemEnd
				}
				sectionEnd
			}
			
			## Display shared folders ##

			if ($sharedfolders) {
				sectionHeading "Shared Folders"
				foreach ($folder in $sharedfolders) {
					itemStart
					itemHeading $folder.Name
					dataout "Path" $folder.Path
					dataout "Description" $folder.Description
					itemEnd
				}
				sectionEnd
			}

			## Display local user acccounts

			if ($userlist) {
				sectionHeading "Local User Accounts"
				foreach ($user in $userlist){
					itemStart
					itemHeading $user.name
					dataout "Description"  $user.description
					itemEnd
				}
				sectionEnd
			}

			## Display server roles and features ##

			if ($Roles) {
				sectionHeading "Server Roles"
				foreach ($role in $Roles){
					itemStart
					itemHeading $role.Displayname
					dataout "Role Name" $role.name
					dataout "Running Service" $role.systemservice
					dataout "Description" $role.description
					itemEnd
				}
				sectionEnd
			}

			if ($Features) {
				sectionheading "Server Features"
				foreach ($feature in $Features){
					itemStart
					itemHeading $feature.displayname
					dataout "Feature Name" $feature.name
					dataout "Description" $feature.description
					itemEnd
				}
				sectionEnd
			}

			## Display installed software ##

			if ($installedsoftware) {
				sectionHeading "Installed Software"
				foreach ($application in $installedsoftware) {
					itemStart
					itemHeading $application.DisplayName
					dataout "Version" $application.displayversion
					dataout "Publisher" $application.publisher
					if ($application.installdate) {
						try {
							$installdate = [datetime]::ParseExact($application.installdate,"yyyyMMdd",$null)
							$installdate = get-date $installdate -format dd/MM/yyy
						} catch {
							$installdate = "Unknown"
						}
						dataout "Installed on" $installdate
					} else {
						dataout "Installed on" "Unknown"
					}
					itemEnd
				}
				sectionEnd
			}

			## Display non-default services ##

			if ($Services) {
				sectionHeading "Services"
				foreach ($service in $Services) {
					itemStart
					$ServiceState = $service.State
					$ServiceStatus = $service.Status
					itemHeading $service.DisplayName # Service Display Name (full name)
					dataout "Service Name" $service.name
					$PathName = $service.PathName -replace "`"", "" #"
					dataout "Path" $PathName # Service Executable
					dataout "Startup Type" $service.StartMode # Service Startup mode
					dataout "Log On As" $service.StartName # Service RunAs Account
					dataout "Status"  "$ServiceState, $ServiceStatus" # Service State (running/stopped etc), Service Status
					dataout "Running" $service.Started # Service Started status
					dataout "Description" $service.Description # Service Description
					itemEnd
				}
				sectionEnd
			}
			
			## Display processes ##
			if ($Processes) {
				if ($TextWiki) {$TextWiki = $false; $TextToggle = $TRUE}
				if ($OutputWiki) {$OutputWiki = $false; $WikiToggle = $TRUE}
				sectionHeading "Processes"
				speak "Running processes at time of report." $TRUE
				speak "Common processes and processes of running services with matching paths and/or process IDs are excluded."
				speak "This information will not be exported to text or wiki as it is only really useful right now." $TRUE
				speak "" $TRUE
				foreach ($process in $Processes) {
					itemstart
					itemHeading $process.name
					dataout "Process ID" $process.id
					itemend
				}
				sectionend
				if ($TextToggle) {$OutputText = $TRUE; $TextToggle = $false}
				if ($WikiToggle) {$OutputWiki = $TRUE; $WikiToggle = $false}
			}

			## Display non-default scheduled tasks ##

			if ($ScheduledTasks) {
				sectionHeading "Scheduled Tasks"
				Foreach($ScheduledTask in $ScheduledTasks)
				{
					$Tasktext = ""
					itemStart
					itemHeading $ScheduledTask.'TaskName'.Substring($ScheduledTask.'TaskName'.IndexOf("\")+1)
					dataout "Start in" $ScheduledTask.'Start In'
					dataout "Task" $ScheduledTask.'Task To Run'
					#In case of W2k12 (and W2k8?)
					If($ScheduledTask.'Schedule Type'){
						Switch($ScheduledTask.'Schedule Type'){
							"Hourly " { $Tasktext = $ScheduledTask.'Schedule Type' + "at " + $ScheduledTask.'Start Time' }
							"Daily " { $Tasktext = $ScheduledTask.'Schedule Type' + "at " + $ScheduledTask.'Start Time' }
							"Weekly" { $Tasktext = $ScheduledTask.'Schedule Type' + " on every " + $ScheduledTask.Days + " at " + $ScheduledTask.'Start Time' }
								"Monthly"
							{
								If($ScheduledTask.Months -eq "Every month") { $Tasktext = $ScheduledTask.'Schedule Type' + " on day " + $ScheduledTask.Days + " at " + $ScheduledTask.'Start Time'}
								Else { $Tasktext = "Yearly on day " + $ScheduledTask.Days + " of " + $ScheduledTask.Months + " at " + $ScheduledTask.'Start Time' }
							}
						}
					}
					#In case of W2k3
						If($ScheduledTask.'Scheduled Type') {
							Switch($ScheduledTask.'Scheduled Type'){
							"Hourly " { $Tasktext = $ScheduledTask.'Scheduled Type' + "at " + $ScheduledTask.'Start Time' }
							"Daily " { $Tasktext = $ScheduledTask.'Scheduled Type' + "at " + $ScheduledTask.'Start Time' }
							"Weekly" { $Tasktext = $ScheduledTask.'Scheduled Type' + " on every " + $ScheduledTask.Days + " at " + $ScheduledTask.'Start Time' }
							"Monthly"
							{
								If($ScheduledTask.Months -eq "JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC") { $Tasktext = $ScheduledTask.'Scheduled Type' + " on day " + $ScheduledTask.Days + " at " + $ScheduledTask.'Start Time' }
								Else { $Tasktext = "Yearly on day " + $ScheduledTask.Days + " of " + $ScheduledTask.Months + " at " + $ScheduledTask.'Start Time' }
							}
						}
					}
					#This line can be removed if the filter excludes disabled tasks
					If($ScheduledTask.'Scheduled Task State' -eq "Disabled") { $Tasktext = "Disabled" }
					dataout "Info" $Tasktext
					itemEnd
				}
				sectionEnd
			}

			## Run extra system checks ##
			
            if ($IncludeSystemChecks) {
			    sectionHeading "System Checks"
			    $OddRow = $TRUE
				#Check for Configuration manager processes
				$CCMProcess = Get-Process -ComputerName $ComputerName | where { $_.Name -eq "CcmExec" }
				If ($CCMProcess) {
					checkresult 0 "Configuration Manager is running"
				} else {
					checkresult 2 "Configuration Manager is not running. Computer cannot be managed."
				}
			    #Check for SNMP service
			    $SNMPService = Get-wmiobject win32_service -computername $ComputerName | where { $_.Caption -eq "SNMP Service" }
			    if ($SNMPService){
				    checkresult 0 "SNMP Service is installed."
			    } else {
				    checkresult 1 "SNMP Service is not installed. Monitoring services will not be able to check this system correctly."
			    }
			    sectionEnd
            }

			if ($OutputText) { $textstream.close() }
			if ($OutputWiki) { 
				$wikiFooter = $(Get-Content "${WikiFooterFile}") -replace "\!\!COMPUTERNAME\!\!", $ComputerName
				foreach ($line in $wikifooter) {
					$wikistream.WriteLine($line)
				}	
				$wikistream.close()
			}
			if ($UploadWiki) { 
				uploadWiki "${wikifile}"
			}
		}
		speak ""
	}
}
