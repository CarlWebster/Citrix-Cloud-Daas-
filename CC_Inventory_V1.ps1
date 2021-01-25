﻿#Requires -Version 5.0
#This File is in Unicode format. Do not edit in an ASCII editor. Notepad++ UTF-8-BOM

#region help text

<#
.SYNOPSIS
	Creates an inventory of a Citrix Cloud Site.
.DESCRIPTION
	Creates an inventory of a Citrix Cloud Site using Microsoft PowerShell, Word, plain 
	text, or HTML.
	
	This script requires at least PowerShell version 5.
	
	This script must run from an elevated PowerShell session.

	The default output is HTML.
	
	Run this script on a computer with the Remote SDK installed.
	
	https://download.apps.cloud.com/CitrixPoshSdk.exe
	
	This script was developed and run from two Windows 10 VMs. One was domain-joined and 
    the other was in a Workgroup.
	
	This script supports only Citrix Cloud, not the on-premises CVAD products.
	
	If you are running XA/XD 7.0 through 7.7, please use: 
	https://carlwebster.com/downloads/download-info/xenappxendesktop-7-x-documentation-script/

	If you are running XA/XD 7.8 through CVAD 2006, please use:
	https://carlwebster.com/downloads/download-info/xenappxendesktop-7-8/

	If you are running CVAD 2006 and later, please use:
	https://carlwebster.com/downloads/download-info/citrix-virtual-apps-and-desktops-v3-script/

	To prevent multiple Citrix Cloud authentication prompts, follow the instructions in 
	the Authentication section of the ReadMe file to create a profile named Default.
	
	ReadMe file: https://carlwebster.sharefile.com/d-sb4e144f9ecc48e7b
	
	By default, only gives summary information for:
		Administrators
		App-V Publishing
		Application Groups
		Applications
		Delivery Groups
		Hosting
		Logging
		Machine Catalogs
		Policies
		StoreFront
		Zones

	The Summary information is what is shown in the top half of Citrix Studio for:
		Machine Catalogs
		Delivery Groups
		Applications
		Policies
		Logging
		Administrators
		Hosting
		StoreFront
		App-V Publishing
		Zones

	Using the MachineCatalogs parameter can cause the report to take a very long time to 
	complete and can generate an extremely long report.
	
	Using the DeliveryGroups parameter can cause the report to take a very long time to 
	complete and can generate an extremely long report.

	Using both the MachineCatalogs and DeliveryGroups parameters can cause the report to 
	take an extremely long time to complete and generate an exceptionally long report.

	Creates an output file named after the CC Site (which by default is cloudxdsite), unless 
	you use the SiteName parameter.
	
	Word and PDF Document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish
		
.PARAMETER HTML
	Creates an HTML file with an.html extension.
	This parameter is set True if no other output format is selected.
.PARAMETER Text
	Creates a formatted text file with a.txt extension.
	This parameter is disabled by default.
.PARAMETER ProfileName
	The profile name to use for Get-XDAuthentication.
	
	The name associated with a set of credentials in the local store that are to be 
	read.
	
	You must follow the process in either the ReadMe file or your own process to capture 
	the Client ID and Client Secret and save them to a CSV credential profile.
	
	ReadMe file: https://carlwebster.sharefile.com/d-sb4e144f9ecc48e7b

	To prevent multiple Citrix Cloud authentication prompts, create a profile named 
	Default.

	This parameter is blank by default.
	This parameter has an alias of PN.
.PARAMETER SiteName
	Site Name to use for the all output files. 

	The default is "cloudxdsite".
	This parameter has an alias of SN.
.PARAMETER Administrators
	Give detailed information for Administrator Scopes and Roles.
	This parameter is disabled by default.
	This parameter has an alias of Admins.
.PARAMETER Applications
	Gives detailed information for all applications.
	This parameter is disabled by default.
	This parameter has an alias of Apps.
.PARAMETER DeliveryGroups
	Gives detailed information on all desktops in all Desktop (Delivery) Groups.
	
	Using the DeliveryGroups parameter can cause the report to take a very long 
	time to complete and can generate an extremely long report.
	
	Using both the MachineCatalogs and DeliveryGroups parameters can cause the 
	report to take an extremely long time to complete and generate an exceptionally 
	long report.
	
	This parameter is disabled by default.
	This parameter has an alias of DG.
.PARAMETER DeliveryGroupsUtilization
	Gives a chart with the delivery group utilization for the last 7 days 
	depending on the information in the database.
	
	This option is only available when the report is generated in Word and requires 
	Microsoft Excel to be locally installed.
	
	Using the DeliveryGroupsUtilization parameter causes the report to take a longer 
	time to complete and generates a longer report.
	
	This parameter is disabled by default.
	This parameter has an alias of DGU.
.PARAMETER Hosting
	Give detailed information for Hosts, Host Connections, and Resources.

	This parameter is disabled by default.
	This parameter has an alias of Host.
.PARAMETER Logging
	Give the Configuration Logging report with, by default, details for the previous 
	seven days.
	
	For Citrix Cloud, there are no Logging preferences or information about the logging 
	database.

	This parameter is disabled by default.
.PARAMETER StartDate
	The start date for the Configuration Logging report.
	
	The format for date only is MM/DD/YYYY.
	
	Format to include a specific time range is "MM/DD/YYYY HH:MM:SS" in 24-hour format.
	The double quotes are needed.
	
	The default is today's date minus seven days.
	This parameter has an alias of SD.
.PARAMETER EndDate
	The end date for the Configuration Logging report.
	
	The format for date only is MM/DD/YYYY.
	
	Format to include a specific time range is "MM/DD/YYYY HH:MM:SS" in 24-hour format.
	The double quotes are needed.
	
	The default is today's date.
	This parameter has an alias of ED.
.PARAMETER MachineCatalogs
	Gives detailed information for all machines in all Machine Catalogs.
	
	Using the MachineCatalogs parameter can cause the report to take a very long 
	time to complete and can generate an extremely long report.
	
	Using both the MachineCatalogs and DeliveryGroups parameters can cause the 
	report to take an extremely long time to complete and generate an exceptionally 
	long report.
	
	This parameter is disabled by default.
	This parameter has an alias of MC.
.PARAMETER NoADPolicies
	Excludes all Citrix AD-based policy information from the output document.
	Includes only Site policies created in Studio.
	
	This Switch is useful in large AD environments, where there may be thousands
	of policies, to keep SYSVOL from being searched.
	
	This parameter is disabled by default.
	This parameter has an alias of NoAD.
.PARAMETER NoPolicies
	Excludes all Site and Citrix AD-based policy information from the output document.
	
	Using the NoPolicies parameter will cause the Policies parameter to be set to False.
	
	This parameter is disabled by default.
	This parameter has an alias of NP.
.PARAMETER NoSessions
	Excludes Machine Catalog, Application and Hosting session data from the report.
	
	Using the MaxDetails parameter does not change this setting.
	
	This parameter is disabled by default.
	This parameter has an alias of NS.
.PARAMETER Policies
	Give detailed information for both Site and Citrix AD-based Policies.
	
	Using the Policies parameter can cause the report to take a very long time 
	to complete and can generate an extremely long report.
	
	Note: The Citrix Group Policy PowerShell module will not load from an elevated 
	PowerShell session. 
	If the module is manually imported, the module is not detected from an elevated 
	PowerShell session.
	
	There are three related parameters: Policies, NoPolicies, and NoADPolicies.
	
	Policies and NoPolicies are mutually exclusive and priority is given to NoPolicies.
	
	This parameter is disabled by default.
	This parameter has an alias of Pol.
.PARAMETER StoreFront
	Give detailed information for StoreFront.
	This parameter is disabled by default.
	This parameter has an alias of SF.
.PARAMETER VDARegistryKeys
	Adds information on registry keys to the Machine Details section.
	
	If this parameter is used, MachineCatalogs is set to True.
	
	This parameter is disabled by default.
	This parameter has an alias of VRK.
.PARAMETER MaxDetails
	Adds maximum detail to the report.
	
	This is the same as using the following parameters:
		Administrators
		Applications
		DeliveryGroups
		Hosting
		Logging
		MachineCatalogs
		Policies
		StoreFront
		VDARegistryKeys

	Does not change the value of NoADPolicies.
	Does not change the value of NoSessions.
	
	WARNING: Using this parameter can create an extremely large report and 
	can take a very long time to run.

	This parameter has an alias of MAX.
.PARAMETER Section
	Processes a specific section of the report.
	Valid options are:
		Admins (Administrators)
		Apps (Applications and Application Group Details)
		AppV
		Catalogs (Machine Catalogs)
		Config (Configuration)
		Groups (Delivery Groups)
		Hosting
		Licensing
		Logging
		Policies
		StoreFront
		Zones
		All
	This parameter defaults to All sections.
	
	Notes:
	Using Logging will force the Logging Switch to True.
	Using Policies will force the Policies Switch to True.
	If Policies is selected and the NoPolicies Switch is used, the script will terminate.
.PARAMETER AddDateTime
	Adds a date timestamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2021 at 6PM is 2021-06-01_1800.
	Output filename will be ReportName_2021-06-01_1800.docx (or.pdf).
	This parameter is disabled by default.
	This parameter has an alias of ADT.
.PARAMETER CSV
	Will create a CSV file for each Appendix.
	The default value is False.
	
	Output CSV filename is in the format:
	
	CCSiteName_Documentation_Appendix#_NameOfAppendix.csv
	
	For example:
		CCSiteName_Documentation_AppendixA_VDARegistryItems.csv
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script runs.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script runs.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is disabled by default.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	This parameter requires Microsoft Word to be installed.
	This parameter uses the Word SaveAs PDF capability.
.PARAMETER CompanyAddress
	Company Address to use for the Cover Page, if the Cover Page has the Address field.
	
	The following Cover Pages have an Address field:
		Banded (Word 2013/2016)
		Contrast (Word 2010)
		Exposure (Word 2010)
		Filigree (Word 2013/2016)
		Ion (Dark) (Word 2013/2016)
		Retrospect (Word 2013/2016)
		Semaphore (Word 2013/2016)
		Tiles (Word 2010)
		ViewMaster (Word 2013/2016)
		
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CA.
.PARAMETER CompanyEmail
	Company Email to use for the Cover Page, if the Cover Page has the Email field. 
	
	The following Cover Pages have an Email field:
		Facet (Word 2013/2016)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax to use for the Cover Page, if the Cover Page has the Fax field. 
	
	The following Cover Pages have a Fax field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CF.
.PARAMETER CompanyName
	Company Name to use for the Cover Page. 
	The default value is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.

	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CN.
.PARAMETER CompanyPhone
	Company Phone to use for the Cover Page if the Cover Page has the Phone field. 
	
	The following Cover Pages have a Phone field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CPh.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)

	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly
		works in 2010, but Subtitle/Subject & Author fields need moving
		after the title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 
		36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually 
		resized or font changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 
		2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)

	The default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UserName
	Username to use for the Cover Page and Footer.
	The default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	The default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	The default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1
	
	Creates an HTML report by default.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -ProfileName "My CC Profile"
	
	Creates an HTML report by default.
	
	Uses the stored profile named "My CC Profile".
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -MSWord
	
	Uses all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Creates a Microsoft Word report.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	****************************************************
	* Example for using the script as a scheduled task *
	****************************************************
	
	First, from an elevated PowerShell session, create a profile named Default.
	
	Set-XDCredentials -CustomerID "CustomerID" -ProfileType CloudAPI -StoreAs "Default" 
	-APIKey "ID" -SecretKey "Secret"

	Note: Review the instructions in the Authentication section of the ReadMe file for the 
	details on creating a profile named Default.
	
	ReadMe file: https://carlwebster.sharefile.com/d-sb4e144f9ecc48e7b
	
	PowerShell.exe -NoLogo -File "C:\PSScript\CC_Inventory_V1.ps1 -MaxDetails 
	-SiteName MyCCSite' -AddDateTime"	
	
	Set the following parameter values:
	
		Administrators      = True
		Applications        = True
		DeliveryGroups      = True
		Hosting             = True
		Logging             = True
		MachineCatalogs     = True
		Policies            = True
		StoreFront          = True
		VDARegistryKeys     = True
		
		NoPolicies          = False
		Section             = "All"

	Creates an HTML report named MyCCSite_YYYY-MM-DD_HHMM.html.

	Uses the profile named Default, which prevents any extraneous Citrix Cloud 
	authentication prompts.

	For more information on running a PowerShell script as a scheduled task, see:	
	https://carlwebster.com/running-powershell-scripts-scheduled-task/
	https://carlwebster.com/running-powershell-scheduled-tasks-third-opinion/	
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -PDF
	
	Uses all default values and saves the document as a PDF file.

	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -Text

	Saves the document as a formatted text file.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -HTML

	Saves the document as an HTML file.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -MachineCatalogs
	
	Creates an HTML report with full details for all machines in all Machine Catalogs.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -DeliveryGroups
	
	Creates an HTML report with full details for all desktops in all Desktop (Delivery) 
	Groups.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -DeliveryGroupsUtilization -MSWord
	
	Note: Using DeliveryGroupsUtilization requires the use of MSWord or PDF. Microsoft Excel 
	is also required.
	
	Creates a Microsoft Word report with utilization details for all Desktop (Delivery) 
	Groups.

	Uses all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -DeliveryGroupsUtilization -PDF
	
	Note: Using DeliveryGroupsUtilization requires the use of MSWord or PDF. Microsoft Excel 
	is also required.
	
	Creates a PDF report with utilization details for all Desktop (Delivery) Groups.

	Uses all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -DeliveryGroups -MachineCatalogs
	
	Creates an HTML report with full details for all machines in all Machine Catalogs and 
	all desktops in all Delivery Groups.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -Applications
	
	Creates an HTML report with full details for all applications.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -Policies
	
	Creates an HTML report with full details for Policies.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
	
	Note: If a profile named Default does not exist, you may be prompted for Citrix Cloud 
	credentials.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -NoPolicies
	
	Creates an HTML report with no Policy information.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -NoADPolicies
	
	Creates an HTML report with no Citrix AD-based Policy information.
	
	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
	
	Note: If a profile named Default does not exist, you may be prompted for Citrix Cloud 
	credentials.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -Policies -NoADPolicies
	
	Creates an HTML report with full details on Site policies created in Studio but 
	no Citrix AD-based Policy information.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
	
	Note: If a profile named Default does not exist, you may be prompted for Citrix Cloud 
	credentials.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -Logging -StartDate 09/01/2021 -EndDate 
	09/30/2021	
	
	Creates an HTML report with Configuration Logging details for the dates 09/01/2021 
	through 09/30/2021.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CVAD_Inventory_V3.ps1 -Logging -StartDate "09/01/2021 10:00:00" 
	-EndDate "09/01/2021 14:00:00" -MSWord
	
	Creates a Microsoft Word report with Configuration Logging details for the time range 
	09/01/2021 10:00:00AM through 09/01/2021 02:00:00PM.
	
	Narrowing the report down to seconds does not work. Seconds must be either 00 or 59.
	
	Uses all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -Administrators
	
	Creates an HTML report with full details on Administrator Scopes and Roles.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -Hosting
	
	Creates an HTML report with full details for Hosts, Host Connections, and Resources.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -StoreFront
	
	Creates an HTML report with full details for StoreFront.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -MachineCatalogs -DeliveryGroups 
	-Applications -Policies -Hosting -StoreFront	
	
	Creates an HTML report with full details for all:
		Machines in all Machine Catalogs
		Desktops in all Delivery Groups
		Applications
		Policies
		Hosts, Host Connections, and Resources
		StoreFront

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
	
	Note: If a profile named Default does not exist, you may be prompted for Citrix Cloud 
	credentials.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -MC -DG -Apps -Policies -Hosting -PDF
	
	Creates a PDF report with full details for all:
		Machines in all Machine Catalogs
		Desktops in all Delivery Groups
		Applications
		Policies
		Hosts, Host Connections, and Resources
		
	Uses all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
	
	Note: If a profile named Default does not exist, you may be prompted for Citrix Cloud 
	credentials.
.EXAMPLE
	PS C:\PSScript.\CC_Inventory_V1.ps1 -MSWord -CompanyName "Carl Webster 
	Consulting" -CoverPage "Mod" -UserName "Carl Webster"
	
	Creates a Microsoft Word report.
	Uses:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript.\CC_Inventory_V1.ps1 -CN "Carl Webster Consulting" -CP "Mod" 
	-UN "Carl Webster" -MSWord
	
	Creates a Microsoft Word report.
	Uses:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript.\CC_Inventory_V1.ps1 -MSWord -CompanyName "Sherlock Holmes 
	Consulting" -CoverPage Exposure -UserName "Dr. Watson" -CompanyAddress "221B Baker 
	Street, London, England" -CompanyFax "+44 1753 276600" -CompanyPhone "+44 1753 276200"
	
	Creates a Microsoft Word report.
	Uses:
		Sherlock Holmes Consulting for the Company Name.
		Exposure for the Cover Page format.
		Dr. Watson for the User Name.
		221B Baker Street, London, England for the Company Address.
		+44 1753 276600 for the Company Fax.
		+44 1753 276200 for the Company Phone.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript.\CC_Inventory_V1.ps1 -MSWord -CompanyName "Sherlock Holmes 
	Consulting" -CoverPage Facet -UserName "Dr. Watson" -CompanyEmail 
	SuperSleuth@SherlockHolmes.com

	Creates a Microsoft Word report.
	Uses:
		Sherlock Holmes Consulting for the Company Name.
		Facet for the Cover Page format.
		Dr. Watson for the User Name.
		SuperSleuth@SherlockHolmes.com for the Company Email.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -AddDateTime
	
	Creates an HTML report.
	Adds a date time stamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2021 at 6PM is 2021-06-01_1800.
	Output filename will be CCSiteName_2021-06-01_1800.docx

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -PDF -AddDateTime
	
	Creates a PDF report.
	Uses all Default values and saves the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2021 at 6PM is 2021-06-01_1800.
	Output filename will be CCSiteName_2021-06-01_1800.pdf

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -Folder \\FileServer\ShareName
	
	Creates an HTML report.
	Output file is saved in the path \\FileServer\ShareName

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -Section Policies
	
	Creates an HTML report that contains only policy information.
	Processes only the Policies section of the report.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
	
	Note: If a profile named Default does not exist, you may be prompted for Citrix Cloud 
	credentials.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -Section Groups -DG
	
	Creates an HTML report.
	Processes only the Delivery Groups section of the report with Delivery Group details.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -Section Groups
	
	Creates an HTML report.
	Processes only the Delivery Groups section of the report with no Delivery Group details.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -VDARegistryKeys
	
	Creates an HTML report.
	Adds the information on VDA registry keys to Appendix A.
	Forces the MachineCatalogs parameter to $True

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -MaxDetails
	
	Set the following parameter values:
		Administrators      = True
		Applications        = True
		DeliveryGroups      = True
		Hosting             = True
		Logging             = True
		MachineCatalogs     = True
		Policies            = True
		StoreFront          = True
		VDARegistryKeys     = True
		
		NoPolicies          = False
		Section             = "All"

	Creates an HTML report.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
	
	Note: If a profile named Default does not exist, you may be prompted for Citrix Cloud 
	credentials.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -MaxDetails -HTML -MSWord -PDF -Text
	
	Creates four reports: HTML, Microsoft Word, PDF, and plain text.
	
	For Microsoft Word and PDF, Uses all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Set the following parameter values:
		Administrators      = True
		Applications        = True
		DeliveryGroups      = True
		Hosting             = True
		Logging             = True
		MachineCatalogs     = True
		Policies            = True
		StoreFront          = True
		VDARegistryKeys     = True
		
		NoPolicies          = False
		Section             = "All"

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
	
	Note: If a profile named Default does not exist, you may be prompted for Citrix Cloud 
	credentials.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -Dev -ScriptInfo -Log
	
	Creates an HTML report.
	
	Creates a text file named CCInventoryScriptErrors_yyyy-MM-dd_HHmm.txt that 
	contains up to the last 250 errors reported by the script.
	
	Creates a text file named CCInventoryScriptInfo_yyyy-MM-dd_HHmm.txt that 
	contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	CCDocScriptTranscript_yyyy-MM-dd_HHmm.txt.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -CSV
	
	Uses all Default values.

	Creates a CSV file for each Appendix.
	For example:
		CCSiteName_Documentation_AppendixA_VDARegistryItems.csv

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -MaxDetails -HTML -MSWord -PDF -Text -Dev 
	-ScriptInfo -Log -CSV
	
	Creates four reports: HTML, Microsoft Word, PDF, and plain text.
	
	Creates a text file named CCInventoryScriptErrors_yyyy-MM-dd_HHmm.txt that 
	contains up to the last 250 errors reported by the script.
	
	Creates a text file named CCInventoryScriptInfo_yyyy-MM-dd_HHmm.txt that 
	contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	CCDocScriptTranscript_yyyy-MM-dd_HHmm.txt.

	Creates a CSV file for each Appendix.
	For example:
		CCSiteName_Documentation_AppendixA_VDARegistryItems.csv

	For Microsoft Word and PDF, Uses all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Set the following parameter values:
		Administrators      = True
		Applications        = True
		DeliveryGroups      = True
		Hosting             = True
		Logging             = True
		MachineCatalogs     = True
		Policies            = True
		StoreFront          = True
		VDARegistryKeys     = True
		
		NoPolicies          = False
		Section             = "All"

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
	
	Note: If a profile named Default does not exist, you may be prompted for Citrix Cloud 
	credentials.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -SmtpServer mail.domain.tld -From 
	CCAdmin@domain.tld -To ITGroup@domain.tld	

	The script Uses the email server mail.domain.tld, sending from CCAdmin@domain.tld and 
	sending to ITGroup@domain.tld.

	The script Uses the default SMTP port 25 and does not use SSL.

	If the current user's credentials are not valid to send email, the script prompts 
	the user to enter valid credentials.

	Creates an HTML report.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -SmtpServer mailrelay.domain.tld -From 
	Anonymous@domain.tld -To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script uses the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld, sending to ITGroup@domain.tld.

	To send an unauthenticated email using an email relay server requires the From email 
	account to use the name Anonymous.

	The script uses the default SMTP port 25 and does not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send an email using a Gmail or g-suite account, you may have to turn ON the "Less 
	secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script will generate an anonymous, secure password for the anonymous@domain.tld 
	account.

	Creates an HTML report.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -SmtpServer 
	labaddomain-com.mail.protection.outlook.com -UseSSL -From 
	SomeEmailAddress@labaddomain.com -To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script uses the email server labaddomain-com.mail.protection.outlook.com, sending 
	from SomeEmailAddress@labaddomain.com and sending to ITGroupDL@labaddomain.com.

	The script uses the default SMTP port 25 and SSL.

	Creates an HTML report.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	The script uses the email server smtp.office365.com on port 587 using SSL, sending from 
	webster@carlwebster.com and sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.

	Creates an HTML report.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.EXAMPLE
	PS C:\PSScript >.\CC_Inventory_V1.ps1 -SmtpServer smtp.gmail.com -SmtpPort 587 
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send an email using a Gmail or g-suite account, you may have to turn ON the "Less 
	secure app access" option on your account.
	*** NOTE ***
	
	The script uses the email server smtp.gmail.com on port 587 using SSL, sending from 
	webster@gmail.com and sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, the script prompts the 
	user to enter valid credentials.

	Creates an HTML report.

	If no authentication profile exists, prompts for Citrix Cloud credentials.
	If a profile named Default exists, uses the credentials stored in the Default profile.
.INPUTS
	None. You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script. 
	This script creates a Word, PDF, plain text, or HTML document.
.NOTES
	NAME: CC_Inventory_V1.ps1
	VERSION: 1.11
	AUTHOR: Carl Webster
	LASTEDIT: January 25, 2021
#>

#endregion

#region script parameters
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "") ]

Param(
	[parameter(Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(Mandatory=$False)] 
	[Alias("PN")]
	[string]$ProfileName="",	
	
	[parameter(Mandatory=$False)] 
	[Alias("SN")]
	[string]$SiteName="",
    
	[parameter(Mandatory=$False)] 
	[Alias("Admins")]
	[Switch]$Administrators=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("Apps")]
	[Switch]$Applications=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("DG")]
	[Switch]$DeliveryGroups=$False,	

	[parameter(Mandatory=$False)] 
	[Alias("DGU")]
	[Switch]$DeliveryGroupsUtilization=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("Host")]
	[Switch]$Hosting=$False,	
	
	[parameter(Mandatory=$False)] 
	[Switch]$Logging=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("SD")]
	[Datetime]$StartDate = ((Get-Date -displayhint date).AddDays(-7)),

	[parameter(Mandatory=$False)] 
	[Alias("ED")]
	[Datetime]$EndDate = (Get-Date -displayhint date),
	
	[parameter(Mandatory=$False)] 
	[Alias("MC")]
	[Switch]$MachineCatalogs=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("NoAD")]
	[Switch]$NoADPolicies=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("NP")]
	[Switch]$NoPolicies=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("NS")]
	[Switch]$NoSessions=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("Pol")]
	[Switch]$Policies=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("SF")]
	[Switch]$StoreFront=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("VRK")]
	[Switch]$VDARegistryKeys=$False,

	[parameter(Mandatory=$False)] 
	[Alias("MAX")]
	[Switch]$MaxDetails=$False,

	[ValidateSet('All', 'Admins', 'Apps', 'AppV', 'Catalogs', 'Config', 'Groups', 
	'Hosting', 'Licensing', 'Logging', 'Policies', 'StoreFront', 'Zones')]
	[parameter(Mandatory=$False)] 
	[string]$Section="All",
	
	[parameter(Mandatory=$False)] 
	[Alias("ADT")]
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[Switch]$CSV=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CA")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CE")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CF")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CPh")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(Mandatory=$False)] 
	[string]$SmtpServer="",

	[parameter(Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(Mandatory=$False)] 
	[Switch]$UseSSL=$False,

	[parameter(Mandatory=$False)] 
	[string]$From="",

	[parameter(Mandatory=$False)] 
	[string]$To=""

	)
#endregion

#region script change log	
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Original script created on October 20, 2013
#started updating for Citrix Cloud on August 28, 2020

# This script is based on the CVAD V3.00 doc script

#Version 1.11 25-Jan-2021
#	Added error checking in Function Check-NeededPSSnapins (Requested by Guy Leech)
#	Updated the error message when Get-XDAuthentication fails
#	Updated the error messages in Function ProcessScriptSetup
#	Updated the help text
#	Updated the ReadMe file
#
#Version 1.10 5-Dec-2020
#	Added CustomerID to function ShowScriptInfo and ProcessScriptEnd
#	Added the missing ReadMe file link to the warning message about the missing Citrix.GroupPolicy.Commands file
#	Added to Hosting Connection section:
#	(Thanks to fellow CTPs Neil Spellings, Kees Baggerman, and Trond Eirik Haavarstein for getting this info for me)
#		Amazon EC2    
#		Google Cloud Platform
#		Microsoft Azure
#		Nutanix AHV
#		Remote PC Wake on LAN
#	Added variables, counters, and text for the following Administrative Roles
#		Cloud Administrator: [int]$Script:TotalCloudAdmins
#		Full Monitor Administrator: [int]$Script:TotalFullMonitorAdmins
#		Probe Agent Administrator: [int]$Script:TotalProbeAdmins
#		Session Administrator: [int]$Script:TotalSessionAdmins
#	Correct the invalid variable name in the ScriptInfo output file for WordFilename
#	Fixed alignment in the Text output for the ScriptInfo output file
#	Fixed bug reported by David Prows in the Hosting section. First, check to see if the hosting connection's 
#		AdditionalStorage.StorageLocations is valid
#	For all calls to Get-AdminAdministrator, remove the -SortBy Name. Sorting by Name is the default behavior.
#	For MCS Machine Catalogs:
#		Check that the catalog's ProvisioningSchemeId is not $Null before retrieving the Provision Scheme's machine data
#		Check that $MachineData is not $Null before checking for HostingUnitName
#	For the Hosting section, for High Availability Servers and Power Actions, handle empty arrays
#	In Function GetAdmins, for Hosting Connections, handle the error "The property 'ScopeId' cannot be found on this object. Verify that the property exists."
#		Also, added some white space to make the function easier for me to read
#	In Function OutputAdminsForDetails, add "No Admins found" to replace blank tables and text output
#	In Function OutputDeliveryGroupCatalogs, handle the case where a Delivery Group has no Machine Catalog(s) assigned
#	In Function OutputMachineDetails, when using Test-NetConnection, add Resolve-DnsName first to see if the machine name is resolvable. 
#		This prevents every call to Test-NetConnection from failing with "<MachineName> was not found in DNS". Add error message:
#		<MachineName> was not found in DNS. VDA Registry Key data cannot be gathered.
#		Otherwise, every machine was reported as offline, which may not be true.
#	In Function OutputPerZoneView, add "There are no zone members for Zone <ZoneName>" to replace blank tables and text output
#	In Function OutputSummaryPage, add text for Cloud, Full Monitor, Probe Agent, and Session Administrators
#	KNOWN ISSUE: 
#		A few users report they get the following errors even in a new elevated PowerShell session:
#			"Import-Module : A drive with the name 'XDHyp' already exists."
#			"ProcessScriptSetup : Unable to import the Citrix.Host.Commands module. Script cannot continue."
#		Even if the script runs for these users, the Hosting section contains no usable data.
#			Hosting
#				Unable to retrieve Hosting Units
#			
#				Unable to retrieve Hosting Connections
#			 
#			My-Hosting-ConnectioName
#				Connection Name		My-Hosting-ConnectioName
#				Type				Hypervisor Type could not be determined:
#				Address	   
#				State				Disabled
#				Username	   
#				Scopes	   
#				Maintenance Mode	Off
#				Zone	   
#				Storage resource name	   
#
#			Advanced
#				Connection Name		My-Hosting-ConnectioName
#				Unable to retrieve Hosting Connections
#		I have been unable to determine the root cause, a resolution, or a workaround for this issue
#	Provide full details in the error message if the Citrix.Common.GroupPolicy snapin is missing (Thanks to Guy Leech for the suggestion)
#	Removed all references to $Script:CCParams1 and @CCParams1
#	Removed all references to AdminAddress
#	Removed Controllers from the ScriptInfo output file and the Section parameter switch statement
#	Removed Licensing from the Summary Page because there are no individual license files in Citrix Cloud
#	Reordered the parameters in an order recommended by Guy Leech
#	Updated the ReadMe file
#
#Version 1.00 21-Sep-2020
#
#	Add a switch statement for the machine/desktop/server Power State
#	Add a test to make sure the script runsning from an elevated PowerShell session
#	Add additional Citrix Cloud administrator permissions:
#		App-V
#			Add App-V applications
#			Remove App-V applications
#		Cloud
#			Read Storefront Configuration 
#			Update Storefront Configuration
#		Director
#			View Filters page Application Instances only
#			View Filters page Connections only
#			View Filters page Machines only
#			View Filters page Sessions only
#		Other Permissions
#			Customer Update Site Configuration
#			Update Site Configuration
#			Read database status information
#			Export Broker Configuration
#		UPM
#			Add UPM Broker Machine Configuration
#			Read UPM Broker Machine Configuration
#			Delete UPM Broker Machine Configuration
#	Added a ValidateSet to the Sections parameter. You can use -Section, press tab, and tab through all the section options. (Credit to Guy Leech)
#	Added a -ProfileName parameter for use by Get-XDAuthentication 
#		For more information, see the Authentication section in the ReadMe file https://carlwebster.sharefile.com/d-sb4e144f9ecc48e7b
#		Thanks to David Prows and Devan Tilly for documenting this process for me to use
#	Added testing to see if the computer running the script is in a Domain or Workgroup
#		If not in a domain, and VDARegistryKeys is set, set it to $False
#		If not in a domain, set NoADPolicies to $True
#	Added testing to see if the Remote SDK is installed (Thanks to Martin Zugec)
#	Change all Write-Verbose $(Get-Date) to add -Format G to put the dates in the user's locale (Thanks to Guy Leech)
#	Change checking some String variables from just $Null to [String]::IsNullOrEmpty
#		Some cmdlet's string properties are sometimes Null and sometimes an empty string
#	Change checking the way a machine is online or offline
#	Change some cmdlets to sort on the left of the pipeline using the cmdlet's -SortBy option
#	Changed testing for existing PSDrives from Get-PSDrive to Test-Path
#	Comment out the code for Session Recording until Citrix enables the cmdlet in the Cloud
#	Fixed administrators output by restricting to UserIdentityType -ne "Sid" -and (-not ([String]::IsNullOrEmpty($_.Name)))
#	Fixed an issue with "Connections meeting any of the following (Access Gateway) filters"
#		If you selected HTML and any other output format in the same run, only the HTML output had any Access Gateway data
#	Fixed issue with PowerShell 5.1.x and empty Hashtables for Ian Brighton's Word Table functions
#		PoSH 3, 4, and 5.0 had no problem with an empty hashtable and would create a blank Word table with only column headings
#		For many tables, before passing the hashtable to Ian's function, test if the hashtable is empty
#		If it is, create a dummy row of data for the hashtable
#		For example, a RemotePC catalog based on OU that contains no machines, or Applications with no administrators
#		Instead of having a missing table, the table will now have a row that says "None found"
#	Fixed issue with the array used for Appendix A and the CSV file when selecting multiple output formats
#		If HTML and Text and MSWord were selected, Appendix A and the CSV file contained three duplicate entries
#		Changed from using only one array to three. Changed from using $Script:ALLVDARegistryItems to
#			$Script:WordALLVDARegistryItems
#			$Script:TextALLVDARegistryItems
#			$Script:HTMLALLVDARegistryItems
#	Fixed output issues with Power Management settings
#	Fixed several more array out of bounds issues when accessing element 0 when the array was empty
#	Fixed Zone output by excluding "Initial Zone" and Zones with a name of "00000000-0000-0000-0000-000000000000"
#	For Appendix A, Text output, change the column heading "DDC Name" to "Computer Name" to match the HTML and MSWord/PDF output
#	For Zones, add Citrix Cloud Connectors
#	For Zones, add MemType into the Sort-Object for Site View and Zone View
#	For Zones, for MSWord and PDF output, change the column widths
#	Removed a lot of the Licensing code as there are no Get-Lic* cmdlets in the Remote SDK
#	Removed all code for BrokerRegistryKeys
#		Removed AppendixB
#	Removed all code for CEIP
#	Removed all code for Hardware
#	Removed all code for Microsoft hotfixes, Citrix hotfixes, and installed Windows Features and Roles
#		Removed AppendixC, AppendixD, and AppendixE
#	Removed all code for the Controllers section
#		Also removed code for checking if StoreFront is installed on the Delivery Controller
#	Removed all code for the three Citrix SQL databases
#	Removed some code for Configuration Logging
#		We don't have access to Preferences or the Logging database details
#		We do have access to the configuration logging high-level action details
#		Remove checking for product code MPS since that doesn't exist in the cloud licenses
#	Removed the help text and parameter for $AdminAddress
#		Hardcode $AdminAddress to LocalHost
#	Removed licensing data that didn't make sense for Citrix Cloud and added some data that did make sense
#	Since the RegistrationState property is an enum, add .ToString() to the machine/desktop/server variable so HTML output is correct
#	When getting the Master VM for an MCS based machine catalog, also check for images ending in .template for Nutanix
#	When getting the provisioning scheme data for a machine, only process machines with a provisioning type of MCS
#		There is no provisioning scheme data for manually or PVS provisioned machines
#	When VDARegistryKeys is used, now test the RemoteRegistry service for its status
#		If the service is not running, add that information into the various *VDARegistryItems arrays
#	Where appropriate, changed all "CVAD" to "CC"
#endregion

#region initial variable testing and setup
Set-StrictMode -Version Latest

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

Write-Verbose "$(Get-Date -Format G): Testing for elevation"
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

If($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator ))
{
	Write-Host "This is an elevated PowerShell session" -ForegroundColor White
}
Else
{
	Write-Error "
	`n
	This is NOT an elevated PowerShell session.
	`n
	Script will exit.
	`n"
	Exit
}

If($Null -eq $HTML)
{
	If($Text -or $MSWord -or $PDF)
	{
		$HTML = $False
	}
	Else
	{
		$HTML = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$HTML = $True
}

Write-Verbose "$(Get-Date -Format G): Testing output parameters"

If($MSWord)
{
	Write-Verbose "$(Get-Date -Format G): MSWord is set"
}
If($PDF)
{
	Write-Verbose "$(Get-Date -Format G): PDF is set"
}
If($Text)
{
	Write-Verbose "$(Get-Date -Format G): Text is set"
}
If($HTML)
{
	Write-Verbose "$(Get-Date -Format G): HTML is set"
}

If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer but did not include a From or To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a To email address but did not include a From email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($To) -and ![String]::IsNullOrEmpty($From))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a From email address but did not include a To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified From and To email addresses but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a From email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a To email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}

#If the MaxDetails parameter is used, set a bunch of stuff true and some stuff false
If($MaxDetails)
{
	$Administrators		= $True
	$Applications		= $True
	$DeliveryGroups		= $True
	$Hosting			= $True
	$Logging			= $True
	$MachineCatalogs	= $True
	$Policies			= $True
	$StoreFront			= $True
	$VDARegistryKeys	= $True
	
	$NoPolicies			= $False
	$Section			= "All"
}

If($NoPolicies)
{
	$Policies = $False
}

If($NoPolicies -and $Section -eq "Policies")
{
	#conflict
	$ErrorActionPreference = $SaveEAPreference
	Write-Error -Message "
	`n`n
	`t`t
	You specified conflicting parameters.
	`n`n
	`t`t
	You specified the Policies section but also selected NoPolicies.
	`n`n
	`t`t
	Please change one of these options and rerun the script.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}

$ValidSection = $False
Switch ($Section)
{
	"Admins"		{$ValidSection = $True; Break}
	"Apps"			{$ValidSection = $True; Break}
	"AppV"			{$ValidSection = $True; Break}
	"Catalogs"		{$ValidSection = $True; Break}
	"Config"		{$ValidSection = $True; Break}
	"Groups"		{$ValidSection = $True; Break}
	"Hosting"		{$ValidSection = $True; Break}
	"Licensing"		{$ValidSection = $True; Break}
	"Logging"		{$ValidSection = $True; $Logging = $True; Break}	#force $logging true if the config logging section is specified
	"Policies"		{$ValidSection = $True; $Policies = $True; Break} #force $policies true if the policies section is specified
	"StoreFront"	{$ValidSection = $True; Break}
	"Zones"			{$ValidSection = $True; Break}
	"All"			{$ValidSection = $True; Break}
}

If($ValidSection -eq $False)
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Error -Message "
	`n`n
	`t`t
	The Section parameter specified, $Section, is an invalid Section option.
	`n`n
	`t`t
	Valid options are:
	
	`tAdmins
	`tApps
	`tAppV
	`tCatalogs (Machine Catalogs)
	`tConfig (Configuration)
	`tGroups (Delivery Groups)
	`tHosting
	`tLicensing
	`tLogging (Configuration Logging)
	`tPolicies
	`tStoreFront
	`tZones
	`tAll
	
	`t`t
	Script cannot continue.
	`n`n
	"
	Exit
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date -Format G): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date -Format G): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
#Do not indent the following write-error lines. Doing so will mess up the console formatting of the error message.
			Write-Error "
			`n`n
	Folder $Folder is a file, not a folder.
			`n`n
	Script cannot continue.
			`n`n"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "
		`n`n
	Folder $Folder does not exist.
		`n`n
	Script cannot continue.
		`n`n
		"
		Exit
	}
}

If($Folder -eq "")
{
	$Script:pwdpath = $pwd.Path
}
Else
{
	$Script:pwdpath = $Folder
}

If($Script:pwdpath.EndsWith("\"))
{
	#remove the trailing \
	$Script:pwdpath = $Script:pwdpath.SubString(0, ($Script:pwdpath.Length - 1))
}

If($Log) 
{
	#start transcript logging
	$Script:LogPath = "$Script:pwdpath\CCDocScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	
	try 
	{
		Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
		Write-Verbose "$(Get-Date -Format G): Transcript/log started at $Script:LogPath"
		$Script:StartLog = $true
	} 
	catch 
	{
		Write-Verbose "$(Get-Date -Format G): Transcript/log failed at $Script:LogPath"
		$Script:StartLog = $false
	}
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$Script:pwdpath\CCInventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

[int]$Script:TotalComputerPolicies       = 0
[int]$Script:TotalUserPolicies           = 0
[int]$Script:TotalSitePolicies           = 0
[int]$Script:TotalADPolicies             = 0
[int]$Script:TotalADPoliciesNotProcessed = 0
[int]$Script:TotalPolicies               = 0

[int]$Script:TotalServerOSCatalogs       = 0
[int]$Script:TotalDesktopOSCatalogs      = 0
[int]$Script:TotalRemotePCCatalogs       = 0

[int]$Script:TotalApplicationGroups      = 0
[int]$Script:TotalDesktopGroups          = 0
[int]$Script:TotalAppsAndDesktopGroups   = 0
[int]$Script:TotalPublishedApplications  = 0
[int]$Script:TotalAppvApplications       = 0

[int]$Script:TotalCloudAdmins            = 0
[int]$Script:TotalDeliveryGroupAdmins    = 0
[int]$Script:TotalFullAdmins             = 0
[int]$Script:TotalFullMonitorAdmins      = 0
[int]$Script:TotalHelpDeskAdmins         = 0
[int]$Script:TotalHostAdmins             = 0
[int]$Script:TotalMachineCatalogAdmins   = 0
[int]$Script:TotalProbeAdmins            = 0
[int]$Script:TotalReadOnlyAdmins         = 0
[int]$Script:TotalSessionAdmins          = 0
[int]$Script:TotalCustomAdmins           = 0

[int]$Script:TotalHostingConnections     = 0

[int]$Script:TotalStoreFrontServers      = 0

[int]$Script:TotalZones                  = 0
[string]$Script:RunningOS                = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption
	
$Script:VDARegistryItems                 = New-Object System.Collections.ArrayList
$Script:WordALLVDARegistryItems          = New-Object System.Collections.ArrayList
$Script:TextALLVDARegistryItems          = New-Object System.Collections.ArrayList
$Script:HTMLALLVDARegistryItems          = New-Object System.Collections.ArrayList

#Final test for variable setting
#is the computer running the script a member of a domain?
#if not, set VDARegistryKeys to False and NoADPolicies to True
#a workgroup computer cannot access the registry on domain-joined computers nor access SYSVOL
#
#http://powershell-guru.com/powershell-tip-63-check-if-a-computer-is-member-of-a-domain-or-workgroup/
If(!((Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain))
{
	#member of workgroup, not a domain
	Write-Host "
	$env:computername is not part of the domain" -ForegroundColor White
	
	If($VDARegistryKeys -eq $True)
	{
		Write-Host "
	Setting VDARegistyryKeys to False" -ForegroundColor White
		$VDARegistryKeys = $False
	}

	If($NoADPolicies -eq $False)
	{
		Write-Host "
	Setting NoADPolicies to True
	" -ForegroundColor White
		$NoADPolicies = $True
	}
}

If($VDARegistryKeys)
{	
	#Force $MachineCatalogs to True
	Write-Verbose "$(Get-Date -Format G): VDARegistryKeys Switch is set True. Forcing MachineCatalogs to True."
	$MachineCatalogs = $True
}

If($ProfileName -eq "")
{
	Get-XDAuthentication -EA 0 *>$Null
	$CCCreds = (Get-XDCredentials -ProfileName default).Credentials
}
Else
{
	Get-XDAuthentication -ProfileName $ProfileName -EA 0 *>$Null
	$CCCreds = (Get-XDCredentials -ProfileName $ProfileName).Credentials
}

$Script:CustomerID = $CCCreds.CustomerID

If(!$?)
{
	Write-Error "
	`n`n
	Get-XDAuthentication failed.
	`n`n
	For more information, see the Authentication section in the ReadMe file at 
	`n`n
	https://carlwebster.sharefile.com/d-sb4e144f9ecc48e7b
	`n`n
	This script is designed for Citrix Cloud/Citrix Virtual Apps and Desktops Service.
	`n`n
	If you are running XA/XD 7.0 through 7.7, please use: 
	https://carlwebster.com/downloads/download-info/xenappxendesktop-7-x-documentation-script/
	`n`n
	If you are running XA/XD 7.8 through CVAD 2006, please use:
	https://carlwebster.com/downloads/download-info/xenappxendesktop-7-8/
	`n`n
	If you are running CVAD 2006 and later, please use:
	https://carlwebster.com/downloads/download-info/citrix-virtual-apps-and-desktops-v3-script/
	`n`n
	Script cannot continue.
	`n`n
	"
	Exit
}

#endregion

#region initialize variables for Word, HTML, and text

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Verbose "$(Get-Date -Format G): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight  = 2
	[int]$wdMove                  = 0
	[int]$wdSeekMainDocument      = 0
	[int]$wdSeekPrimaryFooter     = 4
	[int]$wdStory                 = 6
	[int]$wdColorBlack            = 0
	[int]$wdColorGray05           = 15987699 
	[int]$wdColorGray15           = 14277081
	[int]$wdColorRed              = 255
	[int]$wdColorWhite            = 16777215
	[int]$wdColorYellow           = 65535
	[int]$wdWord2007              = 12
	[int]$wdWord2010              = 14
	[int]$wdWord2013              = 15
	[int]$wdWord2016              = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF             = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	#[int]$wdAlignParagraphLeft   = 0
	#[int]$wdAlignParagraphCenter = 1
	#[int]$wdAlignParagraphRight  = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	#[int]$wdCellAlignVerticalTop    = 0
	#[int]$wdCellAlignVerticalCenter = 1
	#[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed   = 0
	[int]$wdAutoFitContent = 1
	#[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone         = 0
	[int]$wdAdjustProportional = 1
	#[int]$wdAdjustFirstColumn = 2
	#[int]$wdAdjustSameWidth   = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops  = 0 * $PointsPerTabStop
	#[int]$Indent1TabStops = 1 * $PointsPerTabStop
	#[int]$Indent2TabStops = 2 * $PointsPerTabStop
	#[int]$Indent3TabStops = 3 * $PointsPerTabStop
	#[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1         = -2
	[int]$wdStyleHeading2         = -3
	[int]$wdStyleHeading3         = -4
	[int]$wdStyleHeading4         = -5
	[int]$wdStyleNoSpacing        = -158
	[int]$wdTableGrid             = -155
	[int]$wdTableLightListAccent3 = -206

	[int]$wdLineStyleNone       = 0
	[int]$wdLineStyleSingle     = 1
	[int]$wdHeadingFormatTrue   = -1
	#[int]$wdHeadingFormatFalse = 0 
	
	[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption
}
Else
{
	$Script:CoName = ""
}

If($HTML)
{
    $global:htmlredmask       = "#FF0000" 4>$Null
    $global:htmlcyanmask      = "#00FFFF" 4>$Null
    $global:htmlbluemask      = "#0000FF" 4>$Null
    $global:htmldarkbluemask  = "#0000A0" 4>$Null
    $global:htmllightbluemask = "#ADD8E6" 4>$Null
    $global:htmlpurplemask    = "#800080" 4>$Null
    $global:htmlyellowmask    = "#FFFF00" 4>$Null
    $global:htmllimemask      = "#00FF00" 4>$Null
    $global:htmlmagentamask   = "#FF00FF" 4>$Null
    $global:htmlwhitemask     = "#FFFFFF" 4>$Null
    $global:htmlsilvermask    = "#C0C0C0" 4>$Null
    $global:htmlgraymask      = "#808080" 4>$Null
    $global:htmlblackmask     = "#000000" 4>$Null
    $global:htmlorangemask    = "#FFA500" 4>$Null
    $global:htmlmaroonmask    = "#800000" 4>$Null
    $global:htmlgreenmask     = "#008000" 4>$Null
    $global:htmlolivemask     = "#808000" 4>$Null

    $global:htmlbold        = 1 4>$Null
    $global:htmlitalics     = 2 4>$Null
    $global:htmlred         = 4 4>$Null
    $global:htmlcyan        = 8 4>$Null
    $global:htmlblue        = 16 4>$Null
    $global:htmldarkblue    = 32 4>$Null
    $global:htmllightblue   = 64 4>$Null
    $global:htmlpurple      = 128 4>$Null
    $global:htmlyellow      = 256 4>$Null
    $global:htmllime        = 512 4>$Null
    $global:htmlmagenta     = 1024 4>$Null
    $global:htmlwhite       = 2048 4>$Null
    $global:htmlsilver      = 4096 4>$Null
    $global:htmlgray        = 8192 4>$Null
    $global:htmlolive       = 16384 4>$Null
    $global:htmlorange      = 32768 4>$Null
    $global:htmlmaroon      = 65536 4>$Null
    $global:htmlgreen       = 131072 4>$Null
	$global:htmlblack       = 262144 4>$Null

	$global:htmlsb          = ( $htmlsilver -bor $htmlBold ) ## point optimization

	$global:htmlColor = 
	@{
		$htmlred       = $htmlredmask
		$htmlcyan      = $htmlcyanmask
		$htmlblue      = $htmlbluemask
		$htmldarkblue  = $htmldarkbluemask
		$htmllightblue = $htmllightbluemask
		$htmlpurple    = $htmlpurplemask
		$htmlyellow    = $htmlyellowmask
		$htmllime      = $htmllimemask
		$htmlmagenta   = $htmlmagentamask
		$htmlwhite     = $htmlwhitemask
		$htmlsilver    = $htmlsilvermask
		$htmlgray      = $htmlgraymask
		$htmlolive     = $htmlolivemask
		$htmlorange    = $htmlorangemask
		$htmlmaroon    = $htmlmaroonmask
		$htmlgreen     = $htmlgreenmask
		$htmlblack     = $htmlblackmask
	}
}

If($TEXT)
{
	[System.Text.StringBuilder] $global:Output = New-Object System.Text.StringBuilder( 16384 )
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. Smith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 10-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
			'zh-'	{ '自动目录 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_}	{$CultureCode = "ca-"}
		{$ChineseArray -contains $_}	{$CultureCode = "zh-"}
		{$DanishArray -contains $_}		{$CultureCode = "da-"}
		{$DutchArray -contains $_}		{$CultureCode = "nl-"}
		{$EnglishArray -contains $_}	{$CultureCode = "en-"}
		{$FinnishArray -contains $_}	{$CultureCode = "fi-"}
		{$FrenchArray -contains $_}		{$CultureCode = "fr-"}
		{$GermanArray -contains $_}		{$CultureCode = "de-"}
		{$NorwegianArray -contains $_}	{$CultureCode = "nb-"}
		{$PortugueseArray -contains $_}	{$CultureCode = "pt-"}
		{$SpanishArray -contains $_}	{$CultureCode = "es-"}
		{$SwedishArray -contains $_}	{$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		'zh-'	{
				If($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
					'离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
					'切片(深色)', '丝状', '网格', '镶边', '信号灯',
					'运动型')
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		
		If(($MSWord -eq $False) -and ($PDF -eq $True))
		{
			Write-Host "`n`n`t`tThis script uses Microsoft Word's SaveAs PDF function, please install Microsoft Word`n`n" -ForegroundColor White
			Exit
		}
		Else
		{
			Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n" -ForegroundColor White
			Exit
		}
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword runsning in our session
	[bool]$wordrunning = $null –ne ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID})
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n" -ForegroundColor White
		Exit
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function Set-DocumentProperty {
    <#
	.SYNOPSIS
	Function to set the Title Page document properties in MS Word
	.DESCRIPTION
	Long description
	.PARAMETER Document
	Current Document Object
	.PARAMETER DocProperty
	Parameter description
	.PARAMETER Value
	Parameter description
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value 'MyTitle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value 'MyCompany'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value 'Jim Moyle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value 'MySubjectTitle'
	.NOTES
	Function Created by Jim Moyle June 2017
	Twitter : @JimMoyle
	#>
    param (
        [object]$Document,
        [String]$DocProperty,
        [string]$Value
    )
    try {
        $binding = "System.Reflection.BindingFlags" -as [type]
        $builtInProperties = $Document.BuiltInDocumentProperties
        $property = [System.__ComObject].invokemember("item", $binding::GetProperty, $null, $BuiltinProperties, $DocProperty)
        [System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
    }
    catch {
        Write-Warning "Failed to set $DocProperty to $Value"
    }
}

Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Verbose "$(Get-Date -Format G): Setting up Word"
    
	If(!$AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName).pdf"
		}
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
		}
	}

	# Setup word for output
	Write-Verbose "$(Get-Date -Format G): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null

#Do not indent the following write-error lines. Doing so will mess up the console formatting of the error message.
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created. You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	The Word object could not be created. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n"
		Exit
	}

	Write-Verbose "$(Get-Date -Format G): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	Unable to determine the Word language value. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n
		"
		AbortScript
	}
	Write-Verbose "$(Get-Date -Format G): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	Microsoft Word 2007 is no longer supported.`n`n`t`tScript will end.
		`n`n
		"
		AbortScript
	}
	ElseIf($Script:WordVersion -eq 0)
	{
		Write-Error "
		`n`n
	The Word Version is 0. You should run a full online repair of your Office installation.
		`n`n
	Script cannot Continue.
		`n`n
		"
		Exit
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	You are running an untested or unsupported version of Microsoft Word.
		`n`n
	Script will end.
		`n`n
	Please send info on your version of Word to webster@carlwebster.com
		`n`n
		"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($CompanyName))
	{
		Write-Verbose "$(Get-Date -Format G): Company name is blank. Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Host "
		Company Name is blank so Cover Page will not show a Company Name.
		Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value.
		You may want to use the -CompanyName parameter if you need a Company Name on the cover page.
			" -ForegroundColor White
			$Script:CoName = $TmpName
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date -Format G): Updated company name to $($Script:CoName)"
		}
	}
	Else
	{
		$Script:CoName = $CompanyName
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date -Format G): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}

			'zh-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "边线型"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date -Format G): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date -Format G): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date -Format G): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date -Format G): Culture code $($Script:WordCultureCode)"
		Write-Error "
		`n`n
	For $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.
		`n`n
	Script cannot Continue.
		`n`n
		"
		AbortScript
	}

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date -Format G): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where-Object{$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date -Format G): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach-Object {
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date -Format G): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Host "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist." -ForegroundColor White
		Write-Host "This report will not have a Cover Page." -ForegroundColor White
	}

	Write-Verbose "$(Get-Date -Format G): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date -Format G): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	An empty Word document could not be created. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date -Format G): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	An unknown error happened selecting the entire Word document for default formatting options.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 =.50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date -Format G): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date -Format G): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date -Format G): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date -Format G): "
			Write-Host "Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved." -ForegroundColor White
			Write-Host "This report will not have a Table of Contents." -ForegroundColor White
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Host "Table of Contents are not installed." -ForegroundColor White
		Write-Host "Table of Contents are not installed so this report will not have a Table of Contents." -ForegroundColor White
	}

	#set the footer
	Write-Verbose "$(Get-Date -Format G): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date -Format G): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date -Format G): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date -Format G): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#updated 8-Jun-2017 with additional cover page fields
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date -Format G): Set Cover Page Properties"
			#8-Jun-2017 put these 4 items in alpha order
            Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value $UserName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value $Script:CoName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value $SubjectTitle
            Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value $Script:title

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where-Object{$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "Abstract"}
			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyAddress"}
			#set the text
			[string]$abstract = $CompanyAddress
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyEmail"}
			#set the text
			[string]$abstract = $CompanyEmail
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyFax"}
			#set the text
			[string]$abstract = $CompanyFax
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyPhone"}
			#set the text
			[string]$abstract = $CompanyPhone
			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date -Format G): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue2
{
	[CmdletBinding()]
	Param([string]$path, [string]$name, [string]$ComputerName)
	If($ComputerName -eq $env:computername)
	{
		$key = Get-Item -LiteralPath $path -EA 0
		If($key)
		{
			Return $key.GetValue($name, $Null)
		}
		Else
		{
			Return $Null
		}
	}
	Else
	{
		#path needed here is different for remote registry access
		$path = $path.SubString(6)
		$path2 = $path.Replace('\','\\')
        Try
        {
            ## GRL throws an error if can't open
		    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
        }
        Catch
        {
            Write-Debug -Message "Failed to open HKLM on $ComputerName"
            $Reg = $null
        }
        if( $Reg -and ($RegKey = $Reg.OpenSubKey($path2) ) )
		{
			$Results = $RegKey.GetValue($name)

			If($Null -ne $Results)
			{
				Return $Results
			}
			Else
			{
				Return $Null
			}
		}
		Else
		{
			Return $Null
		}
	}
}
#endregion

#region Word, text, and HTML line output functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexch on Twitter
#https://essential.exchange/blog
#for creating the formatted text report
#created March 2011
#updated March 2014
# updated March 2019 to use StringBuilder (about 100 times more efficient than simple strings)
{
	Param
	(
		[Int]    $tabs = 0, 
		[String] $name = '', 
		[String] $value = '', 
		[String] $newline = [System.Environment]::NewLine, 
		[Switch] $nonewline
	)

	while( $tabs -gt 0 )
	{
		$null = $global:Output.Append( "`t" )
		$tabs--
	}

	If( $nonewline )
	{
		$null = $global:Output.Append( $name + $value )
	}
	Else
	{
		$null = $global:Output.AppendLine( $name + $value )
	}
}
	
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 " "

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $Null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlBold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlBold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlBold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML. They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, 
	HTML supports headers h1-h6 and h1-h4 are more commonly used. Unlike word, H1 will not 
	give you a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

$crlf = [System.Environment]::NewLine

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
#errors with $HTMLStyle fixed 7-Dec-2017 by Webster
{
	Param
	(
		[Int]    $style    = 0, 
		[Int]    $tabs     = 0, 
		[String] $name     = '', 
		[String] $value    = '', 
		[String] $fontName = $null,
		[Int]    $fontSize = 1,
		[Int]    $options  = $htmlblack
	)

	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 1024 )

	If( [String]::IsNullOrEmpty( $name ) )	
	{
		$null = $sb.Append( '<p></p>' )
	}
	Else
	{
		[Bool] $ital = $options -band $htmlitalics
		[Bool] $bold = $options -band $htmlBold

		If( $ital ) { $null = $sb.Append( '<i>' ) }
		If( $bold ) { $null = $sb.Append( '<b>' ) } 

		Switch( $style )
		{
			1 { $HTMLOpen = '<h1>'; $HTMLClose = '</h1>'; Break }
			2 { $HTMLOpen = '<h2>'; $HTMLClose = '</h2>'; Break }
			3 { $HTMLOpen = '<h3>'; $HTMLClose = '</h3>'; Break }
			4 { $HTMLOpen = '<h4>'; $HTMLClose = '</h4>'; Break }
			Default { $HTMLOpen = ''; $HTMLClose = ''; Break }
		}

		$null = $sb.Append( $HTMLOpen )

		$null = $sb.Append( ( '&nbsp;&nbsp;&nbsp;&nbsp;' * $tabs ) + $name + $value )

		If( $HTMLClose -eq '' ) { $null = $sb.Append( '<br>' )     }
		Else                    { $null = $sb.Append( $HTMLClose ) }

		If( $ital ) { $null = $sb.Append( '</i>' ) }
		If( $bold ) { $null = $sb.Append( '</b>' ) } 

		If( $HTMLClose -eq '' ) { $null = $sb.Append( '<br />' ) }
	}
	$null = $sb.AppendLine( '' )

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $sb.ToString() 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable function
# Created by Ken Avram
# modified by Jake Rutski
# re-implemented by Michael B. Smith and made the documentation match reality.
#***********************************************************************************************************
Function AddHTMLTable
{
	Param
	(
		[String]   $fontName  = 'Calibri',
		[Int]      $fontSize  = 2,
		[Int]      $colCount  = 0,
		[Int]      $rowCount  = 0,
		[Object[]] $rowInfo   = $null,
		[Object[]] $fixedInfo = $null
	)

	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 8192 )

	If( $rowInfo -and $rowInfo.Length -lt $rowCount )
	{
		$rowCount = $rowInfo.Length
	}

	For( $rowCountIndex = 0; $rowCountIndex -lt $rowCount; $rowCountIndex++ )
	{
		$null = $sb.AppendLine( '<tr>' )

		## reset
		$row = $rowInfo[ $rowCountIndex ]

		$subRow = $row
		If( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
		{
			$subRow = $subRow[ 0 ]
		}

		$subRowLength = $subRow.Length
		For( $columnIndex = 0; $columnIndex -lt $colCount; $columnIndex += 2 )
		{
			$item = If( $columnIndex -lt $subRowLength ) { $subRow[ $columnIndex ] } Else { 0 }

			$text   = If( $item ) { $item.ToString() } Else { '' }
			$format = If( ( $columnIndex + 1 ) -lt $subRowLength ) { $subRow[ $columnIndex + 1 ] } Else { 0 }
			## item, text, and format ALWAYS have values, even if empty values
			$color  = $global:htmlColor[ $format -band 0xffffc ]
			[Bool] $bold = $format -band $htmlBold
			[Bool] $ital = $format -band $htmlitalics

			If( $null -eq $fixedInfo -or $fixedInfo.Length -eq 0 )
			{
				$null = $sb.Append( "<td style=""background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
			}
			Else
			{
				$null = $sb.Append( "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
			}

			If( $bold ) { $null = $sb.Append( '<b>' ) }
			If( $ital ) { $null = $sb.Append( '<i>' ) }

			If( $text -eq ' ' -or $text.length -eq 0)
			{
				##$htmlbody += '&nbsp;&nbsp;&nbsp;'
				$null = $sb.Append( '&nbsp;&nbsp;&nbsp;' )
			}
			Else
			{
				For ($inx = 0; $inx -lt $text.length; $inx++ )
				{
					If( $text[ $inx ] -eq ' ' )
					{
						$null = $sb.Append( '&nbsp;' )
					}
					Else
					{
						Break
					}
				}
				$null = $sb.Append( $text )
			}

			If( $bold ) { $null = $sb.Append( '</b>' ) }
			If( $ital ) { $null = $sb.Append( '</i>' ) }

			$null = $sb.AppendLine( '</font></td>' )
		}

		$null = $sb.AppendLine( '</tr>' )
	}

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $sb.ToString() 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
# reworked by Michael B. Smith
#***********************************************************************************************************

<#
.Synopsis
	Format table for a HTML output document.
.DESCRIPTION
	This function formats a table for HTML from multiple arrays of strings.
.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border = '0'). Otherwise the table will be generated
	with a border (border = '1').
.PARAMETER noHeadCols
	This parameter should be used when generating tables which do not have a separate array containing column headers
	(columnArray is not specified). Set this parameter equal to the number of columns in the table.
.PARAMETER rowArray
	This parameter contains the row data array for the table.
.PARAMETER columnArray
	This parameter contains column header data for the table.
.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $fixedWidth = @("100px","110px","120px","130px","140px")
.PARAMETER tableHeader
	A string containing the header for the table (printed at the top of the table, left justified). The
	default is a blank string.
.PARAMETER tableWidth
	The width of the table in pixels, or 'auto'. The default is 'auto'.
.PARAMETER fontName
	The name of the font to use in the table. The default is 'Calibri'.
.PARAMETER fontSize
	The size of the font to use in the table. The default is 2. Note that this is the HTML size, not the pixel size.

.USAGE
	FormatHTMLTable <Table Header> <Table Width> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file. All of the parameters are optional
	defaults are used if not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column. You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column. Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data. Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array. If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',$htmlsb,'Status',$htmlsb,'Startup Type',$htmlsb)

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics. For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below. As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",$htmlsb,$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',$htmlsb,$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',$htmlsb,$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',$htmlsb,$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',$htmlsb,$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',$htmlsb,$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',$htmlsb,$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',$htmlsb,$ComputerName,$htmlwhite))
	$rowdata += @(,('Filename1',$htmlsb,$Script:FileName1,$htmlwhite))
	$rowdata += @(,('OS Detected',$htmlsb,$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',$htmlsb,$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',$htmlsb,$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' paramater is mandatory to build the table, but it is not set as such in the function - if nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     

#>

Function FormatHTMLTable
{
	Param
	(
		[String]   $tableheader = '',
		[String]   $tablewidth  = 'auto',
		[String]   $fontName    = 'Calibri',
		[Int]      $fontSize    = 2,
		[Switch]   $noBorder    = $false,
		[Int]      $noHeadCols  = 1,
		[Object[]] $rowArray    = $null,
		[Object[]] $fixedWidth  = $null,
		[Object[]] $columnArray = $null
	)

	## FIXME - the help text for this function is wacky wrong - MBS
	## FIXME - Use StringBuilder - MBS - this only builds the table header - benefit relatively small
<#
	If( $SuperVerbose )
	{
		wv "FormatHTMLTable: fontname '$fontname', size $fontSize, tableheader '$tableheader'"
		wv "FormatHTMLTable: noborder $noborder, noheadcols $noheadcols"
		If( $rowarray -and $rowarray.count -gt 0 )
		{
			wv "FormatHTMLTable: rowarray has $( $rowarray.count ) elements"
		}
		Else
		{
			wv "FormatHTMLTable: rowarray is empty"
		}
		If( $columnarray -and $columnarray.count -gt 0 )
		{
			wv "FormatHTMLTable: columnarray has $( $columnarray.count ) elements"
		}
		Else
		{
			wv "FormatHTMLTable: columnarray is empty"
		}
		If( $fixedwidth -and $fixedwidth.count -gt 0 )
		{
			wv "FormatHTMLTable: fixedwidth has $( $fixedwidth.count ) elements"
		}
		Else
		{
			wv "FormatHTMLTable: fixedwidth is empty"
		}
	}
#>

	$HTMLBody = "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>" + $crlf

	If( $null -eq $columnArray -or $columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If( $null -ne $rowArray )
	{
		$NumRows = $rowArray.length + 1
	}
	Else
	{
		$NumRows = 1
	}

	If( $noBorder )
	{
		$HTMLBody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$HTMLBody += "<table border='1' width='" + $tablewidth + "'>"
	}
	$HTMLBody += $crlf

	If( $columnArray -and $columnArray.Length -gt 0 )
	{
		$HTMLBody += '<tr>' + $crlf

		For( $columnIndex = 0; $columnIndex -lt $NumCols; $columnindex += 2 )
		{
			$val = $columnArray[ $columnIndex + 1 ]
			$tmp = $global:htmlColor[ $val -band 0xffffc ]
			[Bool] $bold = $val -band $htmlBold
			[Bool] $ital = $val -band $htmlitalics

			If( $null -eq $fixedWidth -or $fixedWidth.Length -eq 0 )
			{
				$HTMLBody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$HTMLBody += "<td style=""width:$($fixedWidth[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If( $bold ) { $HTMLBody += '<b>' }
			If( $ital ) { $HTMLBody += '<i>' }

			$array = $columnArray[ $columnIndex ]
			If( $array )
			{
				If( $array -eq ' ' -or $array.Length -eq 0 )
				{
					$HTMLBody += '&nbsp;&nbsp;&nbsp;'
				}
				Else
				{
					For( $i = 0; $i -lt $array.Length; $i += 2 )
					{
						If( $array[ $i ] -eq ' ' )
						{
							$HTMLBody += '&nbsp;'
						}
						Else
						{
							break
						}
					}
					$HTMLBody += $array
				}
			}
			Else
			{
				$HTMLBody += '&nbsp;&nbsp;&nbsp;'
			}
			
			If( $bold ) { $HTMLBody += '</b>' }
			If( $ital ) { $HTMLBody += '</i>' }
		}

		$HTMLBody += '</font></td>'
		$HTMLBody += $crlf
	}

	$HTMLBody += '</tr>' + $crlf

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $HTMLBody 4>$Null 
	$HTMLBody = ''

	If( $rowArray )
	{
		AddHTMLTable -fontName $fontName -fontSize $fontSize `
		-colCount $numCols -rowCount $NumRows `
		-rowInfo $rowArray -fixedInfo $fixedWidth
		$rowArray = $null
		$HTMLBody = '</table>'
	}
	Else
	{
		$HTMLBody += '</table>'
	}

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is Returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -eq $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end ElseIf
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date -Format G): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date -Format G): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date -Format G): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date -Format G): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date -Format G): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date -Format G): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end Switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date -Format G): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date -Format G): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date -Format G): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the.Row and.Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells Returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells Returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) Returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$True, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$True, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$True, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $Null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $Null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] [int]$BackgroundColor = $Null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $True; }
					If($Italic) { $Cell.Range.Font.Italic = $True; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $True; }
				If($Italic) { $Cell.Range.Font.Italic = $True; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $True; }
					If($Italic) { $Cell.Range.Font.Italic = $True; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end Switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$True, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$True, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function CheckExcelPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Excel.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tFor the Delivery Groups Utilization option, this script directly outputs to Microsoft Excel, `n`t`tplease install Microsoft Excel or do not use the DeliveryGroupsUtilization (DGU) Switch`n`n" -ForegroundColor White
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if excel runsning in our session
	[bool]$excelrunning = $null –ne ((Get-Process 'Excel' -ea 0) | Where-Object {$_.SessionId -eq $SessionID})
	
	If($excelrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Excel before running this report.`n`n" -ForegroundColor White
		Exit
	}
}

Function Check-NeededPSSnapins
{
	Param([parameter(Mandatory = $True)][alias("Snapin")][string[]]$Snapins)

	#Function specifics
	$MissingSnapins = @()
	[bool]$FoundMissingSnapin = $False
	$LoadedSnapins = @()
	$RegisteredSnapins = @()

	#Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
	$loadedSnapins += Get-PSSnapin | ForEach-Object {$_.name}
	$registeredSnapins += Get-PSSnapin -Registered | ForEach-Object {$_.name}

	ForEach($Snapin in $Snapins)
	{
		#check if the snapin is loaded
		If(!($LoadedSnapins -like $snapin))
		{
			#Check if the snapin is missing
			If(!($RegisteredSnapins -like $Snapin))
			{
				#set the flag if it's not already
				If(!($FoundMissingSnapin))
				{
					$FoundMissingSnapin = $True
				}
				#add the entry to the list
				$MissingSnapins += $Snapin
			}
			Else
			{
				#Snapin is registered, but not loaded, loading it now:
				Write-Host "Loading Windows PowerShell snap-in: $snapin"
				Add-PSSnapin -Name $snapin -EA 0

				If(!($?))
				{
					Write-Error "
	`n`n
	Error loading snapin: $($error[0].Exception.Message)
	`n`n
	Script cannot continue.
	`n`n"
					Return $false
				}				
			}
		}
	}

	If($FoundMissingSnapin)
	{
		Write-Warning "Missing Windows PowerShell snap-ins Detected:"
		Write-Host ""
		$missingSnapins | ForEach-Object {Write-Host "`tMissing Snapin: ($_)"}
		Write-Host ""
		Return $False
	}
	Else
	{
		Return $True
	}
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date -Format G): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date -Format G): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:WordFileName, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date -Format G): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:WordFileName, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date -Format G): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	Write-Verbose "$(Get-Date -Format G): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword runsning in our session
	$wordprocess = $Null
	$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}).Id
	If($null -ne $wordprocess -and $wordprocess -gt 0)
	{
		Write-Verbose "$(Get-Date -Format G): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess)"
		Stop-Process $wordprocess -EA 0
	}
}

Function SetupText
{
	Write-Verbose "$(Get-Date -Format G): Setting up Text"

	[System.Text.StringBuilder] $global:Output = New-Object System.Text.StringBuilder( 16384 )

	If(!$AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName).txt"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}
}

Function SetupHTML
{
	Write-Verbose "$(Get-Date -Format G): Setting up HTML"
	If(!$AddDateTime)
	{
		[string]$Script:HTMLFileName = "$($Script:pwdpath)\$($OutputFileName).html"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:HTMLFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}

	$htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
	out-file -FilePath $Script:HTMLFileName -Force -InputObject $HTMLHead 4>$Null
}

Function SaveandCloseTextDocument
{
	Write-Verbose "$(Get-Date -Format G): Saving Text file"
	Write-Output $global:Output.ToString() | Out-File $Script:TextFileName 4>$Null
}

Function SaveandCloseHTMLDocument
{
	Write-Verbose "$(Get-Date -Format G): Saving HTML file"
	Out-File -FilePath $Script:HTMLFileName -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFileNames
{
	Param([string]$OutputFileName)
	
	If($MSWord -or $PDF)
	{
		CheckWordPreReq
		
		SetupWord
	}
	If($Text)
	{
		SetupText
	}
	If($HTML)
	{
		SetupHTML
	}
	ShowScriptOptions
}

Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	If($Text)
	{
		SaveandCloseTextDocument
	}
	If($HTML)
	{
		SaveandCloseHTMLDocument
	}

	$GotFile = $False

	If($MSWord)
	{
		If(Test-Path "$($Script:WordFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:WordFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Error "Unable to save the output file, $($Script:WordFileName)"
		}
	}
	If($PDF)
	{
		If(Test-Path "$($Script:PDFFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:PDFFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Error "Unable to save the output file, $($Script:PDFFileName)"
		}
	}
	If($Text)
	{
		If(Test-Path "$($Script:TextFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:TextFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Error "Unable to save the output file, $($Script:TextFileName)"
		}
	}
	If($HTML)
	{
		If(Test-Path "$($Script:HTMLFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:HTMLFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Error "Unable to save the output file, $($Script:HTMLFileName)"
		}
	}
	
	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		$emailattachments = @()
		If($MSWord)
		{
			$emailAttachments += $Script:WordFileName
		}
		If($PDF)
		{
			$emailAttachments += $Script:PDFFileName
		}
		If($Text)
		{
			$emailAttachments += $Script:TextFileName
		}
		If($HTML)
		{
			$emailAttachments += $Script:HTMLFileName
		}
		SendEmail $emailAttachments
	}
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): Add DateTime       : $($AddDateTime)"
	Write-Verbose "$(Get-Date -Format G): Administrators     : $($Administrators)"
	Write-Verbose "$(Get-Date -Format G): Applications       : $($Applications)"
	Write-Verbose "$(Get-Date -Format G): Company Name       : $($Script:CoName)"
	Write-Verbose "$(Get-Date -Format G): Company Address    : $($CompanyAddress)"
	Write-Verbose "$(Get-Date -Format G): Company Email      : $($CompanyEmail)"
	Write-Verbose "$(Get-Date -Format G): Company Fax        : $($CompanyFax)"
	Write-Verbose "$(Get-Date -Format G): Company Phone      : $($CompanyPhone)"
	Write-Verbose "$(Get-Date -Format G): Cover Page         : $($CoverPage)"
	Write-Verbose "$(Get-Date -Format G): Customer ID        : $($Script:CustomerID)"
	Write-Verbose "$(Get-Date -Format G): CSV                : $($CSV)"
	Write-Verbose "$(Get-Date -Format G): Dev                : $($Dev)"
	Write-Verbose "$(Get-Date -Format G): DeliveryGroups     : $($DeliveryGroups)"
	If($Dev)
	{
		Write-Verbose "$(Get-Date -Format G): DevErrorFile       : $($Script:DevErrorFile)"
	}
	Write-Verbose "$(Get-Date -Format G): DGUtilization      : $($DeliveryGroupsUtilization)"
	If($HTML)
	{
		Write-Verbose "$(Get-Date -Format G): HTMLFilename       : $($Script:HTMLFilename)"
	}
	If($MSWord)
	{
		Write-Verbose "$(Get-Date -Format G): WordFilename       : $($Script:WordFilename)"
	}
	If($PDF)
	{
		Write-Verbose "$(Get-Date -Format G): PDFFilename        : $($Script:PDFFilename)"
	}
	If($Text)
	{
		Write-Verbose "$(Get-Date -Format G): TextFilename       : $($Script:TextFilename)"
	}
	Write-Verbose "$(Get-Date -Format G): Folder             : $($Script:pwdpath)"
	Write-Verbose "$(Get-Date -Format G): From               : $($From)"
	Write-Verbose "$(Get-Date -Format G): Hosting            : $($Hosting)"
	Write-Verbose "$(Get-Date -Format G): Log                : $($Log)"
	Write-Verbose "$(Get-Date -Format G): Logging            : $($Logging)"
	If($Logging)
	{
		Write-Verbose "$(Get-Date -Format G):    Start Date      : $($StartDate)"
		Write-Verbose "$(Get-Date -Format G):    End Date        : $($EndDate)"
	}
	Write-Verbose "$(Get-Date -Format G): MachineCatalogs    : $($MachineCatalogs)"
	Write-Verbose "$(Get-Date -Format G): MaxDetail          : $($MaxDetails)"
	Write-Verbose "$(Get-Date -Format G): NoADPolicies       : $($NoADPolicies)"
	Write-Verbose "$(Get-Date -Format G): NoPolicies         : $($NoPolicies)"
	Write-Verbose "$(Get-Date -Format G): Policies           : $($Policies)"
	Write-Verbose "$(Get-Date -Format G): Save As PDF        : $($PDF)"
	Write-Verbose "$(Get-Date -Format G): Save As HTML       : $($HTML)"
	Write-Verbose "$(Get-Date -Format G): Save As TEXT       : $($TEXT)"
	Write-Verbose "$(Get-Date -Format G): Save As WORD       : $($MSWORD)"
	Write-Verbose "$(Get-Date -Format G): ScriptInfo         : $($ScriptInfo)"
	Write-Verbose "$(Get-Date -Format G): Section            : $($Section)"
	Write-Verbose "$(Get-Date -Format G): Site Name          : $($CCSiteName)"
	Write-Verbose "$(Get-Date -Format G): Smtp Port          : $($SmtpPort)"
	Write-Verbose "$(Get-Date -Format G): Smtp Server        : $($SmtpServer)"
	Write-Verbose "$(Get-Date -Format G): StoreFront         : $($StoreFront)"
	Write-Verbose "$(Get-Date -Format G): Title              : $($Script:Title)"
	Write-Verbose "$(Get-Date -Format G): To                 : $($To)"
	Write-Verbose "$(Get-Date -Format G): Use SSL            : $($UseSSL)"
	Write-Verbose "$(Get-Date -Format G): User Name          : $($UserName)"
	Write-Verbose "$(Get-Date -Format G): VDA Registry Keys  : $($VDARegistryKeys)"
	Write-Verbose "$(Get-Date -Format G): CC Version         : $($Script:CCSiteVersion)"
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): OS Detected        : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date -Format G): PoSH version       : $($Host.Version)"
	Write-Verbose "$(Get-Date -Format G): PSCulture          : $($PSCulture)"
	Write-Verbose "$(Get-Date -Format G): PSUICulture        : $($PSUICulture)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Word language      : $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date -Format G): Word version       : $($Script:WordProduct)"
	}
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): Script start       : $($Script:StartTime)"
	Write-Verbose "$(Get-Date -Format G): "
}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		$Script:Word.quit()
		Write-Verbose "$(Get-Date -Format G): System Cleanup"
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
		If(Test-Path variable:global:word)
		{
			Remove-Variable -Name word -Scope Global 4>$Null
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date -Format G): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Function OutputWarning
{
	Param([string] $txt)
	Write-Warning $txt
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 $txt
		WriteWordLIne 0 0 ""
	}
	If($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 0 0 $txt
	}
}

Function OutputNotice
{
	Param([string] $txt)
	#Write-Host $txt -ForegroundColor White
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 $txt
		WriteWordLIne 0 0 ""
	}
	If($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 0 0 $txt
	}
}

Function OutputAdminsForDetails
{
	Param([object] $Admins)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Administrators"
		## Create an array of hashtables to store our admins
		[System.Collections.Hashtable[]] $AdminsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Administrators"
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Administrators"
		$rowdata = @()
	}
	
	ForEach($Admin in $Admins)
	{
		$Tmp = ""
		If($Admin.Enabled)
		{
			$Tmp = "Enabled"
		}
		Else
		{
			$Tmp = "Disabled"
		}
		
		If($MSWord -or $PDF)
		{
			$AdminsWordTable += @{ 
			AdminName = $Admin.Name;
			Role = $Admin.Rights[0].RoleName;
			Status = $Tmp;
			}
		}
		If($Text)
		{
			Line 1 "Administrator Name`t: " $Admin.Name
			Line 1 "Role`t`t`t: " $Admin.Rights[0].RoleName
			Line 1 "Status`t`t`t: " $tmp
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,(
			$Admin.Name,$htmlwhite,
			$Admin.Rights[0].RoleName,$htmlwhite,
			$tmp,$htmlwhite))
		}
	}
	
	If($MSWord -or $PDF)
	{
		If($AdminsWordTable.Count -eq 0)
		{
			$AdminsWordTable += @{ 
			AdminName = "No admins found";
			Role = "N/A";
			Status = "N/A";
			}
		}

		$Table = AddWordTable -Hashtable $AdminsWordTable `
		-Columns AdminName, Role, Status `
		-Headers "Administrator Name", "Role", "Status" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 275;
		$Table.Columns.Item(2).Width = 200;
		$Table.Columns.Item(3).Width = 60;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		If($Admins.Count -eq 0)
		{
			Line 0 "No Admins found"
		}
	}
	If($HTML)
	{
		If($rowdata.Count -eq 0)
		{
			$rowdata += @(,(
			"No admins found",$htmlwhite,
			"N/A",$htmlwhite,
			"N/A",$htmlwhite))
		}
		
		$columnHeaders = @(
		'Administrator Name',($global:htmlsb),
		'Role',($global:htmlsb),
		'Status',($global:htmlsb))

		$msg = ""
		$columnWidths = @("275","200","60")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "535"
	}
}

Function TranscriptLogging
{
	If($Log) 
	{
		try 
		{
			If($Script:StartLog -eq $false)
			{
				Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
			}
			Else
			{
				Start-Transcript -Path $Script:LogPath -Append -Verbose:$false | Out-Null
			}
			Write-Verbose "$(Get-Date -Format G): Transcript/log started at $Script:LogPath"
			$Script:StartLog = $true
		} 
		catch 
		{
			Write-Verbose "$(Get-Date -Format G): Transcript/log failed at $Script:LogPath"
			$Script:StartLog = $false
		}
	}
}

Function Get-IPAddress
{
	Param([string]$ComputerName)
	
	$IPAddress = "Unable to determine"
	
	Try
	{
		$IP = Test-Connection -ComputerName $ComputerName -Count 1 -EA 0 | Select-Object IPV4Address
	}
	
	Catch
	{
		$IP = "Unable to resolve IP address"
	}

	If($? -and $Null -ne $IP)
	{
		$IPAddress = $IP.IPV4Address.IPAddressToString
	}
	
	Return $IPAddress
}
#endregion

#region email function
Function SendEmail
{
	Param([array]$Attachments)
	Write-Verbose "$(Get-Date -Format G): Prepare to email"

	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.

"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()
	
	If($From -Like "anonymous@*")
	{
		#https://serverfault.com/questions/543052/sending-unauthenticated-mail-through-ms-exchange-with-powershell-windows-server
		$anonUsername = "anonymous"
		$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
		$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $anonCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $anonCredentials *>$Null 
		}
		
		If($?)
		{
			Write-Verbose "$(Get-Date -Format G): Email successfully sent using anonymous credentials"
		}
		ElseIf(!$?)
		{
			$e = $error[0]

			Write-Verbose "$(Get-Date -Format G): Email was not sent:"
			Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
		}
	}
	Else
	{
		If($UseSSL)
		{
			Write-Verbose "$(Get-Date -Format G): Trying to send email using current user's credentials with SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL *>$Null
		}
		Else
		{
			Write-Verbose  "$(Get-Date -Format G): Trying to send email using current user's credentials without SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
		}

		If(!$?)
		{
			$e = $error[0]
			
			#error 5.7.57 is O365 and error 5.7.0 is gmail
			If($null -ne $e.Exception -and $e.Exception.ToString().Contains("5.7"))
			{
				#The server response was: 5.7.xx SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
				Write-Verbose "$(Get-Date -Format G): Current user's credentials failed. Ask for usable credentials."

				If($Dev)
				{
					Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
				}

				$error.Clear()

				$emailCredentials = Get-Credential -UserName $From -Message "Enter the password to send email"

				If($UseSSL)
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-UseSSL -credential $emailCredentials *>$Null 
				}
				Else
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-credential $emailCredentials *>$Null 
				}

				If($?)
				{
					Write-Verbose "$(Get-Date -Format G): Email successfully sent using new credentials"
				}
				ElseIf(!$?)
				{
					$e = $error[0]

					Write-Verbose "$(Get-Date -Format G): Email was not sent:"
					Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): Email was not sent:"
				Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
			}
		}
	}
}
#endregion

#region getadmins function from Citrix
Function GetAdmins
{
	Param([string]$xType="", [string]$xName="")
	
	Switch ($xType)
	{
		"ApplicationGroup" {
			$scopes = $Null

			$permissions = Get-AdminPermission @CCParams2 | `
			Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "ApplicationGroup"} | `
			Select-Object -ExpandProperty Id

			$roles = Get-AdminRole @CCParams2 | `
			Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | `
			Select-Object -ExpandProperty Id

			#this is an unscoped object type as $admins is done differently than the others
			$Admins = Get-AdminAdministrator @CCParams2 | `
			Where-Object {$_.UserIdentityType -ne "Sid" -and (-not ([String]::IsNullOrEmpty($_.UserIdentityType)))} | `
			Where-Object {$_.Rights | Where-Object {$roles -contains $_.RoleId}}
		}
		"Catalog" {
			$scopes = (Get-BrokerCatalog -Name $xName @CCParams2).Scopes | `
			Select-Object -ExpandProperty ScopeId

			$permissions = Get-AdminPermission @CCParams2 | Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "Catalog" } | `
			Select-Object -ExpandProperty Id

			$roles = Get-AdminRole @CCParams2 | `
			Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | `
			Select-Object -ExpandProperty Id

			$Admins = Get-AdminAdministrator @CCParams2 | `
			Where-Object {$_.UserIdentityType -ne "Sid" -and (-not ([String]::IsNullOrEmpty($_.Name)))} | `
			Where-Object {$_.Rights | `
			Where-Object {($_.ScopeId -eq [guid]::Empty -or $scopes -contains $_.ScopeId) -and $roles -contains $_.RoleId}}
		}
		"DesktopGroup" {
			$scopes = (Get-BrokerDesktopGroup -Name $xName @CCParams2).Scopes | `
			Select-Object -ExpandProperty ScopeId

			$permissions = Get-AdminPermission @CCParams2 | `
			Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "DesktopGroup" } | `
			Select-Object -ExpandProperty Id

			$roles = Get-AdminRole @CCParams2 | `
			Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | `
			Select-Object -ExpandProperty Id

			$Admins = Get-AdminAdministrator @CCParams2 | `
			Where-Object {$_.UserIdentityType -ne "Sid" -and (-not ([String]::IsNullOrEmpty($_.Name)))} | `
			Where-Object {$_.Rights | `
			Where-Object {($_.ScopeId -eq [guid]::Empty -or $scopes -contains $_.ScopeId) -and $roles -contains $_.RoleId}}
		}
		"Host" {
			$scopes = Get-hypscopedobject -ObjectName $xName @CCParams2
            If($null -ne $scopes)
            {
			    $scopes = (Get-hypscopedobject -ObjectName $xName @CCParams2).ScopeId

			    $permissions = Get-AdminPermission @CCParams2 | `
			    Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "Connection" -or `
			    $_.MetadataMap["Citrix_ObjectType"] -eq "Host"} | `
			    Select-Object -ExpandProperty Id		

			    $roles = Get-AdminRole @CCParams2 | `
			    Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | `
			    Select-Object -ExpandProperty Id

			    $Admins = Get-AdminAdministrator @CCParams2 | `
			    Where-Object {$_.UserIdentityType -ne "Sid" -and (-not ([String]::IsNullOrEmpty($_.Name)))} | `
			    Where-Object {$_.Rights | `
			    Where-Object {($_.ScopeId -eq [guid]::Empty -or `
			    $scopes -contains $_.ScopeId) -and $roles -contains $_.RoleId}}
            }
            Else
            {
                #work around issue of "The property 'ScopeId' cannot be found on this object. Verify that the property exists."
			    $scopes = $null

			    $permissions = Get-AdminPermission @CCParams2 | `
			    Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "Connection" -or `
			    $_.MetadataMap["Citrix_ObjectType"] -eq "Host"} | `
			    Select-Object -ExpandProperty Id		

			    $roles = Get-AdminRole @CCParams2 | `
			    Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | `
			    Select-Object -ExpandProperty Id

			    $Admins = Get-AdminAdministrator @CCParams2 | `
			    Where-Object {$_.UserIdentityType -ne "Sid" -and (-not ([String]::IsNullOrEmpty($_.Name)))} | `
			    Where-Object {$_.Rights | `
			    Where-Object {($_.ScopeId -eq [guid]::Empty -or `
			    $scopes -contains $_.ScopeId) -and $roles -contains $_.RoleId}}
            }
		}
		"Storefront" {
			$scopes = $Null

			$permissions = Get-AdminPermission @CCParams2 | `
			Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "Storefront" } | `
			Select-Object -ExpandProperty Id

			$roles = Get-AdminRole @CCParams2 | `
			Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | `
			Select-Object -ExpandProperty Id

			#this is an unscoped object type as $admins is done differently than the others
			$Admins = Get-AdminAdministrator @CCParams2 | `
			Where-Object {$_.UserIdentityType -ne "Sid" -and (-not ([String]::IsNullOrEmpty($_.Name)))} | `
			Where-Object {$_.Rights | `
			Where-Object {$roles -contains $_.RoleId}}
		}
	}
	
	# $scopes = (Get-BrokerCatalog -Name "XenApp 75" -adminaddress xd75 ).Scopes | Select-Object -ExpandProperty ScopeId

	# First, get all the permissions which are relevant to this object type
	# Change "Catalog" here as appropriate for the object type you're interested in
	# $permissions = Get-AdminPermission @CCParams2 | Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "Catalog" } | Select-Object -ExpandProperty Id

	# Now, get all the roles which include at least one of those permissions
	# $roles = Get-AdminRole @CCParams2 | Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | Select-Object -ExpandProperty Id

	# Finally, get all administrators which have a scope/role pair which matches
	#$Admins = Get-AdminAdministrator @CCParams2 | Where-Object {
	#	$_.Rights | Where-Object {
	#		# [guid]::Empty is the GUID for the All scope
	#		# Remove the next line if you're dealing with an unscoped object type
	#		($_.ScopeId -eq [guid]::Empty -or $scopes -contains $_.ScopeId) -and
	#		$roles -contains $_.RoleId
	#	}
	#}
	#$Admins = Get-AdminAdministrator @CCParams2 | Where-Object {$_.Rights | Where-Object {($_.ScopeId -eq [guid]::Empty -or $scopes -contains $_.ScopeId) -and	$roles -contains $_.RoleId}}

	#$Admins = $Admins | Sort-Object Name
	Return ,$Admins
}
#endregion

#region Machine Catalog functions
Function ProcessMachineCatalogs
{
	Write-Verbose "$(Get-Date -Format G): Retrieving Machine Catalogs"

	$txt = "Machine Catalogs"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$AllMachineCatalogs = Get-BrokerCatalog @CCParams2 -SortBy Name 

	If($? -and $Null -ne $AllMachineCatalogs)
	{
		OutputMachines $AllMachineCatalogs
	}
	ElseIf($? -and ($Null -eq $AllMachineCatalogs))
	{
		$txt = "There are no Machine Catalogs"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve Machine Catalogs"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date -Format G): "
}

Function OutputMachines
{
	Param([object]$Catalogs)
	
	Write-Verbose "$(Get-Date -Format G): `tProcessing Machine Catalogs"
	
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $WordTable = @();
	}
	If($Text)
	{
		Line 0 "                                                                              No. of   Allocated Allocation                                        "
		Line 0 "Machine Catalog                          Machine Type                         Machines Machines  Type       User Data     Provisioning Method      "
		Line 0 "==================================================================================================================================================="
		#       1234567890123456789012345678901234567890S123456789012345678901234567890123456S12345678S12345678SS1234567890S1234567890123S1234567890123456789012345
		#                                                Single-session OS (Remote PC Access)                               On local Disk Machine creation services
	}
	If($HTML)
	{
		$rowdata = @()
	}

	ForEach($Catalog in $Catalogs)
	{
		$xCatalogType = ""
		$xAllocationType = ""
		$xPersistType = ""
		$xProvisioningType = ""
		
		If($Catalog.MachinesArePhysical -eq $True -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "SingleSession")
		{
			$xCatalogType = "Single-session OS"
			$Script:TotalDesktopOSCatalogs++
		}
		ElseIf($Catalog.MachinesArePhysical -eq $True -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "MultiSession")
		{
			$xCatalogType = "Multi-session OS"
			$Script:TotalServerOSCatalogs++
		}
		ElseIf($Catalog.MachinesArePhysical -eq $False -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "SingleSession")
		{
			$xCatalogType = "Single-session OS (Virtual)"
			$Script:TotalDesktopOSCatalogs++
		}
		ElseIf($Catalog.MachinesArePhysical -eq $False -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "MultiSession")
		{
			$xCatalogType = "Multi-session OS (Virtual)"
			$Script:TotalServerOSCatalogs++
		}
		ElseIf($Catalog.MachinesArePhysical -eq $True -and $Catalog.IsRemotePC -eq $True)
		{
			$xCatalogType = "Single-session OS (Remote PC Access)"
			$Script:TotalRemotePCCatalogs++
		}
		
		Switch ($Catalog.AllocationType)
		{
			"Static"	{$xAllocationType = "Static"; Break}
			"Permanent"	{$xAllocationType = "Static"; Break}
			"Random"	{$xAllocationType = "Random"; Break}
			Default		{$xAllocationType = "Allocation type could not be determined: $($Catalog.AllocationType)"; Break}
		}
		Switch ($Catalog.PersistUserChanges)
		{
			"OnLocal" {$xPersistType = "On local disk"; Break}
			"Discard" {$xPersistType = "Discard"; Break}
			Default   {$xPersistType = "User data could not be determined: $($Catalog.PersistUserChanges)"; Break}
		}
		Switch ($Catalog.ProvisioningType)
		{
			"Manual" {$xProvisioningType = "Manual"; Break}
			"PVS"    {$xProvisioningType = "Provisioning Services"; Break}
			"MCS"    {$xProvisioningType = "Machine creation services"; Break}
			Default  {$xProvisioningType = "Provisioning method could not be determined: $($Catalog.ProvisioningType)"; Break}
		}

		$Machines = @(Get-BrokerMachine @CCParams2 -CatalogName $Catalog.Name -SortBy DNSName)
		If($? -and ($Null -ne $Machines))
		{
			$NumberOfMachines = $Machines.Count
		}
		
		If($MSWord -or $PDF)
		{
			$WordTable += @{
			MachineCatalogName = $Catalog.Name; 
			MachineType        = $xCatalogType; 
			NoOfMachines       = $NumberOfMachines.ToString();
			AllocatedMachines  = $Catalog.UsedCount.ToString(); 
			AllocationType     = $xAllocationType;
			UserData           = $xPersistType;
			ProvisioningMethod = $xProvisioningType;
			}
		}
		If($Text)
		{
			Line 0 ( "{0,-40} {1,-36} {2,8} {3,8}  {4,-10} {5,-13} {6,-25}" -f `
			$Catalog.Name, 
			$xCatalogType, 
			$NumberOfMachines.ToString(), 
			$Catalog.UsedCount.ToString(), 
			$xAllocationType, 
			$xPersistType, 
			$xProvisioningType)
		}
		If($HTML)
		{
			$rowdata += @(,(
			$Catalog.Name,$htmlwhite,
			$xCatalogType,$htmlwhite,
			$NumberOfMachines.ToString(),$htmlwhite,
			$Catalog.UsedCount.ToString(),$htmlwhite,
			$xAllocationType,$htmlwhite,
			$xPersistType,$htmlwhite,
			$xProvisioningType,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $WordTable `
		-Columns  MachineCatalogName, MachineType, NoOfMachines, AllocatedMachines, AllocationType, UserData, ProvisioningMethod `
		-Headers  "Machine Catalog", "Machine type", "No. of machines", "Allocated machines", "Allocation Type", "User data", "Provisioning method" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 105;
		$Table.Columns.Item(2).Width = 100;
		$Table.Columns.Item(3).Width = 75;
		$Table.Columns.Item(4).Width = 50;
		$Table.Columns.Item(5).Width = 55;
		$Table.Columns.Item(6).Width = 50;
		$Table.Columns.Item(7).Width = 65;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		$columnHeaders = @(
		'Machine Catalog',($global:htmlsb),
		'Machine type',($global:htmlsb),
		'No. of machines',($global:htmlsb),
		'Allocated machines',($global:htmlsb),
		'Allocation Type',($global:htmlsb),
		'User data',($global:htmlsb),
		'Provisioning method',($global:htmlsb)
		)

		$columnWidths = @("125","175","75","50","55","75","145")
		$msg = ""
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "700"
	}
	
	ForEach($Catalog in $Catalogs)
	{
		Write-Verbose "$(Get-Date -Format G): `t`tAdding Catalog $($Catalog.Name)"
		$xCatalogType = ""
		$xAllocationType = ""
		$xPersistType = ""
		$xProvisioningType = ""
		$xVDAVersion = ""
		
		If($Catalog.MachinesArePhysical -eq $True -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "SingleSession")
		{
			$xCatalogType = "Single-session OS"
		}
		ElseIf($Catalog.MachinesArePhysical -eq $True -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "MultiSession")
		{
			$xCatalogType = "Multi-session OS"
		}
		ElseIf($Catalog.MachinesArePhysical -eq $False -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "SingleSession")
		{
			$xCatalogType = "Single-session OS (Virtual)"
		}
		ElseIf($Catalog.MachinesArePhysical -eq $False -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "MultiSession")
		{
			$xCatalogType = "Multi-session OS (Virtual)"
		}
		ElseIf($Catalog.MachinesArePhysical -eq $True -and $Catalog.IsRemotePC -eq $True)
		{
			$xCatalogType = "Single-session OS (Remote PC Access)"
		}

		Switch ($Catalog.AllocationType)
		{
			"Static"	{$xAllocationType = "Permanent"; Break}
			"Permanent"	{$xAllocationType = "Permanent"; Break}
			"Random"	{$xAllocationType = "Random"; Break}
			Default		{$xAllocationType = "Allocation type could not be determined: $($Catalog.AllocationType)"; Break}
		}
		Switch ($Catalog.PersistUserChanges)
		{
			"OnLocal" {$xPersistType = "On local disk"; Break}
			"Discard" {$xPersistType = "Discard"; Break}
			Default   {$xPersistType = "User data could not be determined: $($Catalog.PersistUserChanges)"; Break}
		}
		Switch ($Catalog.ProvisioningType)
		{
			"Manual" {$xProvisioningType = "Manual"; Break}
			"PVS"    {$xProvisioningType = "Provisioning Services"; Break}
			"MCS"    {$xProvisioningType = "Machine creation services"; Break}
			Default  {$xProvisioningType = "Provisioning method could not be determined: $($Catalog.ProvisioningType)"; Break}
		}
		Switch ($Catalog.MinimumFunctionalLevel)
		{
			"L5" 	{$xVDAVersion = "5.6 FP1 (Windows XP and Windows Vista)"; Break}
			"L7"	{$xVDAVersion = "7.0 (or newer)"; Break}
			"L7_6"	{$xVDAVersion = "7.6 (or newer)"; Break}
			"L7_7"	{$xVDAVersion = "7.7 (or newer)"; Break}
			"L7_8"	{$xVDAVersion = "7.8 (or newer)"; Break}
			"L7_9"	{$xVDAVersion = "7.9 (or newer)"; Break}
			"L7_20"	{$xVDAVersion = "1811 (or newer)"; Break}
			"L7_25"	{$xVDAVersion = "2003 (or newer)"; Break}
			Default {$xVDAVersion = "Unable to determine VDA version: $($Catalog.MinimumFunctionalLevel)"; Break}
		}

		If($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -eq $True)
		{
			$RemotePCAccounts = Get-BrokerRemotePCAccount @CCParams2 -CatalogUid $Catalog.Uid
			
			If(!$?)
			{
				$RemotePCOU = "Unable to retrieve"
				$RemotePCSubOU = "Unable to retrieve"
			}
			ElseIf($? -and $Null -eq $RemotePCAccounts)
			{
				$RemotePCOU = "No RemotePC OU configured"
				$RemotePCSubOU = "N/A"
			}
			ElseIf($? -and $Null -ne $RemotePCAccounts)
			{
				#$RemotePCOU = $Results.OU
				#$RemotePCSubOU = $Results.AllowSubfolderMatches.ToString()
				#Handled later
			}
		}
		
		$MachineData = $Null
		$Machines = @(Get-BrokerMachine @CCParams2 -CatalogName $Catalog.Name -SortBy DNSName)
		If($? -and ($Null -ne $Machines))
		{
			$NumberOfMachines = $Machines.Count
			
			#don't process Manually provisioned
			#there is no $Catalog.ProvisioningSchemeId for manually provisioned catalogs or ones based on PVS, only MCS
			If($Catalog.ProvisioningType -eq "MCS")
			{
				If($null -ne $Catalog.ProvisioningSchemeId)
				{
					$MachineData = Get-ProvScheme -ProvisioningSchemeUid $Catalog.ProvisioningSchemeId -EA 0
				}
				Else
				{
					$MachineData = $Null
				}
				
				If($? -and $Null -ne $MachineData)
				{
					$tmp1 = $MachineData.MasterImageVM.Split("\")
					$tmp2 = $tmp1[$tmp1.count -1]
					$tmp3 = $tmp2.Split(".")
					$xDiskImage = $tmp3[0]

					$MasterVM = ""
					ForEach($Item in $tmp1)
					{
						If(($Item.EndsWith(".vm")) -or ($Item.EndsWith(".template"))) #.template for Nutanix
						{
							$MasterVM = $Item
						}
					}
					
					If(($Catalog.MinimumFunctionalLevel -eq "L7_9" -or 
					    $Catalog.MinimumFunctionalLevel -eq "L7_20" -or 
						$Catalog.MinimumFunctionalLevel -eq "L7_25") -and 
					(($xAllocationType -eq "Random") -or 
					($xAllocationType -eq "Permanent" -and $xPersistType -eq "Discard" )))
					{
						$TempDiskCacheSize = $MachineData.WriteBackCacheDiskSize
						$TempMemoryCacheSize = $MachineData.WriteBackCacheMemorySize
					}
					
					If(($Catalog.MinimumFunctionalLevel -eq "L7_9" -or 
					    $Catalog.MinimumFunctionalLevel -eq "L7_20" -or 
						$Catalog.MinimumFunctionalLevel -eq "L7_25") -and 
					($xAllocationType -eq "Permanent" -and $xPersistType -eq "On local disk" ) -and 
					((Get-ConfigEnabledFeature -EA 0) -contains "DedicatedFullDiskClone"))
					{
						If($MachineData.UseFullDiskCloneProvisioning -eq $True)
						{
							$VMCopyMode = "Full Copy"
						}
						Else
						{
							$VMCopyMode = "Fast Clone"
						}
					}
				}
				Else
				{
					$xDiskImage = "Unable to retrieve details"
				}
			}
			Else
			{
				$xDiskImage = "No details for manually or PVS provisioned machines"
			}
		}
		Else
		{
			Write-Host "Unable to retrieve details for Machine Catalog $($Catalog.Name)" -ForegroundColor White
		}
		
		If($Catalog.ProvisioningType -eq "MCS")
		{
			$IdentityPool = @(Get-AcctIdentityPool @CCParams2 -IdentityPoolName $Catalog.Name)
			
			If($? -and $Null -ne $IdentityPool)
			{
				If($IdentityPool.NamingSchemeType -eq "None")
				{
					$IdentityDomain           = "N/A"
					$IdentityNamingScheme     = "Use existing AD accounts"
					$IdentityNamingSchemeType = "None"
					$IdentityOU               = "N/A"
				}
				Else
				{
					$IdentityDomain           = $IdentityPool.Domain
					$IdentityNamingScheme     = $IdentityPool.NamingScheme
					$IdentityNamingSchemeType = $IdentityPool.NamingSchemeType
					$IdentityOU               = $IdentityPool.OU
				}
			}
			Else
			{
				$IdentityDomain           = "Not Found"
				$IdentityNamingScheme     = "Not Found"
				$IdentityNamingSchemeType = "Not Found"
				$IdentityOU               = "Not Found"
			}
		}

		$SessionSupport = "Single-session OS"
		If($Catalog.SessionSupport -eq "SingleSession")
		{
			$SessionSupport = "Single-session OS"
		}
		Else
		{
			$SessionSupport = "Multi-session OS"
		}
		
		If($MSWord -or $PDF)
		{
			$Selection.InsertNewPage()
			WriteWordLine 2 0 "Machine Catalog: $($Catalog.Name)"
			[System.Collections.Hashtable[]] $CatalogInformation = @()
			
			If($Catalog.ProvisioningType -eq "MCS")
			{
				$CatalogInformation += @{Data = "Description"; Value = $Catalog.Description; }
				$CatalogInformation += @{Data = "Machine type"; Value = $xCatalogType; }
				$CatalogInformation += @{Data = "No. of machines"; Value = $NumberOfMachines.ToString(); }
				$CatalogInformation += @{Data = "Allocated machines"; Value = $Catalog.UsedCount.ToString(); }
				$CatalogInformation += @{Data = "Allocation type"; Value = $xAllocationType; }
				$CatalogInformation += @{Data = "User data"; Value = $xPersistType; }
				$CatalogInformation += @{Data = "Provisioning method"; Value = $xProvisioningType; }
				$CatalogInformation += @{Data = "Account naming scheme"; Value = $IdentityNamingScheme; }
				$CatalogInformation += @{Data = "Naming scheme type"; Value = $IdentityNamingSchemeType; }
				$CatalogInformation += @{Data = "AD Domain"; Value = $IdentityDomain; }
				$CatalogInformation += @{Data = "AD Location"; Value = $IdentityOU; }
				$CatalogInformation += @{Data = "Set to VDA version"; Value = $xVDAVersion; }
				If($Null -ne $MachineData)
				{
					If( $MachineData.PSObject.Properties[ 'HostingUnitName' ] )
					{
						## GRL - The property 'HostingUnitName' cannot be found on this object. Verify that the property exists
						$CatalogInformation += @{Data = "Resources"; Value = $MachineData.HostingUnitName; }
					}
				}
				$CatalogInformation += @{Data = "Zone"; Value = $Catalog.ZoneName; }

				If($Null -ne $MachineData)
				{
					$CatalogInformation += @{Data = "Master VM"; Value = $MasterVM; }
					$CatalogInformation += @{Data = "Disk Image"; Value = $xDiskImage; }
					$CatalogInformation += @{Data = "Virtual CPUs"; Value = $MachineData.CpuCount; }
					$CatalogInformation += @{Data = "Memory"; Value = "$($MachineData.MemoryMB) MB"; }
					$CatalogInformation += @{Data = "Hard disk"; Value = "$($MachineData.DiskSize) GB"; }
				}
				ElseIf($Null -eq $MachineData)
				{
					$CatalogInformation += @{Data = "Master VM"; Value = $MasterVM; }
					$CatalogInformation += @{Data = "Disk Image"; Value = $xDiskImage; }
					$CatalogInformation += @{Data = "Virtual CPUs"; Value = "Unable to retrieve details"; }
					$CatalogInformation += @{Data = "Memory"; Value = "Unable to retrieve details"; }
					$CatalogInformation += @{Data = "Hard disk"; Value = "Unable to retrieve details"; }
				}
			
				If(($Catalog.MinimumFunctionalLevel -eq "L7_9" -or 
					$Catalog.MinimumFunctionalLevel -eq "L7_20" -or 
					$Catalog.MinimumFunctionalLevel -eq "L7_25") -and 
				(($xAllocationType -eq "Random") -or 
				($xAllocationType -eq "Permanent" -and $xPersistType -eq "Discard" )))
				{
					$CatalogInformation += @{Data = "Temporary memory cache size"; Value = "$($TempMemoryCacheSize) MB"; }
					$CatalogInformation += @{Data = "Temporary disk cache size"; Value = "$($TempDiskCacheSize) GB"; }
				}

				If(($Catalog.MinimumFunctionalLevel -eq "L7_9" -or 
					$Catalog.MinimumFunctionalLevel -eq "L7_20" -or 
					$Catalog.MinimumFunctionalLevel -eq "L7_25") -and 
				($xAllocationType -eq "Permanent" -and $xPersistType -eq "On local disk" ) -and 
				((Get-ConfigEnabledFeature -EA 0) -contains "DedicatedFullDiskClone"))
				{
					$CatalogInformation += @{Data = "VM copy mode"; Value = $VMCopyMode; }
				}
				
				If($Null -ne $Machines)
				{
					If($Machines -is [array] -and $Machines.Count)
					{
						If([String]::IsNullOrEmpty($Machines[0].AgentVersion))
						{
							$CatalogInformation += @{Data = "Installed VDA version"; Value = "-"; }
							$CatalogInformation += @{Data = "Operating System"; Value = "-"; }
						}
						Else
						{
							$CatalogInformation += @{Data = "Installed VDA version"; Value = $Machines[0].AgentVersion; }
							$CatalogInformation += @{Data = "Operating System"; Value = $Machines[0].OSType; }
						}
					}
					Else
					{
						$CatalogInformation += @{Data = "Installed VDA version"; Value = "Unable to retrieve details"; }
						$CatalogInformation += @{Data = "Operating System"; Value = "Unable to retrieve details"; }
					}
				}
				ElseIf($Null -eq $Machines)
				{
					$CatalogInformation += @{Data = "Installed VDA version"; Value = "Unable to retrieve details"; }
					$CatalogInformation += @{Data = "Operating System"; Value = "Unable to retrieve details"; }
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "PVS")
			{
				$CatalogInformation += @{Data = "Description"; Value = $Catalog.Description; }
				$CatalogInformation += @{Data = "Machine type"; Value = $xCatalogType; }
				$CatalogInformation += @{Data = "Provisioning method"; Value = $xProvisioningType; }
				$CatalogInformation += @{Data = "PVS address"; Value = $Catalog.PvsAddress; }
				$CatalogInformation += @{Data = "Allocation type"; Value = $xAllocationType; }
				$CatalogInformation += @{Data = "Set to VDA version"; Value = $xVDAVersion; }
				$CatalogInformation += @{Data = "Zone"; Value = $Catalog.ZoneName; }
			
				If($Null -ne $Machines)
				{
					If($Machines -is [array] -and $Machines.Count)
					{
						If([String]::IsNullOrEmpty($Machines[0].AgentVersion))
						{
							$CatalogInformation += @{Data = "Installed VDA version"; Value = "-"; }
							$CatalogInformation += @{Data = "Operating System"; Value = "-"; }
						}
						Else
						{
							$CatalogInformation += @{Data = "Installed VDA version"; Value = $Machines[0].AgentVersion; }
							$CatalogInformation += @{Data = "Operating System"; Value = $Machines[0].OSType; }
						}
					}
					Else
					{
						$CatalogInformation += @{Data = "Installed VDA version"; Value = "Unable to retrieve details"; }
						$CatalogInformation += @{Data = "Operating System"; Value = "Unable to retrieve details"; }
					}
				}
				ElseIf($Null -eq $Machines)
				{
					$CatalogInformation += @{Data = "Installed VDA version"; Value = "Unable to retrieve details"; }
					$CatalogInformation += @{Data = "Operating System"; Value = "Unable to retrieve details"; }
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -eq $True)
			{
				$CatalogInformation += @{Data = "Description"; Value = $Catalog.Description; }
				
				If($RemotePCAccounts -is [array])
				{
					ForEach($RemotePCAccount in $RemotePCAccounts)
					{
						$CatalogInformation += @{Data = "Organizational Units"; Value = $RemotePCAccount.OU; }
						$CatalogInformation += @{Data = "     Allow subfolder matches"; Value = $RemotePCAccount.AllowSubfolderMatches.ToString(); }
						If($RemotePCAccount.MachinesExcluded.Count -eq 0)
						{
							$CatalogInformation += @{Data = "     Machines excluded"; Value = "None"; }
						}
						Else
						{
							$cnt = -1
							ForEach($Item in $RemotePCAccount.MachinesExcluded)
							{
								$cnt++
								
								If($cnt -eq 0)
								{
									$CatalogInformation += @{Data = "     Machines excluded"; Value = $Item; }
								}
								Else
								{
									$CatalogInformation += @{Data = ""; Value = $Item; }
								}
							}
						}

						If($RemotePCAccount.MachinesIncluded -eq "*")
						{
							$CatalogInformation += @{Data = "     Machines Included"; Value = "All"; }
						}
						Else
						{
							$cnt = -1
							ForEach($Item in $RemotePCAccount.MachinesIncluded)
							{
								$cnt++
								
								If($cnt -eq 0)
								{
									$CatalogInformation += @{Data = "     Machines Included"; Value = $Item; }
								}
								Else
								{
									$CatalogInformation += @{Data = ""; Value = $Item; }
								}
							}
						}
					}
				}
				Else
				{
					$CatalogInformation += @{Data = "Organizational Units"; Value = $RemotePCOU; }
					$CatalogInformation += @{Data = "     Allow subfolder matches"; Value = $RemotePCSubOU; }
				}

				$CatalogInformation += @{Data = "Machine type"; Value = $xCatalogType; }
				$CatalogInformation += @{Data = "No. of machines"; Value = $NumberOfMachines.ToString(); }
				$CatalogInformation += @{Data = "Allocated machines"; Value = $Catalog.UsedCount.ToString(); }
				$CatalogInformation += @{Data = "Set to VDA version"; Value = $xVDAVersion; }
				$CatalogInformation += @{Data = "Zone"; Value = $Catalog.ZoneName; }

				If($Null -ne $Machines)
				{
					If($Machines -is [array] -and $Machines.Count)
					{
						If([String]::IsNullOrEmpty($Machines[0].AgentVersion))
						{
							$CatalogInformation += @{Data = "Installed VDA version"; Value = "-"; }
							$CatalogInformation += @{Data = "Operating System"; Value = "-"; }
						}
						Else
						{
							$CatalogInformation += @{Data = "Installed VDA version"; Value = $Machines[0].AgentVersion; }
							$CatalogInformation += @{Data = "Operating System"; Value = $Machines[0].OSType; }
						}
					}
					Else
					{
						$CatalogInformation += @{Data = "Installed VDA version"; Value = "Unable to retrieve details"; }
						$CatalogInformation += @{Data = "Operating System"; Value = "Unable to retrieve details"; }
					}
			}
				ElseIf($Null -eq $Machines)
				{
					$CatalogInformation += @{Data = "Installed VDA version"; Value = "Unable to retrieve details"; }
					$CatalogInformation += @{Data = "Operating System"; Value = "Unable to retrieve details"; }
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -ne $True)
			{
				$CatalogInformation += @{Data = "Description"; Value = $Catalog.Description; }
				$CatalogInformation += @{Data = "Machine type"; Value = $xCatalogType; }
				$CatalogInformation += @{Data = "No. of machines"; Value = $NumberOfMachines.ToString(); }
				$CatalogInformation += @{Data = "Allocated machines"; Value = $Catalog.UsedCount.ToString(); }
				$CatalogInformation += @{Data = "Allocation type"; Value = $xAllocationType; }
				$CatalogInformation += @{Data = "User data"; Value = $xPersistType; }
				$CatalogInformation += @{Data = "Provisioning method"; Value = $xProvisioningType; }
				$CatalogInformation += @{Data = "Set to VDA version"; Value = $xVDAVersion; }
				$CatalogInformation += @{Data = "Zone"; Value = $Catalog.ZoneName; }
				
				If($Null -ne $Machines)
				{
					If($Machines -is [array] -and $Machines.Count)
					{
						If([String]::IsNullOrEmpty($Machines[0].AgentVersion))
						{
							$CatalogInformation += @{Data = "Installed VDA version"; Value = "-"; }
							$CatalogInformation += @{Data = "Operating System"; Value = "-"; }
						}
						Else
						{
							$CatalogInformation += @{Data = "Installed VDA version"; Value = $Machines[0].AgentVersion; }
							$CatalogInformation += @{Data = "Operating System"; Value = $Machines[0].OSType; }
						}
					}
					Else
					{
						$CatalogInformation += @{Data = "Installed VDA version"; Value = "Unable to retrieve details"; }
						$CatalogInformation += @{Data = "Operating System"; Value = "Unable to retrieve details"; }
					}
				}
				ElseIf($Null -eq $Machines)
				{
					$CatalogInformation += @{Data = "Installed VDA version"; Value = "Unable to retrieve details"; }
					$CatalogInformation += @{Data = "Operating System"; Value = "Unable to retrieve details"; }
				}
			}
			
			If($SessionSupport -eq "MultiSession")
			{
				$itemKeys = $Catalog.MetadataMap.Keys

				ForEach( $itemKey in $itemKeys )
				{
					If($itemKey.StartsWith("Task"))
					{
						Continue
					}
					If($itemKey.StartsWith("Citrix_DesktopStudio_IdentityPoolUid"))
					{
						Continue
					}

					$value = $Catalog.MetadataMap[ $itemKey ]
					
					If($value -eq "")
					{
						$value = "Not set"
					}
					$CatalogInformation += @{Data = $itemKey; Value = $value; }
				}
			}
			
			$Table = AddWordTable -Hashtable $CatalogInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 225;
			$Table.Columns.Item(2).Width = 275;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 "Machine Catalog: $($Catalog.Name)"
			If($Catalog.ProvisioningType -eq "MCS")
			{
				Line 1 "Description`t`t`t`t: " $Catalog.Description
				Line 1 "Machine type`t`t`t`t: " $xCatalogType
				Line 1 "No. of machines`t`t`t`t: " $NumberOfMachines.ToString()
				Line 1 "Allocated machines`t`t`t: " $Catalog.UsedCount.ToString()
				Line 1 "Allocation type`t`t`t`t: " $xAllocationType
				Line 1 "User data`t`t`t`t: " $xPersistType
				Line 1 "Provisioning method`t`t`t: " $xProvisioningType
				Line 1 "Account naming scheme`t`t`t: " $IdentityNamingScheme
				Line 1 "Naming scheme type`t`t`t: " $IdentityNamingSchemeType
				Line 1 "AD Domain`t`t`t`t: " $IdentityDomain
				Line 1 "AD Location`t`t`t`t: " $IdentityOU
				Line 1 "Set to VDA version`t`t`t: " $xVDAVersion
				If($Null -ne $MachineData)
				{
					If( $MachineData.PSObject.Properties[ 'HostingUnitName' ] )
					{
						## GRL - The property 'HostingUnitName' cannot be found on this object. Verify that the property exists
						Line 1 "Resources`t`t`t`t: " $MachineData.HostingUnitName
					}
				}
				Line 1 "Zone`t`t`t`t`t: " $Catalog.ZoneName
				
				If($Null -ne $MachineData)
				{
					Line 1 "Master VM`t`t`t`t: " $MasterVM
					Line 1 "Disk Image`t`t`t`t: " $xDiskImage
					Line 1 "Virtual CPUs`t`t`t`t: " $MachineData.CpuCount
					Line 1 "Memory`t`t`t`t`t: " "$($MachineData.MemoryMB) MB"
					Line 1 "Hard disk`t`t`t`t: " "$($MachineData.DiskSize) GB"
				}
				ElseIf($Null -eq $MachineData)
				{
					Line 1 "Master VM`t`t`t`t: " $MasterVM
					Line 1 "Disk Image`t`t`t`t: " $xDiskImage
					Line 1 "Virtual CPUs`t`t`t`t: " "Unable to retrieve details"
					Line 1 "Memory`t`t`t`t`t: " "Unable to retrieve details"
					Line 1 "Hard disk`t`t`t`t: " "Unable to retrieve details"
				}
				
				If(($Catalog.MinimumFunctionalLevel -eq "L7_9" -or 
					$Catalog.MinimumFunctionalLevel -eq "L7_20" -or 
					$Catalog.MinimumFunctionalLevel -eq "L7_25") -and 
				(($xAllocationType -eq "Random") -or 
				($xAllocationType -eq "Permanent" -and $xPersistType -eq "Discard" )))
				{
					Line 1 "Temporary memory cache size`t`t: " "$($TempMemoryCacheSize) MB"
					Line 1 "Temporary disk cache size`t`t: " "$($MachineData.WriteBackCacheDiskSize) GB"
				}

				If(($Catalog.MinimumFunctionalLevel -eq "L7_9" -or 
					$Catalog.MinimumFunctionalLevel -eq "L7_20" -or 
					$Catalog.MinimumFunctionalLevel -eq "L7_25") -and 
				($xAllocationType -eq "Permanent" -and $xPersistType -eq "On local disk" ) -and 
				((Get-ConfigEnabledFeature -EA 0) -contains "DedicatedFullDiskClone"))
				{
					Line 1 "VM copy mode`t`t`t`t: " $VMCopyMode
				}
				
				If($Null -ne $Machines)
				{
					If($Machines -is [array] -and $Machines.Count)
					{
						If([String]::IsNullOrEmpty($Machines[0].AgentVersion))
						{
							Line 1 "Installed VDA version`t`t`t: " "-"
							Line 1 "Operating System`t`t`t: " "-"
						}
						Else
						{
							Line 1 "Installed VDA version`t`t`t: " $Machines[0].AgentVersion
							Line 1 "Operating System`t`t`t: " $Machines[0].OSType
						}
					}
					Else
					{
						Line 1 "Installed VDA version`t`t`t: " "Unable to retrieve details"
						Line 1 "Operating System`t`t`t: " "Unable to retrieve details"
					}
				}
				ElseIf($Null -eq $Machines)
				{
					Line 1 "Installed VDA version`t`t`t: " "Unable to retrieve details"
					Line 1 "Operating System`t`t`t: " "Unable to retrieve details"
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "PVS")
			{
				Line 1 "Description`t`t`t`t: " $Catalog.Description
				Line 1 "Machine type`t`t`t`t: " $xCatalogType
				Line 1 "Provisioning method`t`t`t: " $xProvisioningType
				Line 1 "PVS address`t`t`t`t: " $Catalog.PvsAddress
				Line 1 "Allocation type`t`t`t`t: " $xAllocationType
				Line 1 "Set to VDA version`t`t`t: " $xVDAVersion
				Line 1 "Zone`t`t`t`t`t: " $Catalog.ZoneName
				
				If($Null -ne $Machines)
				{
					If($Machines -is [array] -and $Machines.Count)
					{
						If([String]::IsNullOrEmpty($Machines[0].AgentVersion))
						{
							Line 1 "Installed VDA version`t`t`t: " "-"
							Line 1 "Operating System`t`t`t: " "-"
						}
						Else
						{
							Line 1 "Installed VDA version`t`t`t: " $Machines[0].AgentVersion
							Line 1 "Operating System`t`t`t: " $Machines[0].OSType
						}
					}
					Else
					{
						Line 1 "Installed VDA version`t`t`t: " "Unable to retrieve details"
						Line 1 "Operating System`t`t`t: " "Unable to retrieve details"
					}
				}
				ElseIf($Null -eq $Machines)
				{
					Line 1 "Installed VDA version`t`t`t`t: " "Unable to retrieve details"
					Line 1 "Operating System`t`t`t: " "Unable to retrieve details"
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -eq $True)
			{
				Line 1 "Description`t`t`t`t: " $Catalog.Description

				If($RemotePCAccounts -is [array])
				{
					ForEach($RemotePCAccount in $RemotePCAccounts)
					{
						Line 1 "Organizational Units`t`t`t: " $RemotePCAccount.OU
						Line 2 "Allow subfolder matches`t`t: " $RemotePCAccount.AllowSubfolderMatches.ToString()
						If($RemotePCAccount.MachinesExcluded.Count -eq 0)
						{
							Line 2 "Machines excluded`t`t: " "None"
						}
						Else
						{
							$cnt = -1
							ForEach($Item in $RemotePCAccount.MachinesExcluded)
							{
								$cnt++
								
								If($cnt -eq 0)
								{
									Line 2 "Machines excluded`t`t: " $Item
								}
								Else
								{
									Line 6 "  " $Item
								}
							}
						}

						If($RemotePCAccount.MachinesIncluded -eq "*")
						{
							Line 2 "Machines Included`t`t: " "All"
						}
						Else
						{
							$cnt = -1
							ForEach($Item in $RemotePCAccount.MachinesIncluded)
							{
								$cnt++
								
								If($cnt -eq 0)
								{
									Line 2 "Machines Included`t`t: " $Item
								}
								Else
								{
									Line 6 "  " $Item
								}
							}
						}
					}
				}
				Else
				{
					Line 1 "Organizational Units`t`t`t: " $RemotePCOU
					Line 2 "Allow subfolder matches`t`t: " $RemotePCSubOU
				}

				Line 1 "Machine type`t`t`t`t: " $xCatalogType
				Line 1 "No. of machines`t`t`t`t: "$NumberOfMachines.ToString()
				Line 1 "Allocated machines`t`t`t: " $Catalog.UsedCount.ToString()
				Line 1 "Set to VDA version`t`t`t: " $xVDAVersion
				Line 1 "Zone`t`t`t`t`t: " $Catalog.ZoneName
				
				If($Null -ne $Machines)
				{
					If($Machines -is [array] -and $Machines.Count)
					{
						If([String]::IsNullOrEmpty($Machines[0].AgentVersion))
						{
							Line 1 "Installed VDA version`t`t`t: " "-"
							Line 1 "Operating System`t`t`t: " "-"
						}
						Else
						{
							Line 1 "Installed VDA version`t`t`t: " $Machines[0].AgentVersion
							Line 1 "Operating System`t`t`t: " $Machines[0].OSType
						}
					}
					Else
					{
						Line 1 "Installed VDA version`t`t`t: " "Unable to retrieve details"
						Line 1 "Operating System`t`t`t: " "Unable to retrieve details"
					}
				}
				ElseIf($Null -eq $Machines)
				{
					Line 1 "Installed VDA version`t`t`t`t: " "Unable to retrieve details"
					Line 1 "Operating System`t`t`t: " "Unable to retrieve details"
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -ne $True)
			{
				Line 1 "Description`t`t`t`t: " $Catalog.Description
				Line 1 "Machine type`t`t`t`t: " $xCatalogType
				Line 1 "No. of machines`t`t`t`t: "$NumberOfMachines.ToString()
				Line 1 "Allocated machines`t`t`t: " $Catalog.UsedCount.ToString()
				Line 1 "Allocation type`t`t`t`t: " $xAllocationType
				Line 1 "User data`t`t`t`t: " $xPersistType
				Line 1 "Provisioning method`t`t`t: " $xProvisioningType
				Line 1 "Set to VDA version`t`t`t: " $xVDAVersion
				Line 1 "Zone`t`t`t`t`t: " $Catalog.ZoneName
				
				If($Null -ne $Machines)
				{
					If($Machines -is [array] -and $Machines.Count)
					{
						If([String]::IsNullOrEmpty($Machines[0].AgentVersion))
						{
							Line 1 "Installed VDA version`t`t`t: " "-"
							Line 1 "Operating System`t`t`t: " "-"
						}
						Else
						{
							Line 1 "Installed VDA version`t`t`t: " $Machines[0].AgentVersion
							Line 1 "Operating System`t`t`t: " $Machines[0].OSType
						}
					}
					Else
					{
						Line 1 "Installed VDA version`t`t`t: " "Unable to retrieve details"
						Line 1 "Operating System`t`t`t: " "Unable to retrieve details"
					}
				}
				ElseIf($Null -eq $Machines)
				{
					Line 1 "Installed VDA version`t`t`t: " "Unable to retrieve details"
					Line 1 "Operating System`t`t`t: " "Unable to retrieve details"
				}
			}

			If($SessionSupport -eq "MultiSession")
			{
				$itemKeys = $Catalog.MetadataMap.Keys

				ForEach( $itemKey in $itemKeys )
				{
					If($itemKey.StartsWith("Task"))
					{
						Continue
					}
					If($itemKey.StartsWith("Citrix_DesktopStudio_IdentityPoolUid"))
					{
						Continue
					}

					$value = $Catalog.MetadataMap[ $itemKey ]
					
					If($value -eq "")
					{
						$value = "Not set"
					}
					Line 1 "$($itemKey)`t: " $value
				}
			}

			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "Machine Catalog: $($Catalog.Name)"
			$rowdata = @()
			$columnHeaders = @("Machine type",($global:htmlsb),$xCatalogType,$htmlwhite)
			If($Catalog.ProvisioningType -eq "MCS")
			{
				$rowdata += @(,('Description',($global:htmlsb),$Catalog.Description,$htmlwhite))
				$rowdata += @(,('Machine Type',($global:htmlsb),$xCatalogType,$htmlwhite))
				$rowdata += @(,('No. of machines',($global:htmlsb),$NumberOfMachines.ToString(),$htmlwhite))
				$rowdata += @(,('Allocated machines',($global:htmlsb),$Catalog.UsedCount.ToString(),$htmlwhite))
				$rowdata += @(,('Allocation type',($global:htmlsb),$xAllocationType,$htmlwhite))
				$rowdata += @(,('User data',($global:htmlsb),$xPersistType,$htmlwhite))
				$rowdata += @(,('Provisioning method',($global:htmlsb),$xProvisioningType,$htmlwhite))
				$rowdata += @(,("Account naming scheme",($global:htmlsb),$IdentityNamingScheme,$htmlwhite))
				$rowdata += @(,("Naming scheme type",($global:htmlsb),$IdentityNamingSchemeType,$htmlwhite))
				$rowdata += @(,("AD Domain",($global:htmlsb),$IdentityDomain,$htmlwhite))
				$rowdata += @(,("AD Location",($global:htmlsb),$IdentityOU,$htmlwhite))
				$rowdata += @(,('Set to VDA version',($global:htmlsb),$xVDAVersion,$htmlwhite))
				If($Null -ne $MachineData)
				{
					If( $MachineData.PSObject.Properties[ 'HostingUnitName' ] )
					{
						## GRL - The property 'HostingUnitName' cannot be found on this object. Verify that the property exists
						$rowdata += @(,('Resources',($global:htmlsb),$MachineData.HostingUnitName,$htmlwhite))
					}
				}
				$rowdata += @(,('Zone',($global:htmlsb),$Catalog.ZoneName,$htmlwhite))
				
				If($Null -ne $MachineData)
				{
					$rowdata += @(,('Master VM',($global:htmlsb),$MasterVM,$htmlwhite))
					$rowdata += @(,('Disk Image',($global:htmlsb),$xDiskImage,$htmlwhite))
					$rowdata += @(,('Virtual CPUs',($global:htmlsb),$MachineData.CpuCount,$htmlwhite))
					$rowdata += @(,('Memory',($global:htmlsb),"$($MachineData.MemoryMB) MB",$htmlwhite))
					$rowdata += @(,('Hard disk',($global:htmlsb),"$($MachineData.DiskSize) GB",$htmlwhite))
				}
				ElseIf($Null -eq $MachineData)
				{
					$rowdata += @(,('Master VM',($global:htmlsb),$MasterVM,$htmlwhite))
					$rowdata += @(,('Disk Image',($global:htmlsb),$xDiskImage,$htmlwhite))
					$rowdata += @(,('Virtual CPUs',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
					$rowdata += @(,('Memory',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
					$rowdata += @(,('Hard disk',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
				}
				
				If(($Catalog.MinimumFunctionalLevel -eq "L7_9" -or 
					$Catalog.MinimumFunctionalLevel -eq "L7_20" -or 
					$Catalog.MinimumFunctionalLevel -eq "L7_25") -and 
				(($xAllocationType -eq "Random") -or 
				($xAllocationType -eq "Permanent" -and $xPersistType -eq "Discard" )))
				{
					$rowdata += @(,('Temporary memory cache size',($global:htmlsb),"$($TempMemoryCacheSize) MB",$htmlwhite))
					$rowdata += @(,('Temporary disk cache size',($global:htmlsb), "$($TempDiskCacheSize) GB",$htmlwhite))
				}

				If(($Catalog.MinimumFunctionalLevel -eq "L7_9" -or 
					$Catalog.MinimumFunctionalLevel -eq "L7_20" -or 
					$Catalog.MinimumFunctionalLevel -eq "L7_25") -and 
				($xAllocationType -eq "Permanent" -and $xPersistType -eq "On local disk" ) -and 
				((Get-ConfigEnabledFeature -EA 0) -contains "DedicatedFullDiskClone"))
				{
					$rowdata += @(,('VM copy mode',($global:htmlsb),$VMCopyMode,$htmlwhite))
				}
				
				If($Null -ne $Machines)
				{
					If($Machines -is [array] -and $Machines.Count)
					{
						If([String]::IsNullOrEmpty($Machines[0].AgentVersion))
						{
							$rowdata += @(,('Installed VDA version',($global:htmlsb),"-",$htmlwhite))
							$rowdata += @(,('Operating System',($global:htmlsb),"-",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('Installed VDA version',($global:htmlsb),$Machines[0].AgentVersion,$htmlwhite))
							$rowdata += @(,('Operating System',($global:htmlsb),$Machines[0].OSType,$htmlwhite))
						}
					}
					Else
					{
						$rowdata += @(,('Installed VDA version',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
						$rowdata += @(,('Operating System',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
					}
				}
				ElseIf($Null -eq $Machines)
				{
					$rowdata += @(,('Installed VDA version',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
					$rowdata += @(,('Operating System',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "PVS")
			{
				$rowdata += @(,('Description',($global:htmlsb),$Catalog.Description,$htmlwhite))
				$rowdata += @(,('Machine Type',($global:htmlsb),$xCatalogType,$htmlwhite))
				$rowdata += @(,('Provisioning method',($global:htmlsb),$xProvisioningType,$htmlwhite))
				$rowdata += @(,('PVS address',($global:htmlsb),$Catalog.PvsAddress,$htmlwhite))
				$rowdata += @(,('Allocation type',($global:htmlsb),$xAllocationType,$htmlwhite))
				$rowdata += @(,('Set to VDA version',($global:htmlsb),$xVDAVersion,$htmlwhite))
				$rowdata += @(,('Zone',($global:htmlsb),$Catalog.ZoneName,$htmlwhite))
				
				If($Null -ne $Machines)
				{
					If($Machines -is [array] -and $Machines.Count)
					{
						If([String]::IsNullOrEmpty($Machines[0].AgentVersion))
						{
							$rowdata += @(,('Installed VDA version',($global:htmlsb),"-",$htmlwhite))
							$rowdata += @(,('Operating System',($global:htmlsb),"-",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('Installed VDA version',($global:htmlsb),$Machines[0].AgentVersion,$htmlwhite))
							$rowdata += @(,('Operating System',($global:htmlsb),$Machines[0].OSType,$htmlwhite))
						}
					}
					Else
					{
						$rowdata += @(,('Installed VDA version',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
						$rowdata += @(,('Operating System',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
					}
				}
				ElseIf($Null -eq $Machines)
				{
					$rowdata += @(,('Installed VDA version',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
					$rowdata += @(,('Operating System',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -eq $True)
			{
				$rowdata += @(,('Description',($global:htmlsb),$Catalog.Description,$htmlwhite))

				If($RemotePCAccounts -is [array])
				{
					ForEach($RemotePCAccount in $RemotePCAccounts)
					{
						$rowdata += @(,("Organizational Units",($global:htmlsb),$RemotePCAccount.OU,$htmlwhite))
						$rowdata += @(,("     Allow subfolder matches",($global:htmlsb),$RemotePCAccount.AllowSubfolderMatches.ToString(),$htmlwhite))
						If($RemotePCAccount.MachinesExcluded.Count -eq 0)
						{
							$rowdata += @(,("     Machines excluded",($global:htmlsb),"None",$htmlwhite))
						}
						Else
						{
							$cnt = -1
							ForEach($Item in $RemotePCAccount.MachinesExcluded)
							{
								$cnt++
								
								If($cnt -eq 0)
								{
									$rowdata += @(,("     Machines excluded",($global:htmlsb),$Item,$htmlwhite))
								}
								Else
								{
									$rowdata += @(,("",($global:htmlsb),$Item,$htmlwhite))
								}
							}
						}

						If($RemotePCAccount.MachinesIncluded -eq "*")
						{
							$rowdata += @(,("     Machines Included",($global:htmlsb),"All",$htmlwhite))
						}
						Else
						{
							$cnt = -1
							ForEach($Item in $RemotePCAccount.MachinesIncluded)
							{
								$cnt++
								
								If($cnt -eq 0)
								{
									$rowdata += @(,("     Machines Included",($global:htmlsb),$Item,$htmlwhite))
								}
								Else
								{
									$rowdata += @(,("",($global:htmlsb),$Item,$htmlwhite))
								}
							}
						}
					}
				}
				Else
				{
					$rowdata += @(,("Organizational Units",($global:htmlsb),$RemotePCOU,$htmlwhite))
					$rowdata += @(,("     Allow subfolder matches",($global:htmlsb),$RemotePCSubOU,$htmlwhite))
				}

				$rowdata += @(,('Machine Type',($global:htmlsb),$xCatalogType,$htmlwhite))
				$rowdata += @(,('No. of machines',($global:htmlsb),$NumberOfMachines.ToString(),$htmlwhite))
				$rowdata += @(,('Allocated machines',($global:htmlsb),$Catalog.UsedCount.ToString(),$htmlwhite))
				$rowdata += @(,('Set to VDA version',($global:htmlsb),$xVDAVersion,$htmlwhite))
				$rowdata += @(,('Zone',($global:htmlsb),$Catalog.ZoneName,$htmlwhite))
				
				If($Null -ne $Machines)
				{
					If($Machines -is [array] -and $Machines.Count)
					{
						If([String]::IsNullOrEmpty($Machines[0].AgentVersion))
						{
							$rowdata += @(,('Installed VDA version',($global:htmlsb),"-",$htmlwhite))
							$rowdata += @(,('Operating System',($global:htmlsb),"-",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('Installed VDA version',($global:htmlsb),$Machines[0].AgentVersion,$htmlwhite))
							$rowdata += @(,('Operating System',($global:htmlsb),$Machines[0].OSType,$htmlwhite))
						}
					}
					Else
					{
						$rowdata += @(,('Installed VDA version',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
						$rowdata += @(,('Operating System',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
					}
				}
				ElseIf($Null -eq $Machines)
				{
					$rowdata += @(,('Installed VDA version',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
					$rowdata += @(,('Operating System',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -ne $True)
			{
				$rowdata += @(,('Description',($global:htmlsb),$Catalog.Description,$htmlwhite))
				$rowdata += @(,('Machine Type',($global:htmlsb),$xCatalogType,$htmlwhite))
				$rowdata += @(,('No. of machines',($global:htmlsb),$NumberOfMachines.ToString(),$htmlwhite))
				$rowdata += @(,('Allocated machines',($global:htmlsb),$Catalog.UsedCount.ToString(),$htmlwhite))
				$rowdata += @(,('Allocation type',($global:htmlsb),$xAllocationType,$htmlwhite))
				$rowdata += @(,('User data',($global:htmlsb),$xPersistType,$htmlwhite))
				$rowdata += @(,('Provisioning method',($global:htmlsb),$xProvisioningType,$htmlwhite))
				$rowdata += @(,('Set to VDA version',($global:htmlsb),$xVDAVersion,$htmlwhite))
				$rowdata += @(,('Zone',($global:htmlsb),$Catalog.ZoneName,$htmlwhite))
				
				If($Null -ne $Machines)
				{
					If($Machines -is [array] -and $Machines.Count)
					{
						If([String]::IsNullOrEmpty($Machines[0].AgentVersion))
						{
							$rowdata += @(,('Installed VDA version',($global:htmlsb),"-",$htmlwhite))
							$rowdata += @(,('Operating System',($global:htmlsb),"-",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('Installed VDA version',($global:htmlsb),$Machines[0].AgentVersion,$htmlwhite))
							$rowdata += @(,('Operating System',($global:htmlsb),$Machines[0].OSType,$htmlwhite))
						}
					}
					Else
					{
						$rowdata += @(,('Installed VDA version',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
						$rowdata += @(,('Operating System',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
					}
				}
				ElseIf($Null -eq $Machines)
				{
					$rowdata += @(,('Installed VDA version',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
					$rowdata += @(,('Operating System',($global:htmlsb),"Unable to retrieve details",$htmlwhite))
				}
			}
			
			If($Catalog.SessionSupport -eq "MultiSession")
			{
				$itemKeys = $Catalog.MetadataMap.Keys

				ForEach( $itemKey in $itemKeys )
				{
					If($itemKey.StartsWith("Task"))
					{
						Continue
					}
					If($itemKey.StartsWith("Citrix_DesktopStudio_IdentityPoolUid"))
					{
						Continue
					}
					
					$value = $Catalog.MetadataMap[ $itemKey ]
					
					If($value -eq "")
					{
						$value = "Not set"
					}
					$rowdata += @(,($itemKey,($global:htmlsb),$value,$htmlwhite))
				}
			}

			$msg = ""
			$columnWidths = @("200","500")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "700"
		}
			
		#scopes
		$Scopes = (Get-BrokerCatalog -Name $Catalog.Name @CCParams2).Scopes
		
		If($? -and ($Null -eq $Scopes))
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Scopes"
				[System.Collections.Hashtable[]] $ScopesWordTable = @();

				$ScopesWordTable += @{Scope = "All";}

				$Table = AddWordTable -Hashtable $ScopesWordTable `
				-Columns Scope `
				-Headers  "Scopes" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 225;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 1 "Scopes"
				Line 2 "All"
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "Scopes"
				$rowdata = @()
				$rowdata += @(,("All",$htmlwhite))

				$columnHeaders = @(
				'Scopes',($global:htmlsb))

				$msg = ""
				$columnWidths = @("225")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "225"
			}
		}
		ElseIf($? -and ($Null -ne $Scopes))
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Scopes"
				[System.Collections.Hashtable[]] $ScopesWordTable = @();

				$ScopesWordTable += @{Scope = "All";}

				ForEach($Scope in $Scopes)
				{
					$ScopesWordTable += @{Scope = $Scope.ScopeName;}
				}

				$Table = AddWordTable -Hashtable $ScopesWordTable `
				-Columns Scope `
				-Headers  "Scopes" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 225;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 1 "Scopes"
				Line 2 "All"

				ForEach($Scope in $Scopes)
				{
					Line 2 $Scope.ScopeName;
				}
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "Scopes"
				$rowdata = @()
				$rowdata += @(,("All",$htmlwhite))

				ForEach($Scope in $Scopes)
				{
					$rowdata += @(,($Scope.ScopeName,$htmlwhite))
				}
				$columnHeaders = @(
				'Scopes',($global:htmlsb))

				$msg = ""
				$columnWidths = @("225")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "225"
			}
		}
		Else
		{
			$txt = "Unable to retrieve Scopes for Machine Catalog $($Catalog.Name)"
			OutputWarning $txt
		}
		
		If($MachineCatalogs)
		{
			If($Null -ne $Machines)
			{
				Write-Verbose "$(Get-Date -Format G): `t`tProcessing Machines in $($Catalog.Name)"
				
				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "Machines"
					[System.Collections.Hashtable[]] $MachinesWordTable = @()
				}
				If($Text)
				{
					Line 1 "Machines"
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "Machines"
					$rowdata = @()
				}
				
				ForEach($Machine in $Machines)
				{
					If($MSWord -or $PDF)
					{
						$MachinesWordTable += @{
						MachineName = $Machine.MachineName;
						}
					}
					If($Text)
					{
						Line 2 $Machine.MachineName
					}
					If($HTML)
					{
						$rowdata += @(,($Machine.MachineName,$htmlwhite))
					}
				}
				
				If($MSWord -or $PDF)
				{
					If($MachinesWordTable.Count -eq 0)
					{
						$MachinesWordTable += @{
						MachineName = "None found";
						}
					}
					
					$Table = AddWordTable -Hashtable $MachinesWordTable `
					-Columns MachineName `
					-Headers "Machine Names" `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed

					SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 225;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 0 ""
				}
				If($HTML)
				{
					$columnHeaders = @(
					'Machine Names',($global:htmlsb))

					$msg = ""
					$columnWidths = @("225")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "225"
				}
				
				Write-Verbose "$(Get-Date -Format G): `t`tProcessing administrators for Machines in $($Catalog.Name)"
				$Admins = GetAdmins "Catalog" $Catalog.Name
				
				If($? -and ($Null -ne $Admins))
				{
					OutputAdminsForDetails $Admins
				}
				ElseIf($? -and ($Null -eq $Admins))
				{
					$txt = "There are no administrators for Machines in $($Catalog.Name)"
					OutputNotice $txt
				}
				Else
				{
					$txt = "Unable to retrieve administrators for Machines in $($Catalog.Name)"
					OutputWarning $txt
				}
				
				ForEach($Machine in $Machines)
				{
					OutputMachineDetails $Machine
				}
			}
		}
	}
}
#endregion

#region function to output machine/desktop details
Function GetVDARegistryKeys
{
	Param([string]$ComputerName, [string]$xType)

	#Get-VDARegKeyToObject "HKLM:\" "" $ComputerName

	If($xType -eq "Server")
	{
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\Graphics" "BTLLossyThreshold" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Services\CtxDNDSvc" "Enabled" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\AppV" "Features" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\Reconnect" "DisableGPCalculation" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\Reconnect" "FastReconnect" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\UniversalPrintDrivers\PDF" "EnablePostscriptSimulation" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\UniversalPrintDrivers\PDF" "EnableFullFontEmbedding" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Wow6432Node\Citrix\ICAClient\Engine\Configuration\Advanced\Modules\GfxRender" "UseDirect3D" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Wow6432Node\Citrix\ICAClient\Engine\Configuration\Advanced\Modules\GfxRender" "PresentDevice" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Wow6432Node\Citrix\ICAClient\Engine\Configuration\Advanced\Modules\GfxRender" "MaxNumRefFrames" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\System\Currentcontrolset\services\picadm\Parameters" "DisableFullStreamWrite" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\SCMConfig" "EnableSvchostMitigationPolicy" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\ICA" "DisableAppendMouse" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\Citrix Virtual Desktop Agent" "DisableLogonUISuppression" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\SmartCard" "EnableSCardHookVcResponseTimeout" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\Citrix Virtual Desktop Agent" "DisableLogonUISuppressionForSmartCardPublishedApps" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\HDX3D\BitmapRemotingConfig" "EnableDDAPICursor" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\VirtualDesktopAgent" "SupportMultipleForest" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\VirtualDesktopAgent" "ListOfSIDs" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix" "CtxKlMap" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\TerminalServer" "fSingleSessionPerUser" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\ICAClient\GenericUSB" "EnableBloombergHID" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\ICAClient\Engine\Configuration\Advanced\Modules\ClientAudio" "EchoCancellation" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Wow6432Node\Citrix\ICAClient\Engine\Configuration\Advanced\Modules\ClientAudio" "EchoCancellation" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\ica-tcp\AudioConfig" "MaxPolicyAge" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\ica-tcp\AudioConfig" "PolicyTimeout" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\Ica\Thinwire" "EnableDrvTw2NotifyMonitorOrigin" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\Citrix" "EnableVisualEffect" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\StreamingHook" "EnableReadImageFileExecOptionsExclusionList" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Wow6432Node\Citrix\StreamingHook" "EnableReadImageFileExecOptionsExclusionList" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\Citrix\wfshell\TWI" "ApplicationLaunchWaitTimeoutMS" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\Citrix\wfshell\TWI" "LogoffCheckerStartupDelayInSeconds" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\Citrix\wfshell\TWI" "LogoffCheckSysModules" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\Citrix\wfshell\TWI" "SeamlessFlags" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\Citrix\wfshell\TWI" "WorkerWaitInterval" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\Citrix\wfshell\TWI" "WorkerFullCheckInterval" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\Citrix\wfshell\TWI" "AAHookFlags" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\CtxHook\AppInit_Dlls\SHAppBarHook" "FilePathName" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\CtxHook\AppInit_Dlls\SHAppBarHook" "Flag" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\CtxHook\AppInit_Dlls\SHAppBarHook" "Settings" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\CtxHook" "ExcludedImageNames" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Wow6432Node\Citrix\CtxHook" "ExcludedImageNames" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\services\CtxUvi" "UviProcessExcludes" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Services\CtxUvi" "UviEnabled" $ComputerName $xType
	}
	ElseIf($xType -eq "Desktop")
	{
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\Graphics" "BTLLossyThreshold" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Services\CtxDNDSvc" "Enabled" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\AppV" "Features" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\Reconnect" "DisableGPCalculation" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\Reconnect" "FastReconnect" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\UniversalPrintDrivers\PDF" "EnablePostscriptSimulation" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\UniversalPrintDrivers\PDF" "EnableFullFontEmbedding" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Wow6432Node\Citrix\ICAClient\Engine\Configuration\Advanced\Modules\GfxRender" "UseDirect3D" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Wow6432Node\Citrix\ICAClient\Engine\Configuration\Advanced\Modules\GfxRender" "PresentDevice" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Wow6432Node\Citrix\ICAClient\Engine\Configuration\Advanced\Modules\GfxRender" "MaxNumRefFrames" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\Citrix Virtual Desktop Agent" "DisableLogonUISuppression" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\SmartCard" "EnableSCardHookVcResponseTimeout" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\ICA" "DisableAppendMouse" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\Audio" "CleanMappingWhenDisconnect" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\HDX3D\BitmapRemotingConfig" "EnableDDAPICursor" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\VirtualDesktopAgent" "SupportMultipleForest" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\VirtualDesktopAgent" "ListOfSIDs" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix" "CtxKlMap" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\PortICA" "DisableRemotePCSleepPreventer" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\PortICA\RemotePC" "RpcaMode" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\PortICA\RemotePC" "RpcaTimeout" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\Software\Citrix\DesktopServer" "AllowMultipleRemotePCAssignments" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\ICAClient\GenericUSB" "EnableBloombergHID" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\ICAClient\Engine\Configuration\Advanced\Modules\ClientAudio" "EchoCancellation" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Wow6432Node\Citrix\ICAClient\Engine\Configuration\Advanced\Modules\ClientAudio" "EchoCancellation" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\HDX3D\BitmapRemotingConfig" "HKLM_DisableMontereyFBCOnInit" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\Ica\Thinwire" "EnableDrvTw2NotifyMonitorOrigin" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Control\Citrix\wfshell\TWI" "LogoffCheckSysModules" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Citrix\CtxHook" "ExcludedImageNames" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SOFTWARE\Wow6432Node\Citrix\CtxHook" "ExcludedImageNames" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\services\CtxUvi" "UviProcessExcludes" $ComputerName $xType
		Get-VDARegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Services\CtxUvi" "UviEnabled" $ComputerName $xType
	}
}

Function Get-VDARegKeyToObject 
{
	#function contributed by Andrew Williamson @ Fujitsu Services
    param([string]$RegPath,
    [string]$RegKey,
    [string]$ComputerName,
	[string]$xType)
	
    $val = Get-RegistryValue2 $RegPath $RegKey $ComputerName
	
    If($Null -eq $val) 
	{
        $tmp = "Not set"
    } 
	Else 
	{
	    $tmp = $val
    }
	$obj1 = [PSCustomObject] @{
		RegKey       = $RegPath	
		RegValue     = $RegKey	
		VDAType      = $xType	
		ComputerName = $ComputerName	
		Value        = $tmp
	}
	$null = $Script:VDARegistryItems.Add($obj1)
}

Function OutputMachineDetails
{
	Param([object] $Machine)
	
	#if HostedMachineName is empty, like for RemotePC and unregistered machines, use the first part of DNSName
	$tmp = $Machine.DNSName.Split(".")
	$xMachineName = $tmp[0]
	$tmp = $Null
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Machine $xMachineName"
	
	#first see if VDA is Linux
	If($Machine.OSType -Like "*linux*")
	{
		#Linux VDAs do not have an easily accessible registry so skip them
		$LinuxVDA = $True
	}
	Else
	{
		$LinuxVDA = $False
		If($VDARegistryKeys)
		{
			Write-Verbose "$(Get-Date -Format G): `t`t`tTesting $($xMachineName)"
			$MachineIsOnline = $False
			
			If(Resolve-DnsName -Name $xMachineName -EA 0 4>$Null)
			{
				$results = Test-NetConnection -ComputerName $xMachineName -InformationLevel Quiet -EA 0 3>$Null
				If($results)
				{
					Write-Verbose "$(Get-Date -Format G): `t`t`t`t$($xMachineName) is online"
					$MachineIsOnline = $True
				}
				Else
				{
					Write-Verbose "$(Get-Date -Format G): `t`t`t`t$($xMachineName) is offline. VDA Registry Key data cannot be gathered."
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): `t`t`t`t$($xMachineName) was not found in DNS. VDA Registry Key data cannot be gathered."
			}
		}
	}
	
	$xAssociatedUserFullNames = @()
	ForEach($Value in $Machine.AssociatedUserFullNames)
	{
		$xAssociatedUserFullNames += "$($Value)"
	}
		
	$xAssociatedUserNames = @()
	ForEach($Value in $Machine.AssociatedUserNames)
	{
		$xAssociatedUserNames += "$($Value)"
	}
	
	$xAssociatedUserUPNs = @()
	ForEach($Value in $Machine.AssociatedUserUPNs)
	{
		$xAssociatedUserUPNs += "$($Value)"
	}

	$xDesktopConditions = @()
	ForEach($Value in $Machine.DesktopConditions)
	{
		$xDesktopConditions += "$($Value)"
	}
	
	If($xDesktopConditions.Count -eq 0)
	{
		$xDesktopConditions += "-"
	}

	$xAllocationType = ""
	If($Machine.AllocationType -eq "Static")
	{
		$xAllocationType = "Private"
	}
	Else
	{
		$xAllocationType = $Machine.AllocationType
	}

	$xInMaintenanceMode = ""
	If($Machine.InMaintenanceMode)
	{
		$xInMaintenanceMode = "On"
	}
	Else
	{
		$xInMaintenanceMode ="Off"
	}

	$xWindowsConnectionSetting = ""
	If($Machine.SessionSupport -eq "MultiSession")
	{
		Switch ($Machine.WindowsConnectionSetting)
		{
			"LogonEnabled"			{$xWindowsConnectionSetting = "Logon Enabled"}
			"Draining"				{$xWindowsConnectionSetting = "Draining"}
			"DrainingUntilRestart"	{$xWindowsConnectionSetting = "Draining until restart"}
			"LogonDisabled"			{$xWindowsConnectionSetting = "Logons Disabled"}
			Default					{$xWindowsConnectionSetting = "Unable to determine WindowsConnectionSetting: $($Machine.WindowsConnectionSetting)"; Break}
		}
	}

	$xIsPhysical = ""
	If($Machine.IsPhysical)
	{
		$xIsPhysical = "Physical"
	}
	Else
	{
		$xIsPhysical ="Virtual"
	}

	$xSummaryState = ""
	If($Machine.SummaryState -eq "InUse")
	{
		$xSummaryState = "In Use"
	}
	Else
	{
		$xSummaryState = $Machine.SummaryState.ToString()
	}

	$xTags = @()
	ForEach($Value in $Machine.Tags)
	{
		$xTags += "$($Value)"
	}
	
	If($xTags.Count -eq 0)
	{
		$xTags += "-"
	}

	$xApplicationsInUse = @()
	ForEach($value in $Machine.ApplicationsInUse)
	{
		$xApplicationsInUse += "$($value)"
	}
	
	If($xApplicationsInUse.Count -eq 0)
	{
		$xApplicationsInUse += "-"
	}

	$xPublishedApplications = @()
	ForEach($value in $Machine.PublishedApplications)
	{
		$xPublishedApplications += "$($value)"
	}
	
	If($xPublishedApplications.Count -eq 0)
	{
		$xPublishedApplications += "-"
	}

	$xSessionSecureIcaActive = ""
	If($Machine.SessionSecureIcaActive)
	{
		$xSessionSecureIcaActive = "Yes"
	}
	Else
	{
		$xSessionSecureIcaActive = "-"
	}

	$xLastDeregistrationReason = ""
	Switch ($Machine.LastDeregistrationReason)
	{
		$Null									{$xLastDeregistrationReason = "-"; Break}
		"AgentAddressResolutionFailed"			{$xLastDeregistrationReason = "Agent Address Resolution Failed"; Break}
		"AgentNotContactable"					{$xLastDeregistrationReason = "Agent Not Contactable"; Break}
		"AgentRejectedSettingsUpdate"			{$xLastDeregistrationReason = "Agent Rejected Settings Update"; Break}
		"AgentRequested"						{$xLastDeregistrationReason = "Agent Requested"; Break}
		"AgentShutdown"							{$xLastDeregistrationReason = "Agent Shutdown"; Break}
		"AgentSuspended"						{$xLastDeregistrationReason = "Agent Suspended"; Break}
		"AgentWrongActiveDirectoryOU"			{$xLastDeregistrationReason = "Agent Wrong Active Directory OU"; Break}
		"BrokerRegistrationLimitReached"		{$xLastDeregistrationReason = "Broker Registration Limit Reached"; Break}
		"ContactLost"							{$xLastDeregistrationReason = "Contact Lost"; Break}
		"DesktopRemoved"						{$xLastDeregistrationReason = "Desktop Removed"; Break}
		"DesktopRestart"						{$xLastDeregistrationReason = "Desktop Restart"; Break}
		"EmptyRegistrationRequest"				{$xLastDeregistrationReason = "Empty Registration Request"; Break}
		"FunctionalLevelTooLowForCatalog"		{$xLastDeregistrationReason = "Functional Level Too Low For Catalog"; Break}
		"FunctionalLevelTooLowForDesktopGroup"	{$xLastDeregistrationReason = "Functional Level Too Low For Desktop Group"; Break}
		"IncompatibleVersion"					{$xLastDeregistrationReason = "Incompatible Version"; Break}
		"InconsistentRegistrationCapabilities"	{$xLastDeregistrationReason = "Inconsistent Registration Capabilities"; Break}
		"InvalidRegistrationRequest"			{$xLastDeregistrationReason = "Invalid Registration Request"; Break}
		"MissingAgentVersion"					{$xLastDeregistrationReason = "Missing Agent Version"; Break}
		"MissingRegistrationCapabilities"		{$xLastDeregistrationReason = "Missing Registration Capabilities"; Break}
		"NotLicensedForFeature"					{$xLastDeregistrationReason = "Not Licensed For Feature"; Break}
		"PowerOff"								{$xLastDeregistrationReason = "Power Off"; Break}
		"SendSettingsFailure"					{$xLastDeregistrationReason = "Send Settings Failure"; Break}
		"SessionAuditFailure"					{$xLastDeregistrationReason = "Session Audit Failure"; Break}
		"SessionPrepareFailure"					{$xLastDeregistrationReason = "Session Prepare Failure"; Break}
		"SettingsCreationFailure"				{$xLastDeregistrationReason = "Settings Creation Failure"; Break}
		"SingleMultiSessionMismatch"			{$xLastDeregistrationReason = "Single Multi Session Mismatch"; Break}
		"UnknownError"							{$xLastDeregistrationReason = "Unknown Error"; Break}
		"UnsupportedCredentialSecurityVersion"	{$xLastDeregistrationReason = "Unsupported Credential Security Version"; Break} 
		Default {$xLastDeregistrationReason = "Unable to determine LastDeregistrationReason: $($Machine.LastDeregistrationReason)"; Break}
	}

	$xPersistUserChanges = ""
	Switch ($Machine.PersistUserChanges)
	{
		"OnLocal"	{$xPersistUserChanges = "On Local"; Break}
		"Discard"	{$xPersistUserChanges = "Discard"; Break}
		Default		{$xPersistUserChanges = "Unable to determine the value of PersistUserChanges: $($Machine.PersistUserChanges)"; Break}
	}

	$xWillShutdownAfterUse = ""
	If($Machine.WillShutdownAfterUse)
	{
		$xWillShutdownAfterUse = "Yes"
	}
	Else
	{
		$xWillShutdownAfterUse = "No"
	}

	$xSessionSmartAccessTags = @()
	ForEach($value in $Machine.SessionSmartAccessTags)
	{
		$xSessionSmartAccessTags += "$($value)"
	}
	
	If($xSessionSmartAccessTags.Count -eq 0)
	{
		$xSessionSmartAccessTags += "-"
	}
	
	Switch($Machine.FaultState)
	{
		"None"			{$xMachineFaultState = "None"}
		"FailedToStart"	{$xMachineFaultState = "Failed to start"}
		"StuckOnBoot"	{$xMachineFaultState = "Stuck on boot"}
		"Unregistered"	{$xMachineFaultState = "Unregistered"}
		"MaxCapacity"	{$xMachineFaultState = "Maximum capacity"}
		Default			{$xMachineFaultState = "Unable to determine the value of FaultState: $($Machine.FaultState)"; Break}
	}

	If([String]::IsNullOrEmpty($Machine.SessionLaunchedViaHostName))
	{
		$xSessionLaunchedViaHostName = "-"
	}
	Else
	{
		$xSessionLaunchedViaHostName = $Machine.SessionLaunchedViaHostName
	}
	
	
	If([String]::IsNullOrEmpty($Machine.SessionLaunchedViaIP))
	{
		$xSessionLaunchedViaIP = "-"
	}
	Else
	{
		$xSessionLaunchedViaIP = $Machine.SessionLaunchedViaIP
	}
	
	If([String]::IsNullOrEmpty($Machine.SessionClientAddress))
	{
		$xSessionClientAddress = "-"
	}
	Else
	{
		$xSessionClientAddress = $Machine.SessionClientAddress
	}
	
	If([String]::IsNullOrEmpty($Machine.SessionClientName))
	{
		$xSessionClientName = "-"
	}
	Else
	{
		$xSessionClientName = $Machine.SessionClientName
	}
	
	If([String]::IsNullOrEmpty($Machine.SessionClientVersion))
	{
		$xSessionClientVersion = "-"
	}
	Else
	{
		$xSessionClientVersion = $Machine.SessionClientVersion
	}
	
	If([String]::IsNullOrEmpty($Machine.SessionConnectedViaHostName))
	{
		$xSessionConnectedViaHostName = "-"
	}
	Else
	{
		$xSessionConnectedViaHostName = $Machine.SessionConnectedViaHostName
	}
	
	If([String]::IsNullOrEmpty($Machine.SessionConnectedViaIP))
	{
		$xSessionConnectedViaIP = "-"
	}
	Else
	{
		$xSessionConnectedViaIP = $Machine.SessionConnectedViaIP
	}
	
	If([String]::IsNullOrEmpty($Machine.SessionProtocol))
	{
		$xSessionProtocol = "-"
	}
	Else
	{
		$xSessionProtocol = $Machine.SessionProtocol
	}
	
	If([String]::IsNullOrEmpty($Machine.SessionStateChangeTime))
	{
		$xSessionStateChangeTime = "-"
	}
	Else
	{
		$xSessionStateChangeTime = $Machine.SessionStateChangeTime
	}
	
	If([String]::IsNullOrEmpty($Machine.SessionState))
	{
		$xSessionState = "-"
	}
	Else
	{
		$xSessionState = $Machine.SessionState.ToString()
	}
	
	If([String]::IsNullOrEmpty($Machine.SessionUserName))
	{
		$xSessionUserName = "-"
	}
	Else
	{
		$xSessionUserName = $Machine.SessionUserName
	}
	
	If([String]::IsNullOrEmpty($Machine.LastConnectionTime))
	{
		$xLastConnectionTime = "-"
	}
	Else
	{
		$xLastConnectionTime = $Machine.LastConnectionTime.ToString()
	}
	
	If([String]::IsNullOrEmpty($Machine.LastConnectionUser))
	{
		$xLastConnectionUser = "-"
	}
	Else
	{
		$xLastConnectionUser = $Machine.LastConnectionUser
	}

	If([String]::IsNullOrEmpty($Machine.ControllerDNSName))
	{
		$xBroker = "-"
	}
	Else
	{
		$xBroker = $Machine.ControllerDNSName
	}

	If([String]::IsNullOrEmpty($Machine.HostingServerName))
	{
		$xHostingServerName = "-"
	}
	Else
	{
		$xHostingServerName = $Machine.HostingServerName
	}
	
	If([String]::IsNullOrEmpty($Machine.HostedMachineName))
	{
		$xHostedMachineName = "-"
	}
	Else
	{
		$xHostedMachineName = $Machine.HostedMachineName
	}

	If([String]::IsNullOrEmpty($Machine.HypervisorConnectionName))
	{
		$xHypervisorConnectionName = "-"
	}
	Else
	{
		$xHypervisorConnectionName = $Machine.HypervisorConnectionName
	}

	Switch ($Machine.PowerState)
	{
		"Off"			{$xPowerState = "Off"; Break}
		"On"			{$xPowerState = "On"; Break}
        "Resuming"		{$xPowerState = "Resuming"; Break}
		"Suspended"		{$xPowerState = "Suspended"; Break}
		"Suspending"	{$xPowerState = "Suspending"; Break}
		"TurningOff"	{$xPowerState = "Turning Off"; Break}
		"TurningOn"		{$xPowerState = "Turning On"; Break}
		"Unavailable"	{$xPowerState = "Unavailable"; Break}
		"Unknown"		{$xPowerState = "Unknown"; Break}
		"Unmanaged"		{$xPowerState = "Unmanaged"; Break}
		Default			{$xPowerState = "Unabled to determine machine Power State: $($Machine.PowerState)"; Break}
	}
	
	If([String]::IsNullOrEmpty($Machine.IPAddress))
	{
		$xIPAddress = "-"
	}
	Else
	{
		$xIPAddress = $Machine.IPAddress.ToString()
	}

	If([String]::IsNullOrEmpty($Machine.LastDeregistrationTime))
	{
		$xLastDeregistrationTime = "-"
	}
	Else
	{
		$xLastDeregistrationTime = $Machine.LastDeregistrationTime.ToString()
	}

	If([String]::IsNullOrEmpty($Machine.OSType))
	{
		$xOSType = "-"
	}
	Else
	{
		$xOSType = $Machine.OSType
	}
	
	If([String]::IsNullOrEmpty($Machine.AgentVersion))
	{
		$xAgentVersion = "-"
	}
	Else
	{
		$xAgentVersion = $Machine.AgentVersion
	}

	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 3 0 $Machine.DNSName
		If($Machine.SessionSupport -eq "MultiSession")
		{
			WriteWordLine 4 0 "Machine"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Name"; Value = $Machine.DNSName; }
			$ScriptInformation += @{Data = "Machine Catalog"; Value = $Machine.CatalogName; }
			$ScriptInformation += @{Data = "Delivery Group"; Value = $Machine.DesktopGroupName; }
			If($NoSessions -eq $False)
			{
                ## GRL $xAssociatedUserFullNames can have a count of zero so $xAssociatedUserFullNames[0] isn't valid.
                [string]$name = $(If( $xAssociatedUserFullNames -is [array] -and $xAssociatedUserFullNames.Count ) { $xAssociatedUserFullNames[0] } Else { '-' } )
				$ScriptInformation += @{Data = "User Display Name"; Value = $name; }
				$cnt = -1
				ForEach($tmp in $xAssociatedUserFullNames)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}
                ## GRL $xAssociatedUserNames can have a count of zero so $xAssociatedUserNames[0] isn't valid.
                [string]$name = $(If( $xAssociatedUserNames -is [array] -and $xAssociatedUserNames.Count ) { $xAssociatedUserNames[0] } Else { '-' } )
				$ScriptInformation += @{Data = "User"; Value = $name; }
				$cnt = -1
				ForEach($tmp in $xAssociatedUserNames)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}
                ## GRL $xAssociatedUserUPNs can have a count of zero so $xAssociatedUserUPNs[0] isn't valid.
                [string]$upn = $(If( $xAssociatedUserUPNs -is [array] -and $xAssociatedUserUPNs.Count ) { $xAssociatedUserUPNs[0] } Else { '-' } )
				$ScriptInformation += @{Data = "UPN"; Value = $upn; }
				$cnt = -1
				ForEach($tmp in $xAssociatedUserUPNs)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}
			}
			[string]$cond = $(If( $xDesktopConditions -is [array] -and $xDesktopConditions.Count ) { $xDesktopConditions[0] } Else { '-' } )
			$ScriptInformation += @{Data = "Desktop Conditions"; Value = $cond; }
			$cnt = -1
			ForEach($tmp in $xDesktopConditions)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "Allocation Type"; Value = $xAllocationType; }
			$ScriptInformation += @{Data = "Maintenance Mode"; Value = $xInMaintenanceMode; }
			$ScriptInformation += @{Data = "Windows Connection Setting"; Value = $xWindowsConnectionSetting; }
			$ScriptInformation += @{Data = "Is Assigned"; Value = $Machine.IsAssigned.ToString(); }
			$ScriptInformation += @{Data = "Is Physical"; Value = $xIsPhysical; }
			$ScriptInformation += @{Data = "Provisioning Type"; Value = $Machine.ProvisioningType.ToString(); }
			$ScriptInformation += @{Data = "Scheduled Reboot"; Value = $Machine.ScheduledReboot; }
			$ScriptInformation += @{Data = "Zone"; Value = $Machine.ZoneName; }
			$ScriptInformation += @{Data = "Summary State"; Value = $xSummaryState; }
            [string]$TagName = $(If( $xTags -is [array] -and $xTags.Count ) { $xTags[0] } Else { '-' } )
			$ScriptInformation += @{Data = "Tags"; Value = $TagName; }
			$cnt = -1
			ForEach($tmp in $xTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "Load Index"; Value = $Machine.LoadIndex; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			If((!$LinuxVDA) -and $VDARegistryKeys -and $MachineIsOnline)
			{
				#First test if the Remote Registry service is enabled. If not skip the VDA registry keys
				$results = Get-Service -ComputerName $Machine.DNSName -Name "RemoteRegistry" -EA 0
				If($? -and $Null -ne $results)
				{
					If($results.Status -eq "Running")
					{
						GetVDARegistryKeys $Machine.DNSName "Server"
						$Script:WordAllVDARegistryItems += $Script:VDARegistryItems
						$Script:VDARegistryItems = New-Object System.Collections.ArrayList
					}
					Else
					{
						$obj1 = [PSCustomObject] @{
							RegKey       = "N/A"
							RegValue     = "N/A"
							VDAType      = "Server"
							ComputerName = $Machine.DNSName	
							Value        = "The Remote Registry service is not running"
						}
						$null = $Script:VDARegistryItems.Add($obj1)
						$Script:WordAllVDARegistryItems += $Script:VDARegistryItems
						$Script:VDARegistryItems = New-Object System.Collections.ArrayList
					}
				}
				Else
				{
					$obj1 = [PSCustomObject] @{
						RegKey       = "N/A"
						RegValue     = "N/A"
						VDAType      = "Server"
						ComputerName = $Machine.DNSName	
						Value        = "The Remote Registry service is not running"
					}
					$null = $Script:VDARegistryItems.Add($obj1)
					$Script:WordAllVDARegistryItems += $Script:VDARegistryItems
					$Script:VDARegistryItems = New-Object System.Collections.ArrayList
				}
			}
			ElseIf($LinuxVDA -and $VDARegistryKeys)
			{
				#VDA is Linux, skipping
			}
			
			WriteWordLine 4 0 "Machine Details"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Agent Version"; Value = $xAgentVersion; }
			$ScriptInformation += @{Data = "IP Address"; Value = $xIPAddress; }
			$ScriptInformation += @{Data = "Is Assigned"; Value = $Machine.IsAssigned.ToString(); }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Applications"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			
			[string]$AppsInUse = $(If( $xApplicationsInUse -is [array] -and $xApplicationsInUse.Count ) { $xApplicationsInUse[0] } Else { '-' } )
			$ScriptInformation += @{Data = "Applications In Use"; Value = $AppsInUse; }
			$cnt = -1
			ForEach($tmp in $xApplicationsInUse)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			[string]$PubApps = $(If( $xPublishedApplications -is [array] -and $xPublishedApplications.Count ) { $xPublishedApplications[0] } Else { '-' } )
			$ScriptInformation += @{Data = "Published Applications"; Value = $PubApps; }
			$cnt = -1
			ForEach($tmp in $xPublishedApplications)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Registration"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Broker"; Value = $xBroker; }
			$ScriptInformation += @{Data = "Last registration failure"; Value = $xLastDeregistrationReason; }
			$ScriptInformation += @{Data = "Last registration failure time"; Value = $xLastDeregistrationTime; }
			$ScriptInformation += @{Data = "Registration State"; Value = $Machine.RegistrationState.ToString(); }
			$ScriptInformation += @{Data = "Fault State"; Value = $xMachineFaultState; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		
			WriteWordLine 4 0 "Hosting"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "VM"; Value = $xHostedMachineName; }
			$ScriptInformation += @{Data = "Hosting Server Name"; Value = $xHostingServerName; }
			$ScriptInformation += @{Data = "Connection"; Value = $xHypervisorConnectionName ; }
			$ScriptInformation += @{Data = "Pending Update"; Value = $Machine.ImageOutOfDate.ToString(); }
			$ScriptInformation += @{Data = "Persist User Changes"; Value = $xPersistUserChanges; }
			$ScriptInformation += @{Data = "Power Action Pending"; Value = $Machine.PowerActionPending.ToString(); }
			$ScriptInformation += @{Data = "Power State"; Value = $xPowerState; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			If($NoSessions -eq $False)
			{
				WriteWordLine 4 0 "Connection"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{Data = "Last Connection Time"; Value = $xLastConnectionTime ; }
				$ScriptInformation += @{Data = "Last Connection User"; Value = $xLastConnectionUser; }
				$ScriptInformation += @{Data = "Secure ICA Active"; Value = $xSessionSecureIcaActive ; }

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 4 0 "Session Details"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{Data = "Launched Via"; Value = $xSessionLaunchedViaHostName; }
				$ScriptInformation += @{Data = "Launched Via (IP)"; Value = $xSessionLaunchedViaIP; }
				[string]$SSAT = $(If( $xSessionSmartAccessTags -is [array] -and $xSessionSmartAccessTags.Count ) { $xSessionSmartAccessTags[0] } Else { '-' } )
				$ScriptInformation += @{Data = "SmartAccess Filters"; Value = $SSAT; }
				$cnt = -1
				ForEach($tmp in $xSessionSmartAccessTags)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 4 0 "Session"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{Data = "Session Count"; Value = $Machine.SessionCount.ToString(); }

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}
		}
		ElseIf($Machine.SessionSupport -eq "SingleSession")
		{
			WriteWordLine 4 0 "Machine"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Name"; Value = $Machine.DNSName; }
			$ScriptInformation += @{Data = "Machine Catalog"; Value = $Machine.CatalogName; }
			$ScriptInformation += @{Data = "Delivery Group"; Value = $Machine.DesktopGroupName; }
			If($NoSessions -eq $False)
			{
                ## GRL $xAssociatedUserFullNames can have a count of zero so $xAssociatedUserFullNames[0] isn't valid.
                [string]$name = $(If( $xAssociatedUserFullNames -is [array] -and $xAssociatedUserFullNames.Count ) { $xAssociatedUserFullNames[0] } Else { '-' } )
				$ScriptInformation += @{Data = "User Display Name"; Value = $name; }
				$cnt = -1
				ForEach($tmp in $xAssociatedUserFullNames)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}
                ## GRL $xAssociatedUserNames can have a count of zero so $xAssociatedUserNames[0] isn't valid.
                [string]$name = $(If( $xAssociatedUserNames -is [array] -and $xAssociatedUserNames.Count ) { $xAssociatedUserNames[0] } Else { '-' } )
				$ScriptInformation += @{Data = "User"; Value = $name; }
				$cnt = -1
				ForEach($tmp in $xAssociatedUserNames)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}
                ## GRL $xAssociatedUserUPNs can have a count of zero so $xAssociatedUserUPNs[0] isn't valid.
                [string]$upn = $(If( $xAssociatedUserUPNs -is [array] -and $xAssociatedUserUPNs.Count ) { $xAssociatedUserUPNs[0] } Else { '-' } )
				$ScriptInformation += @{Data = "UPN"; Value = $upn; }
				$cnt = -1
				ForEach($tmp in $xAssociatedUserUPNs)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}
			}
			[string]$cond = $(If( $xDesktopConditions -is [array] -and $xDesktopConditions.Count ) { $xDesktopConditions[0] } Else { '-' } )
			$ScriptInformation += @{Data = "Desktop Conditions"; Value = $cond; }
			$cnt = -1
			ForEach($tmp in $xDesktopConditions)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "Allocation Type"; Value = $xAllocationType; }
			$ScriptInformation += @{Data = "Maintenance Mode"; Value = $xInMaintenanceMode; }
			$ScriptInformation += @{Data = "Is Assigned"; Value = $Machine.IsAssigned.ToString(); }
			$ScriptInformation += @{Data = "Is Physical"; Value = $xIsPhysical; }
			$ScriptInformation += @{Data = "Provisioning Type"; Value = $Machine.ProvisioningType.ToString(); }
			$ScriptInformation += @{Data = "Zone"; Value = $Machine.ZoneName; }
			$ScriptInformation += @{Data = "Scheduled Reboot"; Value = $Machine.ScheduledReboot; }
			$ScriptInformation += @{Data = "Summary State"; Value = $xSummaryState; }
            [string]$TagName = $(If( $xTags -is [array] -and $xTags.Count ) { $xTags[0] } Else { '-' } )
			$ScriptInformation += @{Data = "Tags"; Value = $TagName; }
			$cnt = -1
			ForEach($tmp in $xTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			If((!$LinuxVDA) -and $VDARegistryKeys -and $MachineIsOnline)
			{
				#First test if the Remote Registry service is enabled. If not skip the VDA registry keys
				$results = Get-Service -ComputerName $Machine.DNSName -Name "RemoteRegistry" -EA 0
				If($? -and $Null -ne $results)
				{
					If($results.Status -eq "Running")
					{
						GetVDARegistryKeys $Machine.DNSName "Desktop"
						$Script:WordAllVDARegistryItems += $Script:VDARegistryItems
						$Script:VDARegistryItems = New-Object System.Collections.ArrayList
					}
					Else
					{
						$obj1 = [PSCustomObject] @{
							RegKey       = "N/A"
							RegValue     = "N/A"
							VDAType      = "Desktop"
							ComputerName = $Machine.DNSName	
							Value        = "The Remote Registry service is not running"
						}
						$null = $Script:VDARegistryItems.Add($obj1)
						$Script:WordAllVDARegistryItems += $Script:VDARegistryItems
						$Script:VDARegistryItems = New-Object System.Collections.ArrayList
					}
				}
				Else
				{
					$obj1 = [PSCustomObject] @{
						RegKey       = "N/A"
						RegValue     = "N/A"
						VDAType      = "Desktop"
						ComputerName = $Machine.DNSName	
						Value        = "The Remote Registry service is not running"
					}
					$null = $Script:VDARegistryItems.Add($obj1)
					$Script:WordAllVDARegistryItems += $Script:VDARegistryItems
					$Script:VDARegistryItems = New-Object System.Collections.ArrayList
				}
			}
			ElseIf($LinuxVDA -and $VDARegistryKeys)
			{
				#VDA is Linux, skipping
			}
			
			WriteWordLine 4 0 "Machine Details"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Agent Version"; Value = $xAgentVersion; }
			$ScriptInformation += @{Data = "IP Address"; Value = $xIPAddress; }
			$ScriptInformation += @{Data = "Is Assigned"; Value = $Machine.IsAssigned.ToString(); }
			$ScriptInformation += @{Data = "OS Type"; Value = $xOSType; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Applications"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			
			[string]$AppsInUse = $(If( $xApplicationsInUse -is [array] -and $xApplicationsInUse.Count ) { $xApplicationsInUse[0] } Else { '-' } )
			$ScriptInformation += @{Data = "Applications In Use"; Value = $AppsInUse; }
			$cnt = -1
			ForEach($tmp in $xApplicationsInUse)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			[string]$PubApps = $(If( $xPublishedApplications -is [array] -and $xPublishedApplications.Count ) { $xPublishedApplications[0] } Else { '-' } )
			$ScriptInformation += @{Data = "Published Applications"; Value = $PubApps; }
			$cnt = -1
			ForEach($tmp in $xPublishedApplications)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			If($NoSessions -eq $False)
			{
				WriteWordLine 4 0 "Connection"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{Data = "Client (IP)"; Value = $xSessionClientAddress; }
				$ScriptInformation += @{Data = "Client"; Value = $xSessionClientName; }
				$ScriptInformation += @{Data = "Plug-in Version"; Value = $xSessionClientVersion; }
				$ScriptInformation += @{Data = "Connected Via"; Value = $xSessionConnectedViaHostName; }
				$ScriptInformation += @{Data = "Connected Via (IP)"; Value = $xSessionConnectedViaIP; }
				$ScriptInformation += @{Data = "Last Connection Time"; Value = $xLastConnectionTime ; }
				$ScriptInformation += @{Data = "Last Connection User"; Value = $xLastConnectionUser; }
				$ScriptInformation += @{Data = "Connection Type"; Value = $xSessionProtocol; }
				$ScriptInformation += @{Data = "Secure ICA Active"; Value = $xSessionSecureIcaActive ; }

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			
			WriteWordLine 4 0 "Registration"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Broker"; Value = $xBroker; }
			$ScriptInformation += @{Data = "Last registration failure"; Value = $xLastDeregistrationReason; }
			$ScriptInformation += @{Data = "Last registration failure time"; Value = $xLastDeregistrationTime; }
			$ScriptInformation += @{Data = "Registration State"; Value = $Machine.RegistrationState.ToString(); }
			$ScriptInformation += @{Data = "Fault State"; Value = $xMachineFaultState; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Hosting"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "VM"; Value = $xHostedMachineName; }
			$ScriptInformation += @{Data = "Hosting Server Name"; Value = $xHostingServerName; }
			$ScriptInformation += @{Data = "Connection"; Value = $xHypervisorConnectionName ; }
			$ScriptInformation += @{Data = "Pending Update"; Value = $Machine.ImageOutOfDate.ToString(); }
			$ScriptInformation += @{Data = "Persist User Changes"; Value = $xPersistUserChanges; }
			$ScriptInformation += @{Data = "Power Action Pending"; Value = $Machine.PowerActionPending.ToString(); }
			$ScriptInformation += @{Data = "Power State"; Value = $xPowerState; }
			$ScriptInformation += @{Data = "Will Shutdown After Use"; Value = $xWillShutdownAfterUse; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			If($NoSessions -eq $False)
			{
				WriteWordLine 4 0 "Session Details"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{Data = "Launched Via"; Value = $xSessionLaunchedViaHostName; }
				$ScriptInformation += @{Data = "Launched Via (IP)"; Value = $xSessionLaunchedViaIP; }
				$ScriptInformation += @{Data = "Session Change Time"; Value = $xSessionStateChangeTime; }
				[string]$SSAT = $(If( $xSessionSmartAccessTags -is [array] -and $xSessionSmartAccessTags.Count ) { $xSessionSmartAccessTags[0] } Else { '-' } )
				$ScriptInformation += @{Data = "SmartAccess Filters"; Value = $SSAT; }
				$cnt = -1
				ForEach($tmp in $xSessionSmartAccessTags)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 4 0 "Session"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{Data = "Session State"; Value = $xSessionState; }
				$ScriptInformation += @{Data = "Current User"; Value = $xSessionUserName; }
				$ScriptInformation += @{Data = "Start Time"; Value = $xSessionStateChangeTime; }

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}
		}
		
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		If($Machine.SessionSupport -eq "MultiSession")
		{
			Line 1 "Machine"
			Line 2 "Name`t`t`t`t: " $Machine.DNSName
			Line 2 "Machine Catalog`t`t`t: " $Machine.CatalogName
			Line 2 "Delivery Group`t`t`t: " $Machine.DesktopGroupName
			If($NoSessions -eq $False)
			{
                ## GRL $xAssociatedUserFullNames can have a count of zero so $xAssociatedUserFullNames[0] isn't valid.
                [string]$name = $(If( $xAssociatedUserFullNames -is [array] -and $xAssociatedUserFullNames.Count ) { $xAssociatedUserFullNames[0] } Else { '-' } )
				Line 2 "User Display Name`t`t: " $name
				$cnt = -1
				ForEach($tmp in $xAssociatedUserFullNames)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 6 "  " $tmp
					}
				}
                ## GRL $xAssociatedUserNames can have a count of zero so $xAssociatedUserNames[0] isn't valid.
                [string]$name = $(If( $xAssociatedUserNames -is [array] -and $xAssociatedUserNames.Count ) { $xAssociatedUserNames[0] } Else { '-' } )
				Line 2 "User`t`t`t`t: " $name
				$cnt = -1
				ForEach($tmp in $xAssociatedUserNames)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 6 "  " $tmp
					}
				}
                ## GRL $xAssociatedUserUPNs can have a count of zero so $xAssociatedUserUPNs[0] isn't valid.
                [string]$upn = $(If( $xAssociatedUserUPNs -is [array] -and $xAssociatedUserUPNs.Count ) { $xAssociatedUserUPNs[0] } Else { '-' } )
				Line 2 "UPN`t`t`t`t: " $upn
				$cnt = -1
				ForEach($tmp in $xAssociatedUserUPNs)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 6 "  " $tmp
					}
				}
			}
			[string]$cond = $(If( $xDesktopConditions -is [array] -and $xDesktopConditions.Count ) { $xDesktopConditions[0] } Else { '-' } )
			Line 2 "Desktop Conditions`t`t: " $cond
			$cnt = -1
			ForEach($tmp in $xDesktopConditions)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "Allocation Type`t`t`t: " $xAllocationType
			Line 2 "Maintenance Mode`t`t: " $xInMaintenanceMode
			Line 2 "Windows Connection Setting`t: " $xWindowsConnectionSetting
			Line 2 "Is Assigned`t`t`t: " $Machine.IsAssigned.ToString()
			Line 2 "Is Physical`t`t`t: " $xIsPhysical
			Line 2 "Provisioning Type`t`t: " $Machine.ProvisioningType.ToString()
			Line 2 "Scheduled Reboot`t`t: " $Machine.ScheduledReboot
			Line 2 "Zone`t`t`t`t: " $Machine.ZoneName
			Line 2 "Summary State`t`t`t: " $xSummaryState
            [string]$TagName = $(If( $xTags -is [array] -and $xTags.Count ) { $xTags[0] } Else { '-' } )
			Line 2 "Tags`t`t`t`t: " $TagName
			$cnt = -1
			ForEach($tmp in $xTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "Load Index`t`t`t: " $Machine.LoadIndex
			Line 0 ""

			If((!$LinuxVDA) -and $VDARegistryKeys -and $MachineIsOnline)
			{
				#First test if the Remote Registry service is enabled. If not skip the VDA registry keys
				$results = Get-Service -ComputerName $Machine.DNSName -Name "RemoteRegistry" -EA 0
				If($? -and $Null -ne $results)
				{
					If($results.Status -eq "Running")
					{
						GetVDARegistryKeys $Machine.DNSName "Server"
						$Script:TextAllVDARegistryItems += $Script:VDARegistryItems
						$Script:VDARegistryItems = New-Object System.Collections.ArrayList
					}
					Else
					{
						$obj1 = [PSCustomObject] @{
							RegKey       = "N/A"
							RegValue     = "N/A"
							VDAType      = "Server"
							ComputerName = $Machine.DNSName	
							Value        = "The Remote Registry service is not running"
						}
						$null = $Script:VDARegistryItems.Add($obj1)
						$Script:TextAllVDARegistryItems += $Script:VDARegistryItems
						$Script:VDARegistryItems = New-Object System.Collections.ArrayList
					}
				}
				Else
				{
					$obj1 = [PSCustomObject] @{
						RegKey       = "N/A"
						RegValue     = "N/A"
						VDAType      = "Server"
						ComputerName = $Machine.DNSName	
						Value        = "The Remote Registry service is not running"
					}
					$null = $Script:VDARegistryItems.Add($obj1)
					$Script:TextAllVDARegistryItems += $Script:VDARegistryItems
					$Script:VDARegistryItems = New-Object System.Collections.ArrayList
				}
			}
			ElseIf($LinuxVDA -and $VDARegistryKeys)
			{
				#VDA is Linux, skipping
			}
			
			Line 1 "Machine Details"
			Line 2 "Agent Version`t`t`t: " $xAgentVersion
			Line 2 "IP Address`t`t`t: " $xIPAddress
			Line 2 "Is Assigned`t`t`t: " $Machine.IsAssigned.ToString()
			Line 0 ""
			
			Line 1 "Applications"
			[string]$AppsInUse = $(If( $xApplicationsInUse -is [array] -and $xApplicationsInUse.Count ) { $xApplicationsInUse[0] } Else { '-' } )
			Line 2 "Applications In Use`t`t: " $AppsInUse
			$cnt = -1
			ForEach($tmp in $xApplicationsInUse)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			[string]$PubApps = $(If( $xPublishedApplications -is [array] -and $xPublishedApplications.Count ) { $xPublishedApplications[0] } Else { '-' } )
			Line 2 "Published Applications`t`t: " $PubApps
			$cnt = -1
			ForEach($tmp in $xPublishedApplications)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 0 ""
			
			Line 1 "Registration"
			Line 2 "Broker`t`t`t`t: " $xBroker
			Line 2 "Last registration failure`t: " $xLastDeregistrationReason
			Line 2 "Last registration failure time`t: " $xLastDeregistrationTime
			Line 2 "Registration State`t`t: " $Machine.RegistrationState.ToString()
			Line 2 "Fault State`t`t`t: " $xMachineFaultState
			Line 0 ""
			
			Line 1 "Hosting"
			Line 2 "VM`t`t`t`t: " $xHostedMachineName
			Line 2 "Hosting Server Name`t`t: " $xHostingServerName
			Line 2 "Connection`t`t`t: " $xHypervisorConnectionName 
			Line 2 "Pending Update`t`t`t: " $Machine.ImageOutOfDate.ToString()
			Line 2 "Persist User Changes`t`t: " $xPersistUserChanges
			Line 2 "Power Action Pending`t`t: " $Machine.PowerActionPending.ToString()
			Line 2 "Power State`t`t`t: " $xPowerState
			Line 0 ""
			
			If($NoSessions -eq $False)
			{
				Line 1 "Connection"
				Line 2 "Last Connection Time`t`t: " $xLastConnectionTime 
				Line 2 "Last Connection User`t`t: " $xLastConnectionUser
				Line 2 "Secure ICA Active`t`t: " $xSessionSecureIcaActive 
				Line 0 ""
				
				Line 1 "Session Details"
				Line 2 "Launched Via`t`t`t: " $xSessionLaunchedViaHostName
				Line 2 "Launched Via (IP)`t`t: " $xSessionLaunchedViaIP
				[string]$SSAT = $(If( $xSessionSmartAccessTags -is [array] -and $xSessionSmartAccessTags.Count ) { $xSessionSmartAccessTags[0] } Else { '-' } )
				Line 2 "SmartAccess Filters`t`t: " $SSAT
				$cnt = -1
				ForEach($tmp in $xSessionSmartAccessTags)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 5 "  " $tmp
					}
				}
				Line 0 ""
				
				Line 1 "Session"
				Line 2 "Session Count`t`t`t: " $Machine.SessionCount.ToString()
				Line 0 ""
			}
		}
		ElseIf($Machine.SessionSupport -eq "SingleSession")
		{
			Line 1 "Machine"
			Line 2 "Name`t`t`t`t: " $Machine.DNSName
			Line 2 "Machine Catalog`t`t`t: " $Machine.CatalogName
			Line 2 "Delivery Group`t`t`t: " $Machine.DesktopGroupName
			If($NoSessions -eq $False)
			{
                ## GRL $xAssociatedUserFullNames can have a count of zero so $xAssociatedUserFullNames[0] isn't valid.
                [string]$name = $(If( $xAssociatedUserFullNames -is [array] -and $xAssociatedUserFullNames.Count ) { $xAssociatedUserFullNames[0] } Else { '-' } )
				Line 2 "User Display Name`t`t: " $name
				$cnt = -1
				ForEach($tmp in $xAssociatedUserFullNames)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 6 "  " $tmp
					}
				}
                ## GRL $xAssociatedUserNames can have a count of zero so $xAssociatedUserNames[0] isn't valid.
                [string]$name = $(If( $xAssociatedUserNames -is [array] -and $xAssociatedUserNames.Count ) { $xAssociatedUserNames[0] } Else { '-' } )
				Line 2 "User`t`t`t`t: " $name
				$cnt = -1
				ForEach($tmp in $xAssociatedUserNames)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 6 "  " $tmp
					}
				}
                ## GRL $xAssociatedUserUPNs can have a count of zero so $xAssociatedUserUPNs[0] isn't valid.
                [string]$upn = $(If( $xAssociatedUserUPNs -is [array] -and $xAssociatedUserUPNs.Count ) { $xAssociatedUserUPNs[0] } Else { '-' } )
				Line 2 "UPN`t`t`t`t: " $upn
				$cnt = -1
				ForEach($tmp in $xAssociatedUserUPNs)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 6 "  " $tmp
					}
				}
			}
			[string]$cond = $(If( $xDesktopConditions -is [array] -and $xDesktopConditions.Count ) { $xDesktopConditions[0] } Else { '-' } )
			Line 2 "Desktop Conditions`t`t: " $cond
			$cnt = -1
			ForEach($tmp in $xDesktopConditions)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "Allocation Type`t`t`t: " $xAllocationType
			Line 2 "Maintenance Mode`t`t: " $xInMaintenanceMode
			Line 2 "Is Assigned`t`t`t: " $Machine.IsAssigned.ToString()
			Line 2 "Is Physical`t`t`t: " $xIsPhysical
			Line 2 "Provisioning Type`t`t: " $Machine.ProvisioningType.ToString()
			Line 2 "Zone`t`t`t`t: " $Machine.ZoneName
			Line 2 "Scheduled Reboot`t`t: " $Machine.ScheduledReboot
			Line 2 "Summary State`t`t`t: " $xSummaryState
            [string]$TagName = $(If( $xTags -is [array] -and $xTags.Count ) { $xTags[0] } Else { '-' } )
			Line 2 "Tags`t`t`t`t: " $TagName
			$cnt = -1
			ForEach($tmp in $xTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 0 ""

			If((!$LinuxVDA) -and $VDARegistryKeys -and $MachineIsOnline)
			{
				#First test if the Remote Registry service is enabled. If not skip the VDA registry keys
				$results = Get-Service -ComputerName $Machine.DNSName -Name "RemoteRegistry" -EA 0
				If($? -and $Null -ne $results)
				{
					If($results.Status -eq "Running")
					{
						GetVDARegistryKeys $Machine.DNSName "Desktop"
						$Script:TextAllVDARegistryItems += $Script:VDARegistryItems
						$Script:VDARegistryItems = New-Object System.Collections.ArrayList
					}
					Else
					{
						$obj1 = [PSCustomObject] @{
							RegKey       = "N/A"
							RegValue     = "N/A"
							VDAType      = "Desktop"
							ComputerName = $Machine.DNSName	
							Value        = "The Remote Registry service is not running"
						}
						$null = $Script:VDARegistryItems.Add($obj1)
						$Script:TextAllVDARegistryItems += $Script:VDARegistryItems
						$Script:VDARegistryItems = New-Object System.Collections.ArrayList
					}
				}
				Else
				{
					$obj1 = [PSCustomObject] @{
						RegKey       = "N/A"
						RegValue     = "N/A"
						VDAType      = "Desktop"
						ComputerName = $Machine.DNSName	
						Value        = "The Remote Registry service is not running"
					}
					$null = $Script:VDARegistryItems.Add($obj1)
					$Script:TextAllVDARegistryItems += $Script:VDARegistryItems
					$Script:VDARegistryItems = New-Object System.Collections.ArrayList
				}
			}
			ElseIf($LinuxVDA -and $VDARegistryKeys)
			{
				#VDA is Linux, skipping
			}
			
			Line 1 "Machine Details"
			Line 2 "Agent Version`t`t`t: " $xAgentVersion
			Line 2 "IP Address`t`t`t: " $xIPAddress
			Line 2 "Is Assigned`t`t`t: " $Machine.IsAssigned.ToString()
			Line 2 "OS Type`t`t`t`t: " $xOSType
			Line 0 ""
			
			Line 1 "Applications"
			[string]$AppsInUse = $(If( $xApplicationsInUse -is [array] -and $xApplicationsInUse.Count ) { $xApplicationsInUse[0] } Else { '-' } )
			Line 2 "Applications In Use`t`t: " $AppsInUse
			$cnt = -1
			ForEach($tmp in $xApplicationsInUse)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			[string]$PubApps = $(If( $xPublishedApplications -is [array] -and $xPublishedApplications.Count ) { $xPublishedApplications[0] } Else { '-' } )
			Line 2 "Published Applications`t`t: " $PubApps
			$cnt = -1
			ForEach($tmp in $xPublishedApplications)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 0 ""
			
			If($NoSessions -eq $False)
			{
				Line 1 "Connection"
				Line 2 "Client (IP)`t`t`t: " $xSessionClientAddress
				Line 2 "Client`t`t`t`t: " $xSessionClientName
				Line 2 "Plug-in Version`t`t`t: " $xSessionClientVersion
				Line 2 "Connected Via`t`t`t: " $xSessionConnectedViaHostName
				Line 2 "Connect Via (IP)`t`t: " $xSessionConnectedViaIP
				Line 2 "Last Connection Time`t`t: " $xLastConnectionTime 
				Line 2 "Last Connection User`t`t: " $xLastConnectionUser
				Line 2 "Connection Type`t`t`t: " $xSessionProtocol
				Line 2 "Secure ICA Active`t`t: " $xSessionSecureIcaActive 
				Line 0 ""
			}
			
			Line 1 "Registration"
			Line 2 "Broker`t`t`t`t: " $xBroker
			Line 2 "Last registration failure`t: " $xLastDeregistrationReason
			Line 2 "Last registration failure time`t: " $xLastDeregistrationTime
			Line 2 "Registration State`t`t: " $Machine.RegistrationState.ToString()
			Line 2 "Fault State`t`t`t: " $xMachineFaultState
			Line 0 ""
			
			Line 1 "Hosting"
			Line 2 "VM`t`t`t`t: " $xHostedMachineName
			Line 2 "Hosting Server Name`t`t: " $xHostingServerName
			Line 2 "Connection`t`t`t: " $xHypervisorConnectionName 
			Line 2 "Pending Update`t`t`t: " $Machine.ImageOutOfDate.ToString()
			Line 2 "Persist User Changes`t`t: " $xPersistUserChanges
			Line 2 "Power Action Pending`t`t: " $Machine.PowerActionPending.ToString()
			Line 2 "Power State`t`t`t: " $xPowerState
			Line 2 "Will Shutdown After Use`t`t: " $xWillShutdownAfterUse
			Line 0 ""
			
			If($NoSessions -eq $False)
			{
				Line 1 "Session Details"
				Line 2 "Launched Via`t`t`t: " $xSessionLaunchedViaHostName
				Line 2 "Launched Via (IP)`t`t: " $xSessionLaunchedViaIP
				Line 2 "Session Change Time`t`t: " $xSessionStateChangeTime
				[string]$SSAT = $(If( $xSessionSmartAccessTags -is [array] -and $xSessionSmartAccessTags.Count ) { $xSessionSmartAccessTags[0] } Else { '-' } )
				Line 2 "SmartAccess Filters`t`t: " $SSAT
				$cnt = -1
				ForEach($tmp in $xSessionSmartAccessTags)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 5 "  " $tmp
					}
				}
				Line 0 ""
				
				Line 1 "Session"
				Line 2 "Session State`t`t`t: " $xSessionState
				Line 2 "Current User`t`t`t: " $xSessionUserName
				Line 2 "Start Time`t`t`t: " $xSessionStateChangeTime
				Line 0 ""
			}
		}
	}
	If($HTML)
	{
		If($Machine.SessionSupport -eq "MultiSession")
		{
			WriteHTMLLine 4 0 "Machine"
			$rowdata = @()

			$columnHeaders = @("Name",($global:htmlsb),$Machine.DNSName,$htmlwhite)
			$rowdata += @(,('Machine Catalog',($global:htmlsb),$Machine.CatalogName,$htmlwhite))
			$rowdata += @(,('Delivery Group',($global:htmlsb),$Machine.DesktopGroupName,$htmlwhite))
			If($NoSessions -eq $False)
			{
                ## GRL $xAssociatedUserFullNames can have a count of zero so $xAssociatedUserFullNames[0] isn't valid.
                [string]$name = $(If( $xAssociatedUserFullNames -is [array] -and $xAssociatedUserFullNames.Count ) { $xAssociatedUserFullNames[0] } Else { '-' } )
				$rowdata += @(,('User Display Name',($global:htmlsb),$name,$htmlwhite))
				$cnt = -1
				ForEach($tmp in $xAssociatedUserFullNames)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
					}
				}
                ## GRL $xAssociatedUserNames can have a count of zero so $xAssociatedUserNames[0] isn't valid.
                [string]$name = $(If( $xAssociatedUserNames -is [array] -and $xAssociatedUserNames.Count ) { $xAssociatedUserNames[0] } Else { '-' } )
				$rowdata += @(,('User',($global:htmlsb),$name,$htmlwhite))
				$cnt = -1
				ForEach($tmp in $xAssociatedUserNames)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
					}
				}
                ## GRL $xAssociatedUserUPNs can have a count of zero so $xAssociatedUserUPNs[0] isn't valid.
                [string]$upn = $(If( $xAssociatedUserUPNs -is [array] -and $xAssociatedUserUPNs.Count ) { $xAssociatedUserUPNs[0] } Else { '-' } )
				$rowdata += @(,('UPN',($global:htmlsb),$upn,$htmlwhite))
				$cnt = -1
				ForEach($tmp in $xAssociatedUserUPNs)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
					}
				}
			}
			[string]$cond = $(If( $xDesktopConditions -is [array] -and $xDesktopConditions.Count ) { $xDesktopConditions[0] } Else { '-' } )
			$rowdata += @(,('Desktop Conditions',($global:htmlsb),$cond,$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xDesktopConditions)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('Allocation Type',($global:htmlsb),$xAllocationType,$htmlwhite))
			$rowdata += @(,('Maintenance Mode',($global:htmlsb),$xInMaintenanceMode,$htmlwhite))
			$rowdata += @(,('Windows Connection Setting',($global:htmlsb),$xWindowsConnectionSetting,$htmlwhite))
			$rowdata += @(,('Is Assigned',($global:htmlsb),$Machine.IsAssigned.ToString(),$htmlwhite))
			$rowdata += @(,('Is Physical',($global:htmlsb),$xIsPhysical,$htmlwhite))
			$rowdata += @(,('Provisioning Type',($global:htmlsb),$Machine.ProvisioningType.ToString(),$htmlwhite))
			$rowdata += @(,('Scheduled Reboot',($global:htmlsb),$Machine.ScheduledReboot.ToString(),$htmlwhite))
			$rowdata += @(,('Zone',($global:htmlsb),$Machine.ZoneName,$htmlwhite))
			$rowdata += @(,('Summary State',($global:htmlsb),$xSummaryState,$htmlwhite))
            [string]$TagName = $(If( $xTags -is [array] -and $xTags.Count ) { $xTags[0] } Else { '-' } )
			$rowdata += @(,('Tags',($global:htmlsb),$TagName,$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('Load Index',($global:htmlsb),$Machine.LoadIndex.ToString(),$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

			If((!$LinuxVDA) -and $VDARegistryKeys -and $MachineIsOnline)
			{
				#First test if the Remote Registry service is enabled. If not skip the VDA registry keys
				$results = Get-Service -ComputerName $Machine.DNSName -Name "RemoteRegistry" -EA 0
				If($? -and $Null -ne $results)
				{
					If($results.Status -eq "Running")
					{
						GetVDARegistryKeys $Machine.DNSName "Server"
						$Script:HTMLAllVDARegistryItems += $Script:VDARegistryItems
						$Script:VDARegistryItems = New-Object System.Collections.ArrayList
					}
					Else
					{
						$obj1 = [PSCustomObject] @{
							RegKey       = "N/A"
							RegValue     = "N/A"
							VDAType      = "Server"
							ComputerName = $Machine.DNSName	
							Value        = "The Remote Registry service is not running"
						}
						$null = $Script:VDARegistryItems.Add($obj1)
						$Script:HTMLAllVDARegistryItems += $Script:VDARegistryItems
						$Script:VDARegistryItems = New-Object System.Collections.ArrayList
					}
				}
				Else
				{
					$obj1 = [PSCustomObject] @{
						RegKey       = "N/A"
						RegValue     = "N/A"
						VDAType      = "Server"
						ComputerName = $Machine.DNSName	
						Value        = "The Remote Registry service is not running"
					}
					$null = $Script:VDARegistryItems.Add($obj1)
					$Script:HTMLAllVDARegistryItems += $Script:VDARegistryItems
					$Script:VDARegistryItems = New-Object System.Collections.ArrayList
				}
			}
			ElseIf($LinuxVDA -and $VDARegistryKeys)
			{
				#VDA is Linux, skipping
			}
			
			WriteHTMLLine 4 0 "Machine Details"
			$rowdata = @()
			$columnHeaders = @("Agent Version",($global:htmlsb),$xAgentVersion,$htmlwhite)
			$rowdata += @(,('IP Address',($global:htmlsb),$xIPAddress,$htmlwhite))
			$rowdata += @(,('Is Assigned',($global:htmlsb),$Machine.IsAssigned.ToString(),$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

			WriteHTMLLine 4 0 "Applications"
			$rowdata = @()
			[string]$AppsInUse = $(If( $xApplicationsInUse -is [array] -and $xApplicationsInUse.Count ) { $xApplicationsInUse[0] } Else { '-' } )
			$columnHeaders = @("Applications In Use",($global:htmlsb),$AppsInUse,$htmlwhite)
			$cnt = -1
			ForEach($tmp in $xApplicationsInUSe)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
				}
			}
			[string]$PubApps = $(If( $xPublishedApplications -is [array] -and $xPublishedApplications.Count ) { $xPublishedApplications[0] } Else { '-' } )
			$rowdata += @(,('Published Applications',($global:htmlsb),$PubApps,$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xPublishedApplications)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
				}
			}

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

			WriteHTMLLine 4 0 "Registration"
			$rowdata = @()
			$columnHeaders = @("Broker",($global:htmlsb),$xBroker,$htmlwhite)
			$rowdata += @(,('Last registration failure',($global:htmlsb),$xLastDeregistrationReason,$htmlwhite))
			$rowdata += @(,('Last registration failure time',($global:htmlsb),$xLastDeregistrationTime,$htmlwhite))
			$rowdata += @(,('Registration State',($global:htmlsb),$Machine.RegistrationState.ToString(),$htmlwhite))
			$rowdata += @(,('Fault State',($global:htmlsb),$xMachineFaultState,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

			WriteHTMLLine 4 0 "Hosting"
			$rowdata = @()
			$columnHeaders = @("VM",($global:htmlsb),$xHostedMachineName,$htmlwhite)
			$rowdata += @(,('Hosting Server Name',($global:htmlsb),$xHostingServerName,$htmlwhite))
			$rowdata += @(,('Connection',($global:htmlsb),$xHypervisorConnectionName,$htmlwhite))
			$rowdata += @(,('Pending Update',($global:htmlsb),$Machine.ImageOutOfDate.ToString(),$htmlwhite))
			$rowdata += @(,('Persist User Changes',($global:htmlsb),$xPersistUserChanges,$htmlwhite))
			$rowdata += @(,('Power Action Pending',($global:htmlsb),$Machine.PowerActionPending.ToString(),$htmlwhite))
			$rowdata += @(,('Power State',($global:htmlsb),$xPowerState,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

			If($NoSessions -eq $False)
			{
				WriteHTMLLine 4 0 "Connection"
				$rowdata = @()
				$columnHeaders = @("Last Connection Time",($global:htmlsb),$xLastConnectionTime,$htmlwhite)
				$rowdata += @(,('Last Connection User',($global:htmlsb),$xLastConnectionUser,$htmlwhite))
				$rowdata += @(,('Secure ICA Active',($global:htmlsb),$xSessionSecureIcaActive,$htmlwhite))

				$msg = ""
				$columnWidths = @("200px","250px")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

				WriteHTMLLine 4 0 "Session Details"
				$rowdata = @()
				$columnHeaders = @("Launched Via",($global:htmlsb),$xSessionLaunchedViaHostName,$htmlwhite)
				$rowdata += @(,('Launched Via (IP)',($global:htmlsb),$xSessionLaunchedViaIP,$htmlwhite))
				[string]$SSAT = $(If( $xSessionSmartAccessTags -is [array] -and $xSessionSmartAccessTags.Count ) { $xSessionSmartAccessTags[0] } Else { '-' } )
				$rowdata += @(,('SmartAccess Filters',($global:htmlsb),$SSAT,$htmlwhite))
				$cnt = -1
				ForEach($tmp in $xSessionSmartAccessTags)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("200px","250px")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

				WriteHTMLLine 4 0 "Session"
				$rowdata = @()
				$columnHeaders = @("Session Count",($global:htmlsb),$Machine.SessionCount.ToString(),$htmlwhite)

				$msg = ""
				$columnWidths = @("200px","250px")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			}
		}
		ElseIf($Machine.SessionSupport -eq "SingleSession")
		{
			WriteHTMLLine 4 0 "Machine"
			$rowdata = @()

			$columnHeaders = @("Name",($global:htmlsb),$Machine.DNSName,$htmlwhite)
			$rowdata += @(,('Machine Catalog',($global:htmlsb),$Machine.CatalogName,$htmlwhite))
			$rowdata += @(,('Delivery Group',($global:htmlsb),$Machine.DesktopGroupName,$htmlwhite))
			If($NoSessions -eq $False)
			{
                ## GRL $xAssociatedUserFullNames can have a count of zero so $xAssociatedUserFullNames[0] isn't valid.
                [string]$name = $(If( $xAssociatedUserFullNames -is [array] -and $xAssociatedUserFullNames.Count ) { $xAssociatedUserFullNames[0] } Else { '-' } )
				$rowdata += @(,('User Display Name',($global:htmlsb),$name,$htmlwhite))
				$cnt = -1
				ForEach($tmp in $xAssociatedUserFullNames)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
					}
				}
                ## GRL $xAssociatedUserNames can have a count of zero so $xAssociatedUserNames[0] isn't valid.
                [string]$name = $(If( $xAssociatedUserNames -is [array] -and $xAssociatedUserNames.Count ) { $xAssociatedUserNames[0] } Else { '-' } )
				$rowdata += @(,('User',($global:htmlsb),$name,$htmlwhite))
				$cnt = -1
				ForEach($tmp in $xAssociatedUserNames)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
					}
				}
                ## GRL $xAssociatedUserUPNs can have a count of zero so $xAssociatedUserUPNs[0] isn't valid.
                [string]$upn = $(If( $xAssociatedUserUPNs -is [array] -and $xAssociatedUserUPNs.Count ) { $xAssociatedUserUPNs[0] } Else { '-' } )
				$rowdata += @(,('UPN',($global:htmlsb),$upn,$htmlwhite))
				$cnt = -1
				ForEach($tmp in $xAssociatedUserUPNs)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
					}
				}
			}
			[string]$cond = $(If( $xDesktopConditions -is [array] -and $xDesktopConditions.Count ) { $xDesktopConditions[0] } Else { '-' } )
			$rowdata += @(,('Desktop Conditions',($global:htmlsb),$cond,$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xDesktopConditions)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('Allocation Type',($global:htmlsb),$xAllocationType,$htmlwhite))
			$rowdata += @(,('Maintenance Mode',($global:htmlsb),$xInMaintenanceMode,$htmlwhite))
			$rowdata += @(,('Is Assigned',($global:htmlsb),$Machine.IsAssigned.ToString(),$htmlwhite))
			$rowdata += @(,('Is Physical',($global:htmlsb),$xIsPhysical,$htmlwhite))
			$rowdata += @(,('Provisioning Type',($global:htmlsb),$Machine.ProvisioningType.ToString(),$htmlwhite))
			$rowdata += @(,('Zone',($global:htmlsb),$Machine.ZoneName,$htmlwhite))
			$rowdata += @(,('Scheduled Reboot',($global:htmlsb),$Machine.ScheduledReboot.ToString(),$htmlwhite))
			$rowdata += @(,('Summary State',($global:htmlsb),$xSummaryState,$htmlwhite))
            [string]$TagName = $(If( $xTags -is [array] -and $xTags.Count ) { $xTags[0] } Else { '-' } )
			$rowdata += @(,('Tags',($global:htmlsb),$TagName,$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
				}
			}

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

			If((!$LinuxVDA) -and $VDARegistryKeys -and $MachineIsOnline)
			{
				#First test if the Remote Registry service is enabled. If not skip the VDA registry keys
				$results = Get-Service -ComputerName $Machine.DNSName -Name "RemoteRegistry" -EA 0
				If($? -and $Null -ne $results)
				{
					If($results.Status -eq "Running")
					{
						GetVDARegistryKeys $Machine.DNSName "Desktop"
						$Script:HTMLAllVDARegistryItems += $Script:VDARegistryItems
						$Script:VDARegistryItems = New-Object System.Collections.ArrayList
					}
					Else
					{
						$obj1 = [PSCustomObject] @{
							RegKey       = "N/A"
							RegValue     = "N/A"
							VDAType      = "Desktop"
							ComputerName = $Machine.DNSName	
							Value        = "The Remote Registry service is not running"
						}
						$null = $Script:VDARegistryItems.Add($obj1)
						$Script:HTMLAllVDARegistryItems += $Script:VDARegistryItems
						$Script:VDARegistryItems = New-Object System.Collections.ArrayList
					}
				}
				Else
				{
					$obj1 = [PSCustomObject] @{
						RegKey       = "N/A"
						RegValue     = "N/A"
						VDAType      = "Desktop"
						ComputerName = $Machine.DNSName	
						Value        = "The Remote Registry service is not running"
					}
					$null = $Script:VDARegistryItems.Add($obj1)
					$Script:HTMLAllVDARegistryItems += $Script:VDARegistryItems
					$Script:VDARegistryItems = New-Object System.Collections.ArrayList
				}
			}
			ElseIf($LinuxVDA -and $VDARegistryKeys)
			{
				#VDA is Linux, skipping
			}
			
			WriteHTMLLine 4 0 "Machine Details"
			$rowdata = @()
			$columnHeaders = @("Agent Version",($global:htmlsb),$xAgentVersion,$htmlwhite)
			$rowdata += @(,('IP Address',($global:htmlsb),$xIPAddress,$htmlwhite))
			$rowdata += @(,('Is Assigned',($global:htmlsb),$Machine.IsAssigned.ToString(),$htmlwhite))
			$rowdata += @(,('OS Type',($global:htmlsb),$xOSType,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

			WriteHTMLLine 4 0 "Applications"
			$rowdata = @()
			[string]$AppsInUse = $(If( $xApplicationsInUse -is [array] -and $xApplicationsInUse.Count ) { $xApplicationsInUse[0] } Else { '-' } )
			$columnHeaders = @("Applications In Use",($global:htmlsb),$AppsInUse,$htmlwhite)
			$cnt = -1
			ForEach($tmp in $xApplicationsInUSe)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
				}
			}
			[string]$PubApps = $(If( $xPublishedApplications -is [array] -and $xPublishedApplications.Count ) { $xPublishedApplications[0] } Else { '-' } )
			$rowdata += @(,('Published Applications',($global:htmlsb),$PubApps,$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xPublishedApplications)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
				}
			}

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

			If($NoSessions -eq $False)
			{
				WriteHTMLLine 4 0 "Connection"
				$rowdata = @()
				$columnHeaders = @("Client (IP)",($global:htmlsb),$xSessionClientAddress,$htmlwhite)
				$rowdata += @(,('Client',($global:htmlsb),$xSessionClientName,$htmlwhite))
				$rowdata += @(,('Plug-in Version',($global:htmlsb),$xSessionClientVersion,$htmlwhite))
				$rowdata += @(,('Connected Via',($global:htmlsb),$xSessionConnectedViaHostName,$htmlwhite))
				$rowdata += @(,('Connect Via (IP)',($global:htmlsb),$xSessionConnectedViaIP,$htmlwhite))
				$rowdata += @(,('Last Connection Time',($global:htmlsb),$xLastConnectionTime,$htmlwhite))
				$rowdata += @(,('Last Connection User',($global:htmlsb),$xLastConnectionUser,$htmlwhite))
				$rowdata += @(,('Connection Type',($global:htmlsb),$xSessionProtocol,$htmlwhite))
				$rowdata += @(,('Secure ICA Active',($global:htmlsb),$xSessionSecureIcaActive,$htmlwhite))

				$msg = ""
				$columnWidths = @("200px","250px")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			}

			WriteHTMLLine 4 0 "Registration"
			$rowdata = @()
			$columnHeaders = @("Broker",($global:htmlsb),$xBroker,$htmlwhite)
			$rowdata += @(,('Last registration failure',($global:htmlsb),$xLastDeregistrationReason,$htmlwhite))
			$rowdata += @(,('Last registration failure time',($global:htmlsb),$xLastDeregistrationTime,$htmlwhite))
			$rowdata += @(,('Registration State',($global:htmlsb),$Machine.RegistrationState.ToString(),$htmlwhite))
			$rowdata += @(,('Fault State',($global:htmlsb),$xMachineFaultState,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

			WriteHTMLLine 4 0 "Hosting"
			$rowdata = @()
			$columnHeaders = @("VM",($global:htmlsb),$xHostedMachineName,$htmlwhite)
			$rowdata += @(,('Hosting Server Name',($global:htmlsb),$xHostingServerName,$htmlwhite))
			$rowdata += @(,('Connection',($global:htmlsb),$xHypervisorConnectionName,$htmlwhite))
			$rowdata += @(,('Pending Update',($global:htmlsb),$Machine.ImageOutOfDate.ToString(),$htmlwhite))
			$rowdata += @(,('Persist User Changes',($global:htmlsb),$xPersistUserChanges,$htmlwhite))
			$rowdata += @(,('Power Action Pending',($global:htmlsb),$Machine.PowerActionPending.ToString(),$htmlwhite))
			$rowdata += @(,('Power State',($global:htmlsb),$xPowerState,$htmlwhite))
			$rowdata += @(,('Will Shutdown After Use',($global:htmlsb),$xWillShutdownAfterUse,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

			If($NoSessions -eq $False)
			{
				WriteHTMLLine 4 0 "Session Details"
				$rowdata = @()
				$columnHeaders = @("Launched Via",($global:htmlsb),$xSessionLaunchedViaHostName,$htmlwhite)
				$rowdata += @(,('Launched Via (IP)',($global:htmlsb),$xSessionLaunchedViaIP,$htmlwhite))
				$rowdata += @(,('Session Change Time',($global:htmlsb),$xSessionStateChangeTime,$htmlwhite))
				[string]$SSAT = $(If( $xSessionSmartAccessTags -is [array] -and $xSessionSmartAccessTags.Count ) { $xSessionSmartAccessTags[0] } Else { '-' } )
				$rowdata += @(,('SmartAccess Filters',($global:htmlsb),$SSAT,$htmlwhite))
				$cnt = -1
				ForEach($tmp in $xSessionSmartAccessTags)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("200px","250px")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"

				WriteHTMLLine 4 0 "Session"
				$rowdata = @()
				$columnHeaders = @("Session State",($global:htmlsb),$xSessionState,$htmlwhite)
				$rowdata += @(,('Current User',($global:htmlsb),$xSessionUserName,$htmlwhite))
				$rowdata += @(,('Start Time',($global:htmlsb),$xSessionStateChangeTime,$htmlwhite))

				$msg = ""
				$columnWidths = @("200px","250px")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			}
		}
	}
}
#endregion

#region Delivery Group functions
Function ProcessDeliveryGroups
{
	Write-Verbose "$(Get-Date -Format G): Retrieving Delivery Groups"
	
	$txt = "Delivery Groups"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$AllDeliveryGroups = Get-BrokerDesktopGroup @CCParams2 -SortBy Name 

	If($? -and ($Null -ne $AllDeliveryGroups))
	{
		Write-Verbose "$(Get-Date -Format G): `tProcessing Delivery Groups"
		
		OutputDeliveryGroupTable $AllDeliveryGroups
		
		ForEach($Group in $AllDeliveryGroups)
		{
			OutputDeliveryGroup $Group
		}
	}
	ElseIf($? -and ($Null -eq $AllDeliveryGroups))
	{
		$txt = "There are no Delivery Groups"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve Delivery Groups"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date -Format G): "
}

Function OutputDeliveryGroupTable 
{
	Param([object] $AllDeliveryGroups)
	
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $WordTable = @();
	}
	If($Text)
	{
		Line 1 "                                                                                                           No. of   Sessions Machine                                                                "
		Line 1 "Delivery Group                                                                   Delivering                Machines in use   type                                          Unregistered Disconnected"
		Line 1 "============================================================================================================================================================================================================="
		#       12345678901234567890123456789012345678901234567890123456789012345678901234567890S1234567890123456789012345S12345678S12345678S123456789012345678901234567890123456789012345S123456789012S123456789012
		#                                                                                        Applications and Desktops 99999999 99999999 Single-session OS (Static machine assignment) 999999999999 999999999999
	}
	If($HTML)
	{
		$rowdata = @()
	}
	
	ForEach($Group in $AllDeliveryGroups)
	{
		[string]$SessionSupport    = ""
		[string]$xState            = ""
		[string]$xDeliveryType     = ""
		[string]$xGroupName        = ""
		[int]$NumApps              = 0
		[int]$NumAppGroups         = 0
		[int]$NumDesktops          = 0
		
		$SessionSupport = "Single-session OS"
		If($Group.SessionSupport -eq "SingleSession")
		{
			$SessionSupport = "Single-session OS"
		}
		Else
		{
			$SessionSupport = "Multi-session OS"
		}
		
		If($Group.InMaintenanceMode)
		{
			$xState = "(Maint) "
		}
		
		$xGroupName = "$($xState)$($Group.Name)"
		
		If($Group.DesktopKind -eq "Private")
		{
			$SessionSupport += " (Static machine assignment)"
		}
		
		$NumApps      = (@(Get-BrokerApplication @CCParams2 -DesktopGroupUid $Group.Uid)).Count
		$NumAppGroups = (@(Get-BrokerApplicationGroup @CCParams2 -DesktopGroupUid $Group.Uid)).Count
		$NumDesktops  = (@(Get-BrokerEntitlementPolicyRule @CCParams2 -DesktopGroupUid $Group.Uid)).Count

		If($NumApps -gt 0 -or $NumAppGroups -gt 0 -and $NumDesktops -eq 0)
		{
			$xDeliveryType = "Applications"
			$Script:TotalApplicationGroups++
		}
		ElseIf($NumApps -eq 0 -and $NumAppGroups -eq 0 -and $NumDesktops -gt 0)
		{
			$xDeliveryType = "Desktops"
			$Script:TotalDesktopGroups++
		}
		ElseIf($NumApps -gt 0 -or $NumAppGroups -gt 0 -and $NumDesktops -gt 0)
		{
			$xDeliveryType = "Applications and Desktops"
			$Script:TotalAppsAndDesktopGroups++
		}
		Else
		{
			$xDeliveryType = "Delivery type could not be determined: Apps($NumApps) AppGroups($NumAppGroups) Desktops($NumDesktops)"
		}

		If($Group.DeliveryType -eq "DesktopsOnly" -and $Group.DesktopKind -eq "Private")
		{
			$xDeliveryType = "Desktops"
		}
		
		If($MSWord -or $PDF)
		{
			$WordTable += @{
			DeliveryGroupName = $xGroupName; 
			DeliveryType = $xDeliveryType
			NoOfMachines = $Group.TotalDesktops; 
			SessionsInUse = $Group.Sessions; 
			MachineType = $SessionSupport; 
			Unregistered = $Group.DesktopsUnregistered; 
			Disconnected = $Group.DesktopsDisconnected; 
			}
		}
		If($Text)
		{
			Line 1 ( "{0,-80} {1,-25} {2,8} {3,8} {4,-45} {5,12} {6,12}" -f `
			$xGroupName, $xDeliveryType, $Group.TotalDesktops, $Group.Sessions, `
			$SessionSupport, $Group.DesktopsUnregistered, $Group.DesktopsDisconnected)
		}
		If($HTML)
		{
			$rowdata += @(,(
			$xGroupName,$htmlwhite,
			$xDeliveryType,$htmlwhite,
			$Group.TotalDesktops.ToString(),$htmlwhite,
			$Group.Sessions.ToString(),$htmlwhite,
			$SessionSupport,$htmlwhite,
			$Group.DesktopsUnregistered.ToString(),$htmlwhite,
			$Group.DesktopsDisconnected.ToString(),$htmlwhite))
		}
	}
	
	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $WordTable `
		-Columns  DeliveryGroupName, DeliveryType, NoOfMachines, SessionsInUse, MachineType, Unregistered, Disconnected `
		-Headers  "Delivery Group", "Delivering", "No. of machines", "Sessions in use", "Machine type", "Unregistered", "Disconnected" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table -Size 8 -BackgroundColor $wdColorWhite
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 110;
		$Table.Columns.Item(2).Width = 95;
		$Table.Columns.Item(3).Width = 43;
		$Table.Columns.Item(4).Width = 40;
		$Table.Columns.Item(5).Width = 58;
		$Table.Columns.Item(6).Width = 55;
		$Table.Columns.Item(7).Width = 57;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		$columnHeaders = @(
		'Delivery Group',($global:htmlsb),
		'Delivering',($global:htmlsb),
		'No. of machines',($global:htmlsb),
		'Sessions in use',($global:htmlsb),
		'Machine type',($global:htmlsb),
		'Unregistered',($global:htmlsb),
		'Disconnected',($global:htmlsb)
		)

		$columnWidths = @("150","165","50","45","165","60","65")
		$msg = ""
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "700"
	}
}

Function OutputDeliveryGroup
{
	Param([object] $Group)
	
	Write-Verbose "$(Get-Date -Format G): `t`t`tAdding Delivery Group $($Group.Name)"
	$SessionSupport = ""
	$xState = ""

	If($Group.SessionSupport -eq "SingleSession")
	{
		$SessionSupport = "Single-session OS"
	}
	Else
	{
		$SessionSupport = "Multi-session OS"
	}

	If($Group.DesktopKind -eq "Private")
	{
		$SessionSupport += " (Static machine assignment)"
	}
	
	If($Group.Enabled -eq $True -and $Group.InMaintenanceMode -eq $True)
	{
		$xState = "Maintenance Mode"
	}
	ElseIf($Group.Enabled -eq $False -and $Group.InMaintenanceMode -eq $True)
	{
		$xState = "Maintenance Mode"
	}
	ElseIf($Group.Enabled -eq $True -and $Group.InMaintenanceMode -eq $False)
	{
		$xState = "Enabled"
	}
	ElseIf($Group.Enabled -eq $False -and $Group.InMaintenanceMode -eq $False)
	{
		$xState = "Disabled"
	}

	If($MSWord -or$PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 2 0 "Delivery Group: " $Group.Name
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Machine type"; Value = $SessionSupport; }) > $Null
		$ScriptInformation.Add(@{Data = "Number of machines"; Value = $Group.TotalDesktops; }) > $Null
		$ScriptInformation.Add(@{Data = "Sessions in use"; Value = $Group.Sessions; }) > $Null
		$ScriptInformation.Add(@{Data = "Number of applications"; Value = $Group.TotalApplications; }) > $Null
		$ScriptInformation.Add(@{Data = "State"; Value = $xState; }) > $Null
		$ScriptInformation.Add(@{Data = "Unregistered"; Value = $Group.DesktopsUnregistered; }) > $Null
		$ScriptInformation.Add(@{Data = "Disconnected"; Value = $Group.DesktopsDisconnected; }) > $Null
		$ScriptInformation.Add(@{Data = "Available"; Value = $Group.DesktopsAvailable; }) > $Null
		$ScriptInformation.Add(@{Data = "In Use"; Value = $Group.DesktopsInUse; }) > $Null
		$ScriptInformation.Add(@{Data = "Never Registered"; Value = $Group.DesktopsNeverRegistered; }) > $Null
		$ScriptInformation.Add(@{Data = "Preparing"; Value = $Group.DesktopsPreparing; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "Delivery Group: " $Group.Name
		Line 1 "Machine type`t`t: " $SessionSupport
		Line 1 "No. of machines`t`t: " $Group.TotalDesktops
		Line 1 "Sessions in use`t`t: " $Group.Sessions
		Line 1 "No. of applications`t: " $Group.TotalApplications
		Line 1 "State`t`t`t: " $xState
		Line 1 "Unregistered`t`t: " $Group.DesktopsUnregistered
		Line 1 "Disconnected`t`t: " $Group.DesktopsDisconnected
		Line 1 "Available`t`t: " $Group.DesktopsAvailable
		Line 1 "In Use`t`t`t: " $Group.DesktopsInUse
		Line 1 "Never Registered`t: " $Group.DesktopsNeverRegistered
		Line 1 "Preparing`t`t: " $Group.DesktopsPreparing
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		WriteHTMLLine 2 0 "Delivery Group: " $Group.Name
		$columnHeaders = @("Machine type",($global:htmlsb),$SessionSupport,$htmlwhite)
		$rowdata += @(,('No. of machines',($global:htmlsb),$Group.TotalDesktops.ToString(),$htmlwhite))
		$rowdata += @(,('Sessions in use',($global:htmlsb),$Group.Sessions.ToString(),$htmlwhite))
		$rowdata += @(,('No. of applications',($global:htmlsb),$Group.TotalApplications.ToString(),$htmlwhite))
		$rowdata += @(,('State',($global:htmlsb),$xState,$htmlwhite))
		$rowdata += @(,('Unregistered',($global:htmlsb),$Group.DesktopsUnregistered.ToString(),$htmlwhite))
		$rowdata += @(,('Disconnected',($global:htmlsb),$Group.DesktopsDisconnected.ToString(),$htmlwhite))
		$rowdata += @(,('Available',($global:htmlsb),$Group.DesktopsAvailable.ToString(),$htmlwhite))
		$rowdata += @(,('In Use',($global:htmlsb),$Group.DesktopsInUse.ToString(),$htmlwhite))
		$rowdata += @(,('Never Registered',($global:htmlsb),$Group.DesktopsNeverRegistered.ToString(),$htmlwhite))
		$rowdata += @(,('Preparing',($global:htmlsb),$Group.DesktopsPreparing.ToString(),$htmlwhite))

		$msg = ""
		$columnWidths = @("200","275")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "475"
	}
	
	If($DeliveryGroups)
	{
		Write-Verbose "$(Get-Date -Format G): `t`tProcessing details"
		$txt = "Delivery Group Details: "
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 $txt $Group.Name
		}
		If($Text)
		{
			Line 0 $txt $Group.Name
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 $txt $Group.Name
		}
		OutputDeliveryGroupDetails $Group
		
		Write-Verbose "$(Get-Date -Format G): `t`tProcessing applications"
		OutputDeliveryGroupApplicationDetails $Group

		#retrieve machines in delivery group
		$Machines = Get-BrokerMachine -DesktopGroupName $Group.name @CCParams2 -SortBy DNSName
		If($? -and $Null -ne $Machines)
		{
			#if both -MachineCatalogs and -DeliveryGroups parameters are used, only output the machine details for catalogs, not delivery groups
			If($MachineCatalogs -and $DeliveryGroups)
			{
				#do not do machine details for delivery groups if both -MachineCatalogs and -DeliveryGroups parameters are used
			}
			ElseIf($DeliveryGroups -and -not $MachineCatalogs)
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "Desktops"
				}
				If($Text)
				{
					Line 0 "Desktops"
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "Desktops"
				}
				
				ForEach($Machine in $Machines)
				{
					OutputMachineDetails $Machine
				}
			}
			ElseIf(-not $DeliveryGroups -and $MachineCatalogs)
			{
				#shouldn't be here
#Do not indent the following write-error lines. Doing so will mess up the console formatting of the error message.
				Write-Error "
				`n`n
	OOPS!!! An error that should not occur has occured in Function OutputDeliveryGroup.
				`n`n
	Please send an email to webster@carlwebster.com
				`n`n
				"
			}
		}
		ElseIf($? -and $Null -eq $Machines)
		{
			$txt = "There are no Machines for Delivery Group $($Group.name)"
			OutputNotice $txt
		}
		Else
		{
			$txt = "Unable to retrieve Machines for Delivery Group $($Group.name)"
			OutputWarning $txt
		}

		Write-Verbose "$(Get-Date -Format G): `t`tProcessing machine catalogs"
		OutputDeliveryGroupCatalogs $Group

		If($DeliveryGroupsUtilization)
		{
			Write-Verbose "$(Get-Date -Format G): `t`t`tCreating Delivery Group Utilization report"
			OutputDeliveryGroupUtilization $Group
		}

		Write-Verbose "$(Get-Date -Format G): `t`tProcessing Tags"
		OutputDeliveryGroupTags $Group

		Write-Verbose "$(Get-Date -Format G): `t`tProcessing Application Groups"
		OutputDeliveryGroupApplicationGroups $Group

		Write-Verbose "$(Get-Date -Format G): `t`tProcessing administrators"
		$Admins = GetAdmins "DesktopGroup" $Group.Name
		
		If($? -and ($Null -ne $Admins))
		{
			OutputAdminsForDetails $Admins
		}
		ElseIf($? -and ($Null -eq $Admins))
		{
			$txt = "There are no administrators for $($Group.Name)"
			OutputNotice $txt
		}
		Else
		{
			$txt = "Unable to retrieve administrators for $($Group.Name)"
			OutputWarning $txt
		}
	}
	
	If($DeliveryGroupsUtilization)
	{
		Write-Verbose "$(Get-Date -Format G): `t`t`tCreating Delivery Group Utilization report"
		OutputDeliveryGroupUtilization $Group
	}
}

Function OutputDeliveryGroupDetails 
{
	Param([object] $Group)

	[string]$xDGType                   = "Delivery Group Type cannot be determined: $($Group.DeliveryType) $($Group.DesktopKind)"
	[string]$xSessionReconnection      = ""
	[string]$xSecureIcaRequired        = "Use delivery group setting"
	[string]$xVDAVersion               = ""
	[string]$xDeliveryType             = ""
	[string]$xColorDepth               = ""
	[string]$xShutdownDesktopsAfterUse = "No"
	[string]$xTurnOnAddedMachine       = "No"
	$DGIncludedUsers                   = @()
	$DGExcludedUsers                   = @()
	$DGScopes                          = @()
	$DGSFServers                       = @()
	[string]$xSessionPrelaunch         = "Off"
	[int]$xSessionPrelaunchAvgLoad     = 0
	[int]$xSessionPrelaunchAnyLoad     = 0
	[string]$xSessionLinger            = "Off"
	[int]$xSessionLingerAvgLoad        = 0
	[int]$xSessionLingerAnyLoad        = 0
	[string]$xEndPrelaunchSession      = ""
	[string]$xEndLinger                = ""
	[bool]$PwrMgmt1                    = $False
	[bool]$PwrMgmt2                    = $False
	[bool]$PwrMgmt3                    = $False
	[string]$xUsersHomeZone            = "No"
	
	If($Group.DeliveryType -eq "AppsOnly" -and $Group.DesktopKind -eq "Shared")
	{
		$xDGType = "Random Applications"
	}
	ElseIf($Group.DeliveryType -eq "DesktopsOnly" -and $Group.DesktopKind -eq "Shared")
	{
		$xDGType = "Random Desktops"
	}
	ElseIf($Group.DeliveryType -eq "DesktopsOnly" -and $Group.DesktopKind -eq "Private")
	{
		$xDGType = "Static Desktops"
		$xDeliveryType = "Desktops"
	}
	ElseIf($Group.DeliveryType -eq "DesktopsAndApps" -and $Group.DesktopKind -eq "Shared")
	{
		$xDGType = "Random Desktops and applications"
	}
	
	$NumApps      = (@(Get-BrokerApplication @CCParams2 -DesktopGroupUid $Group.Uid)).Count
	$NumAppGroups = (@(Get-BrokerApplicationGroup @CCParams2 -DesktopGroupUid $Group.Uid)).Count
	$NumDesktops  = (@(Get-BrokerEntitlementPolicyRule @CCParams2 -DesktopGroupUid $Group.Uid)).Count

	If($NumApps -gt 0 -or $NumAppGroups -gt 0 -and $NumDesktops -eq 0)
	{
		$xDeliveryType = "Applications"
	}
	ElseIf($NumApps -eq 0 -and $NumAppGroups -eq 0 -and $NumDesktops -gt 0)
	{
		$xDeliveryType = "Desktops"
	}
	ElseIf($NumApps -gt 0 -or $NumAppGroups -gt 0 -and $NumDesktops -gt 0)
	{
		$xDeliveryType = "Applications and Desktops"
	}
	Else
	{
		$xDeliveryType = "Delivery type could not be determined: Apps($NumApps) AppGroups($NumAppGroups) Desktops($NumDesktops)"
	}
	
	If([String]::IsNullOrEmpty($Group.LicenseModel))
	{
		$LicenseModel = "Site Default"
	}
	Else
	{
		$LicenseModel = $Group.LicenseModel
	}
	
	If([String]::IsNullOrEmpty($Group.ProductCode))
	{
		$ProductCode = "Site Default"
	}
	Else
	{
		$ProductCode = $Group.ProductCode
	}

	Switch ($Group.MinimumFunctionalLevel)
	{
		"L5" 	{$xVDAVersion = "5.6 FP1 (Windows XP and Windows Vista)"; Break}
		"L7"	{$xVDAVersion = "7.0 (or newer)"; Break}
		"L7_6"	{$xVDAVersion = "7.6 (or newer)"; Break}
		"L7_7"	{$xVDAVersion = "7.7 (or newer)"; Break}
		"L7_8"	{$xVDAVersion = "7.8 (or newer)"; Break}
		"L7_9"	{$xVDAVersion = "7.9 (or newer)"; Break}
		"L7_20"	{$xVDAVersion = "1811 (or newer)"; Break}
		"L7_25"	{$xVDAVersion = "2003 (or newer)"; Break}
		Default {$xVDAVersion = "Unable to determine VDA version: $($Group.MinimumFunctionalLevel)"; Break}
	}
	
	Switch ($Group.ColorDepth)
	{
		"FourBit"		{$xColorDepth = "4bit - 16 colors"; Break}
		"EightBit"		{$xColorDepth = "8bit - 256 colors"; Break}
		"SixteenBit"	{$xColorDepth = "16bit - High color"; Break}
		"TwentyFourBit"	{$xColorDepth = "24bit - True color"; Break}
		Default			{$xColorDepth = "Unable to determine Color Depth: $($Group.ColorDepth)"; Break}
	}
	
	If($Group.ShutdownDesktopsAfterUse)
	{
		$xShutdownDesktopsAfterUse = "Yes"
	}
	
	If($Group.TurnOnAddedMachine)
	{
		$xTurnOnAddedMachine = "Yes"
	}

	ForEach($Scope in $Group.Scopes)
	{
		$DGScopes += $Scope
	}
	$DGScopes += "All"
	
	ForEach($Server in $Group.MachineConfigurationNames)
	{
		$SFTmp = Get-BrokerMachineConfiguration -Name $Server
		If($? -and $Null -ne $SFTmp)
		{
			$SFByteArray = $SFTmp.Policy
			## GRL add Try/Catch
            try
            {
			    $SFServer = Get-SFStoreFrontAddress -ByteArray $SFByteArray -ErrorAction SilentlyContinue 4>$Null
			    If($? -and $Null -ne $SFServer)
			    {
				    $DGSFServers += $SFServer.Url
			    }
            }
            catch
            {
                Write-Warning -Message "Failed call to Get-SFStoreFrontAddress for $server"
            }
		}
	}
	
	If($DGSFServers.Count -eq 0)
	{
		$DGSFServers += "-"
	}

	$test = Get-BrokerSessionPreLaunch -EA 0
	If($? -and $null -ne $test)
	{
		$SPLUIDs = @()
		ForEach($SPL in $test)
		{
			$SPLUIDs += $SPL.DesktopGroupUid
		}
		If($SPLUIDs -contains $Group.Uid)
		{
			$Results = Get-BrokerSessionPreLaunch -DesktopGroupUid $Group.Uid -EA 0
			If($? -and $Null -ne $Results)
			{
				If($Results.Enabled -and $Results.AssociatedUserFullNames.Count -eq 0)
				{
					$xSessionPrelaunch = "Prelaunch for any user"
				}
				ElseIf($Results.Enabled -and $Results.AssociatedUserFullNames.Count -gt 0)
				{
					$xSessionPrelaunch = "Prelaunch for specific users"
				}
				
				If($Results.MaxAverageLoadThreshold -gt 0)
				{
					$xSessionPrelaunchAvgLoad = ($Results.MaxAverageLoadThreshold/100)
				}
				If($Results.MaxLoadPerMachineThreshold -gt 0)
				{
					$xSessionPrelaunchAnyLoad = ($Results.MaxLoadPerMachineThreshold/100)
				}
				$Mins = $Results.MaxTimeBeforeTerminate.Minutes
				$Hours = $Results.MaxTimeBeforeTerminate.Hours
				$Days = $Results.MaxTimeBeforeTerminate.Days
				If($Mins -gt 0)
				{
					$xEndPrelaunchSession = "$($Mins) Minutes"
				}
				If($Hours -gt 0)
				{
					$xEndPrelaunchSession = "$($Hours) Hours"
				}
				ElseIf($Days -gt 0)
				{
					$xEndPrelaunchSession = "$($Days) Days"
				}
			}
		}
	}
	
	$test = Get-BrokerSessionLinger -EA 0
	If($? -and $null -ne $test)
	{
		$SLUIDs = @()
		ForEach($SL in $test)
		{
			$SLUIDs += $SL.DesktopGroupUid
		}
		If($SLUIDs -contains $Group.Uid)
		{
			$Results = Get-BrokerSessionLinger -DesktopGroupUid $Group.Uid -EA 0
			If($? -and $Null -ne $Results)
			{
				$xSessionLinger = "Keep session active"
				If($Results.MaxAverageLoadThreshold -gt 0)
				{
					$xSessionLingerAvgLoad = ($Results.MaxAverageLoadThreshold/100)
				}
				If($Results.MaxLoadPerMachineThreshold -gt 0)
				{
					$xSessionLingerAnyLoad = ($Results.MaxLoadPerMachineThreshold/100)
				}
				$Mins = $Results.MaxTimeBeforeTerminate.Minutes
				$Hours = $Results.MaxTimeBeforeTerminate.Hours
				$Days = $Results.MaxTimeBeforeTerminate.Days
				If($Mins -gt 0)
				{
					$xEndLinger = "$($Mins) Minutes"
				}
				If($Hours -gt 0)
				{
					$xEndLinger = "$($Hours) Hours"
				}
				ElseIf($Days -gt 0)
				{
					$xEndLinger = "$($Days) Days"
				}
			}
		}
	}

	If($Group.ZonePreferences -Contains "UserHomeOnly")
	{
		$xUsersHomeZone = "Yes, if configured"
	}
	
	#get a desktop in an associated delivery group to get the catalog
	$Desktop = Get-BrokerMachine @CCParams2 -DesktopGroupUid $Group.Uid -Property CatalogName
	
	If($? -and $Null -ne $Desktop)
	{
		$Catalog = Get-BrokerCatalog @CCParams2 -Name $Desktop[0].CatalogName
		
		If($? -and $Null -ne $Catalog)
		{
			If($Catalog.AllocationType -eq "Static" -and $Catalog.PersistUserChanges -eq "Discard" -and $Group.DesktopKind -eq "Private" -and $Group.SessionSupport -eq "SingleSession")
			{
				$PwrMgmt1 = $True
				$PwrMgmt2 = $False
				$PwrMgmt3 = $False
			}
			ElseIf($Catalog.AllocationType -eq "Static" -and $Catalog.PersistUserChanges -ne "Discard" -and $Group.DesktopKind -eq "Private" -and $Group.SessionSupport -eq "SingleSession")
			{
				$PwrMgmt1 = $False
				$PwrMgmt2 = $True
				$PwrMgmt3 = $False
			}
			ElseIf($Catalog.AllocationType -eq "Random" -and $Catalog.PersistUserChanges -eq "Discard" -and $Group.DesktopKind -eq "Shared" -and $Group.SessionSupport -eq "SingleSession")
			{
				$PwrMgmt1 = $False
				$PwrMgmt2 = $False
				$PwrMgmt3 = $True
			}
		}
	}

	If($PwrMgmt2 -or $PwrMgmt3)
	{
		$PwrMgmts = Get-BrokerPowerTimeScheme @CCParams2 -DesktopGroupUid $Group.Uid 
	}
	
	$xOffPeakBufferSizePercent           = $Group.OffPeakBufferSizePercent
	$xOffPeakDisconnectTimeout           = $Group.OffPeakDisconnectTimeout
	$xOffPeakExtendedDisconnectTimeout   = $Group.OffPeakExtendedDisconnectTimeout
	$xOffPeakLogOffTimeout               = $Group.OffPeakLogOffTimeout
	$xPeakBufferSizePercent              = $Group.PeakBufferSizePercent
	$xPeakDisconnectTimeout              = $Group.PeakDisconnectTimeout
	$xPeakExtendedDisconnectTimeout      = $Group.PeakExtendedDisconnectTimeout
	$xPeakLogOffTimeout                  = $Group.PeakLogOffTimeout
	$xSettlementPeriodBeforeAutoShutdown = $Group.SettlementPeriodBeforeAutoShutdown
	$xSettlementPeriodBeforeUse          = $Group.SettlementPeriodBeforeUse
	$xOffPeakDisconnectAction            = ""
	$xOffPeakExtendedDisconnectAction    = ""
	$xOffPeakLogOffAction                = ""
	$xPeakDisconnectAction               = ""
	$xPeakExtendedDisconnectAction       = ""
	$xPeakLogOffAction                   = ""

	Switch ($Group.OffPeakDisconnectAction)
	{
		"Nothing"	{ $xOffPeakDisconnectAction = "No action"; Break}
		"Suspend"	{ $xOffPeakDisconnectAction = "Suspend"; Break}
		"Shutdown"	{ $xOffPeakDisconnectAction = "Shut down"; Break}
		Default		{ $xOffPeakDisconnectAction = "Unable to determine the OffPeakDisconnectAction action: $($Group.OffPeakDisconnectAction)"; Break}
	}
	
	Switch ($Group.OffPeakExtendedDisconnectAction)
	{
		"Nothing"	{ $xOffPeakExtendedDisconnectAction = "No action"; Break}
		"Suspend"	{ $xOffPeakExtendedDisconnectAction = "Suspend"; Break}
		"Shutdown"	{ $xOffPeakExtendedDisconnectAction = "Shut down"; Break}
		Default		{ $xOffPeakExtendedDisconnectAction = "Unable to determine the OffPeakExtendedDisconnectAction action: $($Group.OffPeakExtendedDisconnectAction)"; Break}
	}
	
	Switch ($Group.OffPeakLogOffAction)
	{
		"Nothing"	{ $xOffPeakLogOffAction = "No action"; Break}
		"Suspend"	{ $xOffPeakLogOffAction = "Suspend"; Break}
		"Shutdown"	{ $xOffPeakLogOffAction = "Shut down"; Break}
		Default		{ $xOffPeakLogOffAction = "Unable to determine $xOffPeakLogOffAction action: $($Group.OffPeakLogOffAction)"; Break}
	}
	
	Switch ($Group.PeakDisconnectAction)
	{
		"Nothing"	{ $xPeakDisconnectAction = "No action"; Break}
		"Suspend"	{ $xPeakDisconnectAction = "Suspend"; Break}
		"Shutdown"	{ $xPeakDisconnectAction = "Shut down"; Break}
		Default		{ $xPeakDisconnectAction = "Unable to determine $xPeakDisconnectAction action: $($Group.PeakDisconnectAction)"; Break}
	}
	
	Switch ($Group.PeakExtendedDisconnectAction)
	{
		"Nothing"	{ $xPeakExtendedDisconnectAction = "No action"; Break}
		"Suspend"	{ $xPeakExtendedDisconnectAction = "Suspend"; Break}
		"Shutdown"	{ $xPeakExtendedDisconnectAction = "Shut down"; Break}
		Default		{ $xPeakExtendedDisconnectAction = "Unable to determine $xPeakExtendedDisconnectAction action: $($Group.PeakExtendedDisconnectAction)"; Break}
	}
	
	Switch ($Group.PeakLogOffAction)
	{
		"Nothing"	{ $xPeakLogOffAction = "No action"; Break}
		"Suspend"	{ $xPeakLogOffAction = "Suspend"; Break}
		"Shutdown"	{ $xPeakLogOffAction = "Shut down"; Break}
		Default		{ $xPeakLogOffAction = "Unable to determine $xPeakLogOffAction action: $($Group.PeakLogOffAction)"; Break}
	}

	$xEnabled = "Disabled"
	If($Group.Enabled)
	{
		$xEnabled = "Enabled"
	}

	$xSecureICA = "Disabled"
	If($Group.SecureICARequired)
	{
		$xSecureICA = "Enabled"
	}
	
	$xAutoPowerOnForAssigned = "Disabled"
	$xAutoPowerOnForAssignedDuringPeak = "Disabled"
	
	If($Group.AutomaticPowerOnForAssigned)
	{
		$xAutoPowerOnForAssigned = "Enabled"
	}
	If($Group.AutomaticPowerOnForAssignedDuringPeak)
	{
		$xAutoPowerOnForAssignedDuringPeak = "Enabled"
	}

	$SFAnonymousUsers = $False
	$Results = Get-BrokerAccessPolicyRule -DesktopGroupUid $Group.Uid @CCParams2
	
	If($? -and $Null -ne $Results)
	{
		ForEach($Result in $Results)
		{
			
			If($Result.AllowedUsers -eq "Any" -or $Result.AllowedUsers -eq "FilteredOrAnonymous" -or $Result.AllowedUsers -eq "AnonymousOnly")
			{
				$SFAnonymousUsers = $True
			}
			
			If($Result.IncludedUserFilterEnabled -and $Result.AllowedUsers -eq "Filtered")
			{
				ForEach($User in $Result.IncludedUsers)
				{
					$DGIncludedUsers += $User.Name
				}
			}
			ElseIf($Result.IncludedUserFilterEnabled -and ($Result.AllowedUsers -eq "AnyAuthenticated" -or $Result.AllowedUsers -eq "Any"))
			{
				$DGIncludedUsers += "Allow any authenticated users to use this Delivery Group"
			}
			
			If($Result.ExcludedUserFilterEnabled)
			{
				ForEach($User in $Result.ExcludedUsers)
				{
					$DGExcludedUsers += $User.Name
				}
			}
			
			If($Result.Name -like '*_AG')
			{
				If($Result.AllowedConnections -eq "ViaAG" -and $Result.IncludedSmartAccessFilterEnabled -eq $False -and $Result.Enabled -eq $True)
				{
					$xAllConnections = "Enabled"
					$xNSConnection = "Disabled"
					$xAGFilters = @()
					$xAGFilters += "N/A"
				}
				ElseIf($Result.AllowedConnections -eq "ViaAG" -and $Result.IncludedSmartAccessFilterEnabled -eq $True -and $Result.Enabled -eq $True)
				{
					$xAllConnections = "Enabled"
					$xNSConnection = "Enabled"
					$xAGFilters = @()
					ForEach($AccessCondition in $Result.IncludedSmartAccessTags)
					{
						$xAGFilters += $AccessCondition
					}
					If($xAGFilters.Count -eq 0)
					{
						$xAGFilters += "None"
					}
				}
				ElseIf($Result.AllowedConnections -eq "ViaAG" -and $Result.IncludedSmartAccessFilterEnabled -eq $False -and $Result.Enabled -eq $False)
				{
					$xAllConnections = "Disabled"
					$xNSConnection = "Disabled"
					$xAGFilters = @()
					$xAGFilters += "N/A"
				}
			}
			Else
			{
				$xAllConnections = ""
				$xNSConnection = ""
				$xAGFilters = @()
			}
		}
		
		[array]$DGIncludedUsers = $DGIncludedUsers | Sort-Object -unique
		[array]$DGExcludedUsers = $DGExcludedUsers | Sort-Object -unique
	}
	
	#desktops per user for singlesession OS
	If($Group.SessionSupport -eq "SingleSession")
	{
		If($xDGType -eq "Static Desktops")
		{
			#static desktops have a maxdesktops count stored as a property
			$xMaxDesktops = 0
			$MaxDesktops = Get-BrokerAssignmentPolicyRule @CCParams2 -DesktopGroupUid $Group.Uid
			
			If($? -and $Null -ne $MaxDesktops)
			{
				$xMaxDesktops = $MaxDesktops.MaxDesktops
			}
		}
		ElseIf($xDGType -like "*Random*")
		{
			#random desktops are a count of the number of entitlement policy rules
			$xMaxDesktops = 0
            ## GRL various cmdlets don't return arrays if there is 1 item so put in @( ) to always force to be an array if you are going to treat it as such or test with "if( $maxdesktops -is [array] )"
			$MaxDesktops = @( Get-BrokerEntitlementPolicyRule @CCParams2 -DesktopGroupUid $Group.Uid )
			
			If($? -and $Null -ne $MaxDesktops)
			{
				$xMaxDesktops = $MaxDesktops.Count
			}
		}
	}

	If([String]::IsNullOrEmpty($Group.TimeZone))
	{
		$xTimeZone = "Not Configured"
	}
	Else
	{
		$xTimeZone = $Group.TimeZone
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Details: " $Group.Name
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Description"; Value = $Group.Description; }
		If(![String]::IsNullOrEmpty($Group.PublishedName))
		{
			$ScriptInformation += @{Data = "Display Name"; Value = $Group.PublishedName; }
		}
		$ScriptInformation += @{Data = "Type"; Value = $xDGType; }
		$ScriptInformation += @{Data = "Set to VDA version"; Value = $xVDAVersion; }
		If($Group.SessionSupport -eq "SingleSession" -and ($xDGType -eq "Static Desktops" -or $xDGType -like "*Random*"))
		{
			$ScriptInformation += @{Data = "Desktops per user"; Value = $xMaxDesktops; }
		}
		$ScriptInformation += @{Data = "Time zone"; Value = $xTimeZone; }
		$ScriptInformation += @{Data = "Enable Delivery Group"; Value = $xEnabled; }
		$ScriptInformation += @{Data = "Enable Secure ICA"; Value = $xSecureICA; }
		$ScriptInformation += @{Data = "Color Depth"; Value = $xColorDepth; }
		$ScriptInformation += @{Data = "Shutdown Desktops After Use"; Value = $xShutdownDesktopsAfterUse; }
		$ScriptInformation += @{Data = "Turn On Added Machine"; Value = $xTurnOnAddedMachine; }
		[string]$DGIU = $(If( $DGIncludedUsers -is [array] -and $DGIncludedUsers.Count ) { $DGIncludedUsers[0] } Else { '-' } )
		$ScriptInformation += @{Data = "Included Users"; Value = $DGIU; }
		$cnt = -1
		ForEach($tmp in $DGIncludedUsers)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		
		If($DGExcludedUsers -is [array])
		{
			If($DGExcludedUsers.Count -gt 0)
			{
				$ScriptInformation += @{Data = "Excluded Users"; Value = $DGExcludedUsers[0]; }
				$cnt = -1
				ForEach($tmp in $DGExcludedUsers)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}
			}
		}
		
		If($Group.SessionSupport -eq "MultiSession")
		{
			$ScriptInformation += @{Data = 'Give access to unauthenticated (anonymous) users'; Value = $SFAnonymousUsers; }
		}

		If($xDeliveryType -ne "Applications")
		{
			$DesktopSettings = $Null
			$DesktopSettings = Get-BrokerEntitlementPolicyRule @CCParams2 -DesktopGroupUid $Group.Uid
			
			If($? -and $Null -ne $DesktopSettings)
			{
				ForEach($DesktopSetting in $DesktopSettings)
				{
					$DesktopSettingIncludedUsers = @()
					$DesktopSettingExcludedUsers = @()
					
					If([String]::IsNullOrEmpty($DesktopSetting.RestrictToTag))
					{
						$RestrictedToTag = "-"
					}
					Else
					{
						$RestrictedToTag = $DesktopSetting.RestrictToTag
					}
					
					If($DesktopSetting.IncludedUserFilterEnabled -eq $True)
					{
						ForEach($User in $DesktopSetting.IncludedUsers)
						{
							$DesktopSettingIncludedUsers += $User.Name
						}
						
						[array]$DesktopSettingIncludedUsers = $DesktopSettingIncludedUsers | Sort-Object -unique
					}
					ElseIf($DesktopSetting.IncludedUserFilterEnabled -eq $False)
					{
						$DesktopSettingIncludedUsers += "Allow everyone with access to this Delivery Group to use a desktop"
					}
					
					If($DesktopSetting.ExcludedUserFilterEnabled -eq $True)
					{
						ForEach($User in $DesktopSetting.ExcludedUsers)
						{
							$DesktopSettingExcludedUsers += $User.Name
						}

						[array]$DesktopSettingExcludedUsers = $DesktopSettingExcludedUsers | Sort-Object -unique
					}
					Switch ($DesktopSetting.SessionReconnection)
					{
						"Always" 			{$xSessionReconnection = "Always"; Break}
						"DisconnectedOnly"	{$xSessionReconnection = "Disconnected Only"; Break}
						"SameEndpointOnly"	{$xSessionReconnection = "Same Endpoint Only"; Break}
						Default {$xSessionReconnection = "Unable to determine Session Reconnection value: $($DesktopSetting.SessionReconnection)"; Break}
					}
					If($Null -ne $DesktopSetting.SecureIcaRequired)
					{
						$xSecureIcaRequired = $DesktopSetting.SecureIcaRequired.ToString()
					}
					$ScriptInformation += @{Data = "Desktop Entitlement"; Value = ""; }
					$ScriptInformation += @{Data = "     Display name"; Value = $DesktopSetting.PublishedName; }
					$ScriptInformation += @{Data = "     Description"; Value = $DesktopSetting.Description; }
					$ScriptInformation += @{Data = "     Restrict launches to machines with tag"; Value = $RestrictedToTag; }
					[string]$DSIU = $(If( $DesktopSettingIncludedUsers -is [array] -and $DesktopSettingIncludedUsers.Count ) { $DesktopSettingIncludedUsers[0] } Else { '-' } )
					$ScriptInformation += @{Data = "     Included Users"; Value = $DSIU; }
					$cnt = -1
					ForEach($tmp in $DesktopSettingIncludedUsers)
					{
						$cnt++
						If($cnt -gt 0)
						{
							$ScriptInformation += @{Data = ""; Value = $tmp; }
						}
					}
					
					If($DesktopSettingExcludedUsers.Count -gt 0)
					{
						$ScriptInformation += @{Data = "     Excluded Users"; Value = $DesktopSettingExcludedUsers[0]; }
						$cnt = -1
						ForEach($tmp in $DesktopSettingExcludedUsers)
						{
							$cnt++
							If($cnt -gt 0)
							{
								$ScriptInformation += @{Data = ""; Value = $tmp; }
							}
						}
					}
					$ScriptInformation += @{Data = "     Enable desktop"; Value = $DesktopSetting.Enabled; }
					$ScriptInformation += @{Data = "     Leasing behavior"; Value = $DesktopSetting.LeasingBehavior; }
					$ScriptInformation += @{Data = "     Maximum concurrent instances"; Value = $DesktopSetting.MaxPerEntitlementInstances; }
					$ScriptInformation += @{Data = "     SecureICA required"; Value = $DesktopSetting.SecureIcaRequired; }
					$ScriptInformation += @{Data = "     Session reconnection"; Value = $xSessionReconnection; }
				}
			}
		}
		
		[string]$DGS = $(If( $DGScopes -is [array] -and $DGScopes.Count ) { $DGScopes[0] } Else { '-' } )
		$ScriptInformation += @{Data = "Scopes"; Value = $DGS; }
		$cnt = -1
		ForEach($tmp in $DGScopes)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		
		[string]$DGSFS = $(If( $DGSFServers -is [array] -and $DGSFServers.Count ) { $DGSFServers[0] } Else { '-' } )
		$ScriptInformation += @{Data = "StoreFronts"; Value = $DGSFS; }
		$cnt = -1
		ForEach($tmp in $DGSFServers)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		
		If($Group.SessionSupport -eq "MultiSession" -and $xDeliveryType -like '*App*')
		{
			$ScriptInformation += @{Data = "Session prelaunch"; Value = $xSessionPrelaunch; }
			If($xSessionPrelaunch -ne "Off")
			{
				$ScriptInformation += @{Data = "Prelaunched session will end in"; Value = $xEndPrelaunchSession; }
				
				If($xSessionPrelaunchAvgLoad -gt 0)
				{
					$ScriptInformation += @{Data = "When avg load on all machines exceeds (%)"; Value = $xSessionPrelaunchAvgLoad; }
				}
				If($xSessionPrelaunchAnyLoad -gt 0)
				{
					$ScriptInformation += @{Data = "When load on any machines exceeds (%)"; Value = $xSessionPrelaunchAnyLoad; }
				}
			}
			$ScriptInformation += @{Data = "Session lingering"; Value = $xSessionLinger; }
			If($xSessionLinger -ne "Off")
			{
				$ScriptInformation += @{Data = "Keep sessions active until after"; Value = $xEndLinger; }
				
				If($xSessionLingerAvgLoad -gt 0)
				{
					$ScriptInformation += @{Data = "When avg load on all machines exceeds (%)"; Value = $xSessionPrelaunchAvgLoad; }
				}
				If($xSessionLingerAnyLoad -gt 0)
				{
					$ScriptInformation += @{Data = "When load on any machines exceeds (%)"; Value = $xSessionPrelaunchAnyLoad; }
				}
			}
		}

		$ScriptInformation += @{Data = "Launch in user's home zone"; Value = $xUsersHomeZone; }
		
		If($Group.SessionSupport -eq "MultiSession")
		{
			$RestartSchedules = Get-BrokerRebootScheduleV2 -EA 0 -DesktopGroupUid $Group.Uid
			
			If($? -and $Null -ne $RestartSchedules)
			{
				ForEach($RestartSchedule in $RestartSchedules)
				{
					$ScriptInformation += @{Data = "Restart Schedule"; Value = ""; }
					Switch($RestartSchedule.WarningRepeatInterval)
					{
						0	{$RestartScheduleWarningRepeatInterval = "Do not repeat"}
						5	{$RestartScheduleWarningRepeatInterval = "Every 5 minutes"} 
						Default {$RestartScheduleWarningRepeatInterval = "Notification frequency could not be determined: $($RestartSchedule.WarningRepeatInterval) "}
					}
				
					$ScriptInformation += @{Data = "     Restart machines automatically"; Value = "Yes"; }

					$ScriptInformation += @{Data = "     Restrict to tag"; Value = $RestartSchedule.RestrictToTag; }
					
					$tmp = ""
					If($RestartSchedule.Frequency -eq "Daily")
					{
						$tmp = "Daily"
					}
					ElseIf($RestartSchedule.Frequency -eq "Weekly")
					{
						$tmp = "Every $($RestartSchedule.Day)"
					}
					
					$ScriptInformation += @{Data = "     Restart frequency"; Value = $tmp; }
					$ScriptInformation += @{Data = "     Begin restart at"; Value = "$($RestartSchedule.StartTime.Hours.ToString("00")):$($RestartSchedule.StartTime.Minutes.ToString("00"))"; }
					
					$xTime = 0
					$tmp = ""
					If($RestartSchedule.RebootDuration -eq 0)
					{
						$tmp = "Restart all machines at once"
					}
					ElseIf($RestartSchedule.RebootDuration -eq 30)
					{
						$tmp = "30 minutes"
					}
					Else
					{
						$xTime = $RestartSchedule.RebootDuration / 60
						$tmp = "$($xTime) hours"
					}
					$ScriptInformation += @{Data = "     Restart duration"; Value = $tmp; }
					$xTime = $Null
					$tmp = $Null
					
					$tmp = ""
					If($RestartSchedule.WarningDuration -eq 0)
					{
						$tmp = "Do not send a notification"
						$ScriptInformation += @{Data = "     Send notification to users"; Value = $tmp; }
					}
					Else
					{
						$tmp = "$($RestartSchedule.WarningDuration) minutes before user is logged off"
						$ScriptInformation += @{Data = "     Send notification to users"; Value = $tmp; }
						$ScriptInformation += @{Data = "     Notification message"; Value = $RestartSchedule.WarningMessage; }
					}
					$ScriptInformation += @{Data = "     Notification frequency"; Value = $RestartScheduleWarningRepeatInterval; }
				}
			}
			Else
			{
				$ScriptInformation += @{Data = "Restart machines automatically"; Value = "No"; }
			}
		}

		$ScriptInformation += @{Data = "License model"; Value = $LicenseModel; }
		$ScriptInformation += @{Data = "Product code"; Value = $ProductCode; }
		$ScriptInformation += @{Data = "App Protection Key Logging Required"; Value = $Group.AppProtectionKeyLoggingRequired; }
		$ScriptInformation += @{Data = "App Protection Screen Capture Required"; Value = $Group.AppProtectionScreenCaptureRequired; }
		$ScriptInformation += @{ Data = "Off Peak Buffer Size Percent"; Value = $xOffPeakBufferSizePercent; }
		$ScriptInformation += @{ Data = "Off Peak Disconnect Timeout (Minutes)"; Value = $xOffPeakDisconnectTimeout; }
		$ScriptInformation += @{ Data = "Off Peak Extended Disconnect Timeout (Minutes)"; Value = $xOffPeakExtendedDisconnectTimeout; }
		$ScriptInformation += @{ Data = "Off Peak LogOff Timeout (Minutes)"; Value = $xOffPeakLogOffTimeout; }
		$ScriptInformation += @{ Data = "Peak Buffer Size Percent"; Value = $xPeakBufferSizePercent; }
		$ScriptInformation += @{ Data = "Peak Disconnect Timeout (Minutes)"; Value = $xPeakDisconnectTimeout; }
		$ScriptInformation += @{ Data = "Peak Extended Disconnect Timeout (Minutes)"; Value = $xPeakExtendedDisconnectTimeout; }
		$ScriptInformation += @{ Data = "Peak LogOff Timeout (Minutes)"; Value = $xPeakLogOffTimeout; }
		$ScriptInformation += @{ Data = "Settlement Period Before Autoshutdown (HH:MM:SS)"; Value = $xSettlementPeriodBeforeAutoShutdown; }
		$ScriptInformation += @{ Data = "Settlement Period Before Use (HH:MM:SS)"; Value = $xSettlementPeriodBeforeUse; }

		If($PwrMgmt1)
		{
			$ScriptInformation += @{Data = "During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins"; Value = $xPeakDisconnectAction; }
			$ScriptInformation += @{Data = "During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins"; Value = $xPeakExtendedDisconnectAction; }
			$ScriptInformation += @{Data = "During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins"; Value = $xOffPeakDisconnectAction; }
			$ScriptInformation += @{Data = "During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins"; Value = $xOffPeakExtendedDisconnectAction; }
		}
		If($PwrMgmt2)
		{
			#count if there are any items fist
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				$ScriptInformation += @{Data = "Weekday Peak hours"; Value = "None"; }
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekdays")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PeakHours[$i])
							{
								If($i -eq 0)
								{
									$ScriptInformation += @{Data = "Weekday Peak hours"; Value = "$($i.ToString("00")):00"; }
								}
								Else
								{
									$ScriptInformation += @{Data = ""; Value = "$($i.ToString("00")):00"; }
								}
								$val++
							}
						}
					}
				}
			}

			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				$ScriptInformation += @{Data = "Weekend Peak hours"; Value = "None"; }
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekend")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PeakHours[$i])
							{
								If($val -eq 0)
								{
									$ScriptInformation += @{Data = "Weekend Peak hours"; Value = "$($i.ToString("00")):00"; }
								}
								Else
								{
									$ScriptInformation += @{Data = ""; Value = "$($i.ToString("00")):00"; }
								}
								$val++
							}
						}
					}
				}
			}

			$ScriptInformation += @{Data = "During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins"; Value = $xPeakDisconnectAction; }
			$ScriptInformation += @{Data = "During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins"; Value = $xPeakExtendedDisconnectAction; }
			$ScriptInformation += @{Data = "During peak hours, when logged off $($Group.PeakLogOffTimeout) mins"; Value = $xPeakLogOffAction; }
			$ScriptInformation += @{Data = "During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins"; Value = $xOffPeakDisconnectAction; }
			$ScriptInformation += @{Data = "During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins"; Value = $xOffPeakExtendedDisconnectAction; }
			$ScriptInformation += @{Data = "During off-peak hours, when logged off $($Group.OffPeakLogOffTimeout) mins"; Value = $xOffPeakLogOffAction; }
		}
		If($PwrMgmt3)
		{
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PoolSize[$i] -gt 0)
						{
							$val++
						}
					}
				}
			}
			
			If($val -eq 0 )
			{
				$ScriptInformation += @{Data = "Weekday number machines powered on, and when"; Value = "None"; }
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekdays")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PoolSize[$i] -gt 0)
							{
								If($val -eq 0)
								{
									$ScriptInformation += @{Data = "Weekday number machines powered on, and when"; Value = "$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00"; }
								}
								Else
								{
									$ScriptInformation += @{Data = ""; Value = "$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00"; }
								}
								$val++
							}
						}
					}
				}
			}
			
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PoolSize[$i] -gt 0)
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				$ScriptInformation += @{Data = "Weekend number machines powered on, and when"; Value = "None"; }
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekend")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PoolSize[$i] -gt 0)
							{
								If($val -eq 0)
								{
									$ScriptInformation += @{Data = "Weekend number machines powered on, and when"; Value = "$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00"; }
								}
								Else
								{
									$ScriptInformation += @{Data = ""; Value = "$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00"; }
								}
								$val++
							}
						}
					}
				}
			}
			
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				$ScriptInformation += @{Data = "Weekday Peak hours"; Value = "None"; }
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekdays")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PeakHours[$i])
							{
								If($val -eq 0)
								{
									$ScriptInformation += @{Data = "Weekday Peak hours"; Value = "$($i.ToString("00")):00"; }
								}
								Else
								{
									$ScriptInformation += @{Data = ""; Value = "$($i.ToString("00")):00"; }
								}
								$val++
							}
						}
					}
				}
			}

			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				$ScriptInformation += @{Data = "Weekend Peak hours"; Value = "None"; }
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekend")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PeakHours[$i])
							{
								If($val -eq 0)
								{
									$ScriptInformation += @{Data = "Weekend Peak hours"; Value = "$($i.ToString("00")):00"; }
								}
								Else
								{
									$ScriptInformation += @{Data = ""; Value = "$($i.ToString("00")):00"; }
								}
								$val++
							}
						}
					}
				}
			}

			$ScriptInformation += @{Data = "During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins"; Value = $xPeakDisconnectAction; }
			$ScriptInformation += @{Data = "During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins"; Value = $xPeakExtendedDisconnectAction; }
			$ScriptInformation += @{Data = "During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins"; Value = $xOffPeakDisconnectAction; }
			$ScriptInformation += @{Data = "During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins"; Value = $xOffPeakExtendedDisconnectAction; }
		}

		$ScriptInformation += @{Data = "Automatic power on for assigned"; Value = $xAutoPowerOnForAssigned; }
		$ScriptInformation += @{Data = "Automatic power on for assigned during peak"; Value = $xAutoPowerOnForAssignedDuringPeak; }

		If($Group.ReuseMachinesWithoutShutdownInOutage -eq $Script:CCSite1.ReuseMachinesWithoutShutdownInOutageAllowed)
		{
			$ScriptInformation += @{Data = "Reuse Machines Without Shutdown in Outage"; Value = $Group.ReuseMachinesWithoutShutdownInOutage; }
		}
		Else
		{
			$ScriptInformation += @{Data = "Reuse Machines Without Shutdown in Outage"; Value = "$($Group.ReuseMachinesWithoutShutdownInOutage) (Doesn't match Site setting)"; }
		}
		
		$ScriptInformation += @{Data = "All connections not through NetScaler Gateway"; Value = $xAllConnections; }
		$ScriptInformation += @{Data = "Connections through NetScaler Gateway"; Value = $xNSConnection; }
		[string]$AGF = $(If( $xAGFilters -is [array] -and $xAGFilters.Count ) { $xAGFilters[0] } Else { '-' } )
		$ScriptInformation += @{Data = "Connections meeting any of the following filters"; Value = $AGF; }
		$cnt = -1
		ForEach($tmp in $xAGFilters)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "Details: " $Group.Name
		Line 1 "Description`t`t`t`t`t`t: " $Group.Description
		If(![String]::IsNullOrEmpty($Group.PublishedName))
		{
			Line 1 "Display Name`t`t`t`t`t`t: " $Group.PublishedName
		}
		Line 1 "Type`t`t`t`t`t`t`t: " $xDGType
		Line 1 "Set to VDA version`t`t`t`t`t: " $xVDAVersion
		If($Group.SessionSupport -eq "SingleSession" -and ($xDGType -eq "Static Desktops" -or $xDGType -like "*Random*"))
		{
			Line 1 "Desktops per user`t`t`t`t`t: " $xMaxDesktops
		}
		Line 1 "Time zone`t`t`t`t`t`t: " $xTimeZone
		Line 1 "Enable Delivery Group`t`t`t`t`t: " $xEnabled
		Line 1 "Enable Secure ICA`t`t`t`t`t: " $xSecureICA
		Line 1 "Color Depth`t`t`t`t`t`t: " $xColorDepth
		Line 1 "Shutdown Desktops After Use`t`t`t`t: " $xShutdownDesktopsAfterUse
		Line 1 "Turn On Added Machine`t`t`t`t`t: " $xTurnOnAddedMachine
		[string]$DGIU = $(If( $DGIncludedUsers -is [array] -and $DGIncludedUsers.Count ) { $DGIncludedUsers[0] } Else { '-' } )
		Line 1 "Included Users`t`t`t`t`t`t: " $DGIU
		$cnt = -1
		ForEach($tmp in $DGIncludedUsers)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 8 "  " $tmp
			}
		}
		
		If($DGExcludedUsers -is [array])
		{
			If($DGExcludedUsers.Count -gt 0)
			{
				Line 1 "Excluded Users`t`t`t`t`t`t: " $DGExcludedUsers[0]
				$cnt = -1
				ForEach($tmp in $DGExcludedUsers)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 8 "  " $tmp
					}
				}
			}
		}

		If($Group.SessionSupport -eq "MultiSession")
		{
			Line 1 "Give access to unauthenticated (anonymous) users`t: " $SFAnonymousUsers
		}

		If($xDeliveryType -ne "Applications")
		{
			$DesktopSettings = $Null
			$DesktopSettings = Get-BrokerEntitlementPolicyRule @CCParams2 -DesktopGroupUid $Group.Uid
			
			If($? -and $Null -ne $DesktopSettings)
			{
				ForEach($DesktopSetting in $DesktopSettings)
				{
					$DesktopSettingIncludedUsers = @()
					$DesktopSettingExcludedUsers = @()
					
					If([String]::IsNullOrEmpty($DesktopSetting.RestrictToTag))
					{
						$RestrictedToTag = "-"
					}
					Else
					{
						$RestrictedToTag = $DesktopSetting.RestrictToTag
					}
					
					If($DesktopSetting.IncludedUserFilterEnabled -eq $True)
					{
						ForEach($User in $DesktopSetting.IncludedUsers)
						{
							$DesktopSettingIncludedUsers += $User.Name
						}
						
						[array]$DesktopSettingIncludedUsers = $DesktopSettingIncludedUsers | Sort-Object -unique
					}
					ElseIf($DesktopSetting.IncludedUserFilterEnabled -eq $False)
					{
						$DesktopSettingIncludedUsers += "Allow everyone with access to this Delivery Group to use a desktop"
					}
					
					If($DesktopSetting.ExcludedUserFilterEnabled -eq $True)
					{
						ForEach($User in $DesktopSetting.ExcludedUsers)
						{
							$DesktopSettingExcludedUsers += $User.Name
						}

						[array]$DesktopSettingExcludedUsers = $DesktopSettingExcludedUsers | Sort-Object -unique
					}
					Switch ($DesktopSetting.SessionReconnection)
					{
						"Always" 			{$xSessionReconnection = "Always"; Break}
						"DisconnectedOnly"	{$xSessionReconnection = "Disconnected Only"; Break}
						"SameEndpointOnly"	{$xSessionReconnection = "Same Endpoint Only"; Break}
						Default {$xSessionReconnection = "Unable to determine Session Reconnection value: $($DesktopSetting.SessionReconnection)"; Break}
					}
					If($Null -ne $DesktopSetting.SecureIcaRequired)
					{
						$xSecureIcaRequired = $DesktopSetting.SecureIcaRequired.ToString()
					}
					Line 1 "Desktop Entitlement" ""
					Line 2 "Display name`t`t`t`t`t: " $DesktopSetting.PublishedName
					Line 2 "Description`t`t`t`t`t: " $DesktopSetting.Description
					Line 2 "Restrict launches to machines with tag`t`t: " $RestrictedToTag
					[string]$DSIU = $(If( $DesktopSettingIncludedUsers -is [array] -and $DesktopSettingIncludedUsers.Count ) { $DesktopSettingIncludedUsers[0] } Else { '-' } )
					Line 2 "Included Users`t`t`t`t`t: " $DSIU
					$cnt = -1
					ForEach($tmp in $DesktopSettingIncludedUsers)
					{
						$cnt++
						If($cnt -gt 0)
						{
							Line 8 "  " $tmp
						}
					}
					
					If($DesktopSettingExcludedUsers.Count -gt 0)
					{
						Line 2 "Excluded Users`t`t`t`t`t: " $DesktopSettingExcludedUsers[0]
						$cnt = -1
						ForEach($tmp in $DesktopSettingExcludedUsers)
						{
							$cnt++
							If($cnt -gt 0)
							{
								Line 9 "  " $tmp
							}
						}
					}
					Line 2 "Enable desktop`t`t`t`t`t: " $DesktopSetting.Enabled
					Line 2 "Leasing behavior`t`t`t`t: " $DesktopSetting.LeasingBehavior
					Line 2 "Maximum concurrent instances`t`t`t: " $DesktopSetting.MaxPerEntitlementInstances
					Line 2 "SecureICA required`t`t`t`t: " $DesktopSetting.SecureIcaRequired
					Line 2 "Session reconnection`t`t`t`t: " $xSessionReconnection
				}
			}
		}
		
		[string]$DGS = $(If( $DGScopes -is [array] -and $DGScopes.Count ) { $DGScopes[0] } Else { '-' } )
		Line 1 "Scopes`t`t`t`t`t`t`t: " $DGS
		$cnt = -1
		ForEach($tmp in $DGScopes)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 8 "  " $tmp
			}
		}
		
		[string]$DGSFS = $(If( $DGSFServers -is [array] -and $DGSFServers.Count ) { $DGSFServers[0] } Else { '-' } )
		Line 1 "StoreFronts`t`t`t`t`t`t: " $DGSFS
		$cnt = -1
		ForEach($tmp in $DGSFServers)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 8 "  " $tmp
			}
		}
		
		If($Group.SessionSupport -eq "MultiSession" -and $xDeliveryType -like '*App*')
		{
			Line 1 "Session prelaunch`t`t`t`t`t: " $xSessionPrelaunch
			If($xSessionPrelaunch -ne "Off")
			{
				Line 1 "Prelaunched session will end in`t`t`t`t`t: " $xEndPrelaunchSession
				
				If($xSessionPrelaunchAvgLoad -gt 0)
				{
					Line 1 "When avg load on all machines exceeds (%)`t`t`t: " $xSessionPrelaunchAvgLoad
				}
				If($xSessionPrelaunchAnyLoad -gt 0)
				{
					Line 1 "When load on any machines exceeds (%)`t`t`t`t: " $xSessionPrelaunchAnyLoad
				}
			}
			Line 1 "Session lingering`t`t`t`t`t: " $xSessionLinger
			If($xSessionLinger -ne "Off")
			{
				Line 1 "Keep sessions active until after`t`t`t`t: " $xEndLinger
				
				If($xSessionLingerAvgLoad -gt 0)
				{
					Line 1 "When avg load on all machines exceeds (%)`t`t`t: " $xSessionPrelaunchAvgLoad
				}
				If($xSessionLingerAnyLoad -gt 0)
				{
					Line 1 "When load on any machines exceeds (%)`t`t`t`t: " $xSessionPrelaunchAnyLoad
				}
			}
		}

		Line 1 "Launch in user's home zone`t`t`t`t: " $xUsersHomeZone
		
		If($Group.SessionSupport -eq "MultiSession")
		{
			$RestartSchedules = Get-BrokerRebootScheduleV2 -EA 0 -DesktopGroupUid $Group.Uid
			
			If($? -and $Null -ne $RestartSchedules)
			{
				ForEach($RestartSchedule in $RestartSchedules)
				{
					Line 1 "Restart Schedule"
					Switch($RestartSchedule.WarningRepeatInterval)
					{
						0	{$RestartScheduleWarningRepeatInterval = "Do not repeat"}
						5	{$RestartScheduleWarningRepeatInterval = "Every 5 minutes"} 
						Default {$RestartScheduleWarningRepeatInterval = "Notification frequency could not be determined: $($RestartSchedule.WarningRepeatInterval) "}
					}

					Line 2 "Restart machines automatically`t`t`t: " "Yes"

					Line 2 "Restrict to tag`t`t`t`t`t: " $RestartSchedule.RestrictToTag
					
					$tmp = ""
					If($RestartSchedule.Frequency -eq "Daily")
					{
						$tmp = "Daily"
					}
					ElseIf($RestartSchedule.Frequency -eq "Weekly")
					{
						$tmp = "Every $($RestartSchedule.Day)"
					}
					
					Line 2 "Restart frequency`t`t`t`t: " $tmp
					Line 2 "Begin restart at`t`t`t`t: " "$($RestartSchedule.StartTime.Hours.ToString("00")):$($RestartSchedule.StartTime.Minutes.ToString("00"))"
					
					$xTime = 0
					$tmp = ""
					If($RestartSchedule.RebootDuration -eq 0)
					{
						$tmp = "Restart all machines at once"
					}
					ElseIf($RestartSchedule.RebootDuration -eq 30)
					{
						$tmp = "30 minutes"
					}
					Else
					{
						$xTime = $RestartSchedule.RebootDuration / 60
						$tmp = "$($xTime) hours"
					}
					Line 2 "Restart duration`t`t`t`t: " $tmp
					$xTime = $Null
					$tmp = $Null
					
					$tmp = ""
					If($RestartSchedule.WarningDuration -eq 0)
					{
						$tmp = "Do not send a notification"
						Line 2 "Send notification to users`t`t`t: " $tmp
					}
					Else
					{
						$tmp = "$($RestartSchedule.WarningDuration) minutes before user is logged off"
						Line 2 "Send notification to users`t`t`t: " $tmp
						Line 2 "Notification message`t`t`t`t: " $RestartSchedule.WarningMessage
					}
					Line 2 "Notification frequency`t`t`t`t: " $RestartScheduleWarningRepeatInterval
				}
			}
			Else
			{
				Line 1 "Restart machines automatically`t`t`t: " "No"
			}
		}
		
		Line 1 "License model`t`t`t`t`t`t: " $LicenseModel
		Line 1 "Product code`t`t`t`t`t`t: " $ProductCode
		Line 1 "App Protection Key Logging Required`t`t`t: " $Group.AppProtectionKeyLoggingRequired
		Line 1 "App Protection Screen Capture Required`t`t`t: " $Group.AppProtectionScreenCaptureRequired
		Line 1 "Off Peak Buffer Size Percent`t`t`t`t: " $xOffPeakBufferSizePercent
		Line 1 "Off Peak Disconnect Timeout (Minutes)`t`t`t: " $xOffPeakDisconnectTimeout
		Line 1 "Off Peak Extended Disconnect Timeout (Minutes)`t`t: " $xOffPeakExtendedDisconnectTimeout
		Line 1 "Off Peak LogOff Timeout (Minutes)`t`t`t: " $xOffPeakLogOffTimeout
		Line 1 "Peak Buffer Size Percent`t`t`t`t: " $xPeakBufferSizePercent
		Line 1 "Peak Disconnect Timeout (Minutes)`t`t`t: " $xPeakDisconnectTimeout
		Line 1 "Peak Extended Disconnect Timeout (Minutes)`t`t: " $xPeakExtendedDisconnectTimeout
		Line 1 "Peak LogOff Timeout (Minutes)`t`t`t`t: " $xPeakLogOffTimeout
		Line 1 "Settlement Period Before Autoshutdown (HH:MM:SS)`t: " $xSettlementPeriodBeforeAutoShutdown
		Line 1 "Settlement Period Before Use (HH:MM:SS)`t`t`t: " $xSettlementPeriodBeforeUse
		
		If($PwrMgmt1)
		{
			Line 1 "During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins`t`t: " $xPeakDisconnectAction
			Line 1 "During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins`t: " $xPeakExtendedDisconnectAction
			Line 1 "During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins`t`t: " $xOffPeakDisconnectAction
			Line 1 "During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins: " $xOffPeakExtendedDisconnectAction
		}
		If($PwrMgmt2)
		{
			#count if there are any items fist
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				Line 1 "Weekday Peak hours`t`t`t`t`t: " "None"
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekdays")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PeakHours[$i])
							{
								If($i -eq 0)
								{
									Line 1 "Weekday Peak hours`t`t`t`t`t: " "$($i.ToString("00")):00"
								}
								Else
								{
									Line 8 "  " "$($i.ToString("00")):00"
								}
								$val++
							}
						}
					}
				}
			}

			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				Line 1 "Weekend Peak hours`t`t`t`t`t: " "None"
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekend")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PeakHours[$i])
							{
								If($val -eq 0)
								{
									Line 1 "Weekend Peak hours`t`t`t`t`t: " "$($i.ToString("00")):00"
								}
								Else
								{
									Line 8 "  " "$($i.ToString("00")):00"
								}
								$val++
							}
						}
					}
				}
			}

			Line 1 "During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins`t`t: " $xPeakDisconnectAction
			Line 1 "During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins`t: " $xPeakExtendedDisconnectAction
			Line 1 "During peak hours, when logged off $($Group.PeakLogOffTimeout) mins`t`t: " $xPeakLogOffAction
			Line 1 "During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins`t`t: " $xOffPeakDisconnectAction
			Line 1 "During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins: " $xOffPeakExtendedDisconnectAction
			Line 1 "During off-peak hours, when logged off $($Group.OffPeakLogOffTimeout) mins`t`t: " $xOffPeakLogOffAction
		}
		If($PwrMgmt3)
		{
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PoolSize[$i] -gt 0)
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				Line 1 "Weekday number machines powered on, and when`t`t: " "None"
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekdays")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PoolSize[$i] -gt 0)
							{
								If($val -eq 0)
								{
									Line 1 "Weekday number machines powered on, and when`t`t: " "$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00"
								}
								Else
								{
									Line 8 "  " "$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00"
								}
								$val++
							}
						}
					}
				}
			}

			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PoolSize[$i] -gt 0)
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				Line 1 "Weekend number machines powered on, and when`t`t: " "None"
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekend")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PoolSize[$i] -gt 0)
							{
								If($val -eq 0)
								{
									Line 1 "Weekend number machines powered on, and when`t`t: " "$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00"
								}
								Else
								{
									Line 8 "  " "$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00"
								}
								$val++
							}
						}
					}
				}
			}

			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				Line 1 "Weekday Peak hours`t`t`t`t`t: " "None"
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekdays")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PeakHours[$i])
							{
								If($val -eq 0)
								{
									Line 1 "Weekday Peak hours`t`t`t`t`t: " "$($i.ToString("00")):00"
								}
								Else
								{
									Line 8 "  " "$($i.ToString("00")):00"
								}
								$val++
							}
						}
					}
				}
			}

			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				Line 1 "Weekend Peak hours`t`t`t`t`t: " "None"
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekend")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PeakHours[$i])
							{
								If($val -eq 0)
								{
									Line 1 "Weekend Peak hours`t`t`t`t`t: " "$($i.ToString("00")):00"
								}
								Else
								{
									Line 8 "  " "$($i.ToString("00")):00"
								}
								$val++
							}
						}
					}
				}
			}

			Line 1 "During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins`t`t: " $xPeakDisconnectAction
			Line 1 "During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins`t: " $xPeakExtendedDisconnectAction
			Line 1 "During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins`t`t: " $xOffPeakDisconnectAction
			Line 1 "During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins: " $xOffPeakExtendedDisconnectAction
		}

		Line 1 "Automatic power on for assigned`t`t`t`t: " $xAutoPowerOnForAssigned
		Line 1 "Automatic power on for assigned during peak`t`t: " $xAutoPowerOnForAssignedDuringPeak

		If($Group.ReuseMachinesWithoutShutdownInOutage -eq $Script:CCSite1.ReuseMachinesWithoutShutdownInOutageAllowed)
		{
			Line 1 "Reuse Machines Without Shutdown in Outage`t`t: " $Group.ReuseMachinesWithoutShutdownInOutage
		}
		Else
		{
			Line 1 "Reuse Machines Without Shutdown in Outage`t`t: $($Group.ReuseMachinesWithoutShutdownInOutage) (Doesn't match Site setting)"
		}
		
		Line 1 "All connections not through NetScaler Gateway`t`t: " $xAllConnections
		Line 1 "Connections through NetScaler Gateway`t`t`t: " $xNSConnection
		[string]$AGF = $(If( $xAGFilters -is [array] -and $xAGFilters.Count ) { $xAGFilters[0] } Else { '-' } )
		Line 1 "Connections meeting any of the following filters`t: " $AGF
		$cnt = -1
		ForEach($tmp in $xAGFilters)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 8 "  " $tmp
			}
		}
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Details: " $Group.Name
		$rowdata = @()
		$columnHeaders = @("Description",($global:htmlsb),$Group.Description,$htmlwhite)
		If(![String]::IsNullOrEmpty($Group.PublishedName))
		{
			$rowdata += @(,('Display Name',($global:htmlsb),$Group.PublishedName,$htmlwhite))
		}
		$rowdata += @(,('Type',($global:htmlsb),$xDGType,$htmlwhite))
		$rowdata += @(,('Set to VDA version',($global:htmlsb),$xVDAVersion,$htmlwhite))
		If($Group.SessionSupport -eq "SingleSession" -and ($xDGType -eq "Static Desktops" -or $xDGType -like "*Random*"))
		{
			$rowdata += @(,('Desktops per user',($global:htmlsb),$xMaxDesktops.ToString(),$htmlwhite))
		}
		$rowdata += @(,('Time zone',($global:htmlsb),$xTimeZone,$htmlwhite))
		$rowdata += @(,('Enable Delivery Group',($global:htmlsb),$xEnabled,$htmlwhite))
		$rowdata += @(,('Enable Secure ICA',($global:htmlsb),$xSecureICA,$htmlwhite))
		$rowdata += @(,('Color Depth',($global:htmlsb),$xColorDepth,$htmlwhite))
		$rowdata += @(,("Shutdown Desktops After Use",($global:htmlsb),$xShutdownDesktopsAfterUse,$htmlwhite))
		$rowdata += @(,("Turn On Added Machine",($global:htmlsb),$xTurnOnAddedMachine,$htmlwhite))
		[string]$DGIU = $(If( $DGIncludedUsers -is [array] -and $DGIncludedUsers.Count ) { $DGIncludedUsers[0] } Else { '-' } )
		$rowdata += @(,('Included Users',($global:htmlsb),$DGIU,$htmlwhite))
		$cnt = -1
		ForEach($tmp in $DGIncludedUsers)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
			}
		}
		
		If($DGExcludedUsers -is [array])
		{
			If($DGExcludedUsers.Count -gt 0)
			{
				$rowdata += @(,('Excluded Users',($global:htmlsb), $DGExcludedUsers[0],$htmlwhite))
				$cnt = -1
				ForEach($tmp in $DGExcludedUsers)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
					}
				}
			}
		}

		If($Group.SessionSupport -eq "MultiSession")
		{
			$rowdata += @(,('Give access to unauthenticated (anonymous) users',($global:htmlsb),$SFAnonymousUsers.ToString(),$htmlwhite))
		}

		If($xDeliveryType -ne "Applications")
		{
			$DesktopSettings = $Null
			$DesktopSettings = Get-BrokerEntitlementPolicyRule @CCParams2 -DesktopGroupUid $Group.Uid
			
			If($? -and $Null -ne $DesktopSettings)
			{
				ForEach($DesktopSetting in $DesktopSettings)
				{
					$DesktopSettingIncludedUsers = @()
					$DesktopSettingExcludedUsers = @()
					
					If([String]::IsNullOrEmpty($DesktopSetting.RestrictToTag))
					{
						$RestrictedToTag = "-"
					}
					Else
					{
						$RestrictedToTag = $DesktopSetting.RestrictToTag
					}
					
					If($DesktopSetting.IncludedUserFilterEnabled -eq $True)
					{
						ForEach($User in $DesktopSetting.IncludedUsers)
						{
							$DesktopSettingIncludedUsers += $User.Name
						}
						
						[array]$DesktopSettingIncludedUsers = $DesktopSettingIncludedUsers | Sort-Object -unique
					}
					ElseIf($DesktopSetting.IncludedUserFilterEnabled -eq $False)
					{
						$DesktopSettingIncludedUsers += "Allow everyone with access to this Delivery Group to use a desktop"
					}
					
					If($DesktopSetting.ExcludedUserFilterEnabled -eq $True)
					{
						ForEach($User in $DesktopSetting.ExcludedUsers)
						{
							$DesktopSettingExcludedUsers += $User.Name
						}

						[array]$DesktopSettingExcludedUsers = $DesktopSettingExcludedUsers | Sort-Object -unique
					}
					Switch ($DesktopSetting.SessionReconnection)
					{
						"Always" 			{$xSessionReconnection = "Always"; Break}
						"DisconnectedOnly"	{$xSessionReconnection = "Disconnected Only"; Break}
						"SameEndpointOnly"	{$xSessionReconnection = "Same Endpoint Only"; Break}
						Default {$xSessionReconnection = "Unable to determine Session Reconnection value: $($DesktopSetting.SessionReconnection)"; Break}
					}
					If($Null -ne $DesktopSetting.SecureIcaRequired)
					{
						$xSecureIcaRequired = $DesktopSetting.SecureIcaRequired.ToString()
					}
					$rowdata += @(,("Desktop Entitlement",($global:htmlsb),"",$htmlwhite))
					$rowdata += @(,("     Display name",($global:htmlsb),$DesktopSetting.PublishedName,$htmlwhite))
					$rowdata += @(,("     Description",($global:htmlsb),$DesktopSetting.Description,$htmlwhite))
					$rowdata += @(,("     Restrict launches to machines with tag",($global:htmlsb),$RestrictedToTag,$htmlwhite))
					[string]$DSIU = $(If( $DesktopSettingIncludedUsers -is [array] -and $DesktopSettingIncludedUsers.Count ) { $DesktopSettingIncludedUsers[0] } Else { '-' } )
					$rowdata += @(,("     Included Users",($global:htmlsb),$DSIU,$htmlwhite))
					$cnt = -1
					ForEach($tmp in $DesktopSettingIncludedUsers)
					{
						$cnt++
						If($cnt -gt 0)
						{
							$rowdata += @(,("",($global:htmlsb),$tmp,$htmlwhite))
						}
					}
					
					If($DesktopSettingExcludedUsers.Count -gt 0)
					{
						$rowdata += @(,("     Excluded Users",($global:htmlsb),$DesktopSettingExcludedUsers[0],$htmlwhite))
						$cnt = -1
						ForEach($tmp in $DesktopSettingExcludedUsers)
						{
							$cnt++
							If($cnt -gt 0)
							{
								$rowdata += @(,("",($global:htmlsb),$tmp,$htmlwhite))
							}
						}
					}
					$rowdata += @(,("     Enable desktop",($global:htmlsb),$DesktopSetting.Enabled.ToString(),$htmlwhite))
					$rowdata += @(,("     Leasing behavior",($global:htmlsb),$DesktopSetting.LeasingBehavior.ToString(),$htmlwhite))
					$rowdata += @(,("     Maximum concurrent instances",($global:htmlsb),$DesktopSetting.MaxPerEntitlementInstances.ToString(),$htmlwhite))
					$rowdata += @(,("     SecureICA required",($global:htmlsb),$xSecureIcaRequired,$htmlwhite))
					$rowdata += @(,("     Session reconnection",($global:htmlsb),$xSessionReconnection,$htmlwhite))
				}
			}
		}
		
		[string]$DGS = $(If( $DGScopes -is [array] -and $DGScopes.Count ) { $DGScopes[0] } Else { '-' } )
		$rowdata += @(,('Scopes',($global:htmlsb),$DGS,$htmlwhite))
		$cnt = -1
		ForEach($tmp in $DGScopes)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
			}
		}
		
		[string]$DGSFS = $(If( $DGSFServers -is [array] -and $DGSFServers.Count ) { $DGSFServers[0] } Else { '-' } )
		$rowdata += @(,('StoreFronts',($global:htmlsb),$DGSFS,$htmlwhite))
		$cnt = -1
		ForEach($tmp in $DGSFServers)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
			}
		}
		
		If($Group.SessionSupport -eq "MultiSession" -and $xDeliveryType -like '*App*')
		{
			$rowdata += @(,('Session prelaunch',($global:htmlsb),$xSessionPrelaunch,$htmlwhite))
			If($xSessionPrelaunch -ne "Off")
			{
				$rowdata += @(,('Prelaunched session will end in',($global:htmlsb),$xEndPrelaunchSession,$htmlwhite))
				
				If($xSessionPrelaunchAvgLoad -gt 0)
				{
					$rowdata += @(,('When avg load on all machines exceeds (%)',($global:htmlsb),$xSessionPrelaunchAvgLoad,$htmlwhite))
				}
				If($xSessionPrelaunchAnyLoad -gt 0)
				{
					$rowdata += @(,('When load on any machines exceeds (%)',($global:htmlsb),$xSessionPrelaunchAnyLoad,$htmlwhite))
				}
			}
			$rowdata += @(,('Session lingering',($global:htmlsb),$xSessionLinger,$htmlwhite))
			If($xSessionLinger -ne "Off")
			{
				$rowdata += @(,('Keep sessions active until after',($global:htmlsb),$xEndLinger,$htmlwhite))
				
				If($xSessionLingerAvgLoad -gt 0)
				{
					$rowdata += @(,('When avg load on all machines exceeds (%)',($global:htmlsb),$xSessionPrelaunchAvgLoad,$htmlwhite))
				}
				If($xSessionLingerAnyLoad -gt 0)
				{
					$rowdata += @(,('When load on any machines exceeds (%)',($global:htmlsb),$xSessionPrelaunchAnyLoad,$htmlwhite))
				}
			}
		}
		
		$rowdata += @(,("Launch in user's home zone",($global:htmlsb),$xUsersHomeZone,$htmlwhite)) 
			
		If($Group.SessionSupport -eq "MultiSession")
		{
			$RestartSchedules = Get-BrokerRebootScheduleV2 -EA 0 -DesktopGroupUid $Group.Uid
			
			If($? -and $Null -ne $RestartSchedules)
			{
				ForEach($RestartSchedule in $RestartSchedules)
				{
					$rowdata += @(,('Restart Schedule',($global:htmlsb),"",$htmlwhite))
					Switch($RestartSchedule.WarningRepeatInterval)
					{
						0	{$RestartScheduleWarningRepeatInterval = "Do not repeat"}
						5	{$RestartScheduleWarningRepeatInterval = "Every 5 minutes"} 
						Default {$RestartScheduleWarningRepeatInterval = "Notification frequency could not be determined: $($RestartSchedule.WarningRepeatInterval) "}
					}

					$rowdata += @(,('     Restart machines automatically',($global:htmlsb),"Yes",$htmlwhite))

					$rowdata += @(,('     Restrict to tag',($global:htmlsb),$RestartSchedule.RestrictToTag,$htmlwhite))
					
					$tmp = ""
					If($RestartSchedule.Frequency -eq "Daily")
					{
						$tmp = "Daily"
					}
					ElseIf($RestartSchedule.Frequency -eq "Weekly")
					{
						$tmp = "Every $($RestartSchedule.Day)"
					}
					
					$rowdata += @(,('     Restart frequency',($global:htmlsb),$tmp,$htmlwhite))
					$rowdata += @(,('     Begin restart at',($global:htmlsb),"$($RestartSchedule.StartTime.Hours.ToString("00")):$($RestartSchedule.StartTime.Minutes.ToString("00"))",$htmlwhite))
					
					$xTime = 0
					$tmp = ""
					If($RestartSchedule.RebootDuration -eq 0)
					{
						$tmp = "Restart all machines at once"
					}
					ElseIf($RestartSchedule.RebootDuration -eq 30)
					{
						$tmp = "30 minutes"
					}
					Else
					{
						$xTime = $RestartSchedule.RebootDuration / 60
						$tmp = "$($xTime) hours"
					}
					$rowdata += @(,('     Restart duration',($global:htmlsb),$tmp,$htmlwhite))
					$xTime = $Null
					$tmp = $Null
					
					$tmp = ""
					If($RestartSchedule.WarningDuration -eq 0)
					{
						$tmp = "Do not send a notification"
						$rowdata += @(,('     Send notification to users',($global:htmlsb),$tmp,$htmlwhite))
					}
					Else
					{
						$tmp = "$($RestartSchedule.WarningDuration) minutes before user is logged off"
						$rowdata += @(,('     Send notification to users',($global:htmlsb),$tmp,$htmlwhite))
						$rowdata += @(,('     Notification message',($global:htmlsb),$RestartSchedule.WarningMessage,$htmlwhite))
					}
					$rowdata += @(,('     Notification frequency',($global:htmlsb),$RestartScheduleWarningRepeatInterval,$htmlwhite))
				}
			}
			Else
			{
				$rowdata += @(,('Restart machines automatically',($global:htmlsb),"No",$htmlwhite))
			}
		}
		
		$rowdata += @(,('License model',($global:htmlsb),$LicenseModel,$htmlwhite))
		$rowdata += @(,('Product code',($global:htmlsb),$ProductCode,$htmlwhite))
		$rowdata += @(,("App Protection Key Logging Required",($global:htmlsb),$Group.AppProtectionKeyLoggingRequired.ToString(),$htmlwhite))
		$rowdata += @(,("App Protection Screen Capture Required",($global:htmlsb),$Group.AppProtectionScreenCaptureRequired.ToString(),$htmlwhite))
		$rowdata += @(,( "Off Peak Buffer Size Percent",($global:htmlsb),$xOffPeakBufferSizePercent.ToString(),$htmlwhite))
		$rowdata += @(,( "Off Peak Disconnect Timeout (Minutes)",($global:htmlsb),$xOffPeakDisconnectTimeout.ToString(),$htmlwhite))
		$rowdata += @(,( "Off Peak Extended Disconnect Timeout (Minutes)",($global:htmlsb),$xOffPeakExtendedDisconnectTimeout.ToString(),$htmlwhite))
		$rowdata += @(,( "Off Peak LogOff Timeout (Minutes)",($global:htmlsb),$xOffPeakLogOffTimeout.ToString(),$htmlwhite))
		$rowdata += @(,( "Peak Buffer Size Percent",($global:htmlsb),$xPeakBufferSizePercent.ToString(),$htmlwhite))
		$rowdata += @(,( "Peak Disconnect Timeout (Minutes)",($global:htmlsb),$xPeakDisconnectTimeout.ToString(),$htmlwhite))
		$rowdata += @(,( "Peak Extended Disconnect Timeout (Minutes)",($global:htmlsb),$xPeakExtendedDisconnectTimeout.ToString(),$htmlwhite))
		$rowdata += @(,( "Peak LogOff Timeout (Minutes)",($global:htmlsb),$xPeakLogOffTimeout.ToString(),$htmlwhite))
		$rowdata += @(,( "Settlement Period Before Autoshutdown (HH:MM:SS)",($global:htmlsb),$xSettlementPeriodBeforeAutoShutdown.ToString(),$htmlwhite))
		$rowdata += @(,( "Settlement Period Before Use (HH:MM:SS)",($global:htmlsb),$xSettlementPeriodBeforeUse.ToString(),$htmlwhite))

		If($PwrMgmt1)
		{
			$rowdata += @(,("During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins",($global:htmlsb),$xPeakDisconnectAction,$htmlwhite))
			$rowdata += @(,("During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins",($global:htmlsb),$xPeakExtendedDisconnectAction,$htmlwhite))
			$rowdata += @(,("During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins",($global:htmlsb),$xOffPeakDisconnectAction,$htmlwhite))
			$rowdata += @(,("During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins",($global:htmlsb),$xOffPeakExtendedDisconnectAction,$htmlwhite))
		}
		If($PwrMgmt2)
		{
			#count if there are any items fist
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				$rowdata += @(,('Weekday Peak hours',($global:htmlsb),"None",$htmlwhite))
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekdays")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PeakHours[$i])
							{
								If($i -eq 0)
								{
									$rowdata += @(,('Weekday Peak hours',($global:htmlsb),"$($i.ToString("00")):00",$htmlwhite))
								}
								Else
								{
									$rowdata += @(,('',($global:htmlsb),"$($i.ToString("00")):00",$htmlwhite))
								}
								$val++
							}
						}
					}
				}
			}

			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				$rowdata += @(,('Weekend Peak hours',($global:htmlsb),"None",$htmlwhite))
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekend")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PeakHours[$i])
							{
								If($val -eq 0)
								{
									$rowdata += @(,('Weekend Peak hours',($global:htmlsb),"$($i.ToString("00")):00",$htmlwhite))
								}
								Else
								{
									$rowdata += @(,('',($global:htmlsb),"$($i.ToString("00")):00",$htmlwhite))
								}
								$val++
							}
						}
					}
				}
			}

			$rowdata += @(,("During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins",($global:htmlsb),$xPeakDisconnectAction,$htmlwhite))
			$rowdata += @(,("During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins",($global:htmlsb),$xPeakExtendedDisconnectAction,$htmlwhite))
			$rowdata += @(,("During peak hours, when logged off $($Group.PeakLogOffTimeout) mins",($global:htmlsb),$xPeakLogOffAction,$htmlwhite))
			$rowdata += @(,("During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins",($global:htmlsb),$xOffPeakDisconnectAction,$htmlwhite))
			$rowdata += @(,("During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins",($global:htmlsb),$xOffPeakExtendedDisconnectAction,$htmlwhite))
			$rowdata += @(,("During off-peak hours, when logged off $($Group.OffPeakLogOffTimeout) mins",($global:htmlsb),$xOffPeakLogOffAction,$htmlwhite))
		}
		If($PwrMgmt3)
		{
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PoolSize[$i] -gt 0)
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				$rowdata += @(,('Weekday number machines powered on, and when',($global:htmlsb),"None",$htmlwhite))
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekdays")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PoolSize[$i] -gt 0)
							{
								If($val -eq 0)
								{
									$rowdata += @(,('Weekday number machines powered on, and when',($global:htmlsb),"$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00",$htmlwhite))
								}
								Else
								{
									$rowdata += @(,('',($global:htmlsb),"$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00",$htmlwhite))
								}
								$val++
							}
						}
					}
				}
			}

			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PoolSize[$i] -gt 0)
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				$rowdata += @(,('Weekend number machines powered on, and when',($global:htmlsb),"None",$htmlwhite))
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekend")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PoolSize[$i] -gt 0)
							{
								If($val -eq 0)
								{
									$rowdata += @(,('Weekend number machines powered on, and when',($global:htmlsb),"$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00",$htmlwhite))
								}
								Else
								{
									$rowdata += @(,('',($global:htmlsb),"$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00",$htmlwhite))
								}
								$val++
							}
						}
					}
				}
			}

			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				$rowdata += @(,('Weekday Peak hours',($global:htmlsb),"None",$htmlwhite))
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekdays")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PeakHours[$i])
							{
								If($val -eq 0)
								{
									$rowdata += @(,('Weekday Peak hours',($global:htmlsb),"$($i.ToString("00")):00",$htmlwhite))
								}
								Else
								{
									$rowdata += @(,('',($global:htmlsb),"$($i.ToString("00")):00",$htmlwhite))
								}
								$val++
							}
						}
					}
				}
			}

			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 23;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
						}
					}
				}
			}

			If($val -eq 0)
			{
				$rowdata += @(,('Weekend Peak hours',($global:htmlsb),"None",$htmlwhite))
			}
			Else
			{
				$val = 0
				ForEach($PwrMgmt in $PwrMgmts)
				{
					If($PwrMgmt.DaysOfWeek -eq "Weekend")
					{
						For($i=0;$i -le 23;$i++)
						{
							If($PwrMgmt.PeakHours[$i])
							{
								If($val -eq 0)
								{
									$rowdata += @(,('Weekend Peak hours',($global:htmlsb),"$($i.ToString("00")):00",$htmlwhite))
								}
								Else
								{
									$rowdata += @(,('',($global:htmlsb),"$($i.ToString("00")):00",$htmlwhite))
								}
								$val++
							}
						}
					}
				}
			}

			$rowdata += @(,("During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins",($global:htmlsb),$xPeakDisconnectAction,$htmlwhite))
			$rowdata += @(,("During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins",($global:htmlsb),$xPeakExtendedDisconnectAction,$htmlwhite))
			$rowdata += @(,("During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins",($global:htmlsb),$xOffPeakDisconnectAction,$htmlwhite))
			$rowdata += @(,("During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins",($global:htmlsb),$xOffPeakExtendedDisconnectAction,$htmlwhite))
		}
		
		$rowdata += @(,("Automatic power on for assigned",($global:htmlsb), $xAutoPowerOnForAssigned,$htmlwhite))
		$rowdata += @(,("Automatic power on for assigned during peak",($global:htmlsb), $xAutoPowerOnForAssignedDuringPeak,$htmlwhite))
		If($Group.ReuseMachinesWithoutShutdownInOutage -eq $Script:CCSite1.ReuseMachinesWithoutShutdownInOutageAllowed)
		{
			$rowdata += @(,("Reuse Machines Without Shutdown in Outage",($global:htmlsb),$Group.ReuseMachinesWithoutShutdownInOutage.ToString(),$htmlwhite))
		}
		Else
		{
			$rowdata += @(,("Reuse Machines Without Shutdown in Outage",($global:htmlsb),"$($Group.ReuseMachinesWithoutShutdownInOutage.ToString()) (Doesn't match Site setting)",$htmlwhite))
		}

		$rowdata += @(,('All connections not through NetScaler Gateway',($global:htmlsb),$xAllConnections,$htmlwhite))
		$rowdata += @(,('Connections through NetScaler Gateway',($global:htmlsb),$xNSConnection,$htmlwhite))
		[string]$AGF = $(If( $xAGFilters -is [array] -and $xAGFilters.Count ) { $xAGFilters[0] } Else { '-' } )
		$rowdata += @(,('Connections meeting any of the following filters',($global:htmlsb),$AGF,$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xAGFilters)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
			}
		}

		$msg = ""
		$columnWidths = @("350","350")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "700"
	}
}

Function OutputDeliveryGroupApplicationDetails 
{
	Param([object] $Group)
	
	$AllApplications = Get-BrokerApplication @CCParams2 -AssociatedDesktopGroupUid $Group.Uid -SortBy "ApplicationName"
	
	If($? -and $Null -ne $AllApplications)
	{
		$txt = "Applications"
		If($MSWord -or $PDF)
		{
			WriteWordLine 4 0 $txt
		}
		If($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 4 0 $txt
		}

		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $AllApplicationsWordTable = @();
		}
		If($HTML)
		{
			$rowdata = @()
		}

		ForEach($Application in $AllApplications)
		{
			Write-Verbose "$(Get-Date -Format G): `t`t`tAdding Application $($Application.ApplicationName)"

			$xEnabled = "Enabled"
			If($Application.Enabled -eq $False)
			{
				$xEnabled = "Disabled"
			}
			
			$xLocation = "Master Image"
			If($Application.MetadataKeys.Count -gt 0)
			{
				$xLocation = "App-V"
			}
			
			If($MSWord -or $PDF)
			{
				$AllApplicationsWordTable += @{
				ApplicationName = $Application.ApplicationName; 
				Description = $Application.Description; 
				Location = $xLocation;
				Enabled = $xEnabled; 
				}
			}
			If($Text)
			{
				Line 1 "Name`t`t: " $Application.ApplicationName
				Line 1 "Description`t: " $Application.Description
				Line 1 "Location`t: " $xLocation
				Line 1 "State`t`t: " $xEnabled
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Application.ApplicationName,$htmlwhite,
				$Application.Description,$htmlwhite,
				$xLocation,$htmlwhite,
				$xEnabled,$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $AllApplicationsWordTable `
			-Columns  ApplicationName,Description,Location,Enabled `
			-Headers  "Name","Description","Location","State" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 175;
			$Table.Columns.Item(2).Width = 170;
			$Table.Columns.Item(3).Width = 100;
			$Table.Columns.Item(4).Width = 55;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
			'Name',($global:htmlsb),
			'Description',($global:htmlsb),
			'Location',($global:htmlsb),
			'State',($global:htmlsb))

			$msg = ""
			$columnWidths = @("175","170","100","55")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
		}
	}
}

Function OutputDeliveryGroupCatalogs 
{
	Param([object] $Group)
	
	$MCs = @(Get-BrokerMachine @CCParams2 -DesktopGroupUid $Group.Uid -Property CatalogName -SortBy CatalogName)
	
	If($? -and $Null -ne $MCs)
	{
		If($MCs.Count -gt 1)
		{
			#Adding -Property CatalogName was needed to get the full unique array Returned
			[array]$MCs = $MCs | Sort-Object -Property CatalogName -Unique
		}
		
		$txt = "Machine Catalogs"
		If($MSWord -or $PDF)
		{
			WriteWordLine 4 0 $txt
			[System.Collections.Hashtable[]] $CatalogsWordTable = @();
		}
		If($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 4 0 $txt
			$rowdata = @()
		}

		If($MCs.Count -gt 1)
		{
			ForEach($MC in $MCs)
			{
				Write-Verbose "$(Get-Date -Format G): `t`t`tAdding catalog $($MC.CatalogName)"

				$Catalog = Get-BrokerCatalog @CCParams2 -Name $MC.CatalogName
				If($? -and $Null -ne $Catalog)
				{
					Switch ($Catalog.AllocationType)
					{
						"Static"	{$xAllocationType = "Permanent"; Break}
						"Permanent"	{$xAllocationType = "Permanent"; Break}
						"Random"	{$xAllocationType = "Random"; Break}
						Default		{$xAllocationType = "Allocation type could not be determined: $($Catalog.AllocationType)"; Break}
					}
					
					If($MSWord -or $PDF)
					{
						$CatalogsWordTable += @{
						Name = $Catalog.Name; 
						Type = $xAllocationType; 
						DesktopsTotal = $Catalog.AssignedCount;
						DesktopsFree = $Catalog.AvailableCount; 
						}
					}
					If($Text)
					{
						Line 1 "Machine Catalog name`t: " $Catalog.Name
						Line 1 "Machine Catalog type`t: " $xAllocationType
						Line 1 "Desktops total`t`t: " $Catalog.AssignedCount
						Line 1 "Desktops free`t`t: " $Catalog.AvailableCount
						Line 0 ""
					}
					If($HTML)
					{
						$rowdata += @(,(
						$Catalog.Name,$htmlwhite,
						$xAllocationType,$htmlwhite,
						$Catalog.AssignedCount.ToString(),$htmlwhite,
						$Catalog.AvailableCount.ToString(),$htmlwhite))
					}
				}
			}

			If($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $CatalogsWordTable `
				-Columns  Name,Type,DesktopsTotal,DesktopsFree `
				-Headers  "Machine Catalog name","Machine Catalog type","Desktops total","Desktops free" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 175;
				$Table.Columns.Item(2).Width = 150;
				$Table.Columns.Item(3).Width = 100;
				$Table.Columns.Item(4).Width = 75;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($HTML)
			{
				$columnHeaders = @(
				'Machine Catalog name',($global:htmlsb),
				'Machine Catalog type',($global:htmlsb),
				'Desktops total',($global:htmlsb),
				'Desktops free',($global:htmlsb))

				$msg = ""
				$columnWidths = @("175","150","100","100")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "525"
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "There are no Machine Catalogs for Delivery Group " $Group.Name
			}
			If($Text)
			{
				Line 0 "There are no Machine Catalogs for Delivery Group " $Group.Name
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "There are no Machine Catalogs for Delivery Group " $Group.Name
			}
		}
	}
}

Function OutputDeliveryGroupUtilization
{
	Param([object]$Group)

	#code contributed by Eduardo Molina
	#Twitter: @molikop
	#eduardo@molikop.com
	#www.molikop.com

	$txt = "Delivery Group Utilization Report"
	If($MSWord -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing Utilization for $($Group.Name)"
		WriteWordLine 3 0 $txt
		WriteWordLine 4 0 "Desktop Group Name: " $Group.Name

		$xEnabled = ""
		If($Group.Enabled -eq $True -and $Group.InMaintenanceMode -eq $True)
		{
			$xEnabled = "Maintenance Mode"
		}
		ElseIf($Group.Enabled -eq $False -and $Group.InMaintenanceMode -eq $True)
		{
			$xEnabled = "Maintenance Mode"
		}
		ElseIf($Group.Enabled -eq $True -and $Group.InMaintenanceMode -eq $False)
		{
			$xEnabled = "Enabled"
		}
		ElseIf($Group.Enabled -eq $False -and $Group.InMaintenanceMode -eq $False)
		{
			$xEnabled = "Disabled"
		}

		$xColorDepth = ""
		If($Group.ColorDepth -eq "FourBit")
		{
			$xColorDepth = "4bit - 16 colors"
		}
		ElseIf($Group.ColorDepth -eq "EightBit")
		{
			$xColorDepth = "8bit - 256 colors"
		}
		ElseIf($Group.ColorDepth -eq "SixteenBit")
		{
			$xColorDepth = "16bit - High color"
		}
		ElseIf($Group.ColorDepth -eq "TwentyFourBit")
		{
			$xColorDepth = "24bit - True color"
		}

		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Description"; Value = $Group.Description; }) > $Null
		$ScriptInformation.Add(@{Data = "User Icon Name"; Value = $Group.PublishedName; }) > $Null
		$ScriptInformation.Add(@{Data = "Desktop Type"; Value = $Group.DesktopKind; }) > $Null
		$ScriptInformation.Add(@{Data = "Status"; Value = $xEnabled; }) > $Null
		$ScriptInformation.Add(@{Data = "Automatic reboots when user logs off"; Value = $Group.ShutdownDesktopsAfterUse; }) > $Null
		$ScriptInformation.Add(@{Data = "Color Depth"; Value = $xColorDepth; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
		
		Write-Verbose "$(Get-Date -Format G): `t`t`tInitializing utilization chart for $($Group.Name)"

		$TempFile =  "$($pwd)\emtempgraph_$(Get-Date -UFormat %Y%m%d_%H%M%S).csv"		
		Write-Verbose "$(Get-Date -Format G): `t`t`tGetting utilization data for $($Group.Name)"
		$Results = Get-BrokerDesktopUsage @CCParams2 -DesktopGroupName $Group.Name -SortBy Timestamp | Select-Object Timestamp, InUse

		If($? -and $Null -ne $Results)
		{
			$Results | Export-Csv $TempFile -NoTypeInformation *>$Null

			#Create excel COM object 
			$excel = New-Object -ComObject excel.application 4>$Null

			#Make not visible 
			$excel.Visible  = $False
			$excel.DisplayAlerts  = $False

			#Various Enumerations 
			$xlDirection = [Microsoft.Office.Interop.Excel.XLDirection] 
			$excelChart = [Microsoft.Office.Interop.Excel.XLChartType]
			$excelAxes = [Microsoft.Office.Interop.Excel.XlAxisType]
			$excelCategoryScale = [Microsoft.Office.Interop.Excel.XlCategoryType]
			$excelTickMark = [Microsoft.Office.Interop.Excel.XlTickMark]

			Write-Verbose "$(Get-Date -Format G): `t`t`tOpening Excel with temp file $($TempFile)"

			#Add CSV File into Excel Workbook 
			$Null = $excel.Workbooks.Open($TempFile)
			$worksheet = $excel.ActiveSheet
			$Null = $worksheet.UsedRange.EntireColumn.AutoFit()

			#Assumes that date is always on A column 
			$range = $worksheet.Range("A2")
			$selectionXL = $worksheet.Range($range,$range.end($xlDirection::xlDown))
			#$Start = @($selectionXL)[0].Text
			#$End = @($selectionXL)[-1].Text

			Write-Verbose "$(Get-Date -Format G): `t`t`tCreating chart for $($Group.Name)"
			$chart = $worksheet.Shapes.AddChart().Chart 

			$chart.chartType = $excelChart::xlXYScatterLines
			$chart.HasLegend = $false
			$chart.HasTitle = $true
			$chart.ChartTitle.Text = "$($Group.Name) utilization"

			#Work with the X axis for the Date Stamp 
			$xaxis = $chart.Axes($excelAxes::XlCategory)                                     
			$xaxis.HasTitle = $False
			$xaxis.CategoryType = $excelCategoryScale::xlCategoryScale
			$xaxis.MajorTickMark = $excelTickMark::xlTickMarkCross
			$xaxis.HasMajorGridLines = $true
			$xaxis.TickLabels.NumberFormat = "m/d/yyyy"
			$xaxis.TickLabels.Orientation = 48 #degrees to rotate text

			#Work with the Y axis for the number of desktops in use                                               
			$yaxis = $chart.Axes($excelAxes::XlValue)
			$yaxis.HasTitle = $true                                                       
			$yaxis.AxisTitle.Text = "Desktops in use"
			$yaxis.AxisTitle.Font.Size = 12

			$worksheet.ChartObjects().Item(1).copy()
			$word.Selection.PasteAndFormat(13)  #Pastes an Excel chart as a picture

			Write-Verbose "$(Get-Date -Format G): `t`t`tClosing excel for $($Group.Name)"
			$excel.Workbooks.Close($false)
			$excel.Quit()

			FindWordDocumentEnd
			WriteWordLine 0 0 ""
			
			While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($selectionXL)){}
			While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range)){}
			While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Chart)){}
			While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet)){}
			While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)){}

			If(Test-Path variable:excel)
			{
				Remove-Variable -Name excel 4>$Null
			}

			#If the Excel.exe process is still running for the user's sessionID, kill it
			$SessionID = (Get-Process -PID $PID).SessionId
			(Get-Process 'Excel' -ea 0 | Where-Object {$_.sessionid -eq $Sessionid}) | Stop-Process 4>$Null
			
			Write-Verbose "$(Get-Date -Format G): `t`t`tDeleting temp files $($TempFile)"
			Remove-Item $TempFile *>$Null
		}
		ElseIf($? -and $Null -eq $Results)
		{
			$txt = "There is no Utilization data for $($Group.Name)"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve Utilization data for $($Group.name)"
			OutputWarning $txt
		}
	}
	Write-Verbose "$(Get-Date -Format G): "
}

Function OutputDeliveryGroupTags
{
	Param([object] $Group)

	$txt = "Tags"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 $txt
		[System.Collections.Hashtable[]] $CatalogsWordTable = @();
	}
	If($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 $txt
		$rowdata = @()
	}

	$GroupTags = $Group.Tags
	
	If($GroupTags.Count -gt 0)
	{
		$Tags = New-Object System.Collections.ArrayList
		ForEach($Tag in $GroupTags)
		{
			$Result = Get-BrokerTag @CCParams2 -Name $Tag
			
			If($? -and $Null -ne $Result)
			{
				$obj = [PSCustomObject] @{
					Name        = $Result.Name				
					Description = $Result.Description				
					AppliedTo   = "Delivery Group"				
				}
				$null = $Tags.Add($obj)
			}
		}
		
		$Tags = $Tags | Sort-Object -Property Name
		
		ForEach($Tag in $Tags)
		{
			Write-Verbose "$(Get-Date -Format G): `t`t`tAdding Tag $($Tag.Name)"

			If($MSWord -or $PDF)
			{
				$CatalogsWordTable += @{
				Name = $Tag.Name; 
				Description = $Tag.Description; 
				AppliedTo = $Tag.AppliedTo;
				}
			}
			If($Text)
			{
				Line 1 "Name`t`t: " $Tag.Name
				Line 1 "Description`t: " $Tag.Description
				Line 1 "Applied to`t: " $Tag.AppliedTo
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Tag.Name,$htmlwhite,
				$Tag.Description,$htmlwhite,
				$Tag.AppliedTo,$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $CatalogsWordTable `
			-Columns  Name,Description,AppliedTo `
			-Headers  "Name","Description","Applied to" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 150;
			$Table.Columns.Item(3).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
			'Name',($global:htmlsb),
			'Description',($global:htmlsb),
			'Applied to',($global:htmlsb))

			$msg = ""
			$columnWidths = @("150","150","150")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
		}
	}
	Else
	{
		$txt = "There are no Tags for $($Group.Name)"
		OutputNotice $txt
	}
}
	
Function OutputDeliveryGroupApplicationGroups
{
	Param([object] $Group)

	$txt = "Application Groups"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 $txt
		[System.Collections.Hashtable[]] $CatalogsWordTable = @();
	}
	If($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 $txt
		$rowdata = @()
	}
	
	$ApplicationGroups = Get-BrokerApplicationGroup @CCParams2 -AssociatedDesktopGroupUid $Group.Uid -SortBy Name
	
	If($? -and $Null -ne $ApplicationGroups)
	{
		ForEach($ApplicationGroup in $ApplicationGroups)
		{
			Write-Verbose "$(Get-Date -Format G): `t`t`tAdding Application Group $($ApplicationGroup.Name)"

			If($MSWord -or $PDF)
			{
				$CatalogsWordTable += @{
				Name = $ApplicationGroup.Name; 
				Description = $ApplicationGroup.Description; 
				Applications = $ApplicationGroup.TotalApplications.ToString();
				}
			}
			If($Text)
			{
				Line 1 "Name`t`t: " $ApplicationGroup.Name
				Line 1 "Description`t: " $ApplicationGroup.Description
				Line 1 "Applications`t: " $ApplicationGroup.TotalApplications.ToString()
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata += @(,(
				$ApplicationGroup.Name,$htmlwhite,
				$ApplicationGroup.Description,$htmlwhite,
				$ApplicationGroup.TotalApplications.ToString(),$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $CatalogsWordTable `
			-Columns  Name,Description,Applications `
			-Headers  "Name","Description","Applications" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 150;
			$Table.Columns.Item(3).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
			'Name',($global:htmlsb),
			'Description',($global:htmlsb),
			'Applications',($global:htmlsb))

			$msg = ""
			$columnWidths = @("150","150","150")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
		}
	}
	ElseIf($? -and $Null -eq $ApplicationGroups)
	{
		$txt = "There are no Application Groups for $($Group.Name)"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve Application Groups for $($Group.Name)"
		OutputWarning $txt
	}
}
#endregion

#region process application functions
Function ProcessApplications
{
	Write-Verbose "$(Get-Date -Format G): Retrieving Applications"
	
	$txt = "Applications"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$AllApplications = Get-BrokerApplication @CCParams2 -SortBy "AdminFolderName,ApplicationName"
	If($? -and $Null -ne $AllApplications)
	{
		OutputApplications $AllApplications
	}
	ElseIf($? -and ($Null -eq $AllApplications))
	{
		$txt = "There are no Applications"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve Applications"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date -Format G): "
}

Function OutputApplications
{
	Param([object]$AllApplications)
	
	Write-Verbose "$(Get-Date -Format G): `tProcessing Applications"

	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $AllApplicationsWordTable = @();
	}
	If($HTML)
	{
		$rowdata = @()
	}

	ForEach($Application in $AllApplications)
	{
		Write-Verbose "$(Get-Date -Format G): `t`tAdding Application $($Application.ApplicationName)"

		$xEnabled = "Enabled"
		If($Application.Enabled -eq $False)
		{
			$xEnabled = "Disabled"
		}
		
		$xLocation = "Master Image"
		If($Application.MetadataKeys.Count -gt 0)
		{
			$xLocation = "App-V"
		}

		If($xLocation -eq "Master Image")
		{
			$Script:TotalPublishedApplications++
		}
		Else
		{
			$Script:TotalAppvApplications++
		}
		
		If($MSWord -or $PDF)
		{
			$AllApplicationsWordTable += @{
			FolderName = $Application.AdminFolderName;
			ApplicationName = $Application.ApplicationName; 
			Description = $Application.Description; 
			Location = $xLocation;
			Enabled = $xEnabled; 
			}
		}
		If($Text)
		{
			Line 1 "Folder`t`t: " $Application.AdminFolderName
			Line 1 "Name`t`t: " $Application.ApplicationName
			Line 1 "Description`t: " $Application.Description
			Line 1 "Location`t: " $xLocation
			Line 1 "State`t`t: " $xEnabled
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,(
			$Application.AdminFolderName,$htmlwhite,
			$Application.ApplicationName,$htmlwhite,
			$Application.Description,$htmlwhite,
			$xLocation,$htmlwhite,
			$xEnabled,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $AllApplicationsWordTable `
		-Columns  FolderName,ApplicationName,Description,Location,Enabled `
		-Headers  "Folder","Name","Description","Location","State" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 100;
		$Table.Columns.Item(2).Width = 145;
		$Table.Columns.Item(3).Width = 125;
		$Table.Columns.Item(4).Width = 80;
		$Table.Columns.Item(5).Width = 50;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($HTML)
	{
		$columnHeaders = @(
		'Folder',($global:htmlsb),
		'Name',($global:htmlsb),
		'Description',($global:htmlsb),
		'Location',($global:htmlsb),
		'State',($global:htmlsb))

		$msg = ""
		$columnWidths = @("200","200","125","80","50")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "655"
	}

	If($Applications)
	{
		ForEach($Application in $AllApplications)
		{
			If($MSWord -or $PDF)
			{
				$Selection.InsertNewPage()
				WriteWordLine 2 0 $Application.ApplicationName
			}
			If($Text)
			{
				Line 0 ""
				Line 0 $Application.ApplicationName
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 $Application.ApplicationName
			}
			
			OutputApplicationDetails $Application
			
			If($NoSessions -eq $False)
			{
				OutputApplicationSessions $Application
			}
			OutputApplicationAdministrators $Application
		}
	}
}

Function OutputApplicationDetails
{
	Param([object] $Application)
	
	Write-Verbose "$(Get-Date -Format G): `t`tApplication details for $($Application.ApplicationName)"
	$txt = "Details"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	$xTags = @()
	ForEach($Tag in $Application.Tags)
	{
		$xTags += "$($Tag)"
	}
	$xVisibility = @()
	If($Application.UserFilterEnabled)
	{
		$cnt = -1
		ForEach($tmp in $Application.AssociatedUserFullNames)
		{
			$cnt++
			$xVisibility += "$($tmp) ($($Application.AssociatedUserNames[$cnt]))"
		}
		
	}
	Else
	{
		$xVisibility = {Users inherited from Delivery Group}
	}
	
	$DeliveryGroups = @()
	If($Application.AssociatedDesktopGroupUids.Count -gt 1)
	{
		$cnt = -1
		ForEach($DGUid in $Application.AssociatedDesktopGroupUids)
		{
			$cnt++
			$Results = Get-BrokerDesktopGroup -EA 0 -Uid $DGUid
			If($? -and $Null -ne $Results)
			{
				$DeliveryGroups += "$($Results.Name) Priority: $($Application.AssociatedDesktopGroupPriorities[$cnt])"
			}
		}
	}
	Else
	{
		ForEach($DGUid in $Application.AssociatedDesktopGroupUids)
		{
			$Results = Get-BrokerDesktopGroup -EA 0 -Uid $DGUid
			If($? -and $Null -ne $Results)
			{
				$DeliveryGroups += $Results.Name
			}
		}
	}
	
	$RedirectedFileTypes = @()
	$Results = Get-BrokerConfiguredFTA -ApplicationUid $Application.Uid @CCParams2
	If($? -and $Null -ne $Results)
	{
		ForEach($Result in $Results)
		{
			$RedirectedFileTypes += $Result.ExtensionName
		}
	}
	
	$CPUPriorityLevel = "Normal"
	Switch ($Application.CpuPriorityLevel)
	{
		"Low"			{$CPUPriorityLevel = "Low"; Break}
		"BelowNormal"	{$CPUPriorityLevel = "Below Normal"; Break}
		"Normal"		{$CPUPriorityLevel = "Normal"; Break}
		"AboveNormal"	{$CPUPriorityLevel = "Above Normal"; Break}
		"High"			{$CPUPriorityLevel = "High"; Break}
		Default 		{$CPUPriorityLevel = "Unable to determine CPUPriorityLevel: $($Application.CpuPriorityLevel)"; Break}
	}
	
	$ApplicationType = ""
	Switch ($Application.ApplicationType)
	{
		"HostedOnDesktop"	{$ApplicationType = "Hosted on Desktop"; Break}
		"InstalledOnClient"	{$ApplicationType = "Installed on Client"; Break}
		"PublishedContent"	{$ApplicationType = "Published Content"; Break}
		Default 			{$ApplicationType = "Unable to determine ApplicationType: $($Application.ApplicationType)"; Break}
	}
	
	If([String]::IsNullOrEmpty($Application.HomeZoneName))
	{
		$xHomeZoneName = "Not configured"
	}
	Else
	{
		$xHomeZoneName = $Application.HomeZoneName
	}
	
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Name (for administrator)"; Value = $Application.Name; }
		$ScriptInformation += @{Data = "Name (for user)"; Value = $Application.PublishedName; }
		$ScriptInformation += @{Data = "Description and keywords"; Value = $Application.Description; }
		[string]$xDGs = $(If( $DeliveryGroups -is [array] -and $DeliveryGroups.Count ) { $DeliveryGroups[0] } Else { '-' } )
		$ScriptInformation += @{Data = "Delivery Group"; Value = $xDGs; }
		$cnt = -1
		ForEach($Group in $DeliveryGroups)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $Group; }
			}
		}
		$ScriptInformation += @{Data = "Folder (for administrators)"; Value = $Application.AdminFolderName; }
		$ScriptInformation += @{Data = "Application category (optional)"; Value = $Application.ClientFolder; }
		[string]$xVis = $(If( $xVisibility -is [array] -and $xVisibility.Count ) { $xVisibility[0] } Else { '-' } )
		$ScriptInformation += @{Data = "Visibility"; Value = $xVis; }
		$cnt = -1
		ForEach($tmp in $xVisibility)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $xVisibility[$cnt]; }
			}
		}
		$ScriptInformation += @{Data = "Application Path"; Value = $Application.CommandLineExecutable; }
		$ScriptInformation += @{Data = "Command line arguments"; Value = $Application.CommandLineArguments; }
		$ScriptInformation += @{Data = "Working directory"; Value = $Application.WorkingDirectory; }

		If([String]::IsNullOrEmpty($RedirectedFileTypes))
		{
			$ScriptInformation += @{Data = "Redirected file types"; Value = "-"; }
		}
		Else
		{
			$tmp1 = ""
			ForEach($tmp in $RedirectedFileTypes)
			{
				$tmp1 += "$($tmp); "
			}
			$ScriptInformation += @{Data = "Redirected file types"; Value = $tmp1; }
			$tmp1 = $Null
		}

        [string]$TagName = $(If( $xTags -is [array] -and $xTags.Count ) { $xTags[0] } Else { '-' } )
		$ScriptInformation += @{Data = "Tags"; Value = $TagName; }
		$cnt = -1
		ForEach($tmp in $xTags)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		
		If($Application.Visible -eq $False)
		{
			$ScriptInformation += @{Data = "Hidden"; Value = "Application is hidden"; }
		}
		
		If($Application.MaxTotalInstances -eq 0)
		{
			$ScriptInformation += @{Data = "How do you want to control the use of this application?"; Value = "Allow unlimited use"; }
		}
		Else
		{
			$ScriptInformation += @{Data = "How do you want to control the use of this application?"; Value = ""; }
			$ScriptInformation += @{Data = "     Limit the number of instances running at the same time to"; Value = $Application.MaxTotalInstances.ToString(); }
		}
		
		If($Application.MaxPerUserInstances -eq 0)
		{
		}
		Else
		{
			$ScriptInformation += @{Data = "     Limit to one instance per user"; Value = ""; }
		}
		
		If($Application.MaxPerMachineInstances -eq 0)
		{
			$ScriptInformation += @{Data = "     Limit the number of instances per machine to"; Value = "Unlimited"; }
		}
		Else
		{
			$ScriptInformation += @{Data = "     Limit the number of instances per machine to"; Value = $Application.MaxPerMachineInstances.ToString(); }
		}
		
		$ScriptInformation += @{Data = "Application Type"; Value = $ApplicationType; }
		$ScriptInformation += @{Data = "CPU Priority Level"; Value = $CPUPriorityLevel; }
		$ScriptInformation += @{Data = "Home Zone Name"; Value = $xHomeZoneName; }
		$ScriptInformation += @{Data = "Home Zone Only"; Value = $Application.HomeZoneOnly.ToString(); }
		$ScriptInformation += @{Data = "Ignore User Home Zone"; Value = $Application.IgnoreUserHomeZone.ToString(); }
		$ScriptInformation += @{Data = "Icon from Client"; Value = $Application.IconFromClient; }
		$ScriptInformation += @{Data = "Local Launch Disabled"; Value = $Application.LocalLaunchDisabled.ToString(); }
		$ScriptInformation += @{Data = "Secure Command Line Arguments Enabled"; Value = $Application.SecureCmdLineArgumentsEnabled.ToString(); }
		$ScriptInformation += @{Data = "Add shortcut to user's desktop"; Value = $Application.ShortcutAddedToDesktop.ToString(); }
		$ScriptInformation += @{Data = "Add shortcut to user's Start Menu"; Value = $Application.ShortcutAddedToStartMenu.ToString(); }
		If($Application.ShortcutAddedToStartMenu)
		{
			$ScriptInformation += @{Data = "Start Menu Folder"; Value = $Application.StartMenuFolder ; }
		}
		$ScriptInformation += @{Data = "Wait for Printer Creation"; Value = $Application.WaitForPrinterCreation.ToString(); }
		
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 175;
		$Table.Columns.Item(2).Width = 325;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 1 "Name (for administrator)`t`t: " $Application.Name
		Line 1 "Name (for user)`t`t`t`t: " $Application.PublishedName
		Line 1 "Description and keywords`t`t: " $Application.Description
		[string]$xDGs = $(If( $DeliveryGroups -is [array] -and $DeliveryGroups.Count ) { $DeliveryGroups[0] } Else { '-' } )
		Line 1 "Delivery Group`t`t`t`t: " $xDGs
		$cnt = -1
		ForEach($Group in $DeliveryGroups)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $Group
			}
		}
		Line 1 "Folder (for administrators)`t`t: " $Application.AdminFolderName
		Line 1 "Application category (optional)`t`t: " $Application.ClientFolder
		[string]$xVis = $(If( $xVisibility -is [array] -and $xVisibility.Count ) { $xVisibility[0] } Else { '-' } )
		Line 1 "Visibility`t`t`t`t: " $xVis
		$cnt = -1
		ForEach($tmp in $xVisibility)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 6 "  " $xVisibility[$cnt]
			}
		}
		Line 1 "Application Path`t`t`t: " $Application.CommandLineExecutable
		Line 1 "Command line arguments`t`t`t: " $Application.CommandLineArguments
		Line 1 "Working directory`t`t`t: " $Application.WorkingDirectory

		If([String]::IsNullOrEmpty($RedirectedFileTypes))
		{
			Line 1 "Redirected file types`t`t`t: " "-"
		}
		Else
		{
			$tmp1 = ""
			ForEach($tmp in $RedirectedFileTypes)
			{
				$tmp1 += "$($tmp); "
			}
			Line 1 "Redirected file types`t`t`t: " $tmp1
			$tmp1 = $Null
		}

        [string]$TagName = $(If( $xTags -is [array] -and $xTags.Count ) { $xTags[0] } Else { '-' } )
		Line 1 "Tags`t`t`t`t`t: " $TagName
		$cnt = -1
		ForEach($tmp in $xTags)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 6 "  " $tmp
			}
		}

		If($Application.Visible -eq $False)
		{
			Line 1 "Hidden`t`t`t`t`t: Application is hidden" ""
		}
		
		Line 1 "How do you want to control the use of this application?"
		
		If($Application.MaxTotalInstances -eq 0)
		{
			Line 2 "Allow unlimited use"
		}
		Else
		{
			Line 2 "Limit the number of instances running at the same time to " $Application.MaxTotalInstances.ToString()
		}
		
		If($Application.MaxPerUserInstances -eq 0)
		{
		}
		Else
		{
			Line 2 "Limit to one instance per user"
		}
		
		If($Application.MaxPerMachineInstances -eq 0)
		{
			Line 2 "Limit the number of instances per machine to: Unlimited"
		}
		Else
		{
			Line 2 "Limit the number of instances per machine to " $Application.MaxPerMachineInstances.ToString()
		}

		Line 1 "Application Type`t`t`t: " $ApplicationType
		Line 1 "CPU Priority Level`t`t`t: " $CPUPriorityLevel
		Line 1 "Home Zone Name`t`t`t`t: " $xHomeZoneName
		Line 1 "Home Zone Only`t`t`t`t: " $Application.HomeZoneOnly.ToString()
		Line 1 "Ignore User Home Zone`t`t`t: " $Application.IgnoreUserHomeZone.ToString()
		Line 1 "Icon from Client`t`t`t: " $Application.IconFromClient
		Line 1 "Local Launch Disabled`t`t`t: " $Application.LocalLaunchDisabled.ToString()
		Line 1 "Secure Command Line Arguments Enabled`t: " $Application.SecureCmdLineArgumentsEnabled.ToString()
		Line 1 "Add shortcut to user's desktop`t`t: " $Application.ShortcutAddedToDesktop.ToString()
		Line 1 "Add shortcut to user's Start Menu`t: " $Application.ShortcutAddedToStartMenu.ToString()
		If($Application.ShortcutAddedToStartMenu)
		{
			Line 1 "Start Menu Folder`t`t`t: " $Application.StartMenuFolder 
		}
		Line 1 "Wait for Printer Creation`t`t: " $Application.WaitForPrinterCreation.ToString()
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name (for administrator)",($global:htmlsb),$Application.Name,$htmlwhite)
		$rowdata += @(,('Name (for user)',($global:htmlsb),$Application.PublishedName,$htmlwhite))
		$rowdata += @(,('Description and keywords',($global:htmlsb),$Application.Description,$htmlwhite))
		[string]$xDGs = $(If( $DeliveryGroups -is [array] -and $DeliveryGroups.Count ) { $DeliveryGroups[0] } Else { '-' } )
		$rowdata += @(,('Delivery Group',($global:htmlsb),$xDGs,$htmlwhite))
		$cnt = -1
		ForEach($Group in $DeliveryGroups)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($global:htmlsb),$Group,$htmlwhite))
			}
		}
		$rowdata += @(,('Folder (for administrators)',($global:htmlsb),$Application.AdminFolderName,$htmlwhite))
		$rowdata += @(,('Application category (optional)',($global:htmlsb),$Application.ClientFolder,$htmlwhite))
		[string]$xVis = $(If( $xVisibility -is [array] -and $xVisibility.Count ) { $xVisibility[0] } Else { '-' } )
		$rowdata += @(,('Visibility',($global:htmlsb),$xVis,$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xVisibility)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($global:htmlsb),$xVisibility[$cnt],$htmlwhite))
			}
		}
		$rowdata += @(,('Application Path',($global:htmlsb),$Application.CommandLineExecutable,$htmlwhite))
		$rowdata += @(,('Command Line arguments',($global:htmlsb),$Application.CommandLineArguments,$htmlwhite))
		$rowdata += @(,('Working directory',($global:htmlsb),$Application.WorkingDirectory,$htmlwhite))

		If([String]::IsNullOrEmpty($RedirectedFileTypes))
		{
			$rowdata += @(,("Redirected file types",($global:htmlsb),"-",$htmlwhite))
		}
		Else
		{
			$tmp1 = ""
			ForEach($tmp in $RedirectedFileTypes)
			{
				$tmp1 += "$($tmp); "
			}
			$rowdata += @(,("Redirected file types",($global:htmlsb),$tmp1,$htmlwhite))
			$tmp1 = $Null
		}

        [string]$TagName = $(If( $xTags -is [array] -and $xTags.Count ) { $xTags[0] } Else { '-' } )
		$rowdata += @(,('Tags',($global:htmlsb),$TagName,$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xTags)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
			}
		}

		If($Application.Visible -eq $False)
		{
			$rowdata += @(,('Hidden',($global:htmlsb),"Application is hidden",$htmlwhite))
		}

		If($Application.MaxTotalInstances -eq 0)
		{
			$rowdata += @(,("How do you want to control the use of this application?",($global:htmlsb),"Allow unlimited use",$htmlwhite))
		}
		Else
		{
			$rowdata += @(,("How do you want to control the use of this application?",($global:htmlsb),"",$htmlwhite))
			$rowdata += @(,("     Limit the number of instances running at the same time to",($global:htmlsb),$Application.MaxTotalInstances.ToString(),$htmlwhite))
		}
		
		If($Application.MaxPerUserInstances -eq 0)
		{
		}
		Else
		{
			$rowdata += @(,("     Limit to one instance per user",($global:htmlsb),"",$htmlwhite))
		}
		
		If($Application.MaxPerMachineInstances -eq 0)
		{
			$rowdata += @(,("     Limit the number of instances per machine to",($global:htmlsb),"Unlimited",$htmlwhite))
		}
		Else
		{
			$rowdata += @(,("     Limit the number of instances per machine to",($global:htmlsb),$Application.MaxPerMachineInstances.ToString(),$htmlwhite))
		}

		$rowdata += @(,("Application Type",($global:htmlsb),$ApplicationType,$htmlwhite))
		$rowdata += @(,("CPU Priority Level",($global:htmlsb),$CPUPriorityLevel,$htmlwhite))
		$rowdata += @(,("Home Zone Name",($global:htmlsb),$xHomeZoneName,$htmlwhite))
		$rowdata += @(,("Home Zone Only",($global:htmlsb),$Application.HomeZoneOnly.ToString(),$htmlwhite))
		$rowdata += @(,("Ignore User Home Zone",($global:htmlsb),$Application.IgnoreUserHomeZone.ToString(),$htmlwhite))
		$rowdata += @(,("Icon from Client",($global:htmlsb),$Application.IconFromClient,$htmlwhite))
		$rowdata += @(,("Local Launch Disabled",($global:htmlsb),$Application.LocalLaunchDisabled.ToString(),$htmlwhite))
		$rowdata += @(,("Secure Command Line Arguments Enabled",($global:htmlsb),$Application.SecureCmdLineArgumentsEnabled.ToString(),$htmlwhite))
		$rowdata += @(,("Add shortcut to user's desktop",($global:htmlsb),$Application.ShortcutAddedToDesktop.ToString(),$htmlwhite))
		$rowdata += @(,("Add shortcut to user's Start Menu",($global:htmlsb),$Application.ShortcutAddedToStartMenu.ToString(),$htmlwhite))
		If($Application.ShortcutAddedToStartMenu)
		{
			$rowdata += @(,("Start Menu Folder",($global:htmlsb),$Application.StartMenuFolder,$htmlwhite)) 
		}
		$rowdata += @(,("Wait for Printer Creation",($global:htmlsb),$Application.WaitForPrinterCreation.ToString(),$htmlwhite))

		$msg = ""
		$columnWidths = @("300","325")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "625"
	}
}

Function OutputApplicationSessions
{
	Param([object] $Application)
	
	Write-Verbose "$(Get-Date -Format G): `t`tApplication sessions for $($Application.BrowserName)"
	$txt = "Sessions"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	$Sessions = Get-BrokerSession -ApplicationUid $Application.Uid @CCParams2 -SortBy UserName
	
	If($? -and $Null -ne $Sessions)
	{
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $SessionsWordTable = @();
		}
		If($HTML)
		{
			$rowdata = @()
		}

		#now get the privateappdesktop for each desktopgroup uid
		ForEach($Session in $Sessions)
		{
			#get desktop by Session Uid
			$xMachineName = ""
			$Desktop = Get-BrokerMachine -SessionUid $Session.Uid @CCParams2
			
			If($? -and $Null -ne $Desktop)
			{
				$xMachineName = $Desktop.MachineName
			}
			Else
			{
				If(![String]::IsNullOrEmpty($Session.MachineName))
				{
					$xMachineName = $Session.MachineName
				}
				Else
				{
					$xMachineName = "Not Found"
				}
			}
			
			#$RecordingStatus = "Not supported"
			#$result = Get-BrokerSessionRecordingStatus -Session $Session.Uid
			
			#If($?)
			#{
			#	Switch ($result)
			#	{
			#		"SessionBeingRecorded"	{$RecordingStatus = "Session is being recorded"}
			#		"SessionNotRecorded"	{$RecordingStatus = "Session is not being recorded"}
			#		Default					{$RecordingStatus = "Unable to determine session recording status: $($result)"}
			#	}
			#}
			#Else
			#{
			#	$RecordingStatus = "Unknown"
			#}
			
			If($MSWord -or $PDF)
			{
				$SessionsWordTable += @{
				UserName = $Session.UserName;
				ClientName= $Session.ClientName;
				MachineName = $xMachineName;
				State = $Session.SessionState;
				ApplicationState = $Session.AppState;
				Protocol = $Session.Protocol
				}
				#RecordingStatus = $RecordingStatus
			}
			If($Text)
			{
				Line 2 "User Name`t`t: " $Session.UserName
				Line 2 "Client Name`t`t: " $Session.ClientName
				Line 2 "Machine Name`t`t: " $xMachineName
				Line 2 "State`t`t`t: " $Session.SessionState
				Line 2 "Application State`t: " $Session.AppState
				Line 2 "Protocol`t`t: " $Session.Protocol
				#Line 2 "Recording Status`t: " $RecordingStatus
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Session.UserName,$htmlwhite,
				$Session.ClientName,$htmlwhite,
				$xMachineName,$htmlwhite,
				$Session.SessionState,$htmlwhite,
				$Session.AppState,$htmlwhite,
				$Session.Protocol,$htmlwhite
				))
				#$RecordingStatus,$htmlwhite
			}
		}
		
		If($MSWord -or $PDF)
		{
			#-Columns  UserName,ClientName,MachineName,State,ApplicationState,Protocol,RecordingStatus `
			#-Headers  "User Name","Client Name","Machine Name","State","Application State","Protocol","Recording Status" `
			$Table = AddWordTable -Hashtable $SessionsWordTable `
			-Columns  UserName,ClientName,MachineName,State,ApplicationState,Protocol `
			-Headers  "User Name","Client Name","Machine Name","State","Application State","Protocol" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 120;
			$Table.Columns.Item(2).Width = 70;
			$Table.Columns.Item(3).Width = 125;
			$Table.Columns.Item(4).Width = 35;
			$Table.Columns.Item(5).Width = 55;
			$Table.Columns.Item(6).Width = 45;
			#$Table.Columns.Item(6).Width = 50; #recording status column

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		If($HTML)
		{
			$columnHeaders = @(
			'User Name',($global:htmlsb),
			'Client Name',($global:htmlsb),
			'Machine Name',($global:htmlsb),
			'State',($global:htmlsb),
			'Application State',($global:htmlsb),
			'Protocol',($global:htmlsb)
			)
			#'Recording Status',($global:htmlsb)

			#$columnWidths = @("135","85","135","50","50","55","55")
			#FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "510"
			$msg = ""
			$columnWidths = @("135","85","135","50","50","55")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "455"
			WriteHTMLLine 0 0 ""
		}
	}
	ElseIf($? -and $Null -eq $Sessions)
	{
		$txt = "There are no Sessions for Application $($Application.ApplicationName)"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve Sessions for Application $($Application.ApplicationName)"
		OutputWarning $txt
	}
}

Function OutputApplicationAdministrators
{
	Param([object] $Application)
	
	Write-Verbose "$(Get-Date -Format G): `t`tApplication administrators for $($Application.ApplicationName)"
	$txt = "Administrators"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	If($Text)
	{
		Line 0 ""
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	#get all the delivery groups
	$DeliveryGroups = @()
	ForEach($DGUid in $Application.AssociatedDesktopGroupUids)
	{
		$Results = Get-BrokerDesktopGroup -EA 0 -Uid $DGUid
		If($? -and $Null -ne $Results)
		{
			$DeliveryGroups += $Results.Name
		}
	}
	
	#now get the administrators for each delivery group
	$Admins = @()
	ForEach($Group in $DeliveryGroups)
	{
		$Results = GetAdmins "DesktopGroup" $Group
		If($? -and $Null -ne $Results)
		{
			$Admins += $Results
		}
	}
	
	If($Null -ne $Admins)
	{
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $AdminsWordTable = @();
		}
		If($HTML)
		{
			$rowdata = @()
		}
		
		ForEach($Admin in $Admins)
		{
			$Tmp = ""
			If($Admin.Enabled)
			{
				$Tmp = "Enabled"
			}
			Else
			{
				$Tmp = "Disabled"
			}
			
			If($MSWord -or $PDF)
			{
				$AdminsWordTable += @{ 
				AdminName = $Admin.Name;
				Role = $Admin.Rights[0].RoleName;
				Status = $Tmp;
				}
			}
			If($Text)
			{
				Line 1 "Administrator Name`t: " $Admin.Name
				Line 1 "Role`t`t`t: " $Admin.Rights[0].RoleName
				Line 1 "Status`t`t`t: " $tmp
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Admin.Name,$htmlwhite,
				$Admin.Rights[0].RoleName,$htmlwhite,
				$tmp,$htmlwhite))
			}
		}
		
		If($MSWord -or $PDF)
		{
			If($AdminsWordTable.Count -eq 0)
			{
				$AdminsWordTable += @{ 
				AdminName = "No admins found";
				Role = "N/A";
				Status = "N/A";
				}
			}

			$Table = AddWordTable -Hashtable $AdminsWordTable `
			-Columns AdminName, Role, Status `
			-Headers "Administrator Name", "Role", "Status" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 275;
			$Table.Columns.Item(2).Width = 200;
			$Table.Columns.Item(3).Width = 60;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		If($HTML)
		{
			$columnHeaders = @(
			'Administrator Name',($global:htmlsb),
			'Role',($global:htmlsb),
			'Status',($global:htmlsb))

			$msg = ""
			$columnWidths = @("275","200","60")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "535"
		}
	}
	ElseIf($? -and ($Null -eq $Admins))
	{
		$txt = "There are no administrators for $($Group.Name)"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve administrators for $($Group.Name)"
		OutputWarning $txt
	}
	
}
#endregion

#region application group details
Function ProcessApplicationGroupDetails
{
	Write-Verbose "$(Get-Date -Format G): `tProcessing Application Groups"

	$txt = "Application Groups"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$ApplicationGroups = Get-BrokerApplicationGroup @CCParams2 -SortBy Name
	
	If($? -and $Null -ne $ApplicationGroups)
	{
		ForEach($AppGroup in $ApplicationGroups)
		{
			Write-Verbose "$(Get-Date -Format G): `t`t`tAdding Application Group $($AppGroup.Name)"

			$xEnabled = "No"
			If($AppGroup.Enabled)
			{
				$xEnabled = "Yes"
			}
			
			$xSessionSharing = "Disabled"
			If($AppGroup.SessionSharingEnabled)
			{
				$xSessionSharing = "Enabled"
			}
			
			$xSingleSession = "Disabled"
			If($AppGroup.SingleAppPerSession)
			{
				$xSingleSession = "Enabled"
			}
			
			$DGs = @()
			ForEach($DGUid in $AppGroup.AssociatedDesktopGroupUids)
			{
				$results = Get-BrokerDesktopGroup -EA 0 -Uid $DGUid
				
				If($? -and $Null -ne $results)
				{
					$DGs += $results.Name
				}
			}
			
			[array]$xTags = $AppGroup.Tags
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Name"; Value = $AppGroup.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $AppGroup.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Applications"; Value = $AppGroup.TotalApplications.ToString(); }) > $Null
				
				If([String]::IsNullOrEmpty($AppGroup.Scopes))
				{
					$ScriptInformation.Add(@{Data = "Scopes"; Value = "All"; }) > $Null
				}
				Else
				{
					$ScriptInformation.Add(@{Data = "Scopes"; Value = "All"; }) > $Null
					$cnt = -1
					ForEach($tmp in $AppGroup.Scopes)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $tmp; }) > $Null
					}
				}
				
				$ScriptInformation.Add(@{Data = "Enabled"; Value = $xEnabled; }) > $Null
				$ScriptInformation.Add(@{Data = "Session sharing"; Value = $xSessionSharing; }) > $Null
				$ScriptInformation.Add(@{Data = "Single application per session"; Value = $xSingleSession; }) > $Null
				$ScriptInformation.Add(@{Data = "Restrict launches to machines with tag"; Value = $AppGroup.RestrictToTag; }) > $Null

				[string]$xxDGs = $(If( $DGs -is [array] -and $DGs.Count ) { $DGs[0] } Else { '-' } )
				$ScriptInformation.Add(@{Data = "Delivery Groups"; Value = $xxDGs; }) > $Null
				$cnt = -1
				ForEach($tmp in $DGs)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $tmp; }) > $Null
					}
				}
				
				[string]$TagName = $(If( $xTags -is [array] -and $xTags.Count ) { $xTags[0] } Else { '-' } )
				$ScriptInformation.Add(@{Data = "Tags"; Value = $TagName; }) > $Null
				$cnt = -1
				ForEach($tmp in $xTags)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $tmp; }) > $Null
					}
				}
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 1 "Name`t`t`t`t`t: " $AppGroup.Name
				Line 1 "Description`t`t`t`t: " $AppGroup.Description
				Line 1 "Applications`t`t`t`t: " $AppGroup.TotalApplications.ToString()
				
				If([String]::IsNullOrEmpty($AppGroup.Scopes))
				{
					Line 1 "Scopes`t`t`t`t`t: " "All"
				}
				Else
				{
					Line 1 "Scopes`t`t`t`t`t: " "All"
					$cnt = -1
					ForEach($tmp in $AppGroup.Scopes)
					{
						Line 6 "  " $tmp
					}
				}
				
				Line 1 "Enabled`t`t`t`t`t: " $xEnabled
				Line 1 "Session sharing`t`t`t`t: " $xSessionSharing
				Line 1 "Single application per session`t`t: " $xSingleSession
				Line 1 "Restrict launches to machines with tag`t: " $AppGroup.RestrictToTag

				[string]$xxDGs = $(If( $DGs -is [array] -and $DGs.Count ) { $DGs[0] } Else { '-' } )
				Line 1 "Delivery Groups`t`t`t`t: " $xxDGs
				$cnt = -1
				ForEach($tmp in $DGs)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 6 "  " $tmp
					}
				}
				
				[string]$TagName = $(If( $xTags -is [array] -and $xTags.Count ) { $xTags[0] } Else { '-' } )
				Line 1 "Tags`t`t`t`t`t: " $TagName
				$cnt = -1
				ForEach($tmp in $xTags)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 6 "  " $tmp
					}
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Name",($global:htmlsb),$AppGroup.Name,$htmlwhite)
				$rowdata += @(,('Description',($global:htmlsb),$AppGroup.Description,$htmlwhite))
				$rowdata += @(,('Applications',($global:htmlsb),$AppGroup.TotalApplications.ToString(),$htmlwhite))
				
				If([String]::IsNullOrEmpty($AppGroup.Scopes))
				{
					$rowdata += @(,('Scopes',($global:htmlsb),"All",$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('Scopes',($global:htmlsb),"All",$htmlwhite))
					$cnt = -1
					ForEach($tmp in $AppGroup.Scopes)
					{
						$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
					}
				}
				
				$rowdata += @(,('Enabled',($global:htmlsb),$xEnabled,$htmlwhite))
				$rowdata += @(,('Session sharing',($global:htmlsb),$xSessionSharing,$htmlwhite))
				$rowdata += @(,('Single application per session',($global:htmlsb),$xSingleSession,$htmlwhite))
				$rowdata += @(,('Restrict launches to machines with tag',($global:htmlsb),$AppGroup.RestrictToTag,$htmlwhite))
				
				[string]$xxDGs = $(If( $DGs -is [array] -and $DGs.Count ) { $DGs[0] } Else { '-' } )
				$rowdata += @(,('Delivery Groups',($global:htmlsb),$xxDGs,$htmlwhite))
				$cnt = -1
				ForEach($tmp in $DGs)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
					}
				}
				
				[string]$TagName = $(If( $xTags -is [array] -and $xTags.Count ) { $xTags[0] } Else { '-' } )
				$rowdata += @(,('Tags',($global:htmlsb),$TagName,$htmlwhite))
				$cnt = -1
				ForEach($tmp in $xTags)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("225","275")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
			}
		}
	}
	ElseIf($? -and $Null -eq $ApplicationGroups)
	{
		$txt = "There were no Application Groups found"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Application Groups"
		OutputWarning $txt
	}
}
#endregion

#region policy functions
Function ProcessPolicies
{
	$txt = "Policies"
	$txt1 = "Policies in this report may not match the order shown in Studio."
	$txt2 = "See http://blogs.citrix.com/2013/07/15/merging-of-user-and-computer-policies-in-xendesktop-7-0/"
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 $txt
		WriteWordLine 0 0 $txt1 "" $Null 8 $False $True	
		WriteWordLine 0 0 $txt2 "" $Null 8 $False $True	
	}
	If($Text)
	{
		Line 0 $txt
		Line 0 $txt1
		Line 0 $txt2
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 $txt
		WriteHTMLLine 0 0 $txt1 "" "Calibri" 1
		WriteHTMLLine 0 0 $txt2
	}
	Write-Verbose "$(Get-Date -Format G): Processing CC Policies"
	
	ProcessPolicySummary 
	
	If($Policies)
	{
	
		Write-Verbose "$(Get-Date -Format G): `tDoes localfarmgpo PSDrive already exist?"
		If(Test-Path localfarmgpo:)
		{
			Write-Verbose "$(Get-Date -Format G): `tRemoving the current localfarmgpo PSDrive"
			Remove-PSDrive -Name localfarmgpo -EA 0 4>$Null
		}
		
		Write-Verbose "$(Get-Date -Format G): Creating localfarmgpo PSDrive for Computer policies"
		New-PSDrive -Name localfarmgpo -psprovider citrixgrouppolicy -root \ -controller LocalHost -Scope Global *>$Null
		
		#using Citrix policy stuff and new-psdrive breaks transcript logging so restart transcript logging
		TranscriptLogging
		If(Test-Path localfarmgpo:)
		{
			ProcessCitrixPolicies "localfarmgpo" "Computer" ""
			Write-Verbose "$(Get-Date -Format G): Finished Processing Citrix Site Computer Policies"
			Write-Verbose "$(Get-Date -Format G): "
		}
		Else
		{
			Write-Warning "Unable to create the LocalFarmGPO PSDrive"
		}

		Write-Verbose "$(Get-Date -Format G): Creating localfarmgpo PSDrive for User policies"
		New-PSDrive -Name localfarmgpo -psprovider citrixgrouppolicy -root \ -controller LocalHost -Scope Global *>$Null
		
		#using Citrix policy stuff and new-psdrive breaks transcript logging so restart transcript logging
		TranscriptLogging
		If(Test-Path localfarmgpo:)
		{
			ProcessCitrixPolicies "localfarmgpo" "User" ""
			Write-Verbose "$(Get-Date -Format G): Finished Processing Citrix Site User Policies"
			Write-Verbose "$(Get-Date -Format G): "
		}
		Else
		{
			Write-Warning "Unable to create the LocalFarmGPO PSDrive"
		}
		
		If($NoADPolicies)
		{
			#don't process AD policies
		}
		Else
		{
			#thanks to the Citrix Engineering Team for helping me solve processing Citrix AD-based Policies
			Write-Verbose "$(Get-Date -Format G): "
			Write-Verbose "$(Get-Date -Format G): `tSee if there are any Citrix AD-based policies to process"
			$CtxGPOArray = @()
			$CtxGPOArray = GetCtxGPOsInAD
			If($CtxGPOArray -is [Array] -and $CtxGPOArray.Count -gt 0)
			{
				Write-Verbose "$(Get-Date -Format G): "
				Write-Verbose "$(Get-Date -Format G): `tThere are $($CtxGPOArray.Count) Citrix AD-based policies to process"
				Write-Verbose "$(Get-Date -Format G): "

				[array]$CtxGPOArray = $CtxGPOArray | Sort-Object -unique
				
				ForEach($CtxGPO in $CtxGPOArray)
				{
					Write-Verbose "$(Get-Date -Format G): `tCreating ADGpoDrv PSDrive for Computer Policies"
					New-PSDrive -Name ADGpoDrv -PSProvider CitrixGroupPolicy -Root \ -DomainGpo $($CtxGPO) -Scope Global *>$Null
					
					#using Citrix policy stuff and new-psdrive breaks transcript logging so restart transcript logging
					TranscriptLogging
					If(Test-Path ADGpoDrv:)
					{
						Write-Verbose "$(Get-Date -Format G): `tProcessing Citrix AD Policy $($CtxGPO)"
					
						Write-Verbose "$(Get-Date -Format G): `tRetrieving AD Policy $($CtxGPO)"
						ProcessCitrixPolicies "ADGpoDrv" "Computer" $CtxGPO
						Write-Verbose "$(Get-Date -Format G): Finished Processing Citrix AD Computer Policy $($CtxGPO)"
						Write-Verbose "$(Get-Date -Format G): "
					}
					Else
					{
						Write-Warning "$($CtxGPO) is not readable by this CC Controller"
						Write-Warning "$($CtxGPO) was probably created by an updated Citrix Group Policy Provider"
					}

					Write-Verbose "$(Get-Date -Format G): `tCreating ADGpoDrv PSDrive for UserPolicies"
					New-PSDrive -Name ADGpoDrv -PSProvider CitrixGroupPolicy -Root \ -DomainGpo $($CtxGPO) -Scope Global *>$Null
					
					#using Citrix policy stuff and new-psdrive breaks transcript logging so restart transcript logging
					TranscriptLogging
					If(Test-Path ADGpoDrv:)
					{
						Write-Verbose "$(Get-Date -Format G): `tProcessing Citrix AD Policy $($CtxGPO)"
					
						Write-Verbose "$(Get-Date -Format G): `tRetrieving AD Policy $($CtxGPO)"
						ProcessCitrixPolicies "ADGpoDrv" "User" $CtxGPO
						Write-Verbose "$(Get-Date -Format G): Finished Processing Citrix AD User Policy $($CtxGPO)"
						Write-Verbose "$(Get-Date -Format G): "
					}
					Else
					{
						Write-Warning "$($CtxGPO) is not readable by this CC Controller"
						Write-Warning "$($CtxGPO) was probably created by an updated Citrix Group Policy Provider"
					}
				}
				Write-Verbose "$(Get-Date -Format G): Finished Processing Citrix AD Policies"
				Write-Verbose "$(Get-Date -Format G): "
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): There are no Citrix AD-based policies to process"
				Write-Verbose "$(Get-Date -Format G): "
			}
		}
	}
	Write-Verbose "$(Get-Date -Format G): Finished Processing Citrix Policies"
	Write-Verbose "$(Get-Date -Format G): "
}

Function ProcessPolicySummary
{
	Write-Verbose "$(Get-Date -Format G): `tDoes localfarmgpo PSDrive already exist?"
	If(Test-Path localfarmgpo:)
	{
		Write-Verbose "$(Get-Date -Format G): `tRemoving the current localfarmgpo PSDrive"
		Remove-PSDrive -Name localfarmgpo -EA 0 4>$Null
	}
	Write-Verbose "$(Get-Date -Format G): `tRetrieving Site Policies"
	Write-Verbose "$(Get-Date -Format G): `t`tCreating localfarmgpo PSDrive"
	New-PSDrive -Name localfarmgpo -psprovider citrixgrouppolicy -root \ -controller LocalHost -Scope Global *>$Null
	
	#using Citrix policy stuff and new-psdrive breaks transcript logging so restart transcript logging
	TranscriptLogging

	If(Test-Path localfarmgpo:)
	{
		$HDXPolicies = Get-CtxGroupPolicy -DriveName localfarmgpo -EA 0 `
		| Select-Object PolicyName, Type, Description, Enabled, Priority `
		| Sort-Object Type, Priority
		
		OutputSummaryPolicyTable $HDXPolicies "localfarmgpo"
	}
	Else
	{
		#Write-Warning "Unable to create the LocalFarmGPO PSDrive on the CC Controller $($AdminAddress)"
		Write-Warning "Unable to create the LocalFarmGPO PSDrive"
	}
	
	If($NoADPolicies)
	{
		#don't process AD policies
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): "
		Write-Verbose "$(Get-Date -Format G): See if there are any Citrix AD-based policies to process"
		$CtxGPOArray = @()
		$CtxGPOArray = GetCtxGPOsInAD
		If($CtxGPOArray -is [Array] -and $CtxGPOArray.Count -gt 0)
		{
			[array]$CtxGPOArray = $CtxGPOArray | Sort-Object -unique
			Write-Verbose "$(Get-Date -Format G): "
			Write-Verbose "$(Get-Date -Format G): `tThere are $($CtxGPOArray.Count) Citrix AD-based policies to process"
			Write-Verbose "$(Get-Date -Format G): "
			
			ForEach($CtxGPO in $CtxGPOArray)
			{
				Write-Verbose "$(Get-Date -Format G): `tCreating ADGpoDrv PSDrive"
				New-PSDrive -Name ADGpoDrv -PSProvider CitrixGroupPolicy -Root \ -DomainGpo $($CtxGPO) -Scope "Global" *>$Null
				
				#using Citrix policy stuff and new-psdrive breaks transcript logging so restart transcript logging
				TranscriptLogging
				If(Test-Path ADGpoDrv:)
				{
					Write-Verbose "$(Get-Date -Format G): `tProcessing Citrix AD Policy $($CtxGPO)"
				
					Write-Verbose "$(Get-Date -Format G): `tRetrieving AD Policy $($CtxGPO)"
					$HDXPolicies = Get-CtxGroupPolicy -DriveName ADGpoDrv -EA 0 `
					| Select-Object PolicyName, Type, Description, Enabled, Priority `
					| Sort-Object Type, Priority
			
					OutputSummaryPolicyTable $HDXPolicies "AD" $CtxGPO
					
					Write-Verbose "$(Get-Date -Format G): Finished Processing Citrix AD Policy $($CtxGPO)"
					Write-Verbose "$(Get-Date -Format G): "
				}
				Else
				{
					Write-Warning "$($CtxGPO) is not readable by this CC Controller"
					Write-Warning "$($CtxGPO) was probably created by an updated Citrix Group Policy Provider"
				}
				Remove-PSDrive -Name ADGpoDrv -EA 0 4>$Null
			}
			Write-Verbose "$(Get-Date -Format G): Finished Processing Citrix AD Policies"
			Write-Verbose "$(Get-Date -Format G): "
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): There are no Citrix AD-based policies to process"
			Write-Verbose "$(Get-Date -Format G): "
		}
	}
}

Function OutputSummaryPolicyTable
{
	Param([object] $HDXPolicies, [string] $xLocation, [string] $ADGPOName = "")
	
	If($xLocation -eq "localfarmgpo")
	{
		$txt = "Site Policies"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt
		}
		If($Text)
		{
			Line 0 $txt
		}
		If($HTML)
		{
			WriteHTMLLine 3 0 $txt
		}
	}
	ElseIf($xLocation -eq "AD")
	{
		$txt = "Active Directory Policies ($($ADGpoName))"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt
		}
		If($Text)
		{
			Line 0 $txt
		}
		If($HTML)
		{
			WriteHTMLLine 3 0 $txt
		}
	}

	If($Null -ne $HDXPolicies)
	{
		Write-Verbose "$(Get-Date -Format G): `t`t`tPolicies"
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $PoliciesWordTable = @();
		}
		If($HTML)
		{
			$rowdata = @()
		}

		ForEach($Policy in $HDXPolicies)
		{
			If($MSWord -or $PDF)
			{
				$PoliciesWordTable += @{
				Name = $Policy.PolicyName;
				Description = $Policy.Description;
				Enabled= $Policy.Enabled;
				Type = $Policy.Type;
				Priority = $Policy.Priority;
				}
			}
			If($Text)
			{
				Line 2 "Name`t`t: " $Policy.PolicyName
				Line 2 "Description`t: " $Policy.Description
				Line 2 "Enabled`t`t: " $Policy.Enabled
				Line 2 "Type`t`t: " $Policy.Type
				Line 2 "Priority`t: " $Policy.Priority
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Policy.PolicyName,$htmlwhite,
				$Policy.Description,$htmlwhite,
				$Policy.Enabled.ToString(),$htmlwhite,
				$Policy.Type,$htmlwhite,
				$Policy.Priority,$htmlwhite))
			}
		}
		
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $PoliciesWordTable `
			-Columns  Name,Description,Enabled,Type,Priority `
			-Headers  "Name","Description","Enabled","Type","Priority" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 155
			$Table.Columns.Item(2).Width = 185
			$Table.Columns.Item(3).Width = 55;
			$Table.Columns.Item(4).Width = 60;
			$Table.Columns.Item(5).Width = 45;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		If($HTML)
		{
			$columnHeaders = @(
			'Name',($global:htmlsb),
			'Description',($global:htmlsb),
			'Enabled',($global:htmlsb),
			'Type',($global:htmlsb),
			'Priority',($global:htmlsb))

			$msg = ""
			$columnWidths = @("155","185","55","60","45")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
		}
	}
	ElseIf($Null -eq $HDXPolicies)
	{
		$txt = "There are no Policies"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve Policies"
		OutputWarning $txt
	}
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			If((Get-Member -Name $secondLevel -InputObject $object.$topLevel))
			{
				Return $True
			}
		}
	}
	Return $False
}

Function validObject( [object] $object, [string] $topLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			Return $True
		}
	}
	Return $False
}

Function ProcessCitrixPolicies
{
	Param([string]$xDriveName, [string]$xPolicyType, [string] $ADGpoName = "")

	Write-Verbose "$(Get-Date -Format G): `tRetrieving all $($xPolicyType) policy names"

	$CtxPolicies = Get-CtxGroupPolicy -Type $xPolicyType `
	-DriveName $xDriveName -EA 0 `
	| Select-Object PolicyName, Type, Description, Enabled, Priority `
	| Sort-Object Priority

	If($? -and $Null -ne $CtxPolicies)
	{
		ForEach($Policy in $CtxPolicies)
		{
			Write-Verbose "$(Get-Date -Format G): `tStarted $($Policy.PolicyName) "
			
			If($xDriveName -eq "localfarmgpo")
			{
				$Script:TotalSitePolicies++
			}
			Else
			{
				$Script:TotalADPolicies++
			}
			If($Policy.Type -eq "Computer")
			{
				$Script:TotalComputerPolicies++
			}
			Else
			{
				$Script:TotalUserPolicies++
			}
			$Script:TotalPolicies++
			
			If($MSWord -or $PDF)
			{
				$selection.InsertNewPage()
				If($xDriveName -eq "localfarmgpo")
				{
					WriteWordLine 2 0 "$($Policy.PolicyName) (Site, $($xPolicyType))"
				}
				Else
				{
					WriteWordLine 2 0 "$($Policy.PolicyName) (AD, $($xPolicyType), GPO: $($ADGpoName))"
				}
				[System.Collections.Hashtable[]] $ScriptInformation = @()
			
				$ScriptInformation += @{Data = "Description"; Value = $Policy.Description; }
				$ScriptInformation += @{Data = "Enabled"; Value = $Policy.Enabled; }
				$ScriptInformation += @{Data = "Type"; Value = $Policy.Type; }
				$ScriptInformation += @{Data = "Priority"; Value = $Policy.Priority; }
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 90;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}
			If($Text)
			{
				If($xDriveName -eq "localfarmgpo")
				{
					Line 0 "$($Policy.PolicyName) (Site, $($xPolicyType))"
				}
				Else
				{
					Line 0 "$($Policy.PolicyName) (AD, $($xPolicyType), GPO: $($ADGpoName))"
				}
				If(![String]::IsNullOrEmpty($Policy.Description))
				{
					Line 1 "Description`t: " $Policy.Description
				}
				Line 1 "Enabled`t`t: " $Policy.Enabled
				Line 1 "Type`t`t: " $Policy.Type
				Line 1 "Priority`t: " $Policy.Priority
			}
			If($HTML)
			{
				If($xDriveName -eq "localfarmgpo")
				{
					WriteHTMLLine 2 0 "$($Policy.PolicyName) (Site, $($xPolicyType))"
				}
				Else
				{
					WriteHTMLLine 2 0 "$($Policy.PolicyName) (AD, $($xPolicyType), GPO: $($ADGpoName))"
				}
				$rowdata = @()
				$columnHeaders = @("Description",($global:htmlsb),$Policy.Description,$htmlwhite)
				$rowdata += @(,('Enabled',($global:htmlsb),$Policy.Enabled.ToString(),$htmlwhite))
				$rowdata += @(,('Type',($global:htmlsb),$Policy.Type,$htmlwhite))
				$rowdata += @(,('Priority',($global:htmlsb),$Policy.Priority,$htmlwhite))

				$msg = ""
				$columnWidths = @("90","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "290"
			}

			Write-Verbose "$(Get-Date -Format G): `t`tRetrieving all filters"
			$filters = Get-CtxGroupPolicyFilter -PolicyName $Policy.PolicyName `
			-Type $xPolicyType `
			-DriveName $xDriveName -EA 0 `
			| Sort-Object FilterType, FilterName -Unique

			If($? -and $Null -ne $Filters)
			{
				If(![String]::IsNullOrEmpty($filters))
				{
					Write-Verbose "$(Get-Date -Format G): `t`tProcessing all filters"
					$txt = "Assigned to"
					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 $txt
					}
					If($Text)
					{
						Line 0 $txt
					}
					If($HTML)
					{
						WriteHTMLLine 3 0 $txt
					}
					
					If($MSWord -or $PDF)
					{
						[System.Collections.Hashtable[]] $FiltersWordTable = @();
					}
					If($HTML)
					{
						$rowdata = @()
					}
					
					ForEach($Filter in $Filters)
					{
						$tmp = ""
						Switch($filter.FilterType)
						{
							"AccessControl"  {$tmp = "Access Control"; Break}
							"BranchRepeater" {$tmp = "NetScaler SD-WAN"; Break}
							"ClientIP"       {$tmp = "Client IP Address"; Break}
							"ClientName"     {$tmp = "Client Name"; Break}
							"DesktopGroup"   {$tmp = "Delivery Group"; Break}
							"DesktopKind"    {$tmp = "Delivery GroupType"; Break}
							"DesktopTag"     {$tmp = "Tag"; Break}
							"OU"             {$tmp = "Organizational Unit (OU)"; Break}
							"User"           {$tmp = "User or group"; Break}
							Default {$tmp = "Policy Filter Type could not be determined: $($filter.FilterType)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$FiltersWordTable += @{
							Name = $filter.FilterName;
							Type= $tmp;
							Enabled = $filter.Enabled;
							Mode = $filter.Mode;
							Value = $filter.FilterValue;
							}
						}
						If($Text)
						{
							Line 2 "Name`t: " $filter.FilterName
							Line 2 "Type`t: " $tmp
							Line 2 "Enabled`t: " $filter.Enabled
							Line 2 "Mode`t: " $filter.Mode
							Line 2 "Value`t: " $filter.FilterValue
							Line 2 ""
						}
						If($HTML)
						{
							$rowdata += @(,(
							$filter.FilterName,$htmlwhite,
							$tmp,$htmlwhite,
							$filter.Enabled.ToString(),$htmlwhite,
							$filter.Mode,$htmlwhite,
							$filter.FilterValue,$htmlwhite))
						}
					}
					$tmp = $Null
					If($MSWord -or $PDF)
					{
						$Table = AddWordTable -Hashtable $FiltersWordTable `
						-Columns  Name,Type,Enabled,Mode,Value `
						-Headers  "Name","Type","Enabled","Mode","Value" `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Columns.Item(1).Width = 115;
						$Table.Columns.Item(2).Width = 125;
						$Table.Columns.Item(3).Width = 50;
						$Table.Columns.Item(4).Width = 40;
						$Table.Columns.Item(5).Width = 170;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
					}
					If($HTML)
					{
						$columnHeaders = @(
						'Name',($global:htmlsb),
						'Type',($global:htmlsb),
						'Enabled',($global:htmlsb),
						'Mode',($global:htmlsb),
						'Value',($global:htmlsb))

						$msg = ""
						$columnWidths = @("115","125","50","40","170")
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
					}
				}
				Else
				{
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 1 "Assigned to: None"
					}
					If($Text)
					{
						Line 1 "Assigned to`t`t: None"
					}
					If($HTML)
					{
						WriteHTMLLine 0 1 "Assigned to: None"
					}
				}
			}
			ElseIf($? -and $Null -eq $Filters)
			{
				$txt = "$($Policy.PolicyName) policy applies to all objects in the Site"
				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "Assigned to"
					WriteWordLine 0 1 $txt
				}
				If($Text)
				{
					Line 0 "Assigned to"
					Line 1 $txt
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "Assigned to"
					WriteHTMLLine 0 1 $txt
				}
			}
			ElseIf($? -and $Policy.PolicyName -eq "Unfiltered")
			{
				$txt = "Unfiltered policy applies to all objects in the Site"
				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "Assigned to"
					WriteWordLine 0 1 $txt
				}
				If($Text)
				{
					Line 0 "Assigned to"
					Line 1 $txt
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "Assigned to"
					WriteHTMLLine 0 1 $txt
				}
			}
			Else
			{
				$txt = "Unable to retrieve Filter settings"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 1 $txt
				}
				If($Text)
				{
					Line 1 $txt
				}
				If($HTML)
				{
					WriteHTMLLine 0 1 $txt
				}
			}
			
			Write-Verbose "$(Get-Date -Format G): `t`tRetrieving all policy settings"
			$Settings = Get-CtxGroupPolicyConfiguration -PolicyName $Policy.PolicyName `
			-Type $Policy.Type `
			-DriveName $xDriveName -EA 0
				
			If($? -and $Null -ne $Settings)
			{
				If($MSWord -or $PDF)
				{
					[System.Collections.Hashtable[]] $SettingsWordTable = @();
				}
				If($HTML)
				{
					$rowdata = @()
				}
				
				$First = $True
				ForEach($Setting in $Settings)
				{
					If($First)
					{
						$txt = "Policy settings"
						If($MSWord -or $PDF)
						{
							WriteWordLine 3 0 $txt
						}
						If($Text)
						{
							Line 1 $txt
						}
						If($HTML)
						{
							WriteHTMLLine 3 0 $txt
						}
					}
					$First = $False
					
					Write-Verbose "$(Get-Date -Format G): `t`tPolicy settings"
					Write-Verbose "$(Get-Date -Format G): `t`t`tConnector for Configuration Manager 2012"
					If((validStateProp $Setting AdvanceWarningFrequency State ) -and ($Setting.AdvanceWarningFrequency.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Advance warning frequency interval"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AdvanceWarningFrequency.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AdvanceWarningFrequency.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AdvanceWarningFrequency.Value
						}
					}
					If((validStateProp $Setting AdvanceWarningMessageBody State ) -and ($Setting.AdvanceWarningMessageBody.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Advance warning message box body text"
						$tmpArray = $Setting.AdvanceWarningMessageBody.Value.Split("`n")
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting "`t`t`t`t`t`t`t`t`t      " $tmp
								}
							}
							$txt = ""
						}
						$TmpArray = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting AdvanceWarningMessageTitle State ) -and ($Setting.AdvanceWarningMessageTitle.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Advance warning message box title"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AdvanceWarningMessageTitle.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AdvanceWarningMessageTitle.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AdvanceWarningMessageTitle.Value
						}
					}
					If((validStateProp $Setting AdvanceWarningPeriod State ) -and ($Setting.AdvanceWarningPeriod.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Advance warning time period"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AdvanceWarningPeriod.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AdvanceWarningPeriod.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AdvanceWarningPeriod.Value 
						}
					}
					If((validStateProp $Setting FinalForceLogoffMessageBody State ) -and ($Setting.FinalForceLogoffMessageBody.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Final force logoff message box body text"
						$tmpArray = $Setting.FinalForceLogoffMessageBody.Value.Split("`n")
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting "`t`t`t`t`t`t`t`t`t" $tmp
								}
							}
						}
						$TmpArray = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting FinalForceLogoffMessageTitle State ) -and ($Setting.FinalForceLogoffMessageTitle.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Final force logoff message box title"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FinalForceLogoffMessageTitle.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FinalForceLogoffMessageTitle.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.FinalForceLogoffMessageTitle.Value 
						}
					}
					If((validStateProp $Setting ForceLogoffGracePeriod State ) -and ($Setting.ForceLogoffGracePeriod.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Force logoff grace period"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ForceLogoffGracePeriod.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ForceLogoffGracePeriod.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ForceLogoffGracePeriod.Value 
						}
					}
					If((validStateProp $Setting ForceLogoffMessageBody State ) -and ($Setting.ForceLogoffMessageBody.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Force logoff message box body text"
						$tmpArray = $Setting.ForceLogoffMessageBody.Value.Split("`n")
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting "`t`t`t`t`t`t`t`t`t   " $tmp
								}
							}
							$txt = ""
						}
						$TmpArray = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting ForceLogoffMessageTitle State ) -and ($Setting.ForceLogoffMessageTitle.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Force logoff message box title"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ForceLogoffMessageTitle.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ForceLogoffMessageTitle.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ForceLogoffMessageTitle.Value 
						}
					}
					If((validStateProp $Setting ImageProviderIntegrationEnabled State ) -and ($Setting.ImageProviderIntegrationEnabled.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Image-managed mode"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ImageProviderIntegrationEnabled.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ImageProviderIntegrationEnabled.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ImageProviderIntegrationEnabled.State 
						}
					}
					If((validStateProp $Setting RebootMessageBody State ) -and ($Setting.RebootMessageBody.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Reboot message box body text"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RebootMessageBody.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RebootMessageBody.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.RebootMessageBody.Value 
						}
					}
					If((validStateProp $Setting AgentTaskInterval State ) -and ($Setting.AgentTaskInterval.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Regular time interval at which the agent task is to run"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AgentTaskInterval.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AgentTaskInterval.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AgentTaskInterval.Value 
						}
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`t`tICA"
					If((validStateProp $Setting ApplicationLaunchWaitTimeout State ) -and ($Setting.ApplicationLaunchWaitTimeout.State -ne "NotConfigured"))
					{
						$txt = "ICA\Application Launch Wait Timeout (milliseconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ApplicationLaunchWaitTimeout.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ApplicationLaunchWaitTimeout.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ApplicationLaunchWaitTimeout.Value
						}
					}
					If((validStateProp $Setting ClipboardRedirection State ) -and ($Setting.ClipboardRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Client clipboard redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClipboardRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClipboardRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClipboardRedirection.State 
						}
					}
					If((validStateProp $Setting ClientClipboardWriteAllowedFormats State ) -and ($Setting.ClientClipboardWriteAllowedFormats.State -ne "NotConfigured"))
					{
						$txt = "ICA\Client clipboard write allowed formats"
						If(validStateProp $Setting ClientClipboardWriteAllowedFormats Values )
						{
							$tmpArray = $Setting.ClientClipboardWriteAllowedFormats.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $TmpArray)
							{
								If($Null -eq $Thing)
								{
									$Thing = ''
								}
								$cnt++
								$tmp = "$($Thing) "
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = $txt;
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
								$txt = ""
							}
							$TmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No Client clipboard write allowed formats were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp
							}
						}
					}
					If((validStateProp $Setting ClipboardSelectionUpdateMode State ) -and ($Setting.ClipboardSelectionUpdateMode.State -ne "NotConfigured"))
					{
						$txt = "ICA\Clipboard selection update mode"
						$tmp = ""
						Switch ($Setting.ClipboardSelectionUpdateMode.Value)
						{
							"AllUpdatesAllowed"		{$tmp = "Selection changes are updated on both client and host"; Break}
							"AllUpdatesDenied"		{$tmp = "Select changes are not updated on neither client nor host"; Break}
							"UpdateToClientDenied"	{$tmp = "Host selection changes are not updated to client"; Break}
							"UpdateToHostDenied"	{$tmp = "Client selection changes are not updated to host"; Break}
							Default					{$tmp = "Clipboard selection update mode: $($Setting.ClipboardSelectionUpdateMode.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting DesktopLaunchForNonAdmins State ) -and ($Setting.DesktopLaunchForNonAdmins.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop launches"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DesktopLaunchForNonAdmins.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DesktopLaunchForNonAdmins.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.DesktopLaunchForNonAdmins.State 
						}
					}
					If((validStateProp $Setting HDXAdaptiveTransport State ) -and ($Setting.HDXAdaptiveTransport.State -ne "NotConfigured"))
					{
						$txt = "ICA\HDX Adaptive Transport"
						$tmp = ""
						Switch ($Setting.HDXAdaptiveTransport.Value)
						{
							"DiagnosticMode"	{$tmp = "Diagnostic mode"; Break}
							"Off"				{$tmp = "Off"; Break}
							"Preferred"			{$tmp = "Preferred"; Break}
							Default {$tmp = "HDX Adaptive Transport: $($Setting.HDXAdaptiveTransport.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting IcaListenerTimeout State ) -and ($Setting.IcaListenerTimeout.State -ne "NotConfigured"))
					{
						$txt = "ICA\ICA listener connection timeout (milliseconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IcaListenerTimeout.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaListenerTimeout.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.IcaListenerTimeout.Value 
						}
					}
					If((validStateProp $Setting IcaListenerPortNumber State ) -and ($Setting.IcaListenerPortNumber.State -ne "NotConfigured"))
					{
						$txt = "ICA\ICA listener port number"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IcaListenerPortNumber.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaListenerPortNumber.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.IcaListenerPortNumber.Value 
						}
					}
					If((validStateProp $Setting NonPublishedProgramLaunching State ) -and ($Setting.NonPublishedProgramLaunching.State -ne "NotConfigured"))
					{
						$txt = "ICA\Launching of non-published programs during client connection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.NonPublishedProgramLaunching.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.NonPublishedProgramLaunching.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.NonPublishedProgramLaunching.State
						}
					}
					If((validStateProp $Setting LogoffCheckerStartupDelay State ) -and ($Setting.LogoffCheckerStartupDelay.State -ne "NotConfigured"))
					{
						$txt = "ICA\Logoff Checker Startup Delay (seconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogoffCheckerStartupDelay.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogoffCheckerStartupDelay.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogoffCheckerStartupDelay.Value 
						}
					}
					If((validStateProp $Setting LossTolerantModeAvailable State ) -and ($Setting.LossTolerantModeAvailable.State -ne "NotConfigured"))
					{
						$txt = "ICA\Loss Tolerant Mode"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LossTolerantModeAvailable.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LossTolerantModeAvailable.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LossTolerantModeAvailable.State
						}
					}
					If((validStateProp $Setting PrimarySelectionUpdateMode State ) -and ($Setting.PrimarySelectionUpdateMode.State -ne "NotConfigured"))
					{
						$txt = "ICA\Primary selection update mode"
						$tmp = ""
						Switch ($Setting.PrimarySelectionUpdateMode.Value)
						{
							"AllUpdatesAllowed"		{$tmp = "Selection changes are updated on both client and host"; Break}
							"AllUpdatesDenied"		{$tmp = "Select changes are not updated on neither client nor host"; Break}
							"UpdateToClientDenied"	{$tmp = "Host selection changes are not updated to client"; Break}
							"UpdateToHostDenied"	{$tmp = "Client selection changes are not updated to host"; Break}
							Default					{$tmp = "Clipboard selection update mode: $($Setting.PrimarySelectionUpdateMode.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting RendezvousProtocol State ) -and ($Setting.RendezvousProtocol.State -ne "NotConfigured"))
					{
						$txt = "ICA\Rendezvous Protocol"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RendezvousProtocol.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RendezvousProtocol.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.RendezvousProtocol.State 
						}
					}
					If((validStateProp $Setting RestrictClientClipboardWrite State ) -and ($Setting.RestrictClientClipboardWrite.State -ne "NotConfigured"))
					{
						$txt = "ICA\Restrict client clipboard write"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RestrictClientClipboardWrite.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RestrictClientClipboardWrite.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.RestrictClientClipboardWrite.State
						}
					}
					If((validStateProp $Setting RestrictSessionClipboardWrite State ) -and ($Setting.RestrictSessionClipboardWrite.State -ne "NotConfigured"))
					{
						$txt = "ICA\Restrict session clipboard write"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RestrictSessionClipboardWrite.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RestrictSessionClipboardWrite.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.RestrictSessionClipboardWrite.State 
						}
					}
					If((validStateProp $Setting SessionClipboardWriteAllowedFormats State ) -and ($Setting.SessionClipboardWriteAllowedFormats.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session clipboard write allowed formats"
						If(validStateProp $Setting SessionClipboardWriteAllowedFormats Values )
						{
							$tmpArray = $Setting.SessionClipboardWriteAllowedFormats.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $TmpArray)
							{
								If($Null -eq $Thing)
								{
									$Thing = ''
								}
								$cnt++
								$tmp = "$($Thing) "
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = $txt;
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
								$txt = ""
							}
							$TmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No Session clipboard write allowed formats were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp
							}
						}
					}
					If((validStateProp $Setting VirtualChannelWhiteList State ) -and ($Setting.VirtualChannelWhiteList.State -ne "NotConfigured"))
					{
						$txt = "ICA\Virtual channel white list"
						If(validStateProp $Setting VirtualChannelWhiteList Values )
						{
							$tmpArray = $Setting.VirtualChannelWhiteList.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $TmpArray)
							{
								If($Null -eq $Thing)
								{
									$Thing = ''
								}
								$cnt++
								$tmp = "$($Thing) "
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = $txt;
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
								$txt = ""
							}
							$TmpArray = $Null
							$tmp = $Null
						}
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Audio"
					If((validStateProp $Setting AllowRtpAudio State ) -and ($Setting.AllowRtpAudio.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Audio over UDP real-time transport"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AllowRtpAudio.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowRtpAudio.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AllowRtpAudio.State 
						}
					}
					If((validStateProp $Setting AudioPlugNPlay State ) -and ($Setting.AudioPlugNPlay.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Audio Plug N Play"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AudioPlugNPlay.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AudioPlugNPlay.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AudioPlugNPlay.State 
						}
					}
					If((validStateProp $Setting AudioQuality State ) -and ($Setting.AudioQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Audio quality"
						$tmp = ""
						Switch ($Setting.AudioQuality.Value)
						{
							"Low"		{$tmp = "Low - for low-speed connections"; Break}
							"Medium"	{$tmp = "Medium - optimized for speech"; Break}
							"High"		{$tmp = "High - high definition audio"; Break}
							Default		{$tmp = "Audio quality could not be determined: $($Setting.AudioQuality.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting ClientAudioRedirection State ) -and ($Setting.ClientAudioRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Client audio redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientAudioRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientAudioRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientAudioRedirection.State 
						}
					}
					If((validStateProp $Setting MicrophoneRedirection State ) -and ($Setting.MicrophoneRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Client microphone redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MicrophoneRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MicrophoneRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.MicrophoneRedirection.State 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Auto Client Reconnect"
					If((validStateProp $Setting AutoClientReconnect State ) -and ($Setting.AutoClientReconnect.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Auto client reconnect"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AutoClientReconnect.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AutoClientReconnect.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AutoClientReconnect.State 
						}
					}
					If((validStateProp $Setting AutoClientReconnectAuthenticationRequired  State ) -and ($Setting.AutoClientReconnectAuthenticationRequired.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Auto client reconnect authentication"
						$tmp = ""
						Switch ($Setting.AutoClientReconnectAuthenticationRequired.Value)
						{
							"DoNotRequireAuthentication" {$tmp = "Do not require authentication"; Break}
							"RequireAuthentication"      {$tmp = "Require authentication"; Break}
							Default {$tmp = "Auto client reconnect authentication could not be determined: $($Setting.AutoClientReconnectAuthenticationRequired.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting AutoClientReconnectLogging State ) -and ($Setting.AutoClientReconnectLogging.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Auto client reconnect logging"
						$tmp = ""
						Switch ($Setting.AutoClientReconnectLogging.Value)
						{
							"DoNotLogAutoReconnectEvents" {$tmp = "Do Not Log auto-reconnect events"; Break}
							"LogAutoReconnectEvents"      {$tmp = "Log auto-reconnect events"; Break}
							Default {$tmp = "Auto client reconnect logging could not be determined: $($Setting.AutoClientReconnectLogging.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting ACRTimeout State ) -and ($Setting.ACRTimeout.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Auto client reconnect timeout (seconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ACRTimeout.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ACRTimeout.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ACRTimeout.Value 
						}
					}
					If((validStateProp $Setting ReconnectionUiTransparencyLevel State ) -and ($Setting.ReconnectionUiTransparencyLevel.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Reconnection UI transparency level (%)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ReconnectionUiTransparencyLevel.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ReconnectionUiTransparencyLevel.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ReconnectionUiTransparencyLevel.Value 
						}
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Bandwidth"
					If((validStateProp $Setting AudioBandwidthLimit State ) -and ($Setting.AudioBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Audio redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AudioBandwidthLimit.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AudioBandwidthLimit.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AudioBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting AudioBandwidthPercent State ) -and ($Setting.AudioBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Audio redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AudioBandwidthPercent.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AudioBandwidthPercent.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AudioBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting USBBandwidthLimit State ) -and ($Setting.USBBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Client USB device redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.USBBandwidthLimit.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.USBBandwidthLimit.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.USBBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting USBBandwidthPercent State ) -and ($Setting.USBBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Client USB device redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.USBBandwidthPercent.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.USBBandwidthPercent.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.USBBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting ClipboardBandwidthLimit State ) -and ($Setting.ClipboardBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Clipboard redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClipboardBandwidthLimit.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClipboardBandwidthLimit.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClipboardBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting ClipboardBandwidthPercent State ) -and ($Setting.ClipboardBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Clipboard redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClipboardBandwidthPercent.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClipboardBandwidthPercent.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClipboardBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting ComPortBandwidthLimit State ) -and ($Setting.ComPortBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\COM port redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ComPortBandwidthLimit.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ComPortBandwidthLimit.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ComPortBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting ComPortBandwidthPercent State ) -and ($Setting.ComPortBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\COM port redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ComPortBandwidthPercent.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ComPortBandwidthPercent.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ComPortBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting FileRedirectionBandwidthLimit State ) -and ($Setting.FileRedirectionBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\File redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FileRedirectionBandwidthLimit.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FileRedirectionBandwidthLimit.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.FileRedirectionBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting FileRedirectionBandwidthPercent State ) -and ($Setting.FileRedirectionBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\File redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FileRedirectionBandwidthPercent.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FileRedirectionBandwidthPercent.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.FileRedirectionBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting HDXMultimediaBandwidthLimit State ) -and ($Setting.HDXMultimediaBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\HDX MediaStream Multimedia Acceleration bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.HDXMultimediaBandwidthLimit.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HDXMultimediaBandwidthLimit.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.HDXMultimediaBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting HDXMultimediaBandwidthPercent State ) -and ($Setting.HDXMultimediaBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\HDX MediaStream Multimedia Acceleration bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.HDXMultimediaBandwidthPercent.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HDXMultimediaBandwidthPercent.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.HDXMultimediaBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting LptBandwidthLimit State ) -and ($Setting.LptBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\LPT port redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LptBandwidthLimit.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LptBandwidthLimit.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LptBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting LptBandwidthLimitPercent State ) -and ($Setting.LptBandwidthLimitPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\LPT port redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LptBandwidthLimitPercent.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LptBandwidthLimitPercent.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LptBandwidthLimitPercent.Value 
						}
					}
					If((validStateProp $Setting OverallBandwidthLimit State ) -and ($Setting.OverallBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Overall session bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.OverallBandwidthLimit.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OverallBandwidthLimit.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.OverallBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting PrinterBandwidthLimit State ) -and ($Setting.PrinterBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Printer redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.PrinterBandwidthLimit.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PrinterBandwidthLimit.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.PrinterBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting PrinterBandwidthPercent State ) -and ($Setting.PrinterBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Printer redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.PrinterBandwidthPercent.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PrinterBandwidthPercent.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.PrinterBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting TwainBandwidthLimit State ) -and ($Setting.TwainBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\TWAIN device redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.TwainBandwidthLimit.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TwainBandwidthLimit.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.TwainBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting TwainBandwidthPercent State ) -and ($Setting.TwainBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\TWAIN device redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.TwainBandwidthPercent.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TwainBandwidthPercent.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.TwainBandwidthPercent.Value 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Client Sensors\Location"
					If((validStateProp $Setting AllowLocationServices State ) -and ($Setting.AllowLocationServices.State -ne "NotConfigured"))
					{
						$txt = "ICA\Client Sensors\Location\Allow applications to use the physical location of the client device"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AllowLocationServices.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowLocationServices.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AllowLocationServices.State 
						}
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Desktop UI"
					If((validStateProp $Setting GraphicsQuality State ) -and ($Setting.GraphicsQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Desktop Composition graphics quality"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.GraphicsQuality.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.GraphicsQuality.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.GraphicsQuality.Value 
						}
					}
					If((validStateProp $Setting AeroRedirection State ) -and ($Setting.AeroRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Desktop Composition Redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AeroRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AeroRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AeroRedirection.State 
						}
					}
					If((validStateProp $Setting DesktopWallpaper State ) -and ($Setting.DesktopWallpaper.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Desktop wallpaper"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DesktopWallpaper.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DesktopWallpaper.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.DesktopWallpaper.State 
						}
					}
					If((validStateProp $Setting MenuAnimation State ) -and ($Setting.MenuAnimation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Menu animation"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MenuAnimation.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MenuAnimation.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.MenuAnimation.State 
						}
					}
					If((validStateProp $Setting WindowContentsVisibleWhileDragging State ) -and ($Setting.WindowContentsVisibleWhileDragging.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\View window contents while dragging"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WindowContentsVisibleWhileDragging.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WindowContentsVisibleWhileDragging.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.WindowContentsVisibleWhileDragging.State 
						}
					}
			
					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\End User Monitoring"
					If((validStateProp $Setting IcaRoundTripCalculation State ) -and ($Setting.IcaRoundTripCalculation.State -ne "NotConfigured"))
					{
						$txt = "ICA\End User Monitoring\ICA round trip calculation"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IcaRoundTripCalculation.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaRoundTripCalculation.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.IcaRoundTripCalculation.State 
						}
					}
					If((validStateProp $Setting IcaRoundTripCalculationInterval State ) -and ($Setting.IcaRoundTripCalculationInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\End User Monitoring\ICA round trip calculation interval (seconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IcaRoundTripCalculationInterval.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaRoundTripCalculationInterval.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.IcaRoundTripCalculationInterval.Value 
						}	
					}
					If((validStateProp $Setting IcaRoundTripCalculationWhenIdle State ) -and ($Setting.IcaRoundTripCalculationWhenIdle.State -ne "NotConfigured"))
					{
						$txt = "ICA\End User Monitoring\ICA round trip calculations for idle connections"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IcaRoundTripCalculationWhenIdle.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaRoundTripCalculationWhenIdle.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.IcaRoundTripCalculationWhenIdle.State 
						}	
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Enhanced Desktop Experience"
					If((validStateProp $Setting EnhancedDesktopExperience State ) -and ($Setting.EnhancedDesktopExperience.State -ne "NotConfigured"))
					{
						$txt = "ICA\Enhanced Desktop Experience\Enhanced Desktop Experience"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.EnhancedDesktopExperience.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnhancedDesktopExperience.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.EnhancedDesktopExperience.State 
						}	
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\File Redirection"
					If((validStateProp $Setting AllowFileTransfer State ) -and ($Setting.AllowFileTransfer.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Allow file transfer between desktop and client"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AllowFileTransfer.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowFileTransfer.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AllowFileTransfer.State 
						}
					}
					If((validStateProp $Setting AutoConnectDrives State ) -and ($Setting.AutoConnectDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Auto connect client drives"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AutoConnectDrives.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AutoConnectDrives.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AutoConnectDrives.State 
						}
					}
					If((validStateProp $Setting ClientDriveRedirection State ) -and ($Setting.ClientDriveRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client drive redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientDriveRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientDriveRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientDriveRedirection.State 
						}
					}
					If((validStateProp $Setting ClientFixedDrives State ) -and ($Setting.ClientFixedDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client fixed drives"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientFixedDrives.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientFixedDrives.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientFixedDrives.State 
						}
					}
					If((validStateProp $Setting ClientFloppyDrives State ) -and ($Setting.ClientFloppyDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client floppy drives"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientFloppyDrives.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientFloppyDrives.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientFloppyDrives.State 
						}
					}
					If((validStateProp $Setting ClientNetworkDrives State ) -and ($Setting.ClientNetworkDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client network drives"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientNetworkDrives.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientNetworkDrives.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientNetworkDrives.State 
						}
					}
					If((validStateProp $Setting ClientOpticalDrives State ) -and ($Setting.ClientOpticalDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client optical drives"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientOpticalDrives.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientOpticalDrives.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientOpticalDrives.State 
						}
					}
					If((validStateProp $Setting ClientRemoveableDrives State ) -and ($Setting.ClientRemoveableDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client removable drives"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientRemoveableDrives.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientRemoveableDrives.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientRemoveableDrives.State 
						}
					}
					If((validStateProp $Setting AllowFileDownload State ) -and ($Setting.AllowFileDownload.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Download file from desktop"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AllowFileDownload.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowFileDownload.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AllowFileDownload.State 
						}
					}
					If((validStateProp $Setting HostToClientRedirection State ) -and ($Setting.HostToClientRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Host to client redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.HostToClientRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HostToClientRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.HostToClientRedirection.State 
						}
					}
					If((validStateProp $Setting ClientDriveLetterPreservation State ) -and ($Setting.ClientDriveLetterPreservation.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Preserve client drive letters"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientDriveLetterPreservation.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientDriveLetterPreservation.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientDriveLetterPreservation.State 
						}
					}
					If((validStateProp $Setting ReadOnlyMappedDrive State ) -and ($Setting.ReadOnlyMappedDrive.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Read-only client drive access"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ReadOnlyMappedDrive.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ReadOnlyMappedDrive.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ReadOnlyMappedDrive.State 
						}
					}
					If((validStateProp $Setting SpecialFolderRedirection State ) -and ($Setting.SpecialFolderRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Special folder redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SpecialFolderRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SpecialFolderRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.SpecialFolderRedirection.State 
						}
					}
					If((validStateProp $Setting AllowFileUpload State ) -and ($Setting.AllowFileUpload.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Upload file to desktop"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AllowFileUpload.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowFileUpload.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AllowFileUpload.State 
						}
					}
					If((validStateProp $Setting AsynchronousWrites State ) -and ($Setting.AsynchronousWrites.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Use asynchronous writes"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AsynchronousWrites.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AsynchronousWrites.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AsynchronousWrites.State 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Graphics"
					If((validStateProp $Setting AllowVisuallyLosslessCompression State ) -and ($Setting.AllowVisuallyLosslessCompression.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Allow visually lossless compression"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AllowVisuallyLosslessCompression.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowVisuallyLosslessCompression.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AllowVisuallyLosslessCompression.State 
						}
					}
					If((validStateProp $Setting DisplayMemoryLimit State ) -and ($Setting.DisplayMemoryLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Display memory limit (KB)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DisplayMemoryLimit.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DisplayMemoryLimit.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.DisplayMemoryLimit.Value 
						}	
					}
					If((validStateProp $Setting DynamicPreview State ) -and ($Setting.DynamicPreview.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Dynamic windows preview"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DynamicPreview.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DynamicPreview.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.DynamicPreview.State 
						}	
					}
					If((validStateProp $Setting MaximumColorDepth State ) -and ($Setting.MaximumColorDepth.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Maximum allowed color depth"
						$tmp = ""
						Switch ($Setting.MaximumColorDepth.Value)
						{
							"BitsPerPixel24"	{$tmp = "24 Bits Per Pixel"; Break}
							"BitsPerPixel32"	{$tmp = "32 Bits Per Pixel"; Break}
							"BitsPerPixel16"	{$tmp = "16 Bits Per Pixel"; Break}
							"BitsPerPixel15"	{$tmp = "15 Bits Per Pixel"; Break}
							"BitsPerPixel8"		{$tmp = "8 Bits Per Pixel"; Break}
							Default				{$tmp = "Maximum allowed color depth could not be determined: $($Setting.MaximumColorDepth.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}	
						$tmp = $Null
					}
					If((validStateProp $Setting OptimizeFor3dWorkload State ) -and ($Setting.OptimizeFor3dWorkload.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Optimize for 3D graphics workload"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.OptimizeFor3dWorkload.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OptimizeFor3dWorkload.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.OptimizeFor3dWorkload.State 
						}
					}
					If((validStateProp $Setting UseHardwareEncodingForVideoCodec State ) -and ($Setting.UseHardwareEncodingForVideoCodec.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Use hardware encoding for video codec"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UseHardwareEncodingForVideoCodec.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UseHardwareEncodingForVideoCodec.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UseHardwareEncodingForVideoCodec.State 
						}	
					}
					If((validStateProp $Setting UseVideoCodecForCompression State ) -and ($Setting.UseVideoCodecForCompression.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Use video codec for compression"
						$tmp = ""
						Switch ($Setting.UseVideoCodecForCompression.Value)
						{
							"UseVideoCodecIfAvailable"	{$tmp = "For the entire screen"; Break}
							"DoNotUseVideoCodec"		{$tmp = "Do not use video codec"; Break}
							"UseVideoCodecIfPreferred"	{$tmp = "Use when preferred"; Break}
							"ActivelyChangingRegions"	{$tmp = "For actively changing regions"; Break}
							Default {$tmp = "Use video codec for compression could not be determined: $($Setting.UseVideoCodecForCompression.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}	
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Keep Alive"
					If((validStateProp $Setting IcaKeepAliveTimeout State ) -and ($Setting.IcaKeepAliveTimeout.State -ne "NotConfigured"))
					{
						$txt = "ICA\Keep Alive\ICA keep alive timeout (seconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IcaKeepAliveTimeout.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaKeepAliveTimeout.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.IcaKeepAliveTimeout.Value 
						}
					}
					If((validStateProp $Setting IcaKeepAlives State ) -and ($Setting.IcaKeepAlives.State -ne "NotConfigured"))
					{
						$txt = "ICA\Keep Alive\ICA keep alives"
						$tmp = ""
						Switch ($Setting.IcaKeepAlives.Value)
						{
							"DoNotSendKeepAlives" {$tmp = "Do not send ICA keep alive messages"; Break}
							"SendKeepAlives"      {$tmp = "Send ICA keep alive messages"; Break}
							Default {$tmp = "ICA keep alives could not be determined: $($Setting.IcaKeepAlives.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Keyboard and IME"
					If((validStateProp $Setting ClientKeyboardLayoutSyncAndIME State ) -and ($Setting.ClientKeyboardLayoutSyncAndIME.State -ne "NotConfigured"))
					{
						$txt = "ICA\Keyboard and IME\Client keyboard layout synchronization and IME improvement"
						$tmp = ""
						Switch ($Setting.ClientKeyboardLayoutSyncAndIME.Value)
						{
							"ClientKeyboardLayoutSync"			{$tmp = "Support dynamic client keyboard layout synchronization"; Break}
							"ClientKeyboardLayoutSyncAndIME"	{$tmp = "Support dynamic client keyboard layout synchronization and IME improvement"; Break}
							Default 							{$tmp = "Client keyboard layout synchronization and IME improvement: $($Setting.ClientKeyboardLayoutSyncAndIME.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting EnableUnicodeKeyboardLayoutMapping State ) -and ($Setting.EnableUnicodeKeyboardLayoutMapping.State -ne "NotConfigured"))
					{
						$txt = "ICA\Keyboard and IME\Enable Unicode keyboard layout mapping"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.EnableUnicodeKeyboardLayoutMapping.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnableUnicodeKeyboardLayoutMapping.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.EnableUnicodeKeyboardLayoutMapping.State 
						}
					}
					If((validStateProp $Setting HideKeyboardLayoutSwitchPopupMessageBox State ) -and ($Setting.HideKeyboardLayoutSwitchPopupMessageBox.State -ne "NotConfigured"))
					{
						$txt = "ICA\Keyboard and IME\Hide keyboard layout Switch pop-up message box"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.HideKeyboardLayoutSwitchPopupMessageBox.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HideKeyboardLayoutSwitchPopupMessageBox.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.HideKeyboardLayoutSwitchPopupMessageBox.State
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Local App Access"
					If((validStateProp $Setting AllowLocalAppAccess State ) -and ($Setting.AllowLocalAppAccess.State -ne "NotConfigured"))
					{
						$txt = "ICA\Local App Access\Allow local app access"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AllowLocalAppAccess.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowLocalAppAccess.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AllowLocalAppAccess.State 
						}
					}
					If((validStateProp $Setting URLRedirectionBlackList State ) -and ($Setting.URLRedirectionBlackList.State -ne "NotConfigured"))
					{
						$txt = "ICA\Local App Access\URL redirection blacklist"
						If(validStateProp $Setting URLRedirectionBlackList Values )
						{
							$tmpArray = $Setting.URLRedirectionBlackList.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $TmpArray)
							{
								If($Null -eq $Thing)
								{
									$Thing = ''
								}
								$cnt++
								$tmp = "$($Thing) "
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = $txt;
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$TmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No URL redirection blacklist were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp
							}
						}
					}
					If((validStateProp $Setting URLRedirectionWhiteList State ) -and ($Setting.URLRedirectionWhiteList.State -ne "NotConfigured"))
					{
						$txt = "ICA\Local App Access\URL redirection white list"
						If(validStateProp $Setting URLRedirectionWhiteList Values )
						{
							$tmpArray = $Setting.URLRedirectionWhiteList.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $TmpArray)
							{
								If($Null -eq $Thing)
								{
									$Thing = ''
								}
								$cnt++
								$tmp = "$($Thing) "
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = $txt;
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$TmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No URL redirection white list were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp
							}
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Mobile Experience"
					If((validStateProp $Setting AutoKeyboardPopUp State ) -and ($Setting.AutoKeyboardPopUp.State -ne "NotConfigured"))
					{
						$txt = "ICA\Mobile Experience\Automatic keyboard display"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AutoKeyboardPopUp.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AutoKeyboardPopUp.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AutoKeyboardPopUp.State 
						}
					}
					If((validStateProp $Setting MobileDesktop State ) -and ($Setting.MobileDesktop.State -ne "NotConfigured"))
					{
						$txt = "ICA\Mobile Experience\Launch touch-optimized desktop"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MobileDesktop.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MobileDesktop.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.MobileDesktop.State 
						}
					}
					If((validStateProp $Setting ComboboxRemoting State ) -and ($Setting.ComboboxRemoting.State -ne "NotConfigured"))
					{
						$txt = "ICA\Mobile Experience\Remote the combo box"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ComboboxRemoting.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ComboboxRemoting.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ComboboxRemoting.State 
						}
					}
					If((validStateProp $Setting TabletModeToggle State ) -and ($Setting.TabletModeToggle.State -ne "NotConfigured"))
					{
						$txt = "ICA\Mobile Experience\Tablet Mode Toggle"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.TabletModeToggle.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TabletModeToggle.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.TabletModeToggle.State 
						}
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Multimedia"
					If((validStateProp $Setting WebBrowserRedirection State ) -and ($Setting.WebBrowserRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Browser Content Redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WebBrowserRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WebBrowserRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.WebBrowserRedirection.State
						}
					}
					If((validStateProp $Setting BRUrlWhitelist State ) -and ($Setting.BRUrlWhitelist.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Browser Content Redirection ACL Configuration"
						$array = $Setting.BRUrlWhitelist.Values
						$tmp = $array[0]
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}

						$txt = ""
						$cnt = -1
						ForEach($element in $array)
						{
							$cnt++
							
							If($cnt -ne 0)
							{
								$tmp = "$($element) "
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$array = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting WebBrowserRedirectionAuthenticationSites State ) -and ($Setting.WebBrowserRedirectionAuthenticationSites.State -ne "NotConfigured"))
					{
						If( $Setting.WebBrowserRedirectionAuthenticationSites.PSObject.Properties[ 'Values' ] )
						{
							$txt = "ICA\Multimedia\Browser Content Redirection Authentication Sites"
							$array = $Setting.WebBrowserRedirectionAuthenticationSites.Values
							$tmp = $array[0]
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp 
							}

							$txt = ""
							$cnt = -1
							ForEach($element in $array)
							{
								$cnt++
								
								If($cnt -ne 0)
								{
									$tmp = "$($element) "
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$array = $Null
							$tmp = $Null
						}
					}
					If((validStateProp $Setting WebBrowserRedirectionBlacklist State ) -and ($Setting.WebBrowserRedirectionBlacklist.State -ne "NotConfigured"))
					{
						If( $Setting.WebBrowserRedirectionBlacklist.PSObject.Properties[ 'Values' ] )
						{
							$txt = "ICA\Multimedia\Browser Content Redirection Blacklist Configuration"
							$array = $Setting.WebBrowserRedirectionBlacklist.Values
							$tmp = $array[0]
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp 
							}

							$txt = ""
							$cnt = -1
							ForEach($element in $array)
							{
								$cnt++
								
								If($cnt -ne 0)
								{
									$tmp = "$($element) "
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$array = $Null
							$tmp = $Null
						}
					}
					If((validStateProp $Setting WebBrowserRedirectionProxy State ) -and ($Setting.WebBrowserRedirectionProxy.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Browser Content Redirection Proxy Configuration"
						If($Setting.WebBrowserRedirectionProxy.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.WebBrowserRedirectionProxy.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.WebBrowserRedirectionProxy.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.WebBrowserRedirectionProxy.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.WebBrowserRedirectionProxy.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.WebBrowserRedirectionProxy.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.WebBrowserRedirectionProxy.State 
							}
						}
					}
					If((validStateProp $Setting HTML5VideoRedirection State ) -and ($Setting.HTML5VideoRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\HTML5 video redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.HTML5VideoRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HTML5VideoRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.HTML5VideoRedirection.State
						}
					}
					If((validStateProp $Setting VideoQuality State ) -and ($Setting.VideoQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Limit video quality"
						$tmp = ""
						Switch ($Setting.VideoQuality.Value)
						{
							"P1080"			{$tmp = "Maximum Video Quality 1080p/8.5mbps"; Break}
							"P720"			{$tmp = "Maximum Video Quality 720p/4.0mbps"; Break}
							"P480"			{$tmp = "Maximum Video Quality 480p/720kbps"; Break}
							"P380"			{$tmp = "Maximum Video Quality 380p/400kbps"; Break}
							"P240"			{$tmp = "Maximum Video Quality 240p/200kbps"; Break}
							"Unconfigured"	{$tmp = "Not Configured"; Break}
							Default			{$tmp = "Limit video quality could not be determined: $($Setting.VideoQuality.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting MaxSpeexQuality State ) -and ($Setting.MaxSpeexQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Max Speex quality"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MaxSpeexQuality.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MaxSpeexQuality.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.MaxSpeexQuality.Value 
						}
					}
					If((validStateProp $Setting MSTeamsRedirection State ) -and ($Setting.MSTeamsRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Microsoft Teams redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MSTeamsRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MSTeamsRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.MSTeamsRedirection.State 
						}
					}
					If((validStateProp $Setting MultimediaConferencing State ) -and ($Setting.MultimediaConferencing.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Multimedia conferencing"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MultimediaConferencing.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaConferencing.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaConferencing.State 
						}
					}
					If((validStateProp $Setting MultimediaOptimization State ) -and ($Setting.MultimediaOptimization.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Optimization for Windows Media multimedia redirection over WAN"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MultimediaOptimization.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaOptimization.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaOptimization.State 
						}
					}
					If((validStateProp $Setting UseGPUForMultimediaOptimization State ) -and ($Setting.UseGPUForMultimediaOptimization.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Use GPU for optimizing Windows Media multimedia redirection over WAN"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UseGPUForMultimediaOptimization.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UseGPUForMultimediaOptimization.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UseGPUForMultimediaOptimization.State 
						}
					}
					If((validStateProp $Setting MultimediaAccelerationEnableCSF State ) -and ($Setting.MultimediaAccelerationEnableCSF.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Windows Media client-side content fetching"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MultimediaAccelerationEnableCSF.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaAccelerationEnableCSF.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaAccelerationEnableCSF.State 
						}
					}
					If((validStateProp $Setting VideoLoadManagement State ) -and ($Setting.VideoLoadManagement.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Windows media fallback prevention"
						$tmp = ""
						Switch ($Setting.VideoLoadManagement.Value)
						{
							"SFSR"			{$tmp = "Play all content"; Break}
							"SFCR"			{$tmp = "Play all content only on client"; Break}
							"CFCR"			{$tmp = "Play only client-accessible content on client"; Break}
							"UnConfigured"	{$tmp = "Not Configured"; Break}
							Default			{$tmp = "Windows media fallback prevention could not be determined: $($Setting.VideoLoadManagement.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting MultimediaAcceleration State ) -and ($Setting.MultimediaAcceleration.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Windows Media redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MultimediaAcceleration.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaAcceleration.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaAcceleration.State 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Multi-Stream Connections"
					If((validStateProp $Setting UDPAudioOnServer State ) -and ($Setting.UDPAudioOnServer.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multi-Stream Connections\Audio over UDP"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UDPAudioOnServer.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UDPAudioOnServer.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UDPAudioOnServer.State
						}
					}
					If((validStateProp $Setting RtpAudioPortRange State ) -and ($Setting.RtpAudioPortRange.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multi-Stream Connections\Audio UDP port range"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RtpAudioPortRange.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RtpAudioPortRange.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.RtpAudioPortRange.Value 
						}
					}
					If((validStateProp $Setting MultiStreamPolicy State ) -and ($Setting.MultiStreamPolicy.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multi-Stream Connections\Multi-Stream computer setting"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MultiStreamPolicy.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultiStreamPolicy.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.MultiStreamPolicy.State 
						}
					}
					If((validStateProp $Setting MultiStream State ) -and ($Setting.MultiStream.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multi-Stream Connections\Multi-Stream user setting"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MultiStream.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultiStream.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.MultiStream.State 
						}
					}
					If((validStateProp $Setting MultiStreamAssignment State ) -and ($Setting.MultiStreamAssignment.State -ne "NotConfigured"))
					{
						<#
						Value=		Virtual Channels							Stream Number
						CTXCAM,0; 	Audio                                       0
						CTXEUEM,1;	End User Experience Monitoring              1
						CTXCTL,1;	ICA Control                                 1
						CTXIME,1;	Input Method Editor                         1
						CTXLIC,1;	License Management                          1
						CTXMTOP,1;	Microsoft Teams/WebRTC Redirection          1
						CTXMOB,1;	Mobile Receiver                             1
						CTXMTCH,1;	MultiTouch                                  1
						CTXTWI,1;	Seamless (Tranparent Window Integration)    1
						CTXSENS,1;	Sensor and Location                         1
						CTXSCRD,1;	Smart Card                                  1
						CTXTW,1;	Thinwire Graphics                           1
						CTXDND,1;	CTXDND                                      1
						CTXNSAP,2;	App Flow                                    2
						CTXCSB,2;	Browser Content Redirection                 2
						CTXCDM,2;	Client Drive Mapping                        2
						CTXCLIP,2;	Clipboard                                   2
						CTXFILE,2;	File Transfer (HTML5 Receiver)              2
						CTXGDT,2;	Generic Data Transfer                       2
						CTXPFWD,2;	Port Forwarding                             2
						CTXMM,2;	Remote Audio and Video Extensions (RAVE)    2
						CTXTUI,2;	Transparent UI Integration/Logon Status     2
						CTXTWN,2;	TWAIN Redirection                           2
						CTXGUSB,2;	USB                                         2
						CTXZLFK,2;	Zero Latency Font and Keyboard              2
						CTXZLC,2;	Zero Latency Data Channel                   2
						CTXCCM,3;	Client COM Port Mapping                     3
						CTXCPM,3;	Client Printer Mapping                      3
						CTXCOM1,3;	Legacy Client Printer Mapping (COM1)        3
						CTXCOM2,3;	Legacy Client Printer Mapping (COM2)        3
						CTXLPT1,3;	Legacy Client Printer Mapping (LPT1)        3
						CTXLPT2,3;	Legacy Client Printer Mapping (LPT2)        3
						#>
						$txt = "ICA\Multi-Stream Connections\Multi-Stream virtual channel stream assignment"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = "";
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							"",$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt ""
						}
						$Values = $Setting.MultiStreamAssignment.Value.Split(';')
						$tmp = ""
						ForEach($Value in $Values)
						{
							If($Value -eq "")
							{
								Continue
							}
							
							$tmparray = $Value.Split(",")
							$ChannelName = $tmparray[0]
							
							If($ChannelName -eq "")
							{
								Continue
							}

							$StreamNumber = $tmparray[1]
							
							Switch ($ChannelName)
							{
								"CTXCAM"	
									{
										$tmp = "Virtual Channel: Audio - Stream Number: $StreamNumber"; Break
									}
								"CTXEUEM"	
									{
										$tmp = "Virtual Channel: End User Experience Monitoring - Stream Number: $StreamNumber"; Break
									}
								"CTXCTL"	
									{
										$tmp = "Virtual Channel: ICA Control - Stream Number: $StreamNumber"; Break
									}
								"CTXIME"	
									{
										$tmp = "Virtual Channel: Input Method Editor - Stream Number: $StreamNumber"; Break
									}
								"CTXLIC"	
									{
										$tmp = "Virtual Channel: License Management - Stream Number: $StreamNumber"; Break
									}
								"CTXMTOP"	
									{
										$tmp = "Virtual Channel: Microsoft Teams/WebRTC Redirection - Stream Number: $StreamNumber"; Break
									}
								"CTXMOB"	
									{
										$tmp = "Virtual Channel: Mobile Receiver - Stream Number: $StreamNumber"; Break
									}
								"CTXMTCH"	
									{
										$tmp = "Virtual Channel: MultiTouch - Stream Number: $StreamNumber"; Break
									}
								"CTXTWI"	
									{
										$tmp = "Virtual Channel: Seamless (Tranparent Window Integration) - Stream Number: $StreamNumber"; Break
									}
								"CTXSENS"	
									{
										$tmp = "Virtual Channel: Sensor and Location - Stream Number: $StreamNumber"; Break
									}
								"CTXSCRD"	
									{
										$tmp = "Virtual Channel: Smart Card - Stream Number: $StreamNumber"; Break
									}
								"CTXTW"		
									{
										$tmp = "Virtual Channel: Thinwire Graphics - Stream Number: $StreamNumber"; Break
									}
								"CTXDND"	
									{
										$tmp = "Virtual Channel: CTXDND - Stream Number: $StreamNumber"; Break
									}
								"CTXNSAP"	
									{
										$tmp = "Virtual Channel: App Flow - Stream Number: $StreamNumber"; Break
									}
								"CTXCSB"	
									{
										$tmp = "Virtual Channel: Browser Content Redirection - Stream Number: $StreamNumber"; Break
									}
								"CTXCDM"	
									{
										$tmp = "Virtual Channel: Client Drive Mapping - Stream Number: $StreamNumber"; Break
									}
								"CTXCLIP"	
									{
										$tmp = "Virtual Channel: Clipboard - Stream Number: $StreamNumber"; Break
									}
								"CTXFILE"	
									{
										$tmp = "Virtual Channel: File Transfer (HTML5 Receiver) - Stream Number: $StreamNumber"; Break
									}
								"CTXGDT"	
									{
										$tmp = "Virtual Channel: Generic Data Transfer - Stream Number: $StreamNumber"; Break
									}
								"CTXPFWD"	
									{
										$tmp = "Virtual Channel: Port Forwarding - Stream Number: $StreamNumber"; Break
									}
								"CTXMM"		
									{
										$tmp = "Virtual Channel: Remote Audio and Video Extensions (RAVE) - Stream Number: $StreamNumber"; Break
									}
								"CTXTUI"	
									{
										$tmp = "Virtual Channel: Transparent UI Integration/Logon Status - Stream Number: $StreamNumber"; Break
									}
								"CTXTWN"	
									{
										$tmp = "Virtual Channel: TWAIN Redirection - Stream Number: $StreamNumber"; Break
									}
								"CTXGUSB"	
									{
										$tmp = "Virtual Channel: USB - Stream Number: $StreamNumber"; Break
									}
								"CTXZLFK"	
									{
										$tmp = "Virtual Channel: Zero Latency Font and Keyboard - Stream Number: $StreamNumber"; Break
									}
								"CTXZLC"	
									{
										$tmp = "Virtual Channel: Zero Latency Data Channel - Stream Number: $StreamNumber"; Break
									}
								"CTXCCM"	
									{
										$tmp = "Virtual Channel: Client COM Port Mapping - Stream Number: $StreamNumber"; Break
									}
								"CTXCPM"	
									{
										$tmp = "Virtual Channel: Client Printer Mapping - Stream Number: $StreamNumber"; Break
									}
								"CTXCOM1"	
									{
										$tmp = "Virtual Channel: Legacy Client Printer Mapping (COM1) - Stream Number: $StreamNumber"; Break
									}
								"CTXCOM2"	
									{
										$tmp = "Virtual Channel: Legacy Client Printer Mapping (COM2) - Stream Number: $StreamNumber"; Break
									}
								"CTXLPT1"	
									{
										$tmp = "Virtual Channel: Legacy Client Printer Mapping (LPT1) - Stream Number: $StreamNumber"; Break
									}
								"CTXLPT2"	
									{
										$tmp = "Virtual Channel: Legacy Client Printer Mapping (LPT2) - Stream Number: $StreamNumber"; Break
									}
								Default		
									{
										#assume a custom virtual channel
										$tmp = "Virtual Channel: $($ChannelName) - Stream Number: $StreamNumber"; Break
									}
							}
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = "";
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								"",$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting "`t`t`t`t`t`t" $tmp
							}
						}
						$tmp = $Null
						$Values = $Null
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Port Redirection"
					If((validStateProp $Setting ClientComPortsAutoConnection State ) -and ($Setting.ClientComPortsAutoConnection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Port Redirection\Auto connect client COM ports"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientComPortsAutoConnection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientComPortsAutoConnection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientComPortsAutoConnection.State 
						}
					}
					If((validStateProp $Setting ClientLptPortsAutoConnection State ) -and ($Setting.ClientLptPortsAutoConnection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Port Redirection\Auto connect client LPT ports"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientLptPortsAutoConnection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientLptPortsAutoConnection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientLptPortsAutoConnection.State 
						}
					}
					If((validStateProp $Setting ClientComPortRedirection State ) -and ($Setting.ClientComPortRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Port Redirection\Client COM port redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientComPortRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientComPortRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientComPortRedirection.State 
						}
					}
					If((validStateProp $Setting ClientLptPortRedirection State ) -and ($Setting.ClientLptPortRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Port Redirection\Client LPT port redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientLptPortRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientLptPortRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientLptPortRedirection.State 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Printing"
					If((validStateProp $Setting ClientPrinterRedirection State ) -and ($Setting.ClientPrinterRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client printer redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientPrinterRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientPrinterRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ClientPrinterRedirection.State 
						}
					}
					If((validStateProp $Setting DefaultClientPrinter State ) -and ($Setting.DefaultClientPrinter.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Default printer - Choose client's Default printer"
						$tmp = ""
						Switch ($Setting.DefaultClientPrinter.Value)
						{
							"ClientDefault" {$tmp = "Set Default printer to the client's main printer"; Break}
							"DoNotAdjust"   {$tmp = "Do not adjust the user's Default printer"; Break}
							Default {$tmp = "Default printer could not be determined: $($Setting.DefaultClientPrinter.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp
						}
						$tmp = $Null
					}
					If((validStateProp $Setting PrinterAssignments State ) -and ($Setting.PrinterAssignments.State -ne "NotConfigured"))
					{
						If($Setting.PrinterAssignments.State -eq "Enabled")
						{
							$txt = "ICA\Printing\Printer assignments"
							$PrinterAssign = Get-ChildItem -path "$($xDriveName):\User\$($Policy.PolicyName)\Settings\ICA\Printing\PrinterAssignments" 4>$Null
							If($? -and $Null -ne $PrinterAssign)
							{
								$PrinterAssignments = $PrinterAssign.Contents
								ForEach($PrinterAssignment in $PrinterAssignments)
								{
									$Client = @()
									$DefaultPrinter = ""
									$SessionPrinters = @()
									$tmp1 = ""
									$tmp2 = ""
									$tmp3 = ""
									
									ForEach($Filter in $PrinterAssignment.Filters)
									{
										$Client += "$($Filter); "
									}
									
									Switch ($PrinterAssignment.DefaultPrinterOption)
									{
										"ClientDefault"		{$DefaultPrinter = "Client main printer"; Break}
										"NotConfigured"		{$DefaultPrinter = "<Not set>"; Break}
										"DoNotAdjust"		{$DefaultPrinter = "Do not adjust"; Break}
										"SpecificPrinter"	{$DefaultPrinter = $PrinterAssignment.SpecificDefaultPrinter; Break}
										Default				{$DefaultPrinter = "<Not set>"; Break}
									}
									
									ForEach($SessionPrinter in $PrinterAssignment.SessionPrinters)
									{
										$SessionPrinters += $SessionPrinter
									}
									
									$tmp1 = "Client Names/IP's: $($Client)"
									$tmp2 = "Default Printer  : $($DefaultPrinter)"
									$tmp3 = "Session Printers : $($SessionPrinters)"
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = $txt;
										Value = $tmp1;
										}
										
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp2;
										}
										
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp3;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp1,$htmlwhite))
										
										$rowdata += @(,(
										"",$htmlbold,
										$tmp2,$htmlwhite))
										
										$rowdata += @(,(
										"",$htmlbold,
										$tmp3,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting $txt $tmp1
										OutputPolicySetting "`t`t`t`t" $tmp2
										OutputPolicySetting "`t`t`t`t" $tmp3
									}
									$tmp1 = $Null
									$tmp2 = $Null
									$tmp3 = $Null
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.PrinterAssignments.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.PrinterAssignments.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.PrinterAssignments.State 
							}
						}
					}
					If((validStateProp $Setting AutoCreationEventLogPreference State ) -and ($Setting.AutoCreationEventLogPreference.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Printer auto-creation event log preference"
						$tmp = ""
						Switch ($Setting.AutoCreationEventLogPreference.Value)
						{
							"LogErrorsOnly"        {$tmp = "Log errors only"; Break}
							"LogErrorsAndWarnings" {$tmp = "Log errors and warnings"; Break}
							"DoNotLog"             {$tmp = "Do not log errors or warnings"; Break}
							Default {$tmp = "Printer auto-creation event log preference could not be determined: $($Setting.AutoCreationEventLogPreference.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp
						}
						$tmp = $Null
					}
					If((validStateProp $Setting SessionPrinters State ) -and ($Setting.SessionPrinters.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Session printers"
						If(validStateProp $Setting SessionPrinters Values )
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = "";
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								"",$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt ""
							}
							$valArray = $Setting.SessionPrinters.Values
							$tmp = ""
							ForEach($printer in $valArray)
							{
								$prArray = $printer.Split(',')
								ForEach($element in $prArray)
								{
									If($element.SubString(0, 2) -eq "\\")
									{
										$index = $element.SubString(2).IndexOf('\')
										If($index -ge 0)
										{
											$server = $element.SubString(0, $index + 2)
											$share  = $element.SubString($index + 3)
											$tmp = "Server: $($server)"
											If($MSWord -or $PDF)
											{
												$SettingsWordTable += @{
												Text = "";
												Value = $tmp;
												}
											}
											If($HTML)
											{
												$rowdata += @(,(
												"",$htmlbold,
												$tmp,$htmlwhite))
											}
											If($Text)
											{
												OutputPolicySetting "" $tmp
											}
											$tmp = "Shared Name: $($share)"
											If($MSWord -or $PDF)
											{
												$SettingsWordTable += @{
												Text = "";
												Value = $tmp;
												}
											}
											If($HTML)
											{
												$rowdata += @(,(
												"",$htmlbold,
												$tmp,$htmlwhite))
											}
											If($Text)
											{
												OutputPolicySetting "" $tmp
											}
										}
										$index = $Null
									}
									Else
									{
										$tmp1 = $element.SubString(0, 4)
										$tmp = Get-PrinterModifiedSettings $tmp1 $element
										If(![String]::IsNullOrEmpty($tmp))
										{
											If($MSWord -or $PDF)
											{
												$SettingsWordTable += @{
												Text = "";
												Value = $tmp;
												}
											}
											If($HTML)
											{
												$rowdata += @(,(
												"",$htmlbold,
												$tmp,$htmlwhite))
											}
											If($Text)
											{
												OutputPolicySetting "" $tmp
											}
										}
										$tmp1 = $Null
										$tmp = $Null
									}
								}
							}

							$valArray = $Null
							$prArray = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No Session printers were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp
							}
						}
					}
					If((validStateProp $Setting WaitForPrintersToBeCreated State ) -and ($Setting.WaitForPrintersToBeCreated.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Wait for printers to be created (server desktop)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WaitForPrintersToBeCreated.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WaitForPrintersToBeCreated.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.WaitForPrintersToBeCreated.State 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Printing\Client Printers"
					If((validStateProp $Setting ClientPrinterAutoCreation State ) -and ($Setting.ClientPrinterAutoCreation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Auto-create client printers"
						$tmp = ""
						Switch ($Setting.ClientPrinterAutoCreation.Value)
						{
							"DoNotAutoCreate"    {$tmp = "Do not auto-create client printers"; Break}
							"DefaultPrinterOnly" {$tmp = "Auto-create the client's Default printer only"; Break}
							"LocalPrintersOnly"  {$tmp = "Auto-create local (non-network) client printers only"; Break}
							"AllPrinters"        {$tmp = "Auto-create all client printers"; Break}
							Default {$tmp = "Auto-create client printers could not be determined: $($Setting.ClientPrinterAutoCreation.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp
						}
						$tmp = $Null
					}
					If((validStateProp $Setting GenericUniversalPrinterAutoCreation State ) -and ($Setting.GenericUniversalPrinterAutoCreation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Auto-create generic universal printer"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.GenericUniversalPrinterAutoCreation.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.GenericUniversalPrinterAutoCreation.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.GenericUniversalPrinterAutoCreation.State 
						}
					}
					If((validStateProp $Setting AutoCreatePDFPrinter State ) -and ($Setting.AutoCreatePDFPrinter.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Auto-create PDF Universal Printer"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AutoCreatePDFPrinter.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AutoCreatePDFPrinter.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AutoCreatePDFPrinter.State 
						}
					}
					If((validStateProp $Setting DirectConnectionsToPrintServers State ) -and ($Setting.DirectConnectionsToPrintServers.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Direct connections to print servers"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DirectConnectionsToPrintServers.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DirectConnectionsToPrintServers.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.DirectConnectionsToPrintServers.State 
						}
					}
					If((validStateProp $Setting PrinterDriverMappings State ) -and ($Setting.PrinterDriverMappings.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Printer driver mapping and compatibility"
						If(validStateProp $Setting PrinterDriverMappings Values )
						{
							$array = $Setting.PrinterDriverMappings.Values
							$tmp = $array[0]
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp
							}
							
							$cnt = -1
							ForEach($element in $array)
							{
								$cnt++
								
								If($cnt -ne 0)
								{
									$Items = $element.Split(',')
									$DriverName = $Items[0]
									$Action = $Items[1]
									If($Action -match 'Replace=')
									{
										$ServerDriver = $Action.substring($Action.indexof("=")+1)
										$Action = "Replace "
									}
									Else
									{
										$ServerDriver = ""
										If($Action -eq "Allow")
										{
											$Action = "Allow "
										}
										ElseIf($Action -eq "Deny")
										{
											$Action = "Do not create "
										}
										ElseIf($Action -eq "UPD_Only")
										{
											$Action = "Create with universal driver "
										}
									}
									$tmp = "Driver Name: $($DriverName)"
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "`t`t`t`t`t`t`t`t     " $tmp
									}
									$tmp = "Action     : $($Action)"
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "`t`t`t`t`t`t`t`t     " $tmp
									}
									$tmp = "Settings   : "
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "`t`t`t`t`t`t`t`t     " $tmp
									}
									If($Items.count -gt 2)
									{
										[int]$BeginAt = 2
										[int]$EndAt = $Items.count
										for ($i=$BeginAt;$i -lt $EndAt; $i++) 
										{
											$tmp2 = $Items[$i].SubString(0, 4)
											$tmp = Get-PrinterModifiedSettings $tmp2 $Items[$i]
											If(![String]::IsNullOrEmpty($tmp))
											{
												If($MSWord -or $PDF)
												{
													$SettingsWordTable += @{
													Text = "";
													Value = $tmp;
													}
												}
												If($HTML)
												{
													$rowdata += @(,(
													"",$htmlbold,
													$tmp,$htmlwhite))
												}
												If($Text)
												{
													OutputPolicySetting "`t`t`t`t`t`t`t`t     " $tmp
												}
											}
										}
									}
									Else
									{
										$tmp = "Unmodified "
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "`t`t`t`t`t`t`t`t     " $tmp
										}
									}

									If(![String]::IsNullOrEmpty($ServerDriver))
									{
										$tmp = "Server Driver: $($ServerDriver)"
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "`t`t`t`t`t`t`t`t     " $tmp
										}
									}
									$tmp = $Null
								}
							}
						}
						Else
						{
							$tmp = "No Printer driver mapping and compatibility were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp
							}
						}
					}
					If((validStateProp $Setting PrinterPropertiesRetention State ) -and ($Setting.PrinterPropertiesRetention.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Printer properties retention"
						$tmp = ""
						Switch ($Setting.PrinterPropertiesRetention.Value)
						{
							"SavedOnClientDevice"   {$tmp = "Saved on the client device only"; Break}
							"RetainedInUserProfile" {$tmp = "Retained in user profile only"; Break}
							"FallbackToProfile"     {$tmp = "Held in profile only if not saved on client"; Break}
							"DoNotRetain"           {$tmp = "Do not retain printer properties"; Break}
							Default {$tmp = "Printer properties retention could not be determined: $($Setting.PrinterPropertiesRetention.Value)"; Break}
						}

						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Printing\Drivers"
					If((validStateProp $Setting UniversalDriverPriority State ) -and ($Setting.UniversalDriverPriority.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Drivers\Universal driver preference"
						$Values = $Setting.UniversalDriverPriority.Value.Split(';')
						$tmp = ""
						$cnt = 0
						ForEach($Value in $Values)
						{
							If($Null -eq $Value)
							{
								$Value = ''
							}
							$cnt++
							$tmp = "$($Value)"
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp 
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting "`t`t`t`t`t`t" $tmp
								}
							}
						}
						$tmp = $Null
						$Values = $Null
					}
					If((validStateProp $Setting UniversalPrintDriverUsage State ) -and ($Setting.UniversalPrintDriverUsage.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Drivers\Universal print driver usage"
						$tmp = ""
						Switch ($Setting.UniversalPrintDriverUsage.Value)
						{
							"SpecificOnly"       {$tmp = "Use only printer model specific drivers"; Break}
							"UpdOnly"            {$tmp = "Use universal printing only"; Break}
							"FallbackToUpd"      {$tmp = "Use universal printing only if requested driver is unavailable"; Break}
							"FallbackToSpecific" {$tmp = "Use printer model specific drivers only if universal printing is unavailable"; Break}
							Default {$tmp = "Universal print driver usage could not be determined: $($Setting.UniversalPrintDriverUsage.Value)"; Break}
						}

						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Printing\Universal Print Server"
					If((validStateProp $Setting UpcSslCipherSuite State ) -and ($Setting.UpcSslCipherSuite.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\SSL Cipher Suite"
						Switch ($Setting.UpcSslCipherSuite.Value)
						{
							"All"	{$tmp = "All"; Break}
							"COM"	{$tmp = "COM"; Break}
							"GOV"	{$tmp = "GOV"; Break}
							Default	{$tmp = "Universal Print Server SSL Cipher Suite value could not be determined: $($Setting.UpcSslCipherSuite.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting UpcSslComplianceMode State ) -and ($Setting.UpcSslComplianceMode.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\SSL Compliance Mode"
						Switch ($Setting.UpcSslComplianceMode.Value)
						{
							"None"	{$tmp = "None"; Break}
							"SP800_52"	{$tmp = "SP800-52"; Break}
							Default	{$tmp = "Universal Print Server SSL Compliance Mode value could not be determined: $($Setting.UpcSslComplianceMode.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting UpcSslEnable State ) -and ($Setting.UpcSslEnable.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\SSL Enabled"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UpcSslEnable.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpcSslEnable.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UpcSslEnable.State 
						}
					}
					If((validStateProp $Setting UpcSslFips State ) -and ($Setting.UpcSslFips.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\SSL FIPS Mode"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UpcSslFips.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpcSslFips.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UpcSslFips.State 
						}
					}
					If((validStateProp $Setting UpcSslProtocolVersion State ) -and ($Setting.UpcSslProtocolVersion.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\SSL Protocol Version"
						Switch ($Setting.UpcSslProtocolVersion.Value)
						{
							"All"	{$tmp = "All"; Break}
							"TLS1"	{$tmp = "TLSv1"; Break}
							"TLS11"	{$tmp = "TLSv1.1"; Break}
							"TLS12"	{$tmp = "TLSv1.2"; Break}
							Default	{$tmp = "Universal Print Server SSL Protocol Version value could not be determined: $($Setting.UpcSslProtocolVersion.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting UpcSslCgpPort State ) -and ($Setting.UpcSslCgpPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\SSL Universal Print Server encrypted print data stream (CGP) port"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UpcSslCgpPort.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpcSslCgpPort.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UpcSslCgpPort.Value 
						}
					}
					If((validStateProp $Setting UpcSslHttpsPort State ) -and ($Setting.UpcSslHttpsPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\SSL Universal Print Server encrypted web service (HTTPS/SOAP) port"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UpcSslHttpsPort.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpcSslHttpsPort.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UpcSslHttpsPort.Value 
						}
					}
					If((validStateProp $Setting UpsEnable State ) -and ($Setting.UpsEnable.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Server enable"
						If($Setting.UpsEnable.State)
						{
							$tmp = ""
						}
						Else
						{
							$tmp = "Disabled"
						}
						Switch ($Setting.UpsEnable.Value)
						{
							"UpsEnabledWithFallback"	{$tmp = "Enabled with fallback to Windows' native remote printing"; Break}
							"UpsOnlyEnabled"			{$tmp = "Enabled with no fallback to Windows' native remote printing"; Break}
							"UpsDisabled"				{$tmp = "Disabled"; Break}
							Default	{$tmp = "Universal Print Server enable value could not be determined: $($Setting.UpsEnable.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting UpsCgpPort State ) -and ($Setting.UpsCgpPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Server print data stream (CGP) port"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UpsCgpPort.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpsCgpPort.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UpsCgpPort.Value 
						}
					}
					If((validStateProp $Setting UpsPrintStreamInputBandwidthLimit State ) -and ($Setting.UpsPrintStreamInputBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Server print stream input bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UpsPrintStreamInputBandwidthLimit.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpsPrintStreamInputBandwidthLimit.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UpsPrintStreamInputBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting UpsHttpPort State ) -and ($Setting.UpsHttpPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Server web service (HTTP/SOAP) port"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UpsHttpPort.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpsHttpPort.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UpsHttpPort.Value 
						}
					}
					If((validStateProp $Setting LoadBalancedPrintServers State ) -and ($Setting.LoadBalancedPrintServers.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Servers for load balancing"
						If(validStateProp $Setting LoadBalancedPrintServers Values )
						{
							$array = $Setting.LoadBalancedPrintServers.Values
							$tmp = $array[0]
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp 
							}

							$txt = ""
							$cnt = -1
							
							ForEach($element in $array)
							{
								$cnt++
								
								If($cnt -ne 0)
								{
									$tmp = "$($element) "
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$array = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No Universal Print Servers for load balancing were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp 
							}
						}
					}
					If((validStateProp $Setting PrintServersOutOfServiceThreshold State ) -and ($Setting.PrintServersOutOfServiceThreshold.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Servers out-of-service threshold (seconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.PrintServersOutOfServiceThreshold.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PrintServersOutOfServiceThreshold.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.PrintServersOutOfServiceThreshold.Value 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Printing\Universal Printing"
					If((validStateProp $Setting EMFProcessingMode State ) -and ($Setting.EMFProcessingMode.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing EMF processing mode"
						$tmp = ""
						Switch ($Setting.EMFProcessingMode.Value)
						{
							"ReprocessEMFsForPrinter" {$tmp = "Reprocess EMFs for printer"; Break}
							"SpoolDirectlyToPrinter"  {$tmp = "Spool directly to printer"; Break}
							Default {$tmp = "Universal printing EMF processing mode could not be determined: $($Setting.EMFProcessingMode.Value)"; Break}
						}
						 
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting ImageCompressionLimit State ) -and ($Setting.ImageCompressionLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing image compression limit"
						$tmp = ""
						Switch ($Setting.ImageCompressionLimit.Value)
						{
							"NoCompression"       {$tmp = "No compression"; Break}
							"LosslessCompression" {$tmp = "Best quality (lossless compression)"; Break}
							"MinimumCompression"  {$tmp = "High quality"; Break}
							"MediumCompression"   {$tmp = "Standard quality"; Break}
							"MaximumCompression"  {$tmp = "Reduced quality (maximum compression)"; Break}
							Default {$tmp = "Universal printing image compression limit could not be determined: $($Setting.ImageCompressionLimit.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting UPDCompressionDefaults State ) -and ($Setting.UPDCompressionDefaults.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing optimization defaults"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = "";
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							"",$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt "" 
						}
						
						$TmpArray = $Setting.UPDCompressionDefaults.Value.Split(',')
						$tmp = ""
						ForEach($Thing in $TmpArray)
						{
							$TestLabel = $Thing.substring(0, $Thing.indexof("="))
							$TestSetting = $Thing.substring($Thing.indexof("=")+1)
							$TxtLabel = ""
							$TxtSetting = ""
							Switch($TestLabel)
							{
								"ImageCompression"
								{
									$TxtLabel = "Desired image quality:"
									Switch($TestSetting)
									{
										"StandardQuality"	{$TxtSetting = "Standard quality"; Break}
										"BestQuality"		{$TxtSetting = "Best quality (lossless compression)"; Break}
										"HighQuality"		{$TxtSetting = "High quality"; Break}
										"ReducedQuality"	{$TxtSetting = "Reduced quality (maximum compression)"; Break}
									}
								}
								"HeavyweightCompression"
								{
									$TxtLabel = "Enable heavyweight compression:"
									If($TestSetting -eq "True")
									{
										$TxtSetting = "Yes"
									}
									Else
									{
										$TxtSetting = "No"
									}
								}
								"ImageCaching"
								{
									$TxtLabel = "Allow caching of embedded images:"
									If($TestSetting -eq "True")
									{
										$TxtSetting = "Yes"
									}
									Else
									{
										$TxtSetting = "No"
									}
								}
								"FontCaching"
								{
									$TxtLabel = "Allow caching of embedded fonts:"
									If($TestSetting -eq "True")
									{
										$TxtSetting = "Yes"
									}
									Else
									{
										$TxtSetting = "No"
									}
								}
								"AllowNonAdminsToModify"
								{
									$TxtLabel = "Allow non-administrators to modify these settings:"
									If($TestSetting -eq "True")
									{
										$TxtSetting = "Yes"
									}
									Else
									{
										$TxtSetting = "No"
									}
								}
							}
							$tmp = "$($TxtLabel) $TxtSetting "
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = "";
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								"",$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting "`t`t`t`t`t`t`t`t`t" $tmp
							}
						}
						$TmpArray = $Null
						$tmp = $Null
						$TestLabel = $Null
						$TestSetting = $Null
						$TxtLabel = $Null
						$TxtSetting = $Null
					}
					If((validStateProp $Setting UniversalPrintingPreviewPreference State ) -and ($Setting.UniversalPrintingPreviewPreference.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing preview preference"
						$tmp = ""
						Switch ($Setting.UniversalPrintingPreviewPreference.Value)
						{
							"NoPrintPreview"        {$tmp = "Do not use print preview for auto-created or generic universal printers"; Break}
							"AutoCreatedOnly"       {$tmp = "Use print preview for auto-created printers only"; Break}
							"GenericOnly"           {$tmp = "Use print preview for generic universal printers only"; Break}
							"AutoCreatedAndGeneric" {$tmp = "Use print preview for both auto-created and generic universal printers"; Break}
							Default {$tmp = "Universal printing preview preference could not be determined: $($Setting.UniversalPrintingPreviewPreference.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting DPILimit State ) -and ($Setting.DPILimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing print quality limit"
						$tmp = ""
						Switch ($Setting.DPILimit.Value)
						{
							"Draft"				{$tmp = "Draft (150 DPI)"; Break}
							"LowResolution"		{$tmp = "Low Resolution (300 DPI)"; Break}
							"MediumResolution"	{$tmp = "Medium Resolution (600 DPI)"; Break}
							"HighResolution"	{$tmp = "High Resolution (1200 DPI)"; Break}
							"Unlimited"			{$tmp = "No Limit"; Break}
							Default {$tmp = "Universal printing print quality limit could not be determined: $($Setting.DPILimit.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Security"
					If((validStateProp $Setting MinimumEncryptionLevel State ) -and ($Setting.MinimumEncryptionLevel.State -ne "NotConfigured"))
					{
						$txt = "ICA\Security\SecureICA minimum encryption level" 
						$tmp = ""
						Switch ($Setting.MinimumEncryptionLevel.Value)
						{
							"Unknown"	{$tmp = "Unknown encryption"; Break}
							"Basic"		{$tmp = "Basic"; Break}
							"LogOn"		{$tmp = "RC5 (128 bit) logon only"; Break}
							"Bits40"	{$tmp = "RC5 (40 bit)"; Break}
							"Bits56"	{$tmp = "RC5 (56 bit)"; Break}
							"Bits128"	{$tmp = "RC5 (128 bit)"; Break}
							Default		{$tmp = "SecureICA minimum encryption level could not be determined: $($Setting.MinimumEncryptionLevel.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Server Limits"
					If((validStateProp $Setting IdleTimerInterval State ) -and ($Setting.IdleTimerInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\Server Limits\Server idle timer interval (milliseconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IdleTimerInterval.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IdleTimerInterval.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.IdleTimerInterval.Value 
						}
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Session Interactivity"
					If((validStateProp $Setting LossTolerantThresholds State ) -and ($Setting.LossTolerantThresholds.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Interactivity\Loss Tolerant Mode Thresholds"
						<#
						Value=		Threshold Type             Threshold Value
						loss,n; 	Packet Loss Percentage     n
						latency,n;	Round Trip Latency (ms)    nnn
						#>
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = "";
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							"",$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt ""
						}
						$Values = $Setting.LossTolerantThresholds.Value.Split(';')
						$tmp = ""
						ForEach($Value in $Values)
						{
							If($Value -eq "")
							{
								Continue
							}
							
							$tmparray = $Value.Split(",")
							$ThresholdType  = $tmparray[0]
							$ThresholdValue = $tmparray[1]
							
							If($ThresholdType -eq "")
							{
								Continue
							}
							
							Switch ($ThresholdType)
							{
								"loss"	
									{
										$tmp = "Threshold Type: Packet Loss Percentage     - Threshold Value: $ThresholdValue"; Break
									}
								"latency"	
									{
										$tmp = "Threshold Type: Round Trip Latency (ms)    - Threshold Value: $ThresholdValue"; Break
									}
								Default		
									{
										$tmp = "Unknown Values: Threshold Type: $ThresholdType - Threshold Value: $ThresholdValue"; Break
									}
							}
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = "";
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								"",$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting "`t`t`t`t`t`t" $tmp
							}
						}
						$tmp = $Null
						$Values = $Null
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Session Limits"
					If((validStateProp $Setting SessionDisconnectTimer State ) -and ($Setting.SessionDisconnectTimer.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Disconnected session timer"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionDisconnectTimer.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionDisconnectTimer.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.SessionDisconnectTimer.State 
						}
					}
					If((validStateProp $Setting SessionDisconnectTimerInterval State ) -and ($Setting.SessionDisconnectTimerInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Disconnected session timer interval (minutes)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionDisconnectTimerInterval.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionDisconnectTimerInterval.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.SessionDisconnectTimerInterval.Value 
						}
					}
					If((validStateProp $Setting SessionConnectionTimer State ) -and ($Setting.SessionConnectionTimer.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Session connection timer"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionConnectionTimer.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionConnectionTimer.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.SessionConnectionTimer.State 
						}
					}
					If((validStateProp $Setting SessionConnectionTimerInterval State ) -and ($Setting.SessionConnectionTimerInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Session connection timer interval (minutes)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionConnectionTimerInterval.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionConnectionTimerInterval.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.SessionConnectionTimerInterval.Value 
						}
					}
					If((validStateProp $Setting SessionIdleTimer State ) -and ($Setting.SessionIdleTimer.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Session idle timer"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionIdleTimer.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionIdleTimer.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.SessionIdleTimer.State 
						}
					}
					If((validStateProp $Setting SessionIdleTimerInterval State ) -and ($Setting.SessionIdleTimerInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Session idle timer interval (minutes)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionIdleTimerInterval.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionIdleTimerInterval.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.SessionIdleTimerInterval.Value 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Session Watermark"
					If((validStateProp $Setting EnableSessionWatermark State ) -and ($Setting.EnableSessionWatermark.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Watermark\Enable session watermark"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.EnableSessionWatermark.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnableSessionWatermark.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.EnableSessionWatermark.State 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Session Watermark\Watermark Content"
					If((validStateProp $Setting WatermarkIncludeClientIPAddress State ) -and ($Setting.WatermarkIncludeClientIPAddress.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Watermark\Watermark Content\Include client IP address"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WatermarkIncludeClientIPAddress.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WatermarkIncludeClientIPAddress.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.WatermarkIncludeClientIPAddress.State 
						}
					}
					If((validStateProp $Setting WatermarkIncludeConnectTime State ) -and ($Setting.WatermarkIncludeConnectTime.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Watermark\Watermark Content\Include connection time"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WatermarkIncludeConnectTime.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WatermarkIncludeConnectTime.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.WatermarkIncludeConnectTime.State 
						}
					}
					If((validStateProp $Setting WatermarkIncludeLogonUsername State ) -and ($Setting.WatermarkIncludeLogonUsername.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Watermark\Watermark Content\Include logon user name"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WatermarkIncludeLogonUsername.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WatermarkIncludeLogonUsername.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.WatermarkIncludeLogonUsername.State 
						}
					}
					If((validStateProp $Setting WatermarkIncludeVDAHostName State ) -and ($Setting.WatermarkIncludeVDAHostName.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Watermark\Watermark Content\Include VDA host name"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WatermarkIncludeVDAHostName.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WatermarkIncludeVDAHostName.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.WatermarkIncludeVDAHostName.State 
						}
					}
					If((validStateProp $Setting WatermarkIncludeVDAIPAddress State ) -and ($Setting.WatermarkIncludeVDAIPAddress.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Watermark\Watermark Content\Include VDA IP address"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WatermarkIncludeVDAIPAddress.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WatermarkIncludeVDAIPAddress.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.WatermarkIncludeVDAIPAddress.State 
						}
					}
					If((validStateProp $Setting WatermarkCustomText State ) -and ($Setting.WatermarkCustomText.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Watermark\Watermark Content\Watermark custom text"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WatermarkCustomText.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WatermarkCustomText.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.WatermarkCustomText.Value 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Session Watermark\Watermark Style"
					If((validStateProp $Setting WatermarkStyle State ) -and ($Setting.WatermarkStyle.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Watermark\Watermark Style\Session watermark style"
						$tmp = ""
						Switch ($Setting.WatermarkStyle.Value)
						{
							"StyleMutiple" {$tmp = "Multiple"; Break}
							"StyleSingle"   {$tmp = "Single"; Break}
							Default {$tmp = "Session watermark style could not be determined: $($Setting.WatermarkStyle.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp
						}
						$tmp = $Null
					}
					If((validStateProp $Setting WatermarkTransparency State ) -and ($Setting.WatermarkTransparency.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Watermark\Watermark Style\Watermark transparency"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WatermarkTransparency.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WatermarkTransparency.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.WatermarkTransparency.Value 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Session Reliability"
					If((validStateProp $Setting SessionReliabilityConnections State ) -and ($Setting.SessionReliabilityConnections.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Reliability\Session reliability connections"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionReliabilityConnections.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionReliabilityConnections.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.SessionReliabilityConnections.State 
						}
					}
					If((validStateProp $Setting SessionReliabilityPort State ) -and ($Setting.SessionReliabilityPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Reliability\Session reliability port number"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionReliabilityPort.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionReliabilityPort.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.SessionReliabilityPort.Value 
						}
					}
					If((validStateProp $Setting SessionReliabilityTimeout State ) -and ($Setting.SessionReliabilityTimeout.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Reliability\Session reliability timeout (seconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionReliabilityTimeout.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionReliabilityTimeout.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.SessionReliabilityTimeout.Value 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Time Zone Control"
					If((validStateProp $Setting LocalTimeEstimation State ) -and ($Setting.LocalTimeEstimation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Time Zone Control\Estimate local time for legacy clients"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LocalTimeEstimation.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LocalTimeEstimation.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LocalTimeEstimation.State 
						}
					}
					If((validStateProp $Setting RestoreServerTime State ) -and ($Setting.RestoreServerTime.State -ne "NotConfigured"))
					{
						$txt = "ICA\Time Zone Control\Restore Desktop OS time zone on session disconnect or logoff"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RestoreServerTime.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RestoreServerTime.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.RestoreServerTime.State 
						}
					}
					If((validStateProp $Setting SessionTimeZone State ) -and ($Setting.SessionTimeZone.State -ne "NotConfigured"))
					{
						$txt = "ICA\Time Zone Control\Use local time of client"
						$tmp = ""
						Switch ($Setting.SessionTimeZone.Value)
						{
							"UseServerTimeZone" {$tmp = "Use server time zone"; Break}
							"UseClientTimeZone" {$tmp = "Use client time zone"; Break}
							Default {$tmp = "Use local time of client could not be determined: $($Setting.SessionTimeZone.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\TWAIN Devices"
					If((validStateProp $Setting TwainRedirection State ) -and ($Setting.TwainRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\TWAIN devices\Client TWAIN device redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.TwainRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TwainRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.TwainRedirection.State 
						}
					}
					If((validStateProp $Setting TwainCompressionLevel State ) -and ($Setting.TwainCompressionLevel.State -ne "NotConfigured"))
					{
						$txt = "ICA\TWAIN devices\TWAIN compression level"
						Switch ($Setting.TwainCompressionLevel.Value)
						{
							"None"		{$tmp = "None"; Break}
							"Low"		{$tmp = "Low"; Break}
							"Medium"	{$tmp = "Medium"; Break}
							"High"		{$tmp = "High"; Break}
							Default		{$tmp = "TWAIN compression level could not be determined: $($Setting.TwainCompressionLevel.Value)"; Break}
						}

						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Bidirectional Content Redirection"
					If((validStateProp $Setting AllowURLRedirection State ) -and ($Setting.AllowURLRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bidirectional Content Redirection\Allow Bidirectional Content Redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AllowURLRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowURLRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AllowURLRedirection.State 
						}
					}
					If((validStateProp $Setting AllowedClientURLs State ) -and ($Setting.AllowedClientURLs.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bidirectional Content Redirection\Allowed URLs to be redirected to Client"
						$array = $Setting.AllowedClientURLs.Value.Split(';')
						$tmp = $array[0]
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}

						$txt = ""
						$cnt = -1
						ForEach($element in $array)
						{
							$cnt++
							
							If($cnt -ne 0)
							{
								$tmp = "$($element) "
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$array = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting AllowedVDAURLs State ) -and ($Setting.AllowedVDAURLs.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bidirectional Content Redirection\Allowed URLs to be redirected to VDA"
						$array = $Setting.AllowedVDAURLs.Value.Split(';')
						$tmp = $array[0]
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}

						$txt = ""
						$cnt = -1
						ForEach($element in $array)
						{
							$cnt++
							
							If($cnt -ne 0)
							{
								$tmp = "$($element) "
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$array = $Null
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\USB Devices"
					If((validStateProp $Setting ClientUsbDeviceOptimizationRules State ) -and ($Setting.ClientUsbDeviceOptimizationRules.State -ne "NotConfigured"))
					{
						$txt = "ICA\USB devices\Client USB device optimization rules"
						If(validStateProp $Setting ClientUsbDeviceOptimizationRules Values )
						{
							$array = $Setting.ClientUsbDeviceOptimizationRules.Values
							$tmp = $array[0]
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp 
							}

							$txt = ""
							$cnt = -1
							ForEach($element in $array)
							{
								$cnt++
								
								If($cnt -ne 0)
								{
									$tmp = "$($element) "
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$array = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No Client USB device optimization rules were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp 
							}
						}
					}
					If((validStateProp $Setting UsbDeviceRedirection State ) -and ($Setting.UsbDeviceRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\USB devices\Client USB device redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UsbDeviceRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UsbDeviceRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UsbDeviceRedirection.State 
						}
					}
					If((validStateProp $Setting UsbDeviceRedirectionRules State ) -and ($Setting.UsbDeviceRedirectionRules.State -ne "NotConfigured"))
					{
						$txt = "ICA\USB devices\Client USB device redirection rules"
						If(validStateProp $Setting UsbDeviceRedirectionRules Values )
						{
							$array = $Setting.UsbDeviceRedirectionRules.Values
							$tmp = $array[0]
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp 
							}

							$txt = ""
							$cnt = -1
							ForEach($element in $array)
							{
								$cnt++
								
								If($cnt -ne 0)
								{
									$tmp = "$($element) "
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$array = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No Client USB device redirections rules were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp 
							}
						}
					}
					If((validStateProp $Setting UsbPlugAndPlayRedirection State ) -and ($Setting.UsbPlugAndPlayRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\USB devices\Client USB Plug and Play device redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UsbPlugAndPlayRedirection.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UsbPlugAndPlayRedirection.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UsbPlugAndPlayRedirection.State 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Visual Display"
					If((validStateProp $Setting PreferredColorDepthForSimpleGraphics State ) -and ($Setting.PreferredColorDepthForSimpleGraphics.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Preferred color depth for simple graphics"
						$tmp = ""
						Switch ($Setting.PreferredColorDepthForSimpleGraphics.Value)
						{
							"ColorDepth24Bit"	{$tmp = "24 bits per pixel"; Break}
							"ColorDepth16Bit"	{$tmp = "16 bits per pixel"; Break}
							"ColorDepth8Bit"	{$tmp = "8 bits per pixel"; Break}
							"Default" {$tmp = "Preferred color depth for simple graphics could not be determined: $($Setting.PreferredColorDepthForSimpleGraphics.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting FramesPerSecond State ) -and ($Setting.FramesPerSecond.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Target frame rate (fps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FramesPerSecond.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FramesPerSecond.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.FramesPerSecond.Value 
						}
					}
					If((validStateProp $Setting VisualQuality State ) -and ($Setting.VisualQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Visual quality"
						$tmp = ""
						Switch ($Setting.VisualQuality.Value)
						{
							"BuildToLossless"	{$tmp = "Build to Lossless"; Break}
							"AlwaysLossless"	{$tmp = "Always Lossless"; Break}
							"High"				{$tmp = "High"; Break}
							"Medium"			{$tmp = "Medium"; Break}
							"Low"				{$tmp = "Low"; Break}
							"Default" {$tmp = "Visual quality could not be determined: $($Setting.VisualQuality.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Visual Display\Moving Images"
					If((validStateProp $Setting TargetedMinimumFramesPerSecond State ) -and ($Setting.TargetedMinimumFramesPerSecond.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Moving Images\Target Minimum Frame Rate (fps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.TargetedMinimumFramesPerSecond.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TargetedMinimumFramesPerSecond.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.TargetedMinimumFramesPerSecond.Value 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\Visual Display\Still Images"
					If((validStateProp $Setting ExtraColorCompression State ) -and ($Setting.ExtraColorCompression.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Still Images\Extra Color Compression"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExtraColorCompression.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExtraColorCompression.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExtraColorCompression.State 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tICA\WebSockets"
					If((validStateProp $Setting AcceptWebSocketsConnections State ) -and ($Setting.AcceptWebSocketsConnections.State -ne "NotConfigured"))
					{
						$txt = "ICA\WebSockets\WebSockes connections"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AcceptWebSocketsConnections.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AcceptWebSocketsConnections.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.AcceptWebSocketsConnections.State 
						}
					}
					If((validStateProp $Setting WebSocketsPort State ) -and ($Setting.WebSocketsPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\WebSockets\WebSockets port number"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WebSocketsPort.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WebSocketsPort.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.WebSocketsPort.Value 
						}
					}
					If((validStateProp $Setting WSTrustedOriginServerList State ) -and ($Setting.WSTrustedOriginServerList.State -ne "NotConfigured"))
					{
						$txt = "ICA\WebSockets\WebSockets trusted origin server list"
						$tmpArray = $Setting.WSTrustedOriginServerList.Value.Split(",")
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $tmpArray)
						{
							$cnt++
							$tmp = "$($Thing)"
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$tmpArray = $Null
						$tmp = $Null
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`t`tLoad Management"
					If((validStateProp $Setting ConcurrentLogonsTolerance State ) -and ($Setting.ConcurrentLogonsTolerance.State -ne "NotConfigured"))
					{
						$txt = "Load Management\Concurrent logons tolerance"
						If($Setting.ConcurrentLogonsTolerance.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.ConcurrentLogonsTolerance.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ConcurrentLogonsTolerance.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.ConcurrentLogonsTolerance.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.ConcurrentLogonsTolerance.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ConcurrentLogonsTolerance.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.ConcurrentLogonsTolerance.State 
							}
						}
					}
					If((validStateProp $Setting CPUUsage State ) -and ($Setting.CPUUsage.State -ne "NotConfigured"))
					{
						$txt = "Load Management\CPU usage"
						$tmp = ""
						If($Setting.CPUUsage.State -eq "Enabled")
						{
							If($Setting.CPUUsage.Value -eq -1)
							{
								$tmp = "Disabled"
							}
							Else
							{
								$tmp = "Report full load $($Setting.CPUUsage.Value)(%)"
							}
						}
						Else
						{
							$tmp = "Disabled"
						}
						If($MSWord -or $PDF)
						{
							If($Setting.CPUUsage.State -eq "Enabled")
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp
						}
					}
					If((validStateProp $Setting CPUUsageExcludedProcessPriority State ) -and ($Setting.CPUUsageExcludedProcessPriority.State -ne "NotConfigured"))
					{
						$txt = "Load Management\CPU usage excluded process priority"
						If($Setting.CPUUsageExcludedProcessPriority.State -eq "Enabled")
						{
							$tmp = ""
							Switch ($Setting.CPUUsageExcludedProcessPriority.Value)
							{
								"BelowNormalOrLow"	{$tmp = "Below Normal or Low"; Break}
								"Low"				{$tmp = "Low"; Break}
								Default {$tmp = "CPU usage excluded process priority could not be determined: $($Setting.CPUUsageExcludedProcessPriority.Value)"; Break}
							}
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.CPUUsageExcludedProcessPriority.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPUUsageExcludedProcessPriority.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.CPUUsageExcludedProcessPriority.State 
							}
						}
					}
					If((validStateProp $Setting DiskUsage State ) -and ($Setting.DiskUsage.State -ne "NotConfigured"))
					{
						$txt = "Load Management\Disk usage"
						$tmp = ""
						If($Setting.DiskUsage.State -eq "Enabled")
						{
							If($Setting.DiskUsage.Value -eq -1)
							{
								$tmp = "Disabled"
							}
							Else
							{
								$tmp = "Report 75% load (disk queue length): $($Setting.DiskUsage.Value)"
							}
						}
						Else
						{
							$tmp = "Disabled"
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting MaximumNumberOfSessions State ) -and ($Setting.MaximumNumberOfSessions.State -ne "NotConfigured"))
					{
						If($Setting.MaximumNumberOfSessions.State -eq "Enabled")
						{
							$txt = "Load Management\Maximum number of sessions - Limit"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.MaximumNumberOfSessions.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.MaximumNumberOfSessions.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.MaximumNumberOfSessions.Value 
							}
						}
						Else
						{
							$txt = "Load Management\Maximum number of sessions"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.MaximumNumberOfSessions.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.MaximumNumberOfSessions.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.MaximumNumberOfSessions.Value 
							}
						}
					}
					If((validStateProp $Setting MemoryUsage State ) -and ($Setting.MemoryUsage.State -ne "NotConfigured"))
					{
						$txt = "Load Management\Memory usage"
						$tmp = ""
						If($Setting.MemoryUsage.State -eq "Enabled")
						{
							If($Setting.MemoryUsage.Value -eq -1)
							{
								$tmp = "Disabled"
							}
							Else
							{
								$tmp = "Report full load (%): $($Setting.MemoryUsage.Value)"
							}
						}
						Else
						{
							$tmp = "Disabled"
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting MemoryUsageBaseLoad State ) -and ($Setting.MemoryUsageBaseLoad.State -ne "NotConfigured"))
					{
						$txt = "Load Management\Memory usage base load"
						$tmp = ""
						If($Setting.MemoryUsageBaseLoad.State -eq "Enabled")
						{
							$tmp = "Report zero load (MBs): $($Setting.MemoryUsageBaseLoad.Value)"
						}
						Else
						{
							$tmp = "Disabled"
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management"
					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Advanced settings"
					If((validStateProp $Setting CEIPEnabled State ) -and ($Setting.CEIPEnabled.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Customer Experience Improvement Program"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.CEIPEnabled.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.CEIPEnabled.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.CEIPEnabled.State
						}
					}
					If((validStateProp $Setting DisableDynamicConfig State ) -and ($Setting.DisableDynamicConfig.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Disable automatic configuration"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DisableDynamicConfig.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DisableDynamicConfig.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.DisableDynamicConfig.State
						}
					}
					If((validStateProp $Setting FSLogixProfileContainerSupport State ) -and ($Setting.FSLogixProfileContainerSupport.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Enable multi-session write-back for FSLogix Profile Container"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FSLogixProfileContainerSupport.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FSLogixProfileContainerSupport.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.FSLogixProfileContainerSupport.State
						}
					}
					If((validStateProp $Setting OutlookSearchRoamingEnabled State ) -and ($Setting.OutlookSearchRoamingEnabled.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Enable search index roaming for Outlook"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.OutlookSearchRoamingEnabled.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OutlookSearchRoamingEnabled.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.OutlookSearchRoamingEnabled.State
						}
					}
					If((validStateProp $Setting LogoffRatherThanTempProfile State ) -and ($Setting.LogoffRatherThanTempProfile.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Log off user if a problem is encountered"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogoffRatherThanTempProfile.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogoffRatherThanTempProfile.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogoffRatherThanTempProfile.State
						}
					}
					If((validStateProp $Setting LoadRetries_Part State ) -and ($Setting.LoadRetries_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Number of retries when accessing locked files"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LoadRetries_Part.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LoadRetries_Part.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LoadRetries_Part.Value 
						}
					}
					If((validStateProp $Setting OutlookEdbBackupEnabled State ) -and ($Setting.OutlookEdbBackupEnabled.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Outlook search index database - backup and restore"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.OutlookEdbBackupEnabled.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OutlookEdbBackupEnabled.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.OutlookEdbBackupEnabled.State
						}
					}
					If((validStateProp $Setting ProcessCookieFiles State ) -and ($Setting.ProcessCookieFiles.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Process Internet cookie files on logoff"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ProcessCookieFiles.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProcessCookieFiles.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ProcessCookieFiles.State
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Basic settings"
					If((validStateProp $Setting PSMidSessionWriteBack State ) -and ($Setting.PSMidSessionWriteBack.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Active write back"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.PSMidSessionWriteBack.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PSMidSessionWriteBack.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.PSMidSessionWriteBack.State
						}
					}
					If((validStateProp $Setting PSMidSessionWriteBackReg State ) -and ($Setting.PSMidSessionWriteBackReg.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Active write back registry"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.PSMidSessionWriteBackReg.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PSMidSessionWriteBackReg.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.PSMidSessionWriteBack.State
						}
					}
					If((validStateProp $Setting ServiceActive State ) -and ($Setting.ServiceActive.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Enable Profile management"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ServiceActive.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ServiceActive.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ServiceActive.State
						}
					}
					If((validStateProp $Setting ExcludedGroups_Part State ) -and ($Setting.ExcludedGroups_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Excluded groups"
						If($Setting.ExcludedGroups_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting ExcludedGroups_Part Values )
							{
								$tmpArray = $Setting.ExcludedGroups_Part.Values
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = $txt;
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Excluded groups were found"
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.ExcludedGroups_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ExcludedGroups_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.ExcludedGroups_Part.State
							}
						}
					}
					If((validStateProp $Setting MigrateUserStore_Part State ) -and ($Setting.MigrateUserStore_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Migrate user store"
						If($Setting.MigrateUserStore_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.MigrateUserStore_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.MigrateUserStore_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.MigrateUserStore_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.MigrateUserStore_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.MigrateUserStore_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.MigrateUserStore_Part.State
							}
						}
					}
					If((validStateProp $Setting OfflineSupport State ) -and ($Setting.OfflineSupport.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Offline profile support"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.OfflineSupport.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OfflineSupport.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.OfflineSupport.State
						}
					}
					If((validStateProp $Setting DATPath_Part State ) -and ($Setting.DATPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Path to user store"
						If($Setting.DATPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.DATPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.DATPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.DATPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.DATPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.DATPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.DATPath_Part.State
							}
						}
					}
					If((validStateProp $Setting ProcessAdmins State ) -and ($Setting.ProcessAdmins.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Process logons of local administrators"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ProcessAdmins.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProcessAdmins.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ProcessAdmins.State
						}
					}
					If((validStateProp $Setting ProcessedGroups_Part State ) -and ($Setting.ProcessedGroups_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Processed groups"
						If($Setting.ProcessedGroups_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting ProcessedGroups_Part Values )
							{
								$tmpArray = $Setting.ProcessedGroups_Part.Values
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = $txt;
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Processed groups were found"
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}	
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.ProcessedGroups_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ProcessedGroups_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.ProcessedGroups_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Citrix Virtual Apps Optimization settings"

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Cross-Platform settings"
					If((validStateProp $Setting CPUserGroups_Part State ) -and ($Setting.CPUserGroups_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Cross-Platform settings\Cross-platform settings user groups"
						If($Setting.CPUserGroups_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting CPUserGroups_Part Values )
							{
								$tmpArray = $Setting.CPUserGroups_Part.Values
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = $txt;
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Cross-platform settings user groups were found"
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.CPUserGroups_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPUserGroups_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.CPUserGroups_Part.State
							}
						}
					}
					If((validStateProp $Setting CPEnable State ) -and ($Setting.CPEnable.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Cross-Platform settings\Enable cross-platform settings"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.CPEnable.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.CPEnable.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.CPEnable.State
						}
					}
					If((validStateProp $Setting CPSchemaPathData State ) -and ($Setting.CPSchemaPathData.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Cross-Platform settings\Path to cross-platform definitions"
						If($Setting.CPSchemaPathData.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.CPSchemaPathData.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPSchemaPathData.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.CPSchemaPathData.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.CPSchemaPathData.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPSchemaPathData.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.CPSchemaPathData.State
							}
						}
					}
					If((validStateProp $Setting CPPathData State ) -and ($Setting.CPPathData.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Cross-Platform settings\Path to cross-platform settings store"
						If($Setting.CPPathData.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.CPPathData.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPPathData.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.CPPathData.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.CPPathData.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPPathData.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.CPPathData.State
							}
						}
					}
					If((validStateProp $Setting CPMigrationFromBaseProfileToCPStore State ) -and ($Setting.CPMigrationFromBaseProfileToCPStore.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Cross-Platform settings\Source for creating cross-platform settings"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.CPMigrationFromBaseProfileToCPStore.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.CPMigrationFromBaseProfileToCPStore.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.CPMigrationFromBaseProfileToCPStore.State
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\File system"
					If((validStateProp $Setting LogonExclusionCheck_Part State ) -and ($Setting.LogonExclusionCheck_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Logon Exclusion Check"
						$tmp = ""
						Switch ($Setting.LogonExclusionCheck_Part.Value)
						{
							"Disable"	{$tmp = "Synchronize excluded files or folders"; Break}
							"Ignore"	{$tmp = "Ignore excluded files or folders"; Break}
							"Delete"	{$tmp = "Delete excluded files or folders"; Break}
							Default		{$tmp = "Logon exclusion check could not be determined: $($Setting.LogonExclusionCheck_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\File system\Default Exclusions"
					If((validStateProp $Setting DefaultExclusionListSyncDir State ) -and ($Setting.DefaultExclusionListSyncDir.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\Enable Default Exclusion List - directories"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DefaultExclusionListSyncDir.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DefaultExclusionListSyncDir.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.DefaultExclusionListSyncDir.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir01 State ) -and ($Setting.ExclusionDefaultDir01.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_internetcache!"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir01.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir01.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir01.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir02 State ) -and ($Setting.ExclusionDefaultDir02.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Google\Chrome\User Data\Default\Cache"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir02.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir02.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir02.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir03 State ) -and ($Setting.ExclusionDefaultDir03.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Google\Chrome\User Data\Default\Cache Theme Images"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir03.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir03.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir03.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir04 State ) -and ($Setting.ExclusionDefaultDir04.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Google\Chrome\User Data\Default\JumpListIcons"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir04.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir04.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir04.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir05 State ) -and ($Setting.ExclusionDefaultDir05.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Google\Chrome\User Data\Default\JumpListIconsOld"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir05.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir05.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir05.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir06 State ) -and ($Setting.ExclusionDefaultDir06.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\GroupPolicy"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir06.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir06.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir06.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir07 State ) -and ($Setting.ExclusionDefaultDir07.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\AppV"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir07.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir07.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir07.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir08 State ) -and ($Setting.ExclusionDefaultDir08.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Messenger"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir08.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir08.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir08.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir09 State ) -and ($Setting.ExclusionDefaultDir09.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Office\15.0\Lync\Tracing"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir09.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir09.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir09.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir10 State ) -and ($Setting.ExclusionDefaultDir10.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\OneNote"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir10.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir10.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir10.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir11 State ) -and ($Setting.ExclusionDefaultDir11.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Outlook"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir11.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir11.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir11.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir12 State ) -and ($Setting.ExclusionDefaultDir12.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Terminal Server Client"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir12.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir12.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir12.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir13 State ) -and ($Setting.ExclusionDefaultDir13.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\UEV"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir13.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir13.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir13.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir14 State ) -and ($Setting.ExclusionDefaultDir14.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows Live"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir14.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir14.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir14.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir15 State ) -and ($Setting.ExclusionDefaultDir15.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows Live Contacts"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir15.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir15.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir15.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir16 State ) -and ($Setting.ExclusionDefaultDir16.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows\Application Shortcuts"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir16.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir16.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir16.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir17 State ) -and ($Setting.ExclusionDefaultDir17.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows\Burn"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir17.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir17.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir17.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir18 State ) -and ($Setting.ExclusionDefaultDir18.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows\CD Burning"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir18.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir18.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir18.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir19 State ) -and ($Setting.ExclusionDefaultDir19.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows\Notifications"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir19.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir19.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir19.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir20 State ) -and ($Setting.ExclusionDefaultDir20.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Packages"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir20.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir20.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir20.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir21 State ) -and ($Setting.ExclusionDefaultDir21.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Sun"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir21.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir21.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir21.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir22 State ) -and ($Setting.ExclusionDefaultDir22.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Windows Live"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir22.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir22.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir22.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir23 State ) -and ($Setting.ExclusionDefaultDir23.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localsettings!\Temp"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir23.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir23.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir23.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir24 State ) -and ($Setting.ExclusionDefaultDir24.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_roamingappdata!\Microsoft\AppV\Client\Catalog"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir24.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir24.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir24.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir25 State ) -and ($Setting.ExclusionDefaultDir25.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_roamingappdata!\Sun\Java\Deployment\cache"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir25.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir25.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir25.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir26 State ) -and ($Setting.ExclusionDefaultDir26.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_roamingappdata!\Sun\Java\Deployment\log"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir26.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir26.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir26.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir27 State ) -and ($Setting.ExclusionDefaultDir27.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_roamingappdata!\Sun\Java\Deployment\tmp"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir27.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir27.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir27.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir28 State ) -and ($Setting.ExclusionDefaultDir28.State -ne "NotConfigured"))
					{
						$txt = 'Profile Management\File system\Default Exclusions\UPM - $Recycle.Bin'
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir28.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir28.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir28.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir29 State ) -and ($Setting.ExclusionDefaultDir29.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - AppData\LocalLow"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir29.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir29.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir29.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir30 State ) -and ($Setting.ExclusionDefaultDir30.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - Tracing"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir30.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir30.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir30.State
						}
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\File system\Exclusions"
					If((validStateProp $Setting ExclusionListSyncDir_Part State ) -and ($Setting.ExclusionListSyncDir_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Exclusions\Exclusion list - directories"
						If($Setting.ExclusionListSyncDir_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting ExclusionListSyncDir_Part Values )
							{
								$tmpArray = $Setting.ExclusionListSyncDir_Part.Values
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = $txt;
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Exclusion list - directories were found"
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.ExclusionListSyncDir_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ExclusionListSyncDir_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.ExclusionListSyncDir_Part.State
							}
						}
					}
					If((validStateProp $Setting ExclusionListSyncFiles_Part State ) -and ($Setting.ExclusionListSyncFiles_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Exclusions\Exclusion list - files"
						If($Setting.ExclusionListSyncFiles_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting ExclusionListSyncFiles_Part Values )
							{
								$tmpArray = $Setting.ExclusionListSyncFiles_Part.Values
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = $txt;
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Exclusion list - files were found"
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.ExclusionListSyncFiles_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ExclusionListSyncFiles_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.ExclusionListSyncFiles_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\File system\Synchronization"
					If((validStateProp $Setting SyncDirList_Part State ) -and ($Setting.SyncDirList_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Synchronization\Directories to synchronize"
						If($Setting.SyncDirList_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting SyncDirList_Part Values )
							{
								$tmpArray = $Setting.SyncDirList_Part.Values
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = $txt;
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Directories to synchronize were found"
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.SyncDirList_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.SyncDirList_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.SyncDirList_Part.State
							}
						}
					}
					If((validStateProp $Setting SyncFileList_Part State ) -and ($Setting.SyncFileList_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Synchronization\Files to synchronize"
						If($Setting.SyncFileList_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting SyncFileList_Part Values )
							{
								$tmpArray = $Setting.SyncFileList_Part.Values
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = $txt;
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Files to synchronize were found"
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.SyncFileList_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.SyncFileList_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.SyncFileList_Part.State
							}
						}
					}
					If((validStateProp $Setting MirrorFoldersList_Part State ) -and ($Setting.MirrorFoldersList_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Synchronization\Folders to mirror"
						If($Setting.MirrorFoldersList_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting MirrorFoldersList_Part Values )
							{
								$tmpArray = $Setting.MirrorFoldersList_Part.Values
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = $txt;
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Folders to mirror were found"
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.MirrorFoldersList_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.MirrorFoldersList_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.MirrorFoldersList_Part.State
							}
						}
					}
					If((validStateProp $Setting ProfileContainer_Part State ) -and ($Setting.ProfileContainer_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Synchronization\Profile container - List of folders to be contained in profile disk"
						If($Setting.ProfileContainer_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting ProfileContainer_Part Values )
							{
								$tmpArray = $Setting.ProfileContainer_Part.Values
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = $txt;
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Directories to synchronize were found"
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.ProfileContainer_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ProfileContainer_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.ProfileContainer_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection"
					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\AppData(Roaming)"
					If((validStateProp $Setting FRAppDataPath_Part State ) -and ($Setting.FRAppDataPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\AppData(Roaming)\AppData(Roaming) path"
						If($Setting.FRAppDataPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRAppDataPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRAppDataPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRAppDataPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRAppDataPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRAppDataPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRAppDataPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRAppData_Part State ) -and ($Setting.FRAppData_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\AppData(Roaming)\Redirection settings for AppData(Roaming)"
						$tmp = ""
						Switch ($Setting.FRAppData_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRAppData_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Common settings"
					If((validStateProp $Setting FRAdminAccess_Part State ) -and ($Setting.FRAdminAccess_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Common settings\Grant administrator access"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FRAdminAccess_Part.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FRAdminAccess_Part.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.FRAdminAccess_Part.State
						}
					}
					If((validStateProp $Setting FRIncDomainName_Part State ) -and ($Setting.FRIncDomainName_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Common settings\Include domain name"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FRIncDomainName_Part.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FRIncDomainName_Part.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.FRIncDomainName_Part.State
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Contacts"
					If((validStateProp $Setting FRContactsPath_Part State ) -and ($Setting.FRContactsPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Contacts\Contacts path"
						If($Setting.FRContactsPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRContactsPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRContactsPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRContactsPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRContactsPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRContactsPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRContactsPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRContacts_Part State ) -and ($Setting.FRContacts_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Contacts\Redirection settings for Contacts"
						$tmp = ""
						Switch ($Setting.FRContacts_Part.Value)
						{
							"RedirectUncPath"	{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRContacts_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Desktop"
					If((validStateProp $Setting FRDesktopPath_Part State ) -and ($Setting.FRDesktopPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Desktop\Desktop path"
						If($Setting.FRDesktopPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRDesktopPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRDesktopPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRDesktopPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRDesktopPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRDesktopPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRDesktopPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRDesktop_Part State ) -and ($Setting.FRDesktop_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Desktop\Redirection settings for Desktop"
						$tmp = ""
						Switch ($Setting.FRDesktop_Part.Value)
						{
							"RedirectUncPath"	{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRDesktop_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Documents"
					If((validStateProp $Setting FRDocumentsPath_Part State ) -and ($Setting.FRDocumentsPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Documents\Documents path"
						If($Setting.FRDocumentsPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRDocumentsPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRDocumentsPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRDocumentsPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRDocumentsPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRDocumentsPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRDocumentsPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRDocuments_Part State ) -and ($Setting.FRDocuments_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Documents\Redirection settings for Documents"
						$tmp = ""
						Switch ($Setting.FRDocuments_Part.Value)
						{
							"RedirectUncPath"		{$tmp = "Redirect to the following UNC path"; Break}
							"RedirectRelativeHome"	{$tmp = "Redirect to the users' home directory"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRDocuments_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Downloads"
					If((validStateProp $Setting FRDownloadsPath_Part State ) -and ($Setting.FRDownloadsPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Downloads\Downloads path"
						If($Setting.FRDownloadsPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRDownloadsPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRDownloadsPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRDownloadsPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRDownloadsPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRDownloadsPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRDownloadsPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRDownloads_Part State ) -and ($Setting.FRDownloads_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Downloads\Redirection settings for Downloads"
						$tmp = ""
						Switch ($Setting.FRDownloads_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRDownloads_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Favorites"
					If((validStateProp $Setting FRFavoritesPath_Part State ) -and ($Setting.FRFavoritesPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Favorites\Favorites path"
						If($Setting.FRFavoritesPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRFavoritesPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRFavoritesPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRFavoritesPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRFavoritesPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRFavoritesPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRFavoritesPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRFavorites_Part State ) -and ($Setting.FRFavorites_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Favorites\Redirection settings for Favorites"
						$tmp = ""
						Switch ($Setting.FRFavorites_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRFavorites_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Links"
					If((validStateProp $Setting FRLinksPath_Part State ) -and ($Setting.FRLinksPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Links\Links path"
						If($Setting.FRLinksPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRLinksPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRLinksPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRLinksPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRLinksPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRLinksPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRLinksPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRLinks_Part State ) -and ($Setting.FRLinks_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Links\Redirection settings for Links"
						$tmp = ""
						Switch ($Setting.FRLinks_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRLinks_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Music"
					If((validStateProp $Setting FRMusicPath_Part State ) -and ($Setting.FRMusicPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Music\Music path"
						If($Setting.FRMusicPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRMusicPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRMusicPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRMusicPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRMusicPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRMusicPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRMusicPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRMusic_Part State ) -and ($Setting.FRMusic_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Music\Redirection settings for Music"
						$tmp = ""
						Switch ($Setting.FRMusic_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							"RedirectRelativeDocuments" {$tmp = "Redirect relative to Documents folder"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRMusic_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Pictures"
					If((validStateProp $Setting FRPicturesPath_Part State ) -and ($Setting.FRPicturesPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Pictures\Pictures path"
						If($Setting.FRPicturesPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRPicturesPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRPicturesPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRPicturesPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRPicturesPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRPicturesPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRPicturesPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRPictures_Part State ) -and ($Setting.FRPictures_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Pictures\Redirection settings for Pictures"
						$tmp = ""
						Switch ($Setting.FRPictures_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							"RedirectRelativeDocuments" {$tmp = "Redirect relative to Documents folder"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRPictures_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Saved Games"
					If((validStateProp $Setting FRSavedGames_Part State ) -and ($Setting.FRSavedGames_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Saved Games\Redirection settings for Saved Games"
						$tmp = ""
						Switch ($Setting.FRSavedGames_Part.Value)
						{
							"RedirectUncPath"	{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRSavedGames_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting FRSavedGamesPath_Part State ) -and ($Setting.FRSavedGamesPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Saved Games\Saved Games path"
						If($Setting.FRSavedGamesPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRSavedGamesPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRSavedGamesPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRSavedGamesPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRSavedGamesPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRSavedGamesPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRSavedGamesPath_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Searches"
					If((validStateProp $Setting FRSearches_Part State ) -and ($Setting.FRSearches_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Searches\Redirection settings for Searches"
						$tmp = ""
						Switch ($Setting.FRSearches_Part.Value)
						{
							"RedirectUncPath"	{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRSearches_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting FRSearchesPath_Part State ) -and ($Setting.FRSearchesPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Searches\Searches path"
						If($Setting.FRSearchesPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRSearchesPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRSearchesPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRSearchesPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRSearchesPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRSearchesPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRSearchesPath_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Start Menu"
					If((validStateProp $Setting FRStartMenu_Part State ) -and ($Setting.FRStartMenu_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Start Menu\Redirection settings for Start Menu"
						$tmp = ""
						Switch ($Setting.FRStartMenu_Part.Value)
						{
							"RedirectUncPath"	{$tmp = "Redirect to the following UNC path"; Break}
							Default 			{$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRStartMenu_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting FRStartMenuPath_Part State ) -and ($Setting.FRStartMenuPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Start Menu\Start Menu path"
						If($Setting.FRStartMenuPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRStartMenuPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRStartMenuPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRStartMenuPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRStartMenuPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRStartMenuPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRStartMenuPath_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Folder Redirection\Videos"
					If((validStateProp $Setting FRVideos_Part State ) -and ($Setting.FRVideos_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Videos\Redirection settings for Videos"
						$tmp = ""
						Switch ($Setting.FRVideos_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							"RedirectRelativeDocuments" {$tmp = "Redirect relative to Documents folder"; Break}
							Default 					{$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRVideos_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting FRVideosPath_Part State ) -and ($Setting.FRVideosPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Videos\Videos path"
						If($Setting.FRVideosPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRVideosPath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRVideosPath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRVideosPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.FRVideosPath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRVideosPath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.FRVideosPath_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Log settings"
					If((validStateProp $Setting LogLevel_ActiveDirectoryActions State ) -and ($Setting.LogLevel_ActiveDirectoryActions.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Active Directory actions"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogLevel_ActiveDirectoryActions.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_ActiveDirectoryActions.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_ActiveDirectoryActions.State
						}
					}
					If((validStateProp $Setting LogLevel_Information State ) -and ($Setting.LogLevel_Information.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Common information"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogLevel_Information.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_Information.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_Information.State
						}
					}
					If((validStateProp $Setting LogLevel_Warnings State ) -and ($Setting.LogLevel_Warnings.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Common warnings"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogLevel_Warnings.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_Warnings.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_Warnings.State
						}
					}
					If((validStateProp $Setting DebugMode State ) -and ($Setting.DebugMode.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Enable logging"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DebugMode.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DebugMode.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.DebugMode.State
						}
					}
					If((validStateProp $Setting LogLevel_FileSystemActions State ) -and ($Setting.LogLevel_FileSystemActions.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\File system actions"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogLevel_FileSystemActions.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_FileSystemActions.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_FileSystemActions.State
						}
					}
					If((validStateProp $Setting LogLevel_FileSystemNotification State ) -and ($Setting.LogLevel_FileSystemNotification.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\File system notifications"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogLevel_FileSystemNotification.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_FileSystemNotification.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_FileSystemNotification.State
						}
					}
					If((validStateProp $Setting LogLevel_Logoff State ) -and ($Setting.LogLevel_Logoff.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Logoff"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogLevel_Logoff.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_Logoff.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_Logoff.State
						}
					}
					If((validStateProp $Setting LogLevel_Logon State ) -and ($Setting.LogLevel_Logon.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Logon"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogLevel_Logon.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_Logon.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_Logon.State
						}
					}
					If((validStateProp $Setting MaxLogSize_Part State ) -and ($Setting.MaxLogSize_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Maximum size of the log file (bytes)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MaxLogSize_Part.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MaxLogSize_Part.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.MaxLogSize_Part.Value 
						}
					}
					If((validStateProp $Setting DebugFilePath_Part State ) -and ($Setting.DebugFilePath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Path to log file"
						If($Setting.DebugFilePath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.DebugFilePath_Part.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.DebugFilePath_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.DebugFilePath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.DebugFilePath_Part.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.DebugFilePath_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.DebugFilePath_Part.State
							}
						}
					}
					If((validStateProp $Setting LogLevel_UserName State ) -and ($Setting.LogLevel_UserName.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Personalized user information"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogLevel_UserName.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_UserName.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_UserName.State
						}
					}
					If((validStateProp $Setting LogLevel_PolicyUserLogon State ) -and ($Setting.LogLevel_PolicyUserLogon.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Policy values at logon and logoff"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogLevel_PolicyUserLogon.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_PolicyUserLogon.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_PolicyUserLogon.State
						}
					}
					If((validStateProp $Setting LogLevel_RegistryActions State ) -and ($Setting.LogLevel_RegistryActions.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Registry actions"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogLevel_RegistryActions.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_RegistryActions.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_RegistryActions.State
						}
					}
					If((validStateProp $Setting LogLevel_RegistryDifference State ) -and ($Setting.LogLevel_RegistryDifference.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Registry differences at logoff"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LogLevel_RegistryDifference.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_RegistryDifference.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_RegistryDifference.State
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Profile handling"
					If((validStateProp $Setting ApplicationProfilesAutoMigration State ) -and ($Setting.ApplicationProfilesAutoMigration.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Automatic migration of existing application profiles"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ApplicationProfilesAutoMigration.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ApplicationProfilesAutoMigration.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ApplicationProfilesAutoMigration.State
						}
					}
					If((validStateProp $Setting ProfileDeleteDelay_Part State ) -and ($Setting.ProfileDeleteDelay_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Delay before deleting cached profiles"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ProfileDeleteDelay_Part.Value;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProfileDeleteDelay_Part.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ProfileDeleteDelay_Part.Value 
						}
					}
					If((validStateProp $Setting DeleteCachedProfilesOnLogoff State ) -and ($Setting.DeleteCachedProfilesOnLogoff.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Delete locally cached profiles on logoff"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DeleteCachedProfilesOnLogoff.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DeleteCachedProfilesOnLogoff.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.DeleteCachedProfilesOnLogoff.State
						}
					}
					If((validStateProp $Setting LocalProfileConflictHandling_Part State ) -and ($Setting.LocalProfileConflictHandling_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Local profile conflict handling"
						$tmp = ""
						Switch ($Setting.LocalProfileConflictHandling_Part.Value)
						{
							"Use"		{$tmp = "Use local profile"; Break}
							"Delete"	{$tmp = "Delete local profile"; Break}
							"Rename"	{$tmp = "Rename local profile"; Break}
							Default		{$tmp = "Local profile conflict handling could not be determined: $($Setting.LocalProfileConflictHandling_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting MigrateWindowsProfilesToUserStore_Part State ) -and ($Setting.MigrateWindowsProfilesToUserStore_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Migration of existing profiles"
						$tmp = ""
						Switch ($Setting.MigrateWindowsProfilesToUserStore_Part.Value)
						{
							"All"		{$tmp = "Local and Roaming"; Break}
							"Local"		{$tmp = "Local"; Break}
							"Roaming"	{$tmp = "Roaming"; Break}
							"None"		{$tmp = "None"; Break}
							Default		{$tmp = "Migration of existing profiles could not be determined: $($Setting.MigrateWindowsProfilesToUserStore_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting TemplateProfilePath State ) -and ($Setting.TemplateProfilePath.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Path to the template profile"
						If($Setting.TemplateProfilePath.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.TemplateProfilePath.Value;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.TemplateProfilePath.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.TemplateProfilePath.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.TemplateProfilePath.State;
								}
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.TemplateProfilePath.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.TemplateProfilePath.State
							}
						}
					}
					If((validStateProp $Setting TemplateProfileOverridesLocalProfile State ) -and ($Setting.TemplateProfileOverridesLocalProfile.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Template profile overrides local profile"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.TemplateProfileOverridesLocalProfile.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TemplateProfileOverridesLocalProfile.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.TemplateProfileOverridesLocalProfile.State
						}
					}
					If((validStateProp $Setting TemplateProfileOverridesRoamingProfile State ) -and ($Setting.TemplateProfileOverridesRoamingProfile.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Template profile overrides roaming profile"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.TemplateProfileOverridesRoamingProfile.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TemplateProfileOverridesRoamingProfile.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.TemplateProfileOverridesRoamingProfile.State
						}
					}
					If((validStateProp $Setting TemplateProfileIsMandatory State ) -and ($Setting.TemplateProfileIsMandatory.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Template profile used as a Citrix mandatory profile for all logons"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.TemplateProfileIsMandatory.State;
							}
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TemplateProfileIsMandatory.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.TemplateProfileIsMandatory.State 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Registry"
					If((validStateProp $Setting ExclusionList_Part State ) -and ($Setting.ExclusionList_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\Exclusion list"
						If($Setting.ExclusionList_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting ExclusionList_Part Values )
							{
								$tmpArray = $Setting.ExclusionList_Part.Values
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = $txt;
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = "";
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Exclusion list were found"
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.ExclusionList_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ExclusionList_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.ExclusionList_Part.State
							}
						}
					}
					If((validStateProp $Setting IncludeListRegistry_Part State ) -and ($Setting.IncludeListRegistry_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\Inclusion list"
						If($Setting.IncludeListRegistry_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting IncludeListRegistry_Part Values )
							{
								$tmpArray = $Setting.IncludeListRegistry_Part.Values
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = $txt;
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = "";
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Inclusion list were found"
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.IncludeList_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.IncludeList_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.IncludeList_Part.State
							}
						}
					}
					If((validStateProp $Setting LastKnownGoodRegistry State ) -and ($Setting.LastKnownGoodRegistry.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\NTUSER.DAT backup"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LastKnownGoodRegistry.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LastKnownGoodRegistry.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.LastKnownGoodRegistry.State 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Registry\Default Exclusions"
					If((validStateProp $Setting DefaultExclusionList State ) -and ($Setting.DefaultExclusionList.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\Enable Default Exclusion list"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.DefaultExclusionList.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DefaultExclusionList.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.DefaultExclusionList.State 
						}
					}
					If((validStateProp $Setting ExclusionDefaultReg01 State ) -and ($Setting.ExclusionDefaultReg01.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\UPM - Software\Microsoft\AppV\Client\Integration"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultReg01.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultReg01.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultReg01.State 
						}
					}
					If((validStateProp $Setting ExclusionDefaultReg02 State ) -and ($Setting.ExclusionDefaultReg02.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\UPM - Software\Microsoft\AppV\Client\Publishing"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultReg02.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultReg02.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultReg02.State 
						}
					}
					If((validStateProp $Setting ExclusionDefaultReg03 State ) -and ($Setting.ExclusionDefaultReg03.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\UPM - Software\Microsoft\Speech_OneCore"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultReg03.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultReg03.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultReg03.State 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Streamed user profiles"
					If((validStateProp $Setting PSAlwaysCache State ) -and ($Setting.PSAlwaysCache.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Streamed user profiles\Always cache"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.PSAlwaysCache.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PSAlwaysCache.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.PSAlwaysCache.State 
						}
					}
					If((validStateProp $Setting PSAlwaysCache_Part State ) -and ($Setting.PSAlwaysCache_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Streamed user profiles\Always cache size (MB)"
						If($Setting.PSAlwaysCache_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.PSAlwaysCache_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.PSAlwaysCache_Part.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.PSAlwaysCache_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.PSAlwaysCache_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.PSAlwaysCache_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.PSAlwaysCache_Part.State 
							}
						}
					}
					If((validStateProp $Setting PSEnabled State ) -and ($Setting.PSEnabled.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Streamed user profiles\Profile streaming"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.PSEnabled.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PSEnabled.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.PSEnabled.State 
						}
					}
					If((validStateProp $Setting StreamingExclusionList_Part State ) -and ($Setting.StreamingExclusionList_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Streamed user profiles\Profile Streaming Exclusion list - directories"
						If($Setting.StreamingExclusionList_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting StreamingExclusionList_Part Values )
							{
								$tmpArray = $Setting.StreamingExclusionList_Part.Values.Split(",")
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = $txt;
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = "";
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "`t`t`t`t`t`t`t`t`t`t`t" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Profile Streaming Exclusion list - directories were found"
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.StreamingExclusionList_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.StreamingExclusionList_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.StreamingExclusionList_Part.State 
							}
						}
					}
					If((validStateProp $Setting PSUserGroups_Part State ) -and ($Setting.PSUserGroups_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Streamed user profiles\Streamed user profile groups"
						If($Setting.PSUserGroups_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting PSUserGroups_Part Values )
							{
								$tmpArray = $Setting.PSUserGroups_Part.Values.Split(",")
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = $txt;
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = "";
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No Streamed user profile groups were found"
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.PSUserGroups_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.PSUserGroups_Part.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.PSUserGroups_Part.State 
							}
						}
					}
					If((validStateProp $Setting PSPendingLockTimeout State ) -and ($Setting.PSPendingLockTimeout.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Streamed user profiles\Timeout for pending area lock files (days)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.PSPendingLockTimeout.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PSPendingLockTimeout.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.PSPendingLockTimeout.Value 
						}
					}
					Write-Verbose "$(Get-Date -Format G): `t`t`tProfile Management\Citrix Virtual Apps Optimization settings"
					If((validStateProp $Setting XenAppOptimizationEnable State ) -and ($Setting.XenAppOptimizationEnable.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Citrix Virtual Apps Optimization settings\Enable Citrix Virtual Apps Optimization"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.XenAppOptimizationEnable.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.XenAppOptimizationEnable.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.XenAppOptimizationEnable.State
						}
					}
					If((validStateProp $Setting XenAppOptimizationDefinitionPathData State ) -and ($Setting.XenAppOptimizationDefinitionPathData.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Citrix Virtual Apps Optimization settings\Path to Citrix Virtual Apps optimization definitions:"
						If($Setting.XenAppOptimizationDefinitionPathData.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.XenAppOptimizationDefinitionPathData.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.XenAppOptimizationDefinitionPathData.Value,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.XenAppOptimizationDefinitionPathData.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.XenAppOptimizationDefinitionPathData.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.XenAppOptimizationDefinitionPathData.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.XenAppOptimizationDefinitionPathData.State 
							}
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tReceiver"
					If((validStateProp $Setting StorefrontAccountsList State ) -and ($Setting.StorefrontAccountsList.State -ne "NotConfigured"))
					{
						$txt = "Receiver\Storefront accounts list"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = "";
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							"",$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt ""
						}
						$txt = ""
						If(validStateProp $Setting StorefrontAccountsList Values )
						{
							$cnt = 0
							$tmpArray = $Setting.StorefrontAccountsList.Values
							ForEach($Thing in $TmpArray)
							{
								$cnt++
								$xxx = """$($Thing)"""
								[array]$tmp = $xxx.Split(";").replace('"','')
								$tmp1 = "Name : $($tmp[0])"
								$tmp2 = "URL  : $($tmp[1])"
								$tmp3 = "State: $($tmp[2])"
								$tmp4 = "Desc : $($tmp[3])"
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp1;
									}
									$SettingsWordTable += $WordTableRowHash
									
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp2;
									}
									$SettingsWordTable += $WordTableRowHash
									
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp3;
									}
									$SettingsWordTable += $WordTableRowHash
									
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp4;
									}
									$SettingsWordTable += $WordTableRowHash
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp1,$htmlwhite))
									
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp2,$htmlwhite))
									
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp3,$htmlwhite))
									
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp4,$htmlwhite))
								}
								If($Text)
								{
									$txt = "`t`t`t`t "
									OutputPolicySetting $txt $tmp1
									OutputPolicySetting $txt $tmp2
									OutputPolicySetting $txt $tmp3
									OutputPolicySetting $txt $tmp4
								}
								
								If($cnt -gt 1)
								{
									$tmp = " "
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										"",$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
								$xxx = $Null
								$tmp = $Null
								$tmp1 = $Null
								$tmp2 = $Null
								$tmp3 = $Null
								$tmp4 = $Null
							}
							$TmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No Storefront accounts list were found"
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = "";
								Value = $tmp;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							If($HTML)
							{
								$rowdata += @(,(
								"",$htmlbold,
								"",$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting "" $tmp
							}
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tUser Personalization Layer"
					If((validStateProp $Setting UplRepositoryPath State ) -and ($Setting.UplRepositoryPath.State -ne "NotConfigured"))
					{
						$txt = "User Personalization Layer\User Layer Repository Path"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.UplRepositoryPath.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UplRepositoryPath.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.UplRepositoryPath.Value 
						}
					}
					If((validStateProp $Setting UplUserLayerSizeInGb State ) -and ($Setting.UplUserLayerSizeInGb.State -ne "NotConfigured"))
					{
						$txt = "User Personalization Layer\User Layer Size in GB"
						$UPLSize = 0
						If($Setting.UplUserLayerSizeInGb.Value -eq 0)
						{
							$UPLSize = 10
						}
						Else
						{
							$UPLSize = $Setting.UplUserLayerSizeInGb.Value
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $UPLSize;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$UPLSize,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $UPLSize 
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`tVirtual Delivery Agent Settings"
					If((validStateProp $Setting ControllerRegistrationIPv6Netmask State ) -and ($Setting.ControllerRegistrationIPv6Netmask.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Controller registration IPv6 netmask"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ControllerRegistrationIPv6Netmask.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ControllerRegistrationIPv6Netmask.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ControllerRegistrationIPv6Netmask.Value 
						}
					}
					If((validStateProp $Setting ControllerRegistrationPort State ) -and ($Setting.ControllerRegistrationPort.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Controller registration port"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ControllerRegistrationPort.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ControllerRegistrationPort.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ControllerRegistrationPort.Value 
						}
					}
					If((validStateProp $Setting ControllerSIDs State ) -and ($Setting.ControllerSIDs.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Controller SIDs"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ControllerSIDs.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ControllerSIDs.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.ControllerSIDs.Value 
						}
					}
					If((validStateProp $Setting Controllers State ) -and ($Setting.Controllers.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Controllers"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.Controllers.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.Controllers.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.Controllers.Value 
						}
					}
					If((validStateProp $Setting EnableAutoUpdateOfControllers State ) -and ($Setting.EnableAutoUpdateOfControllers.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\Enable auto update of Controllers"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.EnableAutoUpdateOfControllers.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnableAutoUpdateOfControllers.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.EnableAutoUpdateOfControllers.State 
						}
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`t`tVirtual Delivery Agent Settings\Monitoring"
					If((validStateProp $Setting SelectedFailureLevel State ) -and ($Setting.SelectedFailureLevel.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\Monitoring\Enable monitoring of application failures"
						$tmp = ""
						Switch ($Setting.SelectedFailureLevel.Value)
						{
							"None"	{$tmp = "None"; Break}
							"Both"	{$tmp = "Both application errors and faults"; Break}
							"Fault"	{$tmp = "Application faults only"; Break}
							"Error"	{$tmp = "Application errors only"; Break}
							Default	{$tmp = "Enable monitoring of application failures could not be determined: $($Setting.SelectedFailureLevel.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $tmp
						}
					}
					If((validStateProp $Setting EnableWorkstationVDAFaultMonitoring State ) -and ($Setting.EnableWorkstationVDAFaultMonitoring.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\Monitoring\Enable monitoring of application failures on Desktop OS VDAs"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.EnableWorkstationVDAFaultMonitoring.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnableWorkstationVDAFaultMonitoring.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.EnableWorkstationVDAFaultMonitoring.State 
						}
					}
					If((validStateProp $Setting EnableProcessMonitoring State ) -and ($Setting.EnableProcessMonitoring.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\Monitoring\Enable process monitoring"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.EnableProcessMonitoring.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnableProcessMonitoring.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.EnableProcessMonitoring.State 
						}
					}
					If((validStateProp $Setting EnableResourceMonitoring State ) -and ($Setting.EnableResourceMonitoring.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\Monitoring\Enable resource monitoring"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.EnableResourceMonitoring.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnableResourceMonitoring.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.EnableResourceMonitoring.State 
						}
					}
					If((validStateProp $Setting AppFailureExclusionList State ) -and ($Setting.AppFailureExclusionList.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\Monitoring\List of applications excluded from failure monitoring"
						If($Setting.StreamingExclusionList_Part.State -eq "Enabled")
						{
							If(validStateProp $Setting AppFailureExclusionList Values )
							{
								$tmpArray = $Setting.AppFailureExclusionList.Values.Split(",")
								$tmp = ""
								$cnt = 0
								ForEach($Thing in $tmpArray)
								{
									$cnt++
									$tmp = "$($Thing)"
									If($cnt -eq 1)
									{
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = $txt;
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										If($HTML)
										{
											$rowdata += @(,(
											$txt,$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting $txt $tmp
										}
									}
									Else
									{
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = "";
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										If($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										If($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
								}
								$tmpArray = $Null
								$tmp = $Null
							}
							Else
							{
								$tmp = "No List of applications excluded from failure monitoring were found"
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								If($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								If($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.AppFailureExclusionList.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.AppFailureExclusionList.State,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $Setting.AppFailureExclusionList.State 
							}
						}
					}	
					If((validStateProp $Setting OnlyUseIPv6ControllerRegistration State ) -and ($Setting.OnlyUseIPv6ControllerRegistration.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Only use IPv6 Controller registration"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.OnlyUseIPv6ControllerRegistration.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OnlyUseIPv6ControllerRegistration.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.OnlyUseIPv6ControllerRegistration.State 
						}
					}
					If((validStateProp $Setting SiteGUID State ) -and ($Setting.SiteGUID.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Site GUID"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SiteGUID.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SiteGUID.Value,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.SiteGUID.Value 
						}
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`t`tVirtual IP"
					If((validStateProp $Setting VirtualLoopbackSupport State ) -and ($Setting.VirtualLoopbackSupport.State -ne "NotConfigured"))
					{
						$txt = "Virtual IP\Virtual IP loopback support"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.VirtualLoopbackSupport.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						If($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.VirtualLoopbackSupport.State,$htmlwhite))
						}
						If($Text)
						{
							OutputPolicySetting $txt $Setting.VirtualLoopbackSupport.State 
						}
					}
					If((validStateProp $Setting VirtualLoopbackPrograms State ) -and ($Setting.VirtualLoopbackPrograms.State -ne "NotConfigured"))
					{
						$txt = "Virtual IP\Virtual IP virtual loopback programs list"
						If((validStateProp $Setting VirtualLoopbackPrograms State ) -and ($Setting.VirtualLoopbackPrograms.State -ne "NotConfigured"))
						{
							$tmpArray = $Setting.VirtualLoopbackPrograms.Values
							$array = $Null
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $TmpArray)
							{
								If($Null -eq $Thing)
								{
									$Thing = ''
								}
								$cnt++
								$tmp = "$($Thing) "
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									If($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									If($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									If($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$TmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No Virtual IP virtual loopback programs list were found"
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $tmp;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							If($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							If($Text)
							{
								OutputPolicySetting $txt $tmp
							}
						}
					}
				}
				If($MSWord -or $PDF)
				{
					If($SettingsWordTable.Count -gt 0) #don't process if array is empty
					{
						$Table = AddWordTable -Hashtable $SettingsWordTable `
						-Columns  Text,Value `
						-Headers  "Setting Key","Value"`
						-Format $wdTableLightListAccent3 `
						-NoInternalGridLines `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
						SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Columns.Item(1).Width = 300;
						$Table.Columns.Item(2).Width = 200;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
					}
					Else
					{
						WriteWordLine 0 1 "There are no policy settings"
					}
					FindWordDocumentEnd
					$Table = $Null
				}
				If($Text)
				{
					Line 0 ""
				}
				If($HTML)
				{
					If($rowdata.count -gt 0)
					{
						$columnHeaders = @(
						'Setting Key',($global:htmlsb),
						'Value',($global:htmlsb))

						$msg = ""
						$columnWidths = @("400","300")
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "700"
					}
				}
			}
			Else
			{
				$txt = "Unable to retrieve settings"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 1 $txt
				}
				If($Text)
				{
					Line 2 $txt
				}
				If($HTML)
				{
					WriteHTMLLine 0 1 $txt
				}
			}
			$Filter = $Null
			$Settings = $Null
			Write-Verbose "$(Get-Date -Format G): `t`tFinished $($Policy.PolicyName)"
			Write-Verbose "$(Get-Date -Format G): "
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "Citrix Policy information could not be retrieved"
	}
	Else
	{
		Write-Warning "No results Returned for Citrix Policy information"
	}
	
	$CtxPolicies = $Null
	Write-Verbose "$(Get-Date -Format G): `tRemoving $($xDriveName) PSDrive"
	Remove-PSDrive -Name $xDriveName -EA 0 4>$Null
	Write-Verbose "$(Get-Date -Format G): "
}

Function OutputPolicySetting
{
	Param([string] $outputText, [string] $outputData)

	If($outputText -ne "")
	{
		$xLength = $outputText.Length
		If($outputText.Substring($xLength-2,2) -ne ": ")
		{
			$outputText += ": "
		}
	}
	Line 2 $outputText $outputData
}

Function Get-PrinterModifiedSettings
{
	Param([string]$Value, [string]$xelement)
	
	[string]$ReturnStr = ""

	Switch ($Value)
	{
		"copi" 
		{
			$txt="Copies: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"coll"
		{
			$txt="Collate: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"scal"
		{
			$txt="Scale (%): "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"colo"
		{
			$txt="Color: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1 {$tmp2 = "Monochrome"; Break}
					2 {$tmp2 = "Color"; Break}
					Default {$tmp2 = "Color could not be determined: $($xelement) "; Break}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"prin"
		{
			$txt="Print Quality: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					-1 {$tmp2 = "150 dpi"; Break}
					-2 {$tmp2 = "300 dpi"; Break}
					-3 {$tmp2 = "600 dpi"; Break}
					-4 {$tmp2 = "1200 dpi"; Break}
					Default {$tmp2 = "Custom...X resolution: $tmp1"; Break}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"yres"
		{
			$txt="Y resolution: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"orie"
		{
			$txt="Orientation: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					"portrait"  {$tmp2 = "Portrait"; Break}
					"landscape" {$tmp2 = "Landscape"; Break}
					Default {$tmp2 = "Orientation could not be determined: $($xelement) ; Break"}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"dupl"
		{
			$txt="Duplex: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1 {$tmp2 = "Simplex"; Break}
					2 {$tmp2 = "Vertical"; Break}
					3 {$tmp2 = "Horizontal"; Break}
					Default {$tmp2 = "Duplex could not be determined: $($xelement) "; Break}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"pape"
		{
			$txt="Paper Size: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1   {$tmp2 = "Letter"; Break}
					2   {$tmp2 = "Letter Small"; Break}
					3   {$tmp2 = "Tabloid"; Break}
					4   {$tmp2 = "Ledger"; Break}
					5   {$tmp2 = "Legal"; Break}
					6   {$tmp2 = "Statement"; Break}
					7   {$tmp2 = "Executive"; Break}
					8   {$tmp2 = "A3"; Break}
					9   {$tmp2 = "A4"; Break}
					10  {$tmp2 = "A4 Small"; Break}
					11  {$tmp2 = "A5"; Break}
					12  {$tmp2 = "B4 (JIS)"; Break}
					13  {$tmp2 = "B5 (JIS)"; Break}
					14  {$tmp2 = "Folio"; Break}
					15  {$tmp2 = "Quarto"; Break}
					16  {$tmp2 = "10X14"; Break}
					17  {$tmp2 = "11X17"; Break}
					18  {$tmp2 = "Note"; Break}
					19  {$tmp2 = "Envelope #9"; Break}
					20  {$tmp2 = "Envelope #10"; Break}
					21  {$tmp2 = "Envelope #11"; Break}
					22  {$tmp2 = "Envelope #12"; Break}
					23  {$tmp2 = "Envelope #14"; Break}
					24  {$tmp2 = "C Size Sheet"; Break}
					25  {$tmp2 = "D Size Sheet"; Break}
					26  {$tmp2 = "E Size Sheet"; Break}
					27  {$tmp2 = "Envelope DL"; Break}
					28  {$tmp2 = "Envelope C5"; Break}
					29  {$tmp2 = "Envelope C3"; Break}
					30  {$tmp2 = "Envelope C4"; Break}
					31  {$tmp2 = "Envelope C6"; Break}
					32  {$tmp2 = "Envelope C65"; Break}
					33  {$tmp2 = "Envelope B4"; Break}
					34  {$tmp2 = "Envelope B5"; Break}
					35  {$tmp2 = "Envelope B6"; Break}
					36  {$tmp2 = "Envelope Italy"; Break}
					37  {$tmp2 = "Envelope Monarch"; Break}
					38  {$tmp2 = "Envelope Personal"; Break}
					39  {$tmp2 = "US Std Fanfold"; Break}
					40  {$tmp2 = "German Std Fanfold"; Break}
					41  {$tmp2 = "German Legal Fanfold"; Break}
					42  {$tmp2 = "B4 (ISO)"; Break}
					43  {$tmp2 = "Japanese Postcard"; Break}
					44  {$tmp2 = "9X11"; Break}
					45  {$tmp2 = "10X11"; Break}
					46  {$tmp2 = "15X11"; Break}
					47  {$tmp2 = "Envelope Invite"; Break}
					48  {$tmp2 = "Reserved - DO NOT USE"; Break}
					49  {$tmp2 = "Reserved - DO NOT USE"; Break}
					50  {$tmp2 = "Letter Extra"; Break}
					51  {$tmp2 = "Legal Extra"; Break}
					52  {$tmp2 = "Tabloid Extra"; Break}
					53  {$tmp2 = "A4 Extra"; Break}
					54  {$tmp2 = "Letter Transverse"; Break}
					55  {$tmp2 = "A4 Transverse"; Break}
					56  {$tmp2 = "Letter Extra Transverse"; Break}
					57  {$tmp2 = "A Plus"; Break}
					58  {$tmp2 = "B Plus"; Break}
					59  {$tmp2 = "Letter Plus"; Break}
					60  {$tmp2 = "A4 Plus"; Break}
					61  {$tmp2 = "A5 Transverse"; Break}
					62  {$tmp2 = "B5 (JIS) Transverse"; Break}
					63  {$tmp2 = "A3 Extra"; Break}
					64  {$tmp2 = "A5 Extra"; Break}
					65  {$tmp2 = "B5 (ISO) Extra"; Break}
					66  {$tmp2 = "A2"; Break}
					67  {$tmp2 = "A3 Transverse"; Break}
					68  {$tmp2 = "A3 Extra Transverse"; Break}
					69  {$tmp2 = "Japanese Double Postcard"; Break}
					70  {$tmp2 = "A6"; Break}
					71  {$tmp2 = "Japanese Envelope Kaku #2"; Break}
					72  {$tmp2 = "Japanese Envelope Kaku #3"; Break}
					73  {$tmp2 = "Japanese Envelope Chou #3"; Break}
					74  {$tmp2 = "Japanese Envelope Chou #4"; Break}
					75  {$tmp2 = "Letter Rotated"; Break}
					76  {$tmp2 = "A3 Rotated"; Break}
					77  {$tmp2 = "A4 Rotated"; Break}
					78  {$tmp2 = "A5 Rotated"; Break}
					79  {$tmp2 = "B4 (JIS) Rotated"; Break}
					80  {$tmp2 = "B5 (JIS) Rotated"; Break}
					81  {$tmp2 = "Japanese Postcard Rotated"; Break}
					82  {$tmp2 = "Double Japanese Postcard Rotated"; Break}
					83  {$tmp2 = "A6 Rotated"; Break}
					84  {$tmp2 = "Japanese Envelope Kaku #2 Rotated"; Break}
					85  {$tmp2 = "Japanese Envelope Kaku #3 Rotated"; Break}
					86  {$tmp2 = "Japanese Envelope Chou #3 Rotated"; Break}
					87  {$tmp2 = "Japanese Envelope Chou #4 Rotated"; Break}
					88  {$tmp2 = "B6 (JIS)"; Break}
					89  {$tmp2 = "B6 (JIS) Rotated"; Break}
					90  {$tmp2 = "12X11"; Break}
					91  {$tmp2 = "Japanese Envelope You #4"; Break}
					92  {$tmp2 = "Japanese Envelope You #4 Rotated"; Break}
					93  {$tmp2 = "PRC 16K"; Break}
					94  {$tmp2 = "PRC 32K"; Break}
					95  {$tmp2 = "PRC 32K(Big)"; Break}
					96  {$tmp2 = "PRC Envelope #1"; Break}
					97  {$tmp2 = "PRC Envelope #2"; Break}
					98  {$tmp2 = "PRC Envelope #3"; Break}
					99  {$tmp2 = "PRC Envelope #4"; Break}
					100 {$tmp2 = "PRC Envelope #5"; Break}
					101 {$tmp2 = "PRC Envelope #6"; Break}
					102 {$tmp2 = "PRC Envelope #7"; Break}
					103 {$tmp2 = "PRC Envelope #8"; Break}
					104 {$tmp2 = "PRC Envelope #9"; Break}
					105 {$tmp2 = "PRC Envelope #10"; Break}
					106 {$tmp2 = "PRC 16K Rotated"; Break}
					107 {$tmp2 = "PRC 32K Rotated"; Break}
					108 {$tmp2 = "PRC 32K(Big) Rotated"; Break}
					109 {$tmp2 = "PRC Envelope #1 Rotated"; Break}
					110 {$tmp2 = "PRC Envelope #2 Rotated"; Break}
					111 {$tmp2 = "PRC Envelope #3 Rotated"; Break}
					112 {$tmp2 = "PRC Envelope #4 Rotated"; Break}
					113 {$tmp2 = "PRC Envelope #5 Rotated"; Break}
					114 {$tmp2 = "PRC Envelope #6 Rotated"; Break}
					115 {$tmp2 = "PRC Envelope #7 Rotated"; Break}
					116 {$tmp2 = "PRC Envelope #8 Rotated"; Break}
					117 {$tmp2 = "PRC Envelope #9 Rotated"; Break}
					Default {$tmp2 = "Paper Size could not be determined: $($xelement) "; Break}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"form"
		{
			$txt="Form Name: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				If($tmp2.length -gt 0)
				{
					$ReturnStr = "$txt $tmp2"
				}
			}
		}
		"true"
		{
			$txt="TrueType: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1 {$tmp2 = "Bitmap"; Break}
					2 {$tmp2 = "Download"; Break}
					3 {$tmp2 = "Substitute"; Break}
					4 {$tmp2 = "Outline"; Break}
					Default {$tmp2 = "TrueType could not be determined: $($xelement) "; Break}
				}
			}
			$ReturnStr = "$txt $tmp2"
		}
		"mode" 
		{
			$txt="Printer Model: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"loca" 
		{
			$txt="Location: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				If($tmp2.length -gt 0)
				{
					$ReturnStr = "$txt $tmp2"
				}
			}
		}
		Default {$ReturnStr = "Session printer setting could not be determined: $($xelement) "}
	}
	Return $ReturnStr
}

Function GetCtxGPOsInAD
{
	#thanks to the Citrix Engineering Team for pointers and for Michael B. Smith for creating the function
	#updated 07-Nov-13 to work in a Windows Workgroup environment
	Write-Verbose "$(Get-Date -Format G): Testing for an Active Directory environment"
	$root = [ADSI]"LDAP://RootDSE"
	If([String]::IsNullOrEmpty($root.PSBase.Name))
	{
		Write-Verbose "$(Get-Date -Format G): `tNot in an Active Directory environment"
		$root = $Null
		$xArray = @()
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): `tIn an Active Directory environment"
		$domainNC = $root.defaultNamingContext.ToString()
		$root = $Null
		$xArray = @()

		$domain = $domainNC.Replace( 'DC=', '' ).Replace( ',', '.' )
		Write-Verbose "$(Get-Date -Format G): `tSearching \\$($domain)\sysvol\$($domain)\Policies"
		$sysvolFiles = @()
		$sysvolFiles = Get-ChildItem -Recurse ( '\\' + $domain  + '\sysvol\' + $domain + '\Policies' ) -EA 0
		If($sysvolFiles.Count -eq 0)
		{
			Write-Verbose "$(Get-Date -Format G): `tSearch timed out. Retrying. Searching \\ + $($domain)\sysvol\$($domain)\Policies a second time."
			$sysvolFiles = Get-ChildItem -Recurse ( '\\' + $domain  + '\sysvol\' + $domain + '\Policies' ) -EA 0
		}
		ForEach( $file in $sysvolFiles )
		{
			If( -not $file.PSIsContainer )
			{
				#$file.FullName  ### name of the policy file
				If( $file.FullName -like "*\Citrix\GroupPolicy\Policies.gpf" )
				{
					#"have match " + $file.FullName ### name of the Citrix policies file
					$array = $file.FullName.Split( '\' )
					If( $array.Length -gt 7 )
					{
						$gp = $array[ 6 ].ToString()
						$gpObject = [ADSI]( "LDAP://" + "CN=" + $gp + ",CN=Policies,CN=System," + $domainNC )
						If(!$xArray.Contains($gpObject.DisplayName))
						{
							$xArray += $gpObject.DisplayName	### name of the group policy object
						}
					}
				}
			}
		}
	}
	Return ,$xArray
}
#endregion

#region configuration Logging functions
Function ProcessConfigLogging
{
	#do not show config logging if not Details AND
	# if Virtual Desktops must be Platinum or Enterprise OR
	# if Virtual Apps must be Platinum or Enterprise
	# all Citrix Cloud license editions are eligible

	If($Logging)
	{
		Write-Verbose "$(Get-Date -Format G): Processing Configuration Logging"
		$txt1 = "Logging"
		If($MSword -or $PDF)
		{
			$Selection.InsertNewPage()
			WriteWordLine 1 0 $txt1
		}
		If($Text)
		{
			Line 0 $txt1
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 1 0 $txt1
		}
		
		If(
		($Script:CCSite2.ProductCode -eq "XDT" -and ($Script:CCSite2.ProductEdition -eq "PLT" -or $Script:CCSite2.ProductEdition -eq "ENT")) -or 
		(
		$Script:CCSite2.ProductCode -eq "CVADS" -or 
		$Script:CCSite2.ProductCode -eq "VADS" -or 
		$Script:CCSite2.ProductCode -eq "VAS" -or 
		$Script:CCSite2.ProductCode -eq "VAD")
		)
		{
			Write-Verbose "$(Get-Date -Format G): `tConfiguration Logging Details"
			$ConfigLogItems = Get-LogHighLevelOperation @CCParams2 -Filter {StartTime -ge $StartDate -and EndTime -le $EndDate} -SortBy "-StartTime"
			If($? -and $Null -ne $ConfigLogItems)
			{
				OutputConfigLog $ConfigLogItems
			}
			ElseIf($? -and ($Null -eq $ConfigLogItems))
			{
				$txt = "There are no Configuration Logging actions recorded for $($StartDate) through $($EndDate)."
				OutputNotice $txt
			}
			Else
			{
				$txt = "Configuration Logging information could not be retrieved."
				OutputWarning $txt
			}
			Write-Verbose "$(Get-Date -Format G): "
		}
		Else
		{
			$txt = "Not licensed for Configuration Logging"
			OutputNotice $txt
			Write-Verbose "$(Get-Date -Format G): "
		}
	}
}

Function OutputConfigLog
{
	Param([object] $ConfigLogItems)
	
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Configuration Logging Details"
	$txt2 = " For date range $($StartDate) through $($EndDate)"
	If($MSword -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 0 0 $txt2
	}
	If($Text)
	{
		Line 0 $txt2
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 0 0 $txt2
	}
	
	If($MSWord -or $PDF)
	{
		$ItemsWordTable = @()
	}
	If($HTML)
	{
		$rowdata = @()
	}
	
	ForEach($Item in $ConfigLogItems)
	{
		$Tmp = $Null
		If($Item.IsSuccessful)
		{
			$Tmp = "Success"
		}
		Else
		{
			$Tmp = "Failed"
		}
		
		If($MSWord -or $PDF)
		{
			$ItemsWordTable += @{ 
			Administrator = $Item.User;
			MainTask = $Item.Text;
			Start = $Item.StartTime;
			End = $Item.EndTime;
			Status = $tmp;
			}
		}
		If($Text)
		{
			Line 1 "Administrator`t: " $Item.User
			Line 1 "Main task`t: " $Item.Text
			Line 1 "Start`t`t: " $Item.StartTime
			Line 1 "End`t`t: " $Item.EndTime
			Line 1 "Status`t`t: " $Tmp
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,(
			$Item.User,$htmlwhite,
			$Item.Text,$htmlwhite,
			$Item.StartTime,$htmlwhite,
			$Item.EndTime,$htmlwhite,
			$Tmp,$htmlwhite
			))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ItemsWordTable `
		-Columns Administrator, MainTask, Start, End, Status `
		-Headers  "Administrator", "Main task", "Start", "End", "Status" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 120;
		$Table.Columns.Item(2).Width = 210;
		$Table.Columns.Item(3).Width = 60;
		$Table.Columns.Item(4).Width = 60;
		$Table.Columns.Item(5).Width = 50;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($HTML)
	{
		$columnHeaders = @(
		'Administrator',($global:htmlsb),
		'Main task',($global:htmlsb),
		'Start',($global:htmlsb),
		'End',($global:htmlsb),
		'Status',($global:htmlsb))

		$msg = ""
		$columnWidths = @("125","275","125","125","50")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "700"
	}
}
#endregion

#region site configuration functions
Function ProcessConfiguration
{
	Write-Verbose "$(Get-Date -Format G): Process Configuration Settings"
	OutputSiteSettings
	Write-Verbose "$(Get-Date -Format G): "
}

Function OutputSiteSettings
{
	Switch ($Script:CCSite1.ColorDepth)
	{
		"FourBit"		{$xColorDepth = "4bit - 16 colors"; Break}
		"EightBit"		{$xColorDepth = "8bit - 256 colors"; Break}
		"SixteenBit"	{$xColorDepth = "16bit - High color"; Break}
		"TwentyFourBit"	{$xColorDepth = "24bit - True color"; Break}
		Default			{$xColorDepth = "Unable to determine Color Depth: $($Script:CCSite1.ColorDepth)"; Break}
	}
	Switch ($Script:CCSite1.DefaultMinimumFunctionalLevel)
	{
		"L5" 	{$xVDAVersion = "5.6 FP1 (Windows XP and Windows Vista)"; Break}
		"L7"	{$xVDAVersion = "7.0 (or newer)"; Break}
		"L7_6"	{$xVDAVersion = "7.6 (or newer)"; Break}
		"L7_7"	{$xVDAVersion = "7.7 (or newer)"; Break}
		"L7_8"	{$xVDAVersion = "7.8 (or newer)"; Break}
		"L7_9"	{$xVDAVersion = "7.9 (or newer)"; Break}
		"L7_20"	{$xVDAVersion = "1811 (or newer)"; Break}
		"L7_25"	{$xVDAVersion = "2003 (or newer)"; Break}
		Default {$xVDAVersion = "Unable to determine VDA version: $($Script:CCSite1.DefaultMinimumFunctionalLevel)"; Break}
	}

	Write-Verbose "$(Get-Date -Format G): `tOutput Site Settings"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Configuration"
		WriteWordLine 2 0 "Site Settings"
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Site name"; Value = $CCSiteName; }) > $Null
		$ScriptInformation.Add(@{Data = "Base OU"; Value = $Script:CCSite1.BaseOU; }) > $Null
		$ScriptInformation.Add(@{Data = "Color Depth"; Value = $xColorDepth; }) > $Null
		$ScriptInformation.Add(@{Data = "Default Minimum Functional Level"; Value = $xVDAVersion; }) > $Null
		$ScriptInformation.Add(@{Data = "DNS Resolution Enabled"; Value = $Script:CCSite1.DnsResolutionEnabled.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Is Secondary Broker"; Value = $Script:CCSite1.IsSecondaryBroker.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Local Host Cache Enabled"; Value = $Script:CCSite1.LocalHostCacheEnabled.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Reuse Machines Without Shutdown in Outage Allowed"; Value = $Script:CCSite1.ReuseMachinesWithoutShutdownInOutageAllowed.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Secure ICA Required"; Value = $Script:CCSite1.SecureIcaRequired.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Trust Managed Anonymous XML Service Requests"; Value = $Script:CCSite1.TrustManagedAnonymousXmlServiceRequests.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Trust Requests Sent to the XML Service Port"; Value = $Script:CCSite1.TrustRequestsSentToTheXmlServicePort.ToString(); }) > $Null
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "Configuration"
		Line 0 ""
		Line 0 "Site Settings"
		Line 0 ""
		Line 1 "Site name`t`t`t`t`t`t: " $CCSiteName
		Line 1 "Base OU`t`t`t`t`t`t`t: " $Script:CCSite1.BaseOU
		Line 1 "Color Depth`t`t`t`t`t`t: " $xColorDepth
		Line 1 "Default Minimum Functional Level`t`t`t: " $xVDAVersion
		Line 1 "DNS Resolution Enabled`t`t`t`t`t: " $Script:CCSite1.DnsResolutionEnabled.ToString()
		Line 1 "Is Secondary Broker`t`t`t`t`t: " $Script:CCSite1.IsSecondaryBroker.ToString()
		Line 1 "Local Host Cache Enabled`t`t`t`t: " $Script:CCSite1.LocalHostCacheEnabled.ToString()
		Line 1 "Reuse Machines Without Shutdown in Outage Allowed`t: " $Script:CCSite1.ReuseMachinesWithoutShutdownInOutageAllowed.ToString()
		Line 1 "Secure ICA Required`t`t`t`t`t: " $Script:CCSite1.SecureIcaRequired.ToString()
		Line 1 "Trust Managed Anonymous XML Service Requests`t`t: " $Script:CCSite1.TrustManagedAnonymousXmlServiceRequests.ToString()
		Line 1 "Trust Requests Sent to the XML Service Port`t`t: " $Script:CCSite1.TrustRequestsSentToTheXmlServicePort.ToString()
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Configuration"
		WriteHTMLLine 2 0 "Site Settings"
		$rowdata = @()
		$columnHeaders = @("Site name",($global:htmlsb),$CCSiteName,$htmlwhite)
		$rowdata += @(,("Base OU",($global:htmlsb),$Script:CCSite1.BaseOU,$htmlwhite))
		$rowdata += @(,("Color Depth",($global:htmlsb),$xColorDepth,$htmlwhite))
		$rowdata += @(,("Default Minimum Functional Level",($global:htmlsb),$xVDAVersion,$htmlwhite))
		$rowdata += @(,("DNS Resolution Enabled",($global:htmlsb),$Script:CCSite1.DnsResolutionEnabled.ToString(),$htmlwhite))
		$rowdata += @(,("Is Secondary Broker",($global:htmlsb),$Script:CCSite1.IsSecondaryBroker.ToString(),$htmlwhite))
		$rowdata += @(,("Local Host Cache Enabled",($global:htmlsb),$Script:CCSite1.LocalHostCacheEnabled.ToString(),$htmlwhite))
		$rowdata += @(,("Reuse Machines Without Shutdown in Outage Allowed",($global:htmlsb),$Script:CCSite1.ReuseMachinesWithoutShutdownInOutageAllowed.ToString(),$htmlwhite))
		$rowdata += @(,("Secure ICA Required",($global:htmlsb),$Script:CCSite1.SecureIcaRequired.ToString(),$htmlwhite))
		$rowdata += @(,("Trust Managed Anonymous XML Service Requests",($global:htmlsb),$Script:CCSite1.TrustManagedAnonymousXmlServiceRequests.ToString(),$htmlwhite))
		$rowdata += @(,("Trust Requests Sent to the XML Service Port",($global:htmlsb),$Script:CCSite1.TrustRequestsSentToTheXmlServicePort.ToString(),$htmlwhite))
		
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
}

#region Administrator, Scope and Roles functions
Function ProcessAdministrators
{
	Write-Verbose "$(Get-Date -Format G): Processing Administrators"
	Write-Verbose "$(Get-Date -Format G): `tRetrieving Administrator data"
	
	$Admins = Get-AdminAdministrator @CCParams2 | `
	Where-Object {$_.UserIdentityType -ne "Sid" -and (-not [String]::IsNullOrEmpty($_.Name))}

	If($? -and ($Null -ne $Admins))
	{
		OutputAdministrators $Admins
	}
	ElseIf($? -and ($Null -eq $Admins))
	{
		$txt = "There are no Administrators"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve Administrators"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date -Format G): "
}

Function OutputAdministrators
{
	Param([object] $Admins)

	#fix for when $Admin.Rights.ScopeName and $Admin.Rights.RoleName are arrays
	Write-Verbose "$(Get-Date -Format G): `tOutput Administrator data"
	
	ForEach($Admin in $Admins)
	{
		Switch ($Admin.Rights.RoleName)
		{
			"Cloud Administrator"			{$Script:TotalCloudAdmins++}
			"Delivery Group Administrator"	{$Script:TotalDeliveryGroupAdmins++}
			"Full Administrator"			{$Script:TotalFullAdmins++}
			"Full Monitor Administrator"	{$Script:TotalFullMonitorAdmins++}
			"Help Desk Administrator"		{$Script:TotalHelpDeskAdmins++}
			"Host Administrator"			{$Script:TotalHostAdmins++}
			"Machine Catalog Administrator"	{$Script:TotalMachineCatalogAdmins++}
			"Probe Agent Administrator"		{$Script:TotalProbeAdmins++}
			"Read Only Administrator"		{$Script:TotalReadOnlyAdmins++}
			"Session Administrator"			{$Script:TotalSessionAdmins++}
			Default							{$Script:TotalCustomAdmins++}
		}
	}
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Administrators"
		$AdminsWordTable = @()
	}
	If($Text)
	{
		Line 0 "Administrators"
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Administrators"
		$rowdata = @()
	}
	
	If($MSWord -or $PDF -or $HTML)
	{
		ForEach($Admin in $Admins)
		{
			$Tmp = $Null
			$xScopeName = ""
			$xRoleName = ""

			If($Admin.Enabled)
			{
				$Tmp = "Enabled"
			}
			Else
			{
				$Tmp = "Disabled"
			}

			If($Admin.Rights.ScopeName -is [array])
			{
				$cnt = 0
				ForEach($xScope in $Admin.Rights.ScopeName)
				{
					$cnt++
					If($cnt -lt $Admin.Rights.ScopeName.Count)
					{
						$xScopeName += "$xScope; "
					}
					Else
					{
						$xScopeName += "$xScope"
					}
				}
			}
			Else
			{
				$xScopeName = $Admin.Rights.ScopeName
			}

			If($Admin.Rights.RoleName -is [array])
			{
				$cnt = 0
				ForEach($xRole in $Admin.Rights.RoleName)
				{
					$cnt++
					If($cnt -lt $Admin.Rights.RoleName.Count)
					{
						$xRoleName += "$xRole; "
					}
					Else
					{
						$xRoleName += "$xRole"
					}
				}
			}
			Else
			{
				$xRoleName = $Admin.Rights.RoleName
			}			

			If($MSWord -or $PDF)
			{
				$AdminsWordTable += @{
				Name = $Admin.Name; 
				Scope = $xScopeName; 
				Role = $xRoleName; 
				Status = $Tmp;
				}
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Admin.Name,$htmlwhite,
				$xScopeName,$htmlwhite,
				$xRoleName,$htmlwhite,
				$Tmp,$htmlwhite))
			}
		}
		
		If($MSWord -or $PDF)
		{
			If($AdminsWordTable.Count -eq 0)
			{
				$AdminsWordTable += @{ 
				Name = "No admins found";
				Scope = "N/A"
				Role = "N/A";
				Status = "N/A";
				}
			}

			$Table = AddWordTable -Hashtable $AdminsWordTable `
			-Columns Name, Scope, Role, Status `
			-Headers "Name", "Scope", "Role", "Status" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		If($HTML)
		{
			$columnHeaders = @(
			'Name',($global:htmlsb),
			'Scope',($global:htmlsb),
			'Role',($global:htmlsb),
			'Status',($global:htmlsb))

			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		}
	}
	If($Text)
	{
		ForEach($Admin in $Admins)
		{
			Line 1 "Name`t: " $Admin.Name

			If($Admin.Rights.ScopeName -is [array])
			{
				Line 1 "Scope`t: " $Admin.Rights.ScopeName[0]
				$cnt = -1
				ForEach($xScope in $Admin.Rights.ScopeName)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 2 "  " $Admin.Rights.ScopeName[$cnt]
					}
				}
			}
			Else
			{
				Line 1 "Scope`t: " $Admin.Rights.ScopeName
			}

			If($Admin.Rights.RoleName -is [array])
			{
				Line 1 "Role`t: " $Admin.Rights.RoleName[0]
				$cnt = -1
				ForEach($xRole in $Admin.Rights.RoleName)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 2 "  " $Admin.Rights.RoleName[$cnt]
					}
				}
			}
			Else
			{
				Line 1 "Role`t: " $Admin.Rights.RoleName
			}

			Line 1 "Status`t: " -NoNewLine
			If($Admin.Enabled)
			{
				Line 0 "Enabled"
			}
			Else
			{
				Line 0 "Disabled"
			}
			Line 0 ""
		}
	}
}

Function ProcessScopes
{
	Write-Verbose "$(Get-Date -Format G): Processing Administrator Scopes"
	$Scopes = Get-AdminScope @CCParams2 -SortBy Name
	
	If($? -and ($Null -ne $Scopes))
	{
		OutputScopes $Scopes
		If($Administrators)
		{
			OutputScopeObjects $Scopes
			OutputScopeAdministrators $Scopes
		}
	}
	ElseIf($? -and ($Null -eq $Scopes))
	{
		$txt = "There are no Administrator Scopes"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve Administrator Scopes"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date -Format G): "
}

Function OutputScopes
{
	Param([object] $Scopes)
	
	Write-Verbose "$(Get-Date -Format G): `tOutput Scopes"
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Administrative Scopes"
		$ScopesWordTable = @()
	}
	If($Text)
	{
		Line 0 "Administrative Scopes"
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Administrative Scopes"
		$rowdata = @()
	}
	ForEach($Scope in $Scopes)
	{
		If($MSWord -or $PDF)
		{
			$ScopesWordTable += @{ 
			Name = $Scope.Name; 
			Description = $Scope.Description;
			}
		}
		If($Text)
		{
			Line 1 "Name`t`t: " $Scope.Name
			Line 1 "Description`t: " $Scope.Description
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,(
			$Scope.Name,$htmlwhite,
			$Scope.Description,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScopesWordTable `
		-Columns Name, Description `
		-Headers "Name", "Description" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($Text)
	{
		#nothing to do
	}
	If($HTML)
	{
		$columnHeaders = @(
		'Name',($global:htmlsb),
		'Description',($global:htmlsb))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
}

Function OutputScopeObjects
{
	Param([object] $Scopes)
	
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Scope Objects"

	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
	}
		
	ForEach($Scope in $Scopes)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 "Scope Objects for $($Scope.Name)"
		}
		If($Text)
		{
			Line 0 "Scope Objects for $($Scope.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 3 0 "Scope Objects for $($Scope.Name)"
		}

		$Results = GetScopeDG $Scope
		
		If($Results.Count -gt 0)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Delivery Groups"
				[System.Collections.Hashtable[]] $WordTable = @()
			}
			If($Text)
			{
				Line 0 "Delivery Groups"
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "Delivery Groups"
				$rowdata = @()
			}
			
			ForEach($Result in $Results)
			{
				If($MSWord -or $PDF)
				{
					$WordTable += @{ 
					GroupName = $Result.Name; 
					GroupDesc = $Result.Description; 
					}
				}
				If($Text)
				{
					Line 1 "Name: " $Result.Name
					Line 1 "Description: " $Result.Description
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata += @(,(
					$Result.Name,$htmlwhite,
					$Result.Description,$htmlwhite))
				}
			}

			If($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $WordTable `
				-Columns GroupName, GroupDesc `
				-Headers "Name", "Description" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 250;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}
			If($Text)
			{
				#nothing to do
			}
			If($HTML)
			{
				$msg = ""
				$columnHeaders = @('Name',($global:htmlsb),'Description',($global:htmlsb))
				$ColumnWidths = @("250","250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $ColumnWidths -tablewidth "500"
			}
		}

		$Results = GetScopeMC $Scope

		If($Results.Count -gt 0)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Machine Catalogs"
				[System.Collections.Hashtable[]] $WordTable = @()
			}
			If($Text)
			{
				Line 0 "Machine Catalogs"
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "Machine Catalogs"
				$rowdata = @()
			}

			ForEach($Result in $Results)
			{
				If($MSWord -or $PDF)
				{
					$WordTable += @{ 
					CatalogName = $Result.Name; 
					CatalogDesc = $Result.Description; 
					}
				}
				If($Text)
				{
					Line 1 "Name: " $Result.Name
					Line 1 "Description: " $Result.Description
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata += @(,(
					$Result.Name,$htmlwhite,
					$Result.Description,$htmlwhite))
				}
			}
			
			If($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $WordTable `
				-Columns CatalogName, CatalogDesc `
				-Headers "Name", "Description" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 250;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}
			If($Text)
			{
				#nothing to do
			}
			If($HTML)
			{
				$msg = ""
				$columnHeaders = @('Name',($global:htmlsb),'Description',($global:htmlsb))
				$ColumnWidths = @("250","250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $ColumnWidths -tablewidth "500"
			}
		}

		$Results = GetScopeHyp $Scope

		If($Results.Count -gt 0)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Hosting"
				[System.Collections.Hashtable[]] $WordTable = @()
			}
			If($Text)
			{
				Line 0 "Hosting"
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "Hosting"
				$rowdata = @()
			}

			ForEach($Result in $Results)
			{
				If($MSWord -or $PDF)
				{
					$WordTable += @{ 
					HypName = $Result.Name; 
					HypDesc = $Result.Description; 
					}
				}
				If($Text)
				{
					Line 1 "Name: " $Result.Name
					Line 1 "Description: " $Result.Description
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata += @(,(
					$Result.Name,$htmlwhite,
					$Result.Description,$htmlwhite))
				}
			}
			
			If($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $WordTable `
				-Columns HypName, HypDesc `
				-Headers "Name", "Description" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 250;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}
			If($Text)
			{
				#nothing to do
			}
			If($HTML)
			{
				$msg = ""
				$columnHeaders = @('Name',($global:htmlsb),'Description',($global:htmlsb))
				$ColumnWidths = @("250","250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $ColumnWidths -tablewidth "500"
			}
		}

		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
		}
	}
}

Function GetScopeDG
{
	Param([object] $Scope)
	
	$DG = New-Object System.Collections.ArrayList
	#get delivery groups
	If($Scope.Name -eq "All")
	{
		$Results = @(Get-BrokerDesktopGroup @CCParams2 | `
		Select-Object Name, Description, Scopes | `
		Sort-Object Name -unique)
	}
	Else
	{
		$Results = @(Get-BrokerDesktopGroup @CCParams2 | `
		Select-Object Name, Description, Scopes | `
		Where-Object {$_.Scopes -like $Scope.Name} | `
		Sort-Object Name -unique)
	}
	
	If($? -and $Null -ne $Results)
	{
		ForEach($Result in $Results)
		{
			$obj = [PSCustomObject] @{
				Name        = $Result.Name			
				Description = $Result.Description			
			}
			$null = $DG.Add($obj)
		}
	}

	Return ,$DG
}

Function GetScopeMC
{
	Param([object] $Scope)
	
	#get machine catalogs
	$MC = New-Object System.Collections.ArrayList
	
	If($Scope.Name -eq "All")
	{
		$Results = @(Get-BrokerCatalog @CCParams2 | `
		Select-Object Name, Description, Scopes | `
		Sort-Object Name -unique)
	}
	Else
	{
		$Results = @(Get-BrokerCatalog @CCParams2 | `
		Select-Object Name, Description, Scopes | `
		Where-Object {$_.Scopes -like $Scope.Name} | `
		Sort-Object Name -unique)
	}

	If($? -and $Null -ne $Results)
	{
		ForEach($Result in $Results)
		{
			$obj = [PSCustomObject] @{
				Name        = $Result.Name			
				Description = $Result.Description			
			}
			$null = $MC.Add($obj)
		}
	}

	Return ,$MC
}

Function GetScopeHyp
{
	Param([object] $Scope)
	
	#get hypervisor connections
	$Hyp = New-Object System.Collections.ArrayList
	
	If($Scope.Name -eq "All")
	{
		$Results = @(Get-HypScopedObject @CCParams2 | `
		Select-Object ObjectName, Description, ScopeName | `
		Sort-Object ObjectName -unique)
	}
	Else
	{
		$Results = @(Get-HypScopedObject @CCParams2 | `
		Select-Object ObjectName, Description, ScopeName | `
		Where-Object {$_.ScopeName -like $Scope.Name} | `
		Sort-Object ObjectName -unique)
	}

	If($? -and $Null -ne $Results)
	{
		ForEach($Result in $Results)
		{
			$obj = [PSCustomObject] @{
				Name        = $Result.ObjectName			
				Description = $Result.Description			
			}
			$null = $Hyp.Add($obj)
		}
	}

	Return ,$Hyp
}

Function OutputScopeAdministrators 
{
	Param([object] $Scopes)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Scope Administrators"

	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
	}
	
	ForEach($Scope in $Scopes)
	{
		If($MSword -or $PDF)
		{
			[System.Collections.Hashtable[]] $WordTable = @()
			WriteWordLine 3 0 "Administrators for Scope: $($Scope.Name)"
		}
		If($Text)
		{
			Line 1 "Administrators for Scope: $($Scope.Name)"
		}
		If($HTML)
		{
			$rowdata = @()
			WriteHTMLLine 3 0 "Administrators for Scope: $($Scope.Name)"
		}
	
		$Admins = Get-AdminAdministrator -EA 0 | `
		Where-Object {$_.UserIdentityType -ne "Sid" -and (-not [String]::IsNullOrEmpty($_.Name))} | `
		Where-Object {$_.Rights.ScopeName -Contains $Scope.Name}
		
		If($? -and $Null -ne $Admins)
		{
			ForEach($Admin in $Admins)
			{
				$xEnabled = "Disabled"
				If($Admin.Enabled)
				{
					$xEnabled = "Enabled"
				}

				$xRoleName = ""
				ForEach($Right in $Admin.Rights)
				{
					If($Right.ScopeName -eq $Scope.Name -or $Right.ScopeName -eq "All")
					{
						$xRoleName = $Right.RoleName
					}
				}
				
				If($MSWord -or $PDF)
				{
					$WordTable += @{ 
					AdminName = $Admin.Name; 
					Role = $xRoleName; 
					Type = $xEnabled;
					}
				}
				If($Text)
				{
					Line 2 "Administrator Name`t: " $Admin.Name
					Line 2 "Role`t`t`t: " $xRoleName
					Line 2 "Status`t`t`t: " $xEnabled
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata += @(,(
					$Admin.Name,$htmlwhite,
					$xRoleName,$htmlwhite,
					$xEnabled,$htmlwhite))
				}
			}
			
			If($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $WordTable `
				-Columns AdminName, Role, Type `
				-Headers "Administrator Name", "Role", "Status" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 275;
				$Table.Columns.Item(2).Width = 225;
				$Table.Columns.Item(3).Width = 55;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				#nothing to do
			}
			If($HTML)
			{
				$columnHeaders = @(
				'Administrator Name',($global:htmlsb),
				'Role',($global:htmlsb),
				'Status',($global:htmlsb))

				$msg = ""
				$columnWidths = @("275","225","55")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "555"
			}
		}
		ElseIf($? -and $Null -eq $Admins)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "No administrators defined"
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "No administrators defined"
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "No administrators defined"
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "Unable to retrieve administrators"
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "Unable to retrieve administrators"
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "Unable to retrieve administrators"
			}
		}
	}
}

Function ProcessRoles
{
	Write-Verbose "$(Get-Date -Format G): Processing Administrator Roles"
	$Roles = Get-AdminRole @CCParams2 -SortBy Name

	If($? -and ($Null -ne $Roles))
	{
		OutputRoles $Roles
		If($Administrators)
		{
			OutputRoleDefinitions $Roles
			OutputRoleAdministrators $Roles
		}
	}
	ElseIf($? -and ($Null -eq $Roles))
	{
		$txt = "There are no Administrator Roles"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve Administrator Roles"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date -Format G): "
}

Function OutputRoles
{
	Param([object] $Roles)
	
	Write-Verbose "$(Get-Date -Format G): `tOutput Roles"
	
	If($MSWord -or $PDF)
	{
		If($Administrators)
		{
			$Selection.InsertNewPage()
		}
		WriteWordLine 2 0 "Administrative Roles"
		[System.Collections.Hashtable[]] $RolesWordTable = @()
	}
	If($Text)
	{
		Line 0 "Administrative Roles"
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		WriteHTMLLine 2 0 "Administrative Roles"
	}
	
	ForEach($Role in $Roles)
	{
		$Tmp = $Null
		If($Role.BuiltIn)
		{
			$Tmp = "Built In"
		}
		Else
		{
			$Tmp = "Custom"
		}

		If($MSWord -or $PDF)
		{
			$RolesWordTable += @{ 
			Role = $Role.Name; 
			Description = $Role.Description; 
			Type = $Tmp;
			}
		}
		If($Text)
		{
			Line 1 "Role`t`t: " $Role.Name
			Line 1 "Description`t: " $Role.Description
			Line 1 "Type`t`t: " $tmp
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,(
			$Role.Name,$htmlwhite,
			$Role.Description,$htmlwhite,
			$tmp,$htmlwhite))
		}
	}
	
	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $RolesWordTable `
		-Columns Role, Description, Type `
		-Headers "Role", "Description", "Type" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 300;
		$Table.Columns.Item(3).Width = 50;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($Text)
	{
		#nothing to do
	}
	If($HTML)
	{
		$columnHeaders = @(
		'Role',($global:htmlsb),
		'Description',($global:htmlsb),
		'Type',($global:htmlsb))

		$msg = ""
		$columnWidths = @("200","450","50")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "700"
	}
}

Function OutputRoleDefinitions
{
	Param([object] $Roles)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Role Definitions"
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
	}
	
	ForEach($Role in $Roles)
	{
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $WordTable = @()
			WriteWordLine 3 0 "Role definition for $($Role.Name)"
			WriteWordLine 0 0 "Details - " $Role.Name
			WriteWordLine 0 0 $Role.Description

			$comp = ""
			$w = 0
		}
		If($Text)
		{
			Line 0 "Role definition for $($Role.Name)"
			Line 0 "Details - " $Role.Name
			Line 0 $Role.Description
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			WriteHTMLLine 3 0 "Role definition for $($Role.Name)"
			WriteHTMLLine 0 0 "Details - " $Role.Name
			WriteHTMLLine 0 0 $Role.Description

			$comp = ""
			$h = 0
		}

		$Permissions = $Role.Permissions
		$Results = GetRolePermissions $Permissions

		ForEach($Result in $Results)
		{
			If($MSWord -or $PDF)
			{
				If($w -eq 0)
				{
					$comp = $Result.Value

					$WordTable += @{ 
					FolderName = $Result.Value; 
					Permission = $Result.Name; 
					}
				}
				Else
				{
					If($comp -eq $Result.value)
					{
						$WordTable += @{ 
						FolderName = ""; 
						Permission = $Result.Name; 
						}
					}
					Else
					{
						$comp = $Result.Value

						$WordTable += @{ 
						FolderName = $Result.Value; 
						Permission = $Result.Name; 
						}
					}
				}
				$w++
			}
			If($Text)
			{
				Line 1 "Folder Name`t: " $Result.Value
				Line 1 "Permission`t: " $Result.Name
				Line 0 ""
			}
			If($HTML)
			{
				If($h -eq 0)
				{
					$comp = $Result.Value
					$rowdata += @(,(
					$Result.Value,$htmlwhite,
					$Result.Name,$htmlwhite))
				}
				Else
				{
					If($comp -eq $Result.value)
					{
						$rowdata += @(,(
						"",$htmlwhite,
						$Result.Name,$htmlwhite))
					}
					Else
					{
						$comp = $Result.Value
						$rowdata += @(,(
						$Result.Value,$htmlwhite,
						$Result.Name,$htmlwhite))
					}
				}
				$h++
			}
		}

		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $WordTable `
			-Columns FolderName, Permission `
			-Headers "Folder Name", "Permissions" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 100;
			$Table.Columns.Item(2).Width = 400;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
			'Folder Name',($global:htmlsb),
			'Permissions',($global:htmlsb))

			$msg = ""
			$ColumnWidths = @("100","500")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders	-fixedWidth $columnWidths -tablewidth "600"
		}
	}
}

Function GetRolePermissions
{
	Param([object] $Permissions)
	
	$Results = @{}
	
	ForEach($Permission in $Permissions)
	{
		Switch ($Permission)
		{
			"Admin_FullControl"											{$Results.Add("Manage Administrators", "Administrators")}
			"Admin_Read"												{$Results.Add("View Administrators", "Administrators")}
			"Admin_RoleControl"											{$Results.Add("Manage Administrator Custom Roles", "Administrators")}
			"Admin_ScopeControl"										{$Results.Add("Manage Administrator Scopes", "Administrators")}
			"Manage_ServiceConfigurationData"							{$Results.Add("Manage ServiceSettings", "Administrators")}
			
			"AppGroupApplications_ChangeTags"							{$Results.Add("Edit Application tags (Application Group)", "Application Groups")}
			"AppGroupApplications_Create"								{$Results.Add("Create Application (Application Group)", "Application Groups")}
			"AppGroupApplications_CreateFolder"							{$Results.Add("Create Application Folder (Application Group)", "Application Groups")}
			"AppGroupApplications_Delete"								{$Results.Add("Delete Application (Application Group)", "Application Groups")}
			"AppGroupApplications_EditFolder"							{$Results.Add("Edit Application Folder (Application Group)", "Application Groups")}
			"AppGroupApplications_EditProperties"						{$Results.Add("Edit Application Properties (Application Group)", "Application Groups")}
			"AppGroupApplications_MoveFolder"							{$Results.Add("Move Application Folder (Application Group)", "Application Groups")}
			"AppGroupApplications_Read"									{$Results.Add("View Applications (Application Group)", "Application Groups")}
			"AppGroupApplications_RemoveFolder"							{$Results.Add("Remove Application Folder (Application Group)", "Application Groups")}
			"AppGroupApplications_ChangeUserAssignment"					{$Results.Add("Change users assigned to an application (Application Group)", "Application Groups")}
			"ApplicationGroup_AddApplication"							{$Results.Add("Add Application to Application Group", "Application Groups")}
			"ApplicationGroup_AddScope"									{$Results.Add("Add Application Group to Scope", "Application Groups")}
			"ApplicationGroup_AddToDesktopGroup"						{$Results.Add("Add Delivery Group to Application Group", "Application Groups")}
			"ApplicationGroup_ChangeTags"								{$Results.Add("Change Tags on Application Group", "Application Groups")}
			"ApplicationGroup_ChangeUserAssignment"						{$Results.Add("Edit User Assignment on Application Group", "Application Groups")}
			"ApplicationGroup_Create"									{$Results.Add("Create Application Group", "Application Groups")}
			"ApplicationGroup_Delete"									{$Results.Add("Delete Application Group", "Application Groups")}
			"ApplicationGroup_EditProperties"							{$Results.Add("Edit Application Group Properties", "Application Groups")}
			"ApplicationGroup_Read"										{$Results.Add("View Application Groups", "Application Groups")}
			"ApplicationGroup_RemoveApplication"						{$Results.Add("Remove Application from Application Group", "Application Groups")}
			"ApplicationGroup_RemoveFromDesktopGroup"					{$Results.Add("Remove Delivery Group from Application Group", "Application Groups")}
			"ApplicationGroup_RemoveScope"								{$Results.Add("Remove Application Group from Scope", "Application Groups")}
			
			"AppLib_AddApplication"										{$Results.Add("Add App-V applications", "App-V")}
			"AppLib_AddPackage"											{$Results.Add("Add App-V Application Libraries and Packages", "App-V")}
			"AppLib_IsolationGroup_Create"								{$Results.Add("Create App-V Isolation Group", "App-V")}
			"AppLib_IsolationGroup_Remove"								{$Results.Add("Remove App-V Isolation Groups", "App-V")}
			"AppLib_Read"												{$Results.Add("Read App-V Application Libraries and Packages", "App-V")}
			"AppLib_RemoveApplication"									{$Results.Add("Remove App-V applications", "App-V")}
			"AppLib_RemovePackage"										{$Results.Add("Remove App-V Application Libraries and Packages", "App-V")}
			"AppV_AddServer"											{$Results.Add("Add App-V publishing server", "App-V")}
			"AppV_DeleteServer"											{$Results.Add("Delete App-V publishing server", "App-V")}
			"AppV_Read"													{$Results.Add("Read App-V servers", "App-V")}
			
			"Cloud_Storefront_Read"										{$Results.Add("Read Storefront Configuration", "Cloud")}
			"Cloud_Storefront_Write"									{$Results.Add("Update Storefront Configuration", "Cloud")}

			"Controller_EditProperties"									{$Results.Add("Edit Controller", "Controllers")}
			"Controllers_Remove"										{$Results.Add("Remove Delivery Controller", "Controllers")}

			"Applications_AttachClientHostedApplicationToDesktopGroup"	{$Results.Add("Attach Local Access Application to Delivery Group", "Delivery Groups")}
			"Applications_ChangeMaintenanceMode"						{$Results.Add("Enable/disable maintenance mode of an Application", "Delivery Groups")}
			"Applications_ChangeTags"									{$Results.Add("Edit Application tags", "Delivery Groups")}
			"Applications_ChangeUserAssignment"							{$Results.Add("Change users assigned to an application", "Delivery Groups")}
			"Applications_Create"										{$Results.Add("Create Application", "Delivery Groups")}
			"Applications_CreateFolder"									{$Results.Add("Create Application Folder", "Delivery Groups")}
			"Applications_Delete"										{$Results.Add("Delete Application", "Delivery Groups")}
			"Applications_DetachClientHostedApplicationToDesktopGroup"	{$Results.Add("Detach Local Access Application from Delivery Group", "Delivery Groups")}
			"Applications_EditFolder"									{$Results.Add("Edit Application Folder", "Delivery Groups")}
			"Applications_EditProperties"								{$Results.Add("Edit Application Properties", "Delivery Groups")}
			"Applications_MoveFolder"									{$Results.Add("Move Application Folder", "Delivery Groups")}
			"Applications_Read"											{$Results.Add("View Applications", "Delivery Groups")}
			"Applications_RemoveFolder"									{$Results.Add("Remove Application Folder", "Delivery Groups")}
			"DesktopGroup_AddApplication"								{$Results.Add("Add Application to Delivery Group", "Delivery Groups")}
			"DesktopGroup_AddApplicationGroup"							{$Results.Add("Add Application Group to Delivery Group", "Delivery Groups")}
			"DesktopGroup_AddMachines"									{$Results.Add("Add Machines to Delivery Group", "Delivery Groups")}
			"DesktopGroup_AddScope"										{$Results.Add("Add Delivery Group to Scope", "Delivery Groups")}
			"DesktopGroup_AddWebhook"									{$Results.Add("Add Webhooks to Delivery Group", "Delivery Groups")}
			"DesktopGroup_ChangeMachineMaintenanceMode"					{$Results.Add("Enable/disable maintenance mode of a machine via Delivery Group membership", "Delivery Groups")}
			"DesktopGroup_ChangeMaintenanceMode"						{$Results.Add("Enable/disable maintenance mode of a Delivery Group", "Delivery Groups")}
			"DesktopGroup_ChangeTags"									{$Results.Add("Edit Delivery Group tags", "Delivery Groups")}
			"DesktopGroup_ChangeUserAssignment"							{$Results.Add("Change users assigned to a desktop", "Delivery Groups")}
			"DesktopGroup_Create"										{$Results.Add("Create Delivery Group", "Delivery Groups")}
			"DesktopGroup_Delete"										{$Results.Add("Delete Delivery Group", "Delivery Groups")}
			"DesktopGroup_EditProperties"								{$Results.Add("Edit Delivery Group Properties", "Delivery Groups")}
			"DesktopGroup_Machine_ChangeTags"							{$Results.Add("Edit Delivery Group machine tags", "Delivery Groups")}
			"DesktopGroup_PowerOperations_RDS"							{$Results.Add("Perform power operations on Windows Server machines via Delivery Group membership", "Delivery Groups")}
			"DesktopGroup_PowerOperations_VDI"							{$Results.Add("Perform power operations on Windows Desktop machines via Delivery Group membership", "Delivery Groups")}
			"DesktopGroup_Read"											{$Results.Add("View Delivery Groups", "Delivery Groups")}
			"DesktopGroup_RemoveApplication"							{$Results.Add("Remove Application from Delivery Group", "Delivery Groups")}
			"DesktopGroup_RemoveApplicationGroup"						{$Results.Add("Remove Application Group from Delivery Group", "Delivery Groups")}
			"DesktopGroup_RemoveDesktop"								{$Results.Add("Remove Desktop from Delivery Group", "Delivery Groups")}
			"DesktopGroup_RemoveScope"									{$Results.Add("Remove Delivery Group from Scope", "Delivery Groups")}
			"DesktopGroup_SessionManagement"							{$Results.Add("Perform session management on machines via Delivery Group membership", "Delivery Groups")}
			"Machine_ChangeTagsBase"									{$Results.Add("Edit machine tags", "Delivery Groups")}
			
			"Director_AlertPolicy_Edit"									{$Results.Add("Create\Edit\Delete Alert Policies", "Director")}
			"Director_AlertPolicy_Read"									{$Results.Add("View Alert Policies", "Director")}
			"Director_Alerts_Read"										{$Results.Add("View Alerts", "Director")}
			"Director_ApplicationDashboard"								{$Results.Add("View Applications page", "Director")}
			"Director_ClientDetails_Read"								{$Results.Add("View Client Details page", "Director")}
			"Director_ClientHelpDesk_Read"								{$Results.Add("View Client Activity Manager page", "Director")}
			"Director_CloudAnalyticsConfiguration"						{$Results.Add("Create\Edit\Remove Cloud Analytics Configurations", "Director")}
			"Director_Configuration"									{$Results.Add("View Configurations page", "Director")}
			"Director_Dashboard_Read"									{$Results.Add("View Dashboard page", "Director")}
			"Director_DesktopHardwareInformation_Edit"					{$Results.Add("Edit Machine Hardware related Broker machine command properties", "Director")}
			"Director_DiskMetrics_Edit"									{$Results.Add("Edit Disk metrics related Broker machine command properties", "Director")}
			"Director_DismissAlerts"									{$Results.Add("Dismiss Alerts", "Director")}
			"Director_EmailserverConfiguration_Edit"					{$Results.Add("Create\Edit\Remove Alert Email Server Configuration", "Director")}
			"Director_Filters_ApplicationInstances"						{$Results.Add("View Filters page Application Instances only", "Director")}
			"Director_Filters_Connections"								{$Results.Add("View Filters page Connections only", "Director")}
			"Director_Filters_Machines"									{$Results.Add("View Filters page Machines only", "Director")}
			"Director_Filters_Sessions"									{$Results.Add("View Filters page Sessions only", "Director")}
			"Director_GPOData_Edit"										{$Results.Add("Edit GPO Data related Broker machine command properties", "Director")}
			"Director_GpuMetrics_Edit"									{$Results.Add("Edit Gpu metrics related Broker machine command properties", "Director")}
			"Director_HDXInformation_Edit"								{$Results.Add("Edit HDX related Broker machine command properties", "Director")}
			"Director_HDXProtocol_Edit"									{$Results.Add("Edit HDX Protocol related Broker machine command properties", "Director")}
			"Director_HelpDesk_Read"									{$Results.Add("View Activity Manager page", "Director")}
			"Director_KillApplication"									{$Results.Add("Perform Kill Application running on a machine", "Director")}
			"Director_KillApplication_Edit"								{$Results.Add("Edit Kill Application related Broker machine command properties", "Director")}
			"Director_KillProcess"										{$Results.Add("Perform Kill Process running on a machine", "Director")}
			"Director_KillProcess_Edit"									{$Results.Add("Edit Kill Process related Broker machine command properties", "Director")}
			"Director_LatencyInformation_Edit"							{$Results.Add("Edit Latency related Broker machine command properties", "Director")}
			"Director_MachineDetails_Read"								{$Results.Add("View Machine Details page", "Director")}
			"Director_MachineMetricValues_Edit"							{$Results.Add("Edit Machine metric related Broker machine command properties", "Director")}
			"Director_PersonalizationInformation_Edit"					{$Results.Add("Edit Personalization related Broker machine command properties", "Director")}
			"Director_PoliciesInformation_Edit"							{$Results.Add("Edit Policies related Broker machine command properties", "Director")}
			"Director_ProbeConfigurationActions"						{$Results.Add("Create\Edit\Remove Probe Configurations", "Director")}
			"Director_ProfileLoadData_Edit"								{$Results.Add("Edit Profile Load Data related Broker machine command properties", "Director")}
			"Director_ResetVDisk"										{$Results.Add("Perform Reset VDisk operation", "Director")}
			"Director_ResetVDisk_Edit"									{$Results.Add("Edit Reset VDisk related Broker machine command properties", "Director")}
			"Director_RoundTripInformation_Edit"						{$Results.Add("Edit Roundtrip Time related Broker machine command properties", "Director")}
			"Director_SCOM_Read"										{$Results.Add("View SCOM Notifications", "Director")}
			"Director_ShadowSession"									{$Results.Add("Perform Remote Assistance on a machine", "Director")}
			"Director_ShadowSession_Edit"								{$Results.Add("Edit Remote Assistance related Broker machine command properties", "Director")}
			"Director_SliceAndDice_Read"								{$Results.Add("View Filters page", "Director")}
			"Director_StartupMetrics_Edit"								{$Results.Add("Edit Startup related Broker machine command properties", "Director")}
			"Director_TaskManagerInformation_Edit"						{$Results.Add("Edit Task Manager related Broker machine command properties", "Director")}
			"Director_Trends_Read"										{$Results.Add("View Trends page", "Director")}
			"Director_UserDetails_Read"									{$Results.Add("View User Details page", "Director")}
			"Director_WindowsSessionId_Edit"							{$Results.Add("Edit Windows Sessionid related Broker machine command properties", "Director")}
			"UPM_Reset_Profiles"										{$Results.Add("Reset user profiles", "Director")}
			"UPM_Reset_Profiles_Edit"									{$Results.Add("Edit Reset User Profiles related Broker machine command properties", "Director")}
			
			"Director_ProbeAgentConfigurationActions"					{$Results.Add("Access to Probe Agent APIs", "DirectorProbeAgent")}

			"Hosts_AddScope"											{$Results.Add("Add Host Connection to Scope", "Hosts")}
			"Hosts_AddStorage"											{$Results.Add("Add storage to Resources", "Hosts")}
			"Hosts_ChangeMaintenanceMode"								{$Results.Add("Enable/disable maintenance mode of a Host Connection", "Hosts")}
			"Hosts_Consume"												{$Results.Add("Use Host Connection or Resources to Create Catalog", "Hosts")}
			"Hosts_CreateHost"											{$Results.Add("Add Host Connection or Resources", "Hosts")}
			"Hosts_DeleteConnection"									{$Results.Add("Delete Host Connection", "Hosts")}
			"Hosts_DeleteHost"											{$Results.Add("Delete Resources", "Hosts")}
			"Hosts_EditConnectionProperties"							{$Results.Add("Edit Host Connection properties", "Hosts")}
			"Hosts_EditHostProperties"									{$Results.Add("Edit Resources", "Hosts")}
			"Hosts_Read"												{$Results.Add("View Host Connections and Resources", "Hosts")}
			"Hosts_RemoveScope"											{$Results.Add("Remove Host Connection from Scope", "Hosts")}

			"Licensing_ChangeLicenseServer"								{$Results.Add("Change licensing server", "Licensing")}
			"Licensing_EditLicensingProperties"							{$Results.Add("Edit product edition", "Licensing")}
			"Licensing_Read"											{$Results.Add("View Licensing", "Licensing")}

			"Logging_Delete"											{$Results.Add("Delete Configuration Logs", "Logging")}
			"Logging_EditPreferences"									{$Results.Add("Edit Logging Preferences", "Logging")}
			"Logging_Read"												{$Results.Add("View Configuration Logs", "Logging")}

			"Catalog_AddMachines"										{$Results.Add("Add Machines to Machine Catalog", "Machine Catalogs")}
			"Catalog_AddScope"											{$Results.Add("Add Machine Catalog to Scope", "Machine Catalogs")}
			"Catalog_CancelProvTask"									{$Results.Add("Cancel Provisioning Task", "Machine Catalogs")}
			"Catalog_ChangeMachineMaintenanceMode"						{$Results.Add("Enable/disable maintenance mode of a machine via Machine Catalog membership", "Machine Catalogs")}
			"Catalog_ChangeMaintenanceMode"								{$Results.Add("Enable/disable maintenance mode on Desktop via Machine Catalog membership", "Machine Catalogs")}
			"Catalog_ChangeTags"										{$Results.Add("Edit Catalog tags", "Machine Catalogs")}
			"Catalog_ChangeUserAssignment"								{$Results.Add("Change users assigned to a machine", "Machine Catalogs")}
			"Catalog_ConsumeMachines"									{$Results.Add("Allow machines to be consumed by a Delivery Group", "Machine Catalogs")}
			"Catalog_Create"											{$Results.Add("Create Machine Catalog", "Machine Catalogs")}
			"Catalog_Delete"											{$Results.Add("Delete Machine Catalog", "Machine Catalogs")}
			"Catalog_EditProperties"									{$Results.Add("Edit Machine Catalog Properties", "Machine Catalogs")}
			"Catalog_Manage_ChangeTags"									{$Results.Add("Edit Catalog machine tags", "Machine Catalogs")}
			"Catalog_ManageAccounts"									{$Results.Add("Manage Active Directory Accounts", "Machine Catalogs")}
			"Catalog_PowerOperations_RDS"								{$Results.Add("Perform power operations on Windows Server machines via Machine Catalog membership", "Machine Catalogs")}
			"Catalog_PowerOperations_VDI"								{$Results.Add("Perform power operations on Windows Desktop machines via Machine Catalog membership", "Machine Catalogs")}
			"Catalog_Read"												{$Results.Add("View Machine Catalogs", "Machine Catalogs")}
			"Catalog_RemoveMachine"										{$Results.Add("Remove Machines from Machine Catalog", "Machine Catalogs")}
			"Catalog_RemoveScope"										{$Results.Add("Remove Machine Catalog from Scope", "Machine Catalogs")}
			"Catalog_SessionManagement"									{$Results.Add("Perform session management on machines via Machine Catalog membership", "Machine Catalogs")}
			"Catalog_UpdateMasterImage"									{$Results.Add("Perform Machine update", "Machine Catalogs")}

			"Configuration_Read"										{$Results.Add("Read Site Configuration (Configuration_Read)", "Other permissions")}
			"Configuration_Restricted_Write"							{$Results.Add("Customer Update Site Configuration", "Other permissions")}
			"Configuration_Unrestricted_Write"							{$Results.Add("Update Site Configuration", "Other permissions")}
			#"Configuration_Write"										{$Results.Add("Update Site Configuration (Configuration_Write)", "Other permissions")}
			"Database_Read"												{$Results.Add("Read database status information", "Other permissions")}
			"EnvTest"													{$Results.Add("Run environment tests", "Other permissions")}
			"Export_BrokerConfiguration"								{$Results.Add("Export Broker Configuration", "Other permissions")}
			"Global_Read"												{$Results.Add("Read Site Configuration (Global_Read)", "Other permissions")}
			"Global_Write"												{$Results.Add("Update Site Configuration (Global_Write)", "Other permissions")}
			"Orchestration_RestApi"										{$Results.Add("Manage Orchestration Service REST API", "Other permissions")}
			"PerformUpgrade"											{$Results.Add("Perform upgrade", "Other permissions")}
			"Tag_Create"												{$Results.Add("Create tags", "Other permissions")}
			"Tag_Delete"												{$Results.Add("Delete tags", "Other permissions")}
			"Tag_Edit"													{$Results.Add("Edit tags", "Other permissions")}
			"Tag_Read"													{$Results.Add("Read tags", "Other permissions")}
			"Trust_ServiceKeys"											{$Results.Add("Manage Trust Service Keys", "Other permissions")}

			"Policies_Manage"											{$Results.Add("Manage Policies", "Policies")}
			"Policies_Read"												{$Results.Add("View Policies", "Policies")}

			"Storefront_Create"											{$Results.Add("Create a new StoreFront definition", "StoreFronts")}
			"Storefront_Delete"											{$Results.Add("Delete a StoreFront definition", "StoreFronts")}
			"Storefront_Read"											{$Results.Add("Read StoreFront definitions", "StoreFronts")}
			"Storefront_Update"											{$Results.Add("Update a StoreFront definition", "StoreFronts")}

			"UPM_NewConfiguration"										{$Results.Add("Add UPM Broker Machine Configuration", "UPM")}
			"UPM_Read"													{$Results.Add("Read UPM Broker Machine Configuration", "UPM")}
			"UPM_RemoveConfiguration"									{$Results.Add("Delete UPM Broker Machine Configuration", "UPM")}

			"EdgeServer_Manage"											{$Results.Add("Manage Citrix Cloud Connector", "Zones")}
			"EdgeServer_Read"											{$Results.Add("View Citrix Cloud Connector", "Zones")}
			"Zone_Create"												{$Results.Add("Create Zone", "Zones")}
			"Zone_Delete"												{$Results.Add("Delete Zone", "Zones")}
			"Zone_EditProperties"										{$Results.Add("Edit Zone", "Zones")}
			"Zone_Read"													{$Results.Add("View Zones", "Zones")}
		}
	}

	$Results = $Results.GetEnumerator() | Sort-Object Value
	Return $Results
}

Function OutputRoleAdministrators 
{
	Param([object] $Roles)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Role Administrators"
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
	}
	
	ForEach($Role in $Roles)
	{
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $WordTable = @()
			WriteWordLine 3 0 "Administrators for Role: $($Role.Name)"
		}
		If($Text)
		{
			Line 1 "Administrators for Role: $($Role.Name)"
		}
		If($HTML)
		{
			$rowdata = @()
			WriteHTMLLine 3 0 "Administrators for Role: $($Role.Name)"
		}
		
		$Admins = Get-AdminAdministrator -EA 0 | `
		Where-Object {$_.UserIdentityType -ne "Sid" -and (-not [String]::IsNullOrEmpty($_.Name))} | `
		Where-Object {$_.Rights.RoleName -Contains $Role.Name}
		
		If($? -and $Null -ne $Admins)
		{
			ForEach($Admin in $Admins)
			{
				$xEnabled = "Disabled"
				If($Admin.Enabled)
				{
					$xEnabled = "Enabled"
				}

				$xScopeName = ""
				ForEach($Right in $Admin.Rights)
				{
					If($Right.RoleName -eq $Role.Name)
					{
						$xScopeName = $Right.ScopeName
					}
				}
				
				If($MSWord -or $PDF)
				{
					$WordTable += @{ 
					AdminName = $Admin.Name; 
					Scope = $xScopeName; 
					Type = $xEnabled;
					}
				}
				If($Text)
				{
					Line 2 "Administrator Name`t: " $Admin.Name
					Line 2 "Scope`t`t`t: " $xScopeName
					Line 2 "Status`t`t`t: " $xEnabled
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata += @(,(
					$Admin.Name,$htmlwhite,
					$xScopeName,$htmlwhite,
					$xEnabled,$htmlwhite))
				}
			}
			If($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $WordTable `
				-Columns AdminName, Scope, Type `
				-Headers "Administrator Name", "Scope", "Status" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 275;
				$Table.Columns.Item(2).Width = 225;
				$Table.Columns.Item(3).Width = 55;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				#nothing to do
			}
			If($HTML)
			{
				$columnHeaders = @(
				'Administrator Name',($global:htmlsb),
				'Scope',($global:htmlsb),
				'Status',($global:htmlsb))

				$msg = ""
				$columnWidths = @("275","225","55")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "555"
			}
		}
		ElseIf($? -and $Null -eq $Admins)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "No administrators defined"
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "No administrators defined"
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "No administrators defined"
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "Unable to retrieve administrators"
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "Unable to retrieve administrators"
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "Unable to retrieve administrators"
			}
		}
	}
}
#endregion

#region Hosting functions
Function ProcessHosting
{
	#original work on the Hosting was done by Kenny Baldwin
	Write-Verbose "$(Get-Date -Format G): Processing Hosting"
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Hosting"
	}
	If($Text)
	{
		Line 0 "Hosting"
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Hosting"
	}

	$vmstorage = @()
	$tmpstorage = @()
	$vmnetwork = @()
	$IntelliCache = New-Object System.Collections.ArrayList

	Write-Verbose "$(Get-Date -Format G): `tProcessing Hosting Units"
	$HostingUnits = Get-ChildItem -EA 0 -path 'xdhyp:\hostingunits' 4>$Null
	If($? -and $Null -ne $HostingUnits)
	{
		ForEach($item in $HostingUnits)
		{	
			$Script:TotalHostingConnections++
			ForEach($storage in $item.Storage)
			{	
				$vmstorage += $storage.StoragePath
			}
			If( $item.AdditionalStorage.Length -gt 0 )
			{
				ForEach($storage in $item.AdditionalStorage.StorageLocations)
				{
					$tmpstorage += $storage.StoragePath
				}
			}
			ForEach($network in $item.PermittedNetworks)
			{	
				$vmnetwork += $network
			}
			
			$obj1 = [PSCustomObject] @{
				hypName = $item.RootPath			
				IC      = $item.UseLocalStorageCaching			
			}
			$null = $IntelliCache.Add($obj1)
		}
	}
	ElseIf($? -and $Null -eq $HostingUnits)
	{
		$txt = "No Hosting Units found"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Hosting Units"
		OutputWarning $txt
	}

	Write-Verbose "$(Get-Date -Format G): `tProcessing Hypervisors"
	$Hypervisors = Get-BrokerHypervisorConnection @CCParams2
	If($? -and $Null -ne $Hypervisors)
	{
		ForEach($Hypervisor in $Hypervisors)
		{
			$hypvmstorage = @()
			$hypnetwork = @()
			$hypIntelliCache = @()
			$hyptmpstorage = @()
			ForEach($storage in $vmstorage)
			{
                $tmpArray = $storage.Split("\")
                $tmpHypName = $tmpArray[2]
				If($tmpHypName -eq $Hypervisor.Name)
				{		
					$hypvmstorage += $storage		
				}
				$tmpArray = $Null
				$tmpHypName = $Null
			}
			ForEach($storage in $tmpstorage)
			{
                $tmpArray = $storage.Split("\")
                $tmpHypName = $tmpArray[2]
				If($tmpHypName -eq $Hypervisor.Name)
				{		
					$hyptmpstorage += $storage		
				}
				$tmpArray = $Null
				$tmpHypName = $Null
			}
			ForEach($network in $vmnetwork)
			{
                $tmpArray = $network.NetworkPath.Split("\")
                $tmpHypName = $tmpArray[2]
				If($tmpHypName -eq $Hypervisor.Name)
				{
					$hypnetwork += $network
				}
				$tmpArray = $Null
				$tmpHypName = $Null
			}
			ForEach($ICItem in $IntelliCache)
			{
                $tmpArray = $ICItem.hypName.Split("\")
                $tmpHypName = $tmpArray[2]
				If($tmpHypName -eq $Hypervisor.Name)
				{
					$hypIntelliCache += $ICItem
				}
				$tmpArray = $Null
				$tmpHypName = $Null
			}
			$xStorageName = ""
			ForEach($Unit in $HostingUnits)
			{
				If($Unit.HypervisorConnection.HypervisorConnectionName -eq $Hypervisor.Name)
				{
					$xStorageName = $Unit.HostingUnitName
				}
			}
			$xAddress = ""
			$xHAAddress = @()
			$xUserName = ""
			$xScopes = ""
			$xMaintMode = $False
			$xConnectionType = ""
			$xConnectionPluginID = ""
			$xState = ""
			$xZoneName = ""
			$xPowerActions = @()
			Write-Verbose "$(Get-Date -Format G): `tProcessing Hosting Connections"
			$Connections = Get-ChildItem -EA 0 -path 'xdhyp:\connections' 4>$Null
			
			If($? -and $Null -ne $Connections)
			{
				ForEach($Connection in $Connections)
				{
					If($Connection.HypervisorConnectionName -eq $Hypervisor.Name)
					{
						$xAddress = $Connection.HypervisorAddress[0]
						ForEach($tmpaddress in $Connection.HypervisorAddress)
						{
							$xHAAddress += $tmpaddress
						}
						$xUserName = $Connection.UserName
						ForEach($Scope in $Connection.Scopes)
						{
							$xScopes += $Scope.ScopeName + "; "
						}
						$xScopes += "All"
						$xMaintMode = $Connection.MaintenanceMode
						$xConnectionType = $Connection.ConnectionType
						$xConnectionPluginID = $Connection.PluginID
						$xState = $Hypervisor.State
						$xZoneName = $Connection.ZoneName
						$xPowerActions = $Connection.metadata
					}
				}
			}
			ElseIf($? -and $Null -eq $Connections)
			{
				$txt = "No Hosting Connections found"
				OutputWarning $txt
			}
			Else
			{
				$txt = "Unable to retrieve Hosting Connections"
				OutputWarning $txt
			}
			OutputHosting $Hypervisor $xConnectionType $xConnectionPluginID $xAddress $xState `
			$xUserName $xMaintMode $xStorageName $xHAAddress `
			$xPowerActions $xScopes $xZoneName $hypvmstorage `
			$hypnetwork $hyptmpstorage $hypIntelliCache
		}
	}
	ElseIf($? -and $Null -eq $Hypervisors)
	{
		$txt = "No Hypervisors found"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Hypervisors"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date -Format G):"
}

Function OutputHosting
{
	Param([object] $Hypervisor, 
	[string] $xConnectionType, 
	[string] $xConnectionPluginID, 
	[string] $xAddress, 
	[string] $xState, 
	[string] $xUserName, 
	[bool] $xMaintMode, 
	[string] $xStorageName, 
	[array] $xHAAddress, 
	[array] $xPowerActions, 
	[string] $xScopes, 
	[string] $xZoneName, 
	[array] $hypvmstorage, 
	[array] $hypnetwork,
	[array] $hyptmpstorage,
	[array] $hypIntelliCache)

	$xHAAddress = $xHAAddress | Sort-Object 
	
	#get array of standard storage
	$HypStdStorage = @()
	ForEach($path in $hypvmstorage)
	{
		$tmp1 = $path.split("\")
		$cnt1 = $tmp1.Length
		$tmp2 = $tmp1[$cnt1-1]

		ForEach($tmp in $tmp2)
		{
			$tmp3 = $tmp.Split(".")
			$HypStdStorage += $tmp3[0]
		}
	}
	
	$HypTempStorage = @()
	ForEach($path in $hyptmpstorage)
	{
		$tmp1 = $path.split("\")
		$cnt1 = $tmp1.Length
		$tmp2 = $tmp1[$cnt1-1]

		ForEach($tmp in $tmp2)
		{
			$tmp3 = $tmp.Split(".")
			$HypTempStorage += $tmp3[0]
		}
	}
	
	$HypNetworkName = @()
	If($hypnetwork.length -gt 0)
	{
		ForEach($path in $hypnetwork.NetworkPath)
		{
			$tmp1 = $path.split("\")
			$cnt1 = $tmp1.Length
			$tmp2 = $tmp1[$cnt1-1]

			ForEach($tmp in $tmp2)
			{
				$tmp3 = $tmp.Split(".")
				$HypNetworkName += $tmp3[0]
			}
		}
	}
	
	$HypICName = @()
	ForEach($item in $hypIntelliCache)
	{
		If($item.IC)
		{
			$ICState = "Enabled"
		}
		Else
		{
			$ICState = "Disabled"
		}
		$HypICName += $ICState
	}
	
	#to get all the Connection Types and PluginIDs, use Get-HypHypervisorPlugin
	#Thanks to fellow CTPs Neil Spellings, Kees Baggerman, and Trond Eirik Haavarstein for getting this info for me
	#For Citrix Cloud, the values are:
	<#
		ConnectionType DisplayName                                      PluginFactoryName           UsesCloudInfrastructure
		-------------- -----------                                      -----------------           -----------------------
				   AWS Amazon EC2                                       AWSMachineManagerFactory                       True
				 SCVMM Microsoft® System Center Virtual Machine Manager MicrosoftPSFactory                            False
			   VCenter VMware vSphere®                                  VmwareFactory                                 False
			 XenServer Citrix Hypervisor®                               XenFactory                                    False
				Custom Google Cloud Platform                            GcpPluginFactory                              False
				Custom Microsoft® Azure™                                AzureRmFactory                                False
				Custom Nutanix AHV                                      AcropolisFactory                              False
				Custom Remote PC Wake on LAN                            VdaWOLMachineManagerFactory                   False
	#>
	$xxConnectionType = ""
	Switch ($xConnectionType)
	{
		"AWS"   		{$xxConnectionType = "Amazon EC2"; Break}
		"SCVMM"     	{$xxConnectionType = "Microsoft System Center Virtual Machine Manager"; Break}
		"vCenter"   	{$xxConnectionType = "VMware vSphere"; Break}
		"XenServer" 	{$xxConnectionType = "Citrix Hypervisor"; Break}
		"Custom"    	{
							Switch ($xConnectionPluginID)
							{
								"AcropolisFactory"				{$xxConnectionType = "Nutanix AHV"; Break}
								"AzureRmFactory" 				{$xxConnectionType = "Microsoft Azure"; Break}
								"GcpPluginFactory" 				{$xxConnectionType = "Google Cloud Platform"; Break}
								"VdaWOLMachineManagerFactory"	{$xxConnectionType = "Remote PC Wake on LAN"; Break}
								Default     					{$xxConnectionType = "Custom Hypervisor Type PluginID could not be determined: $($xConnectionPluginID)"; Break}
							}
							Break
						}
		Default     {$xxConnectionType = "Hypervisor Type could not be determined: $($xConnectionType)"; Break}
	}

	$xxState = ""
	If($xState -eq "On")
	{
		$xxState = "Enabled"
	}
	Else
	{
		$xxState = "Disabled"
	}

	$xxMaintMode = ""
	If($xMaintMode)
	{
		$xxMaintMode = "On"
	}
	Else
	{
		$xxMaintMode = "Off"
	}
	
	Write-Verbose "$(Get-Date -Format G): `t`t`tOutput $($Hypervisor.Name)"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 $Hypervisor.Name
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Connection Name"; Value = $Hypervisor.Name; }) > $Null
		$ScriptInformation.Add(@{Data = "Type"; Value = $xxConnectionType; }) > $Null
		$ScriptInformation.Add(@{Data = "Address"; Value = $xAddress; }) > $Null
		$ScriptInformation.Add(@{Data = "State"; Value = $xxState; }) > $Null
		$ScriptInformation.Add(@{Data = "Username"; Value = $xUserName; }) > $Null
		$ScriptInformation.Add(@{Data = "Scopes"; Value = $xScopes; }) > $Null
		$ScriptInformation.Add(@{Data = "Maintenance Mode"; Value = $xxMaintMode; }) > $Null
		$ScriptInformation.Add(@{Data = "Zone"; Value = $xZoneName; }) > $Null
		$ScriptInformation.Add(@{Data = "Storage resource name"; Value = $xStorageName; }) > $Null
		If($HypNetworkName.Length -gt 0)
		{
			$ScriptInformation.Add(@{Data = "Networks"; Value = $HypNetworkName[0]; }) > $Null
			$cnt = -1
			ForEach($item in $HypNetworkName)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					$ScriptInformation.Add(@{Data = ""; Value = $item; }) > $Null
				}
			}
		}
		If($HypStdStorage.Length -gt 0)
		{
			$ScriptInformation.Add(@{Data = "Standard storage"; Value = $HypStdStorage[0]; }) > $Null
			$cnt = -1
			ForEach($item in $HypStdStorage)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					$ScriptInformation.Add(@{Data = ""; Value = $item; }) > $Null
				}
			}
		}
		If($HypTempStorage.Length -gt 0)
		{
			$ScriptInformation.Add(@{Data = "Temporary storage"; Value = $HypTempStorage[0]; }) > $Null
			$cnt = -1
			ForEach($item in $HypTempStorage)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					$ScriptInformation.Add(@{Data = ""; Value = $item; }) > $Null
				}
			}
		}
		If($HypICName.Length -gt 0)
		{
			$ScriptInformation.Add(@{Data = "IntelliCache:"; Value = $HypICName[0]; }) > $Null
		}

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		
		WriteWordLine 4 0 "Advanced"
		$ScriptInformation = New-Object System.Collections.ArrayList
		If($xHAAddress -is [array])
		{
			$ScriptInformation.Add(@{Data = "High Availability Servers"; Value = $xHAAddress[0]; }) > $Null
			$cnt = 0
			ForEach($tmpaddress in $xHAAddress)
			{
				If($cnt -gt 0)
				{
					$ScriptInformation.Add(@{Data = ""; Value = $tmpaddress; }) > $Null
				}
				$cnt++
			}
		}
		Else
		{
			$ScriptInformation.Add(@{Data = "High Availability Servers"; Value = "N/A"; }) > $Null
		}
		
		If($xPowerActions.Length -gt 0)
		{
			$ScriptInformation.Add(@{Data = "Simultaneous actions (all types) [Absolute]"; Value = $xPowerActions[0].Value; }) > $Null
			$ScriptInformation.Add(@{Data = "Simultaneous actions (all types) [Percentage]"; Value = $xPowerActions[2].Value; }) > $Null
			$ScriptInformation.Add(@{Data = "Maximum new actions per minute"; Value = $xPowerActions[1].Value; }) > $Null
			If($xPowerActions.Count -gt 5)
			{
				$ScriptInformation.Add(@{Data = "Connection options"; Value = $xPowerActions[5].Value; }) > $Null
			}
		}
		Else
		{
			$ScriptInformation.Add(@{Data = "Simultaneous actions (all types) [Absolute]"; Value = "N/A"; }) > $Null
			$ScriptInformation.Add(@{Data = "Simultaneous actions (all types) [Percentage]"; Value = "N/A"; }) > $Null
			$ScriptInformation.Add(@{Data = "Maximum new actions per minute"; Value = "N/A"; }) > $Null
			$ScriptInformation.Add(@{Data = "Connection options"; Value = "N/A"; }) > $Null
		}
		
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 225;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($Text)
	{
		Line 0 $Hypervisor.Name
		Line 0 ""
		Line 1 "Connection Name`t`t: " $Hypervisor.Name
		Line 1 "Type`t`t`t: " $xxConnectionType
		Line 1 "Address`t`t`t: " $xAddress
		Line 1 "State`t`t`t: " $xxState
		Line 1 "Username`t`t: " $xUserName
		Line 1 "Scopes`t`t`t: " $xScopes
		Line 1 "Maintenance Mode`t: " $xxMaintMode
		Line 1 "Zone`t`t`t: " $xZoneName
		Line 1 "Storage resource name`t: " $xStorageName
		If($HypNetworkName.Length -gt 0)
		{
			Line 1 "Networks`t`t: " $HypNetworkName[0]
			$cnt = -1
			ForEach($item in $HypNetworkName)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					Line 4 "  " $item
				}
			}
		}
		If($HypStdStorage.Length -gt 0)
		{
			Line 1 "Standard storage`t: " $HypStdStorage[0]
			$cnt = -1
			ForEach($item in $HypStdStorage)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					Line 4 "  " $item
				}
			}
		}
		If($HypTempStorage.Length -gt 0)
		{
			Line 1 "Temporary storage`t: " $HypTempStorage[0]
			$cnt = -1
			ForEach($item in $HypTempStorage)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					Line 4 "  " $item
				}
			}
		}
		If($HypICName.Length -gt 0)
		{
			Line 1 "IntelliCache`t`t: " $HypICName[0]
		}
		Line 0 ""
		
		Line 1 "Advanced"
		If($xHAAddress -is [array])
		{
			Line 2 "High Availability Servers`t`t`t: " $xHAAddress[0]

			$cnt = 0
			ForEach($tmpaddress in $xHAAddress)
			{
				If($cnt -gt 0)
				{
					Line 8 "  " $tmpaddress
				}
				$cnt++
			}
		}
		Else
		{
			Line 2 "High Availability Servers`t`t`t: N/A"
		}
		
		If($xPowerActions.Length -gt 0)
		{
			Line 2 "Simultaneous actions (all types) [Absolute]`t: " $xPowerActions[0].Value
			Line 2 "Simultaneous actions (all types) [Percentage]`t: " $xPowerActions[2].Value
			Line 2 "Maximum new actions per minute`t`t`t: " $xPowerActions[1].Value
			If($xPowerActions.Count -gt 5)
			{
				Line 2 "Connection options`t`t`t`t: " $xPowerActions[5].Value
			}
		}
		Else
		{
			Line 2 "Simultaneous actions (all types) [Absolute]`t: N/A"
			Line 2 "Simultaneous actions (all types) [Percentage]`t: N/A"
			Line 2 "Maximum new actions per minute`t`t`t: N/A"
			Line 2 "Connection options`t`t`t`t: N/A"
		}
		Line 0 ""
		
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 $Hypervisor.Name
		$rowdata = @()
		$columnHeaders = @("Connection Name",($global:htmlsb),$Hypervisor.Name,$htmlwhite)
		$rowdata += @(,('Type',($global:htmlsb),$xxConnectionType,$htmlwhite))
		$rowdata += @(,('Address',($global:htmlsb),$xAddress,$htmlwhite))
		$rowdata += @(,('State',($global:htmlsb),$xxState,$htmlwhite))
		$rowdata += @(,('Username',($global:htmlsb),$xUserName,$htmlwhite))
		$rowdata += @(,('Scopes',($global:htmlsb),$xScopes,$htmlwhite))
		$rowdata += @(,('Maintenance Mode',($global:htmlsb),$xxMaintMode,$htmlwhite))
		$rowdata += @(,('Zone',($global:htmlsb),$xZoneName,$htmlwhite))
		$rowdata += @(,('Storage resource name',($global:htmlsb),$xStorageName,$htmlwhite))
		If($HypNetworkName.Length -gt 0)
		{
			$rowdata += @(,('Network',($global:htmlsb),$HypNetworkName[0],$htmlwhite))
			$cnt = -1
			ForEach($item in $HypNetworkName)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$item,$htmlwhite))
				}
			}
		}
		If($HypStdStorage.Length -gt 0)
		{
			$rowdata += @(,('Standard storage',($global:htmlsb),$HypStdStorage[0],$htmlwhite))
			$cnt = -1
			ForEach($item in $HypStdStorage)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$item,$htmlwhite))
				}
			}
		}
		If($HypTempStorage.Length -gt 0)
		{
			$rowdata += @(,('Temporary storage',($global:htmlsb),$HypTempStorage[0],$htmlwhite))
			$cnt = -1
			ForEach($item in $HypTempStorage)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$item,$htmlwhite))
				}
			}
		}
		If($HypICName.Length -gt 0)
		{
			$rowdata += @(,('IntelliCache:',($global:htmlsb),$HypICName[0],$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("150","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		
		WriteHTMLLine 4 0 "Advanced"
		$rowdata = @()
		If($xHAAddress -is [array])
		{
			$columnHeaders = @("High Availability Servers",($global:htmlsb),$xHAAddress[0],$htmlwhite)
			$cnt = 0
			ForEach($tmpaddress in $xHAAddress)
			{
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmpaddress,$htmlwhite))
				}
				$cnt++
			}
		}
		Else
		{
			$columnHeaders = @("High Availability Servers",($global:htmlsb),"N/A",$htmlwhite)
		}
		
		If($xPowerActions.Length -gt 0)
		{
			$rowdata += @(,('Simultaneous actions (all types) [Absolute]',($global:htmlsb),$xPowerActions[0].Value,$htmlwhite))
			$rowdata += @(,('Simultaneous actions (all types) [Percentage]',($global:htmlsb),$xPowerActions[2].Value,$htmlwhite))
			$rowdata += @(,('Maximum new actions per minute',($global:htmlsb),$xPowerActions[1].Value,$htmlwhite))
			If($xPowerActions.Count -gt 5)
			{
				$rowdata += @(,('Connection options',($global:htmlsb),$xPowerActions[5].Value,$htmlwhite))
			}
		}
		Else
		{
			$rowdata += @(,('Simultaneous actions (all types) [Absolute]',($global:htmlsb),"N/A",$htmlwhite))
			$rowdata += @(,('Simultaneous actions (all types) [Percentage]',($global:htmlsb),"N/A",$htmlwhite))
			$rowdata += @(,('Maximum new actions per minute',($global:htmlsb),"N/A",$htmlwhite))
			$rowdata += @(,('Connection options',($global:htmlsb),"N/A",$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("300","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
	}
	
	If($Hosting)
	{	
		Write-Verbose "$(Get-Date -Format G): `tProcessing Host Administrators"
		$Admins = GetAdmins "Host" $Hypervisor.Name
		
		If($? -and ($Null -ne $Admins))
		{
			OutputAdminsForDetails $Admins
		}
		ElseIf($? -and ($Null -eq $Admins))
		{
			$txt = "There are no administrators for Host $($Hypervisor.Name)"
			OutputNotice $txt
		}
		Else
		{
			$txt = "Unable to retrieve administrators for Host $($Hypervisor.Name)"
			OutputWarning $txt
		}

		Write-Verbose "$(Get-Date -Format G): `tProcessing Single-session OS Data"
		$DesktopOSMachines = @(Get-BrokerMachine @CCParams2 -hypervisorconnectionname $Hypervisor.Name -sessionsupport "SingleSession")

		If($? -and ($Null -ne $DesktopOSMachines))
		{
			[int]$cnt = $DesktopOSMachines.Count
			
			If($MSWord -or $PDF)
			{
				$Selection.InsertNewPage()
				WriteWordLine 4 0 "Single-session OS Machines ($($cnt))"
			}
			If($Text)
			{
				Line 0 "Single-session OS Machines ($($cnt))"
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "Single-session OS Machines ($($cnt))"
			}

			ForEach($Desktop in $DesktopOSMachines)
			{
				OutputDesktopOSMachine $Desktop
			}
		}
		ElseIf($? -and ($Null -eq $DesktopOSMachines))
		{
			$txt = "There are no Single-session OS Machines"
			OutputNotice $txt
		}
		Else
		{
			$txt = "Unable to retrieve Single-session OS Machines"
			OutputWarning $txt
		}

		Write-Verbose "$(Get-Date -Format G): `tProcessing Multi-session OS Data"
		$ServerOSMachines = @(Get-BrokerMachine @CCParams2 -hypervisorconnectionname $Hypervisor.Name -sessionsupport "MultiSession")
		
		If($? -and ($Null -ne $ServerOSMachines))
		{
			[int]$cnt = $ServerOSMachines.Count

			If($MSWord -or $PDF)
			{
				$Selection.InsertNewPage()
				WriteWordLine 4 0 "Multi-session OS Machines ($($cnt))"
			}
			If($Text)
			{
				Line 0 ""
				Line 0 "Multi-session OS Machines ($($cnt))"
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "Multi-session OS Machines ($($cnt))"
			}
			
			ForEach($Server in $ServerOSMachines)
			{
				OutputServerOSMachine $Server
			}
		}
		ElseIf($? -and ($Null -eq $ServerOSMachines))
		{
			$txt = "There are no Multi-session OS Machines"
			OutputNotice $txt
		}
		Else
		{
			$txt = "Unable to retrieve Multi-session OS Machines"
			OutputWarning $txt
		}

		If($NoSessions -eq $False)
		{
			Write-Verbose "$(Get-Date -Format G): `tProcessing Sessions Data"
			$Sessions = @(Get-BrokerSession @CCParams2 -hypervisorconnectionname $Hypervisor.Name -SortBy UserName)
			If($? -and ($Null -ne $Sessions))
			{
				[int]$cnt = $Sessions.Count

				If($MSWord -or $PDF)
				{
					$Selection.InsertNewPage()
					WriteWordLine 4 0 "Sessions ($($cnt))"
				}
				If($Text)
				{
					Line 0 ""
					Line 0 "Sessions ($($cnt))"
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "Sessions ($($cnt))"
				}
				
				OutputHostingSessions $Sessions
			}
			ElseIf($? -and ($Null -eq $Sessions))
			{
				$txt = "There are no Sessions"
				OutputNotice $txt
			}
			Else
			{
				$txt = "Unable to retrieve Sessions"
				OutputWarning $txt
			}
		}
	}
}

Function OutputDesktopOSMachine 
{
	Param([object]$Desktop)

	Write-Verbose "$(Get-Date -Format G): `t`t`tOutput desktop $($Desktop.DNSName)"

	$xMaintMode = ""
	$xUserChanges = ""
	
	If($Desktop.InMaintenanceMode)
	{
		$xMaintMode = "On"
	}
	Else
	{
		$xMaintMode = "Off"
	}
	Switch($Desktop.PersistUserChanges)
	{
		"OnLocal" {$xUserChanges = "On Local"; Break}
		"Discard" {$xUserChanges = "Discard"; Break}
		Default   {$xUserChanges = "Unknown: $($Desktop.PersistUserChanges)"; Break}
	}

	Switch ($Desktop.PowerState)
	{
		"Off"			{$xPowerState = "Off"; Break}
		"On"			{$xPowerState = "On"; Break}
        "Resuming"		{$xPowerState = "Resuming"; Break}
		"Suspended"		{$xPowerState = "Suspended"; Break}
		"Suspending"	{$xPowerState = "Suspending"; Break}
		"TurningOff"	{$xPowerState = "Turning Off"; Break}
		"TurningOn"		{$xPowerState = "Turning On"; Break}
		"Unavailable"	{$xPowerState = "Unavailable"; Break}
		"Unknown"		{$xPowerState = "Unknown"; Break}
		"Unmanaged"		{$xPowerState = "Unmanaged"; Break}
		Default			{$xPowerState = "Unabled to determine desktop Power State: $($Desktop.PowerState)"; Break}
	}
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Name"; Value = $Desktop.DNSName; }) > $Null
		$ScriptInformation.Add(@{Data = "Machine Catalog"; Value = $Desktop.CatalogName; }) > $Null
		$ScriptInformation.Add(@{Data = "Delivery Group"; Value = $Desktop.DesktopGroupName; }) > $Null
		If(![String]::IsNullOrEmpty($Desktop.AssociatedUserNames))
		{
			$cnt = -1
			ForEach($AssociatedUserName in $Desktop.AssociatedUserNames)
			{
				If($cnt -eq 0)
				{
					$ScriptInformation.Add(@{Data = "User"; Value = $AssociatedUserName; }) > $Null
				}
				Else
				{
					$ScriptInformation.Add(@{Data = ""; Value = $AssociatedUserName; }) > $Null
				}
			}
		}
		$ScriptInformation.Add(@{Data = "Maintenance Mode"; Value = $xMaintMode; }) > $Null
		$ScriptInformation.Add(@{Data = "Persist User Changes"; Value = $xUserChanges; }) > $Null
		$ScriptInformation.Add(@{Data = "Power State"; Value = $xPowerState; }) > $Null
		$ScriptInformation.Add(@{Data = "Registration State"; Value = $Desktop.RegistrationState.ToString(); }) > $Null
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 1 "Name`t`t`t: " $Desktop.DNSName
		Line 1 "Machine Catalog`t`t: " $Desktop.CatalogName
		If(![String]::IsNullOrEmpty($Desktop.DesktopGroupName))
		{
			Line 1 "Delivery Group`t`t: " $Desktop.DesktopGroupName
		}
		If(![String]::IsNullOrEmpty($Desktop.AssociatedUserNames))
		{
			$cnt = -1
			ForEach($AssociatedUserName in $Desktop.AssociatedUserNames)
			{
				$cnt++
				If($cnt -eq 0)
				{
					Line 1 "User`t`t`t: " $AssociatedUserName
				}
				Else
				{
					Line 4 "  " $AssociatedUserName
				}
			}
			
		}
		Line 1 "Maintenance Mode`t: " $xMaintMode
		Line 1 "Persist User Changes`t: " $xUserChanges
		Line 1 "Power State`t`t: " $xPowerState
		Line 1 "Registration State`t: " $Desktop.RegistrationState.ToString()
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($global:htmlsb),$Desktop.DNSName,$htmlwhite)
		$rowdata += @(,('Machine Catalog',($global:htmlsb),$Desktop.CatalogName,$htmlwhite))
		If(![String]::IsNullOrEmpty($Desktop.DesktopGroupName))
		{
			$rowdata += @(,('Delivery Group',($global:htmlsb),$Desktop.DesktopGroupName,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($Desktop.AssociatedUserNames))
		{
			$cnt = -1
			ForEach($AssociatedUserName in $Desktop.AssociatedUserNames)
			{
				$cnt++
				If($cnt -eq 0)
				{
					$rowdata += @(,('User',($global:htmlsb),$AssociatedUserName,$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('',($global:htmlsb),$AssociatedUserName,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('Maintenance Mode',($global:htmlsb),$xMaintMode,$htmlwhite))
		$rowdata += @(,('Persist User Changes',($global:htmlsb),$xUserChanges,$htmlwhite))
		$rowdata += @(,('Power State',($global:htmlsb),$xPowerState,$htmlwhite))
		$rowdata += @(,('Registration State',($global:htmlsb),$Desktop.RegistrationState.ToString(),$htmlwhite))

		$msg = ""
		$columnWidths = @("150","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 ""
	}
}

Function OutputServerOSMachine 
{
	Param([object]$Server)
	
	Write-Verbose "$(Get-Date -Format G): `t`t`tOutput server $($Server.DNSName)"
	$xMaintMode = ""
	$xUserChanges = ""

	If($Server.InMaintenanceMode)
	{
		$xMaintMode = "On"
	}
	Else
	{
		$xMaintMode = "Off"
	}

	Switch($Server.PersistUserChanges)
	{
		"OnLocal" {$xUserChanges = "On Local"; Break}
		"Discard" {$xUserChanges = "Discard"; Break}
		Default   {$xUserChanges = "Unknown: $($Server.PersistUserChanges)"; Break}
	}

	Switch ($Server.PowerState)
	{
		"Off"			{$xPowerState = "Off"; Break}
		"On"			{$xPowerState = "On"; Break}
        "Resuming"		{$xPowerState = "Resuming"; Break}
		"Suspended"		{$xPowerState = "Suspended"; Break}
		"Suspending"	{$xPowerState = "Suspending"; Break}
		"TurningOff"	{$xPowerState = "Turning Off"; Break}
		"TurningOn"		{$xPowerState = "Turning On"; Break}
		"Unavailable"	{$xPowerState = "Unavailable"; Break}
		"Unknown"		{$xPowerState = "Unknown"; Break}
		"Unmanaged"		{$xPowerState = "Unmanaged"; Break}
		Default			{$xPowerState = "Unabled to determine desktop Power State: $($Server.PowerState)"; Break}
	}
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Name"; Value = $Server.DNSName; }) > $Null
		$ScriptInformation.Add(@{Data = "Machine Catalog"; Value = $Server.CatalogName; }) > $Null
		$ScriptInformation.Add(@{Data = "Delivery Group"; Value = $Server.DesktopGroupName; }) > $Null
		If(![String]::IsNullOrEmpty($Server.AssociatedUserNames))
		{
			$cnt = -1
			ForEach($AssociatedUserName in $Server.AssociatedUserNames)
			{
				If($cnt -eq 0)
				{
					$ScriptInformation.Add(@{Data = "User"; Value = $AssociatedUserName; }) > $Null
				}
				Else
				{
					$ScriptInformation.Add(@{Data = ""; Value = $AssociatedUserName; }) > $Null
				}
			}
		}
		$ScriptInformation.Add(@{Data = "Maintenance Mode"; Value = $xMaintMode; }) > $Null
		$ScriptInformation.Add(@{Data = "Persist User Changes"; Value = $xUserChanges; }) > $Null
		$ScriptInformation.Add(@{Data = "Power State"; Value = $xPowerState; }) > $Null
		$ScriptInformation.Add(@{Data = "Registration State"; Value = $Server.RegistrationState.ToString(); }) > $Null
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 1 "Name`t`t`t: " $Server.DNSName
		Line 1 "Machine Catalog`t`t: " $Server.CatalogName
		If(![String]::IsNullOrEmpty($Server.DesktopGroupName))
		{
			Line 1 "Delivery Group`t`t: " $Server.DesktopGroupName
		}
		If(![String]::IsNullOrEmpty($Server.AssociatedUserNames))
		{
			$cnt = -1
			ForEach($AssociatedUserName in $Server.AssociatedUserNames)
			{
				$cnt++
				If($cnt -eq 0)
				{
					Line 1 "User`t`t`t: " $AssociatedUserName
				}
				Else
				{
					Line 4 "  " $AssociatedUserName
				}
			}
			
		}
		Line 1 "Maintenance Mode`t: " $xMaintMode
		Line 1 "Persist User Changes`t: " $xUserChanges
		Line 1 "Power State`t`t: " $xPowerState
		Line 1 "Registration State`t: " $Server.RegistrationState.ToString()
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($global:htmlsb),$Server.DNSName,$htmlwhite)
		$rowdata += @(,('Machine Catalog',($global:htmlsb),$Server.CatalogName,$htmlwhite))
		If(![String]::IsNullOrEmpty($Server.DesktopGroupName))
		{
			$rowdata += @(,('Delivery Group',($global:htmlsb),$Server.DesktopGroupName,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($Server.AssociatedUserNames))
		{
			$cnt = -1
			ForEach($AssociatedUserName in $Server.AssociatedUserNames)
			{
				$cnt++
				If($cnt -eq 0)
				{
					$rowdata += @(,('User',($global:htmlsb),$AssociatedUserName,$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('',($global:htmlsb),$AssociatedUserName,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('Maintenance Mode',($global:htmlsb),$xMaintMode,$htmlwhite))
		$rowdata += @(,('Persist User Changes',($global:htmlsb),$xUserChanges,$htmlwhite))
		$rowdata += @(,('Power State',($global:htmlsb),$xPowerState,$htmlwhite))
		$rowdata += @(,('Registration State',($global:htmlsb),$Server.RegistrationState.ToString(),$htmlwhite))

		$msg = ""
		$columnWidths = @("150","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 ""
	}
}

Function OutputHostingSessions 
{
	Param([object] $Sessions)
	
	ForEach($Session in $Sessions)
	{
		Write-Verbose "$(Get-Date -Format G): `t`t`tOutput session $($Session.UserName)"
		
		If($Session.SessionSupport -eq "SingleSession")
		{
			$xSessionType = "Single"
		}
		Else
		{
			$xSessionType = "Multi"
		}
		
		#$RecordingStatus = "Not supported"
		#$result = Get-BrokerSessionRecordingStatus -Session $Session.Uid -EA 0
		
		#If($?)
		#{
		#	Switch ($result)
		#	{
		#		"SessionBeingRecorded"	{$RecordingStatus = "Session is being recorded"}
		#		"SessionNotRecorded"	{$RecordingStatus = "Session is not being recorded"}
		#		Default					{$RecordingStatus = "Unable to determine session recording status: $($result)"}
		#	}
		#}
		#Else
		#{
		#	$RecordingStatus = "Unknown"
		#}

		If([String]::IsNullOrEmpty($Session.ClientName))
		{
			$xClientName = "-"
		}
		Else
		{
			$xClientName = $Session.ClientName
		}
		
		If([String]::IsNullOrEmpty($Session.BrokeringTime))
		{
			$xBrokeringTime = "-"
		}
		Else
		{
			$xBrokeringTime = $Session.BrokeringTime
		}
		
		If($MSWord -or $PDF)
		{
			$ScriptInformation = New-Object System.Collections.ArrayList
			$ScriptInformation.Add(@{Data = "Current User"; Value = $Session.UserName; }) > $Null
			$ScriptInformation.Add(@{Data = "Name"; Value = $xClientName; }) > $Null
			$ScriptInformation.Add(@{Data = "Delivery Group"; Value = $Session.DesktopGroupName; }) > $Null
			$ScriptInformation.Add(@{Data = "Machine Catalog"; Value = $Session.CatalogName; }) > $Null
			$ScriptInformation.Add(@{Data = "Brokering Time"; Value = $xBrokeringTime; }) > $Null
			$ScriptInformation.Add(@{Data = "Session State"; Value = $Session.SessionState; }) > $Null
			$ScriptInformation.Add(@{Data = "Application State"; Value = $Session.AppState; }) > $Null
			$ScriptInformation.Add(@{Data = "Session Support"; Value = $xSessionType; }) > $Null
			#$ScriptInformation.Add(@{Data = "Recording Status"; Value = $RecordingStatus; }) > $Null
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 200;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 "Current User`t`t: " $Session.UserName
			Line 1 "Name`t`t`t: " $xClientName
			Line 1 "Delivery Group`t`t: " $Session.DesktopGroupName
			Line 1 "Machine Catalog`t`t: " $Session.CatalogName
			Line 1 "Brokering Time`t`t: " $xBrokeringTime
			Line 1 "Session State`t`t: " $Session.SessionState
			Line 1 "Application State`t: " $Session.AppState
			Line 1 "Session Support`t`t: " $xSessionType
			#Line 1 "Recording Status`t`t: " $RecordingStatus
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Current User",($global:htmlsb),$Session.UserName,$htmlwhite)
			$rowdata += @(,('Name',($global:htmlsb),$xClientName,$htmlwhite))
			$rowdata += @(,('Delivery Group',($global:htmlsb),$Session.DesktopGroupName,$htmlwhite))
			$rowdata += @(,('Machine Catalog',($global:htmlsb),$Session.CatalogName,$htmlwhite))
			$rowdata += @(,('Brokering Time',($global:htmlsb),$xBrokeringTime,$htmlwhite))
			$rowdata += @(,('Session State',($global:htmlsb),$Session.SessionState,$htmlwhite))
			$rowdata += @(,('Application State',($global:htmlsb),$Session.AppState,$htmlwhite))
			$rowdata += @(,('Session Support',($global:htmlsb),$xSessionType,$htmlwhite))
			#$rowdata += @(,('Recording Status',($global:htmlsb),$RecordingStatus,$htmlwhite))

			$msg = ""
			$columnWidths = @("150","200")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
			WriteHTMLLine 0 0 ""
		}
	}
}
#endregion

#region Licensing functions
Function ProcessLicensing
{
	Write-Verbose "$(Get-Date -Format G): Processing Licensing"
	
	$Script:Licenses = New-Object System.Collections.ArrayList
	OutputLicensingOverview
	
	Write-Verbose "$(Get-Date -Format G): "
}

Function OutputLicensingOverview
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Licensing Overview"

	$LicenseEditionType = ""
	$LicenseModelType   = ""
	$LicensedProduct    = ""

	Switch ($Script:CCSite2.ProductCode)
	{
		"CVADS"	{$LicensedProduct = "Citrix virtual apps and desktops service"; Break}
		"XDT"	{$LicensedProduct = "XenDesktop"; Break}
		"VADS"	{$LicensedProduct = "Virtual apps and desktops service"; Break}
		"VAS"	{$LicensedProduct = "Virtual apps service"; Break}
		"VAD"	{$LicensedProduct = "Virtual desktops service"; Break}
		Default	{$LicensedProduct = "Unable to determine licensed product: $Script:CCSite2.ProductCode"; Break}
	}
	
	If($Script:CCSite2.ProductCode -eq "XDT")
	{
		Switch ($Script:CCSite2.ProductEdition)
		{
			"ADV" 	{$LicenseEditionType = "Advanced Edition"; Break}
			"APP" 	{$LicenseEditionType = "App Edition"; Break}
			"BAS" 	{$LicenseEditionType = "Basic Edition"; Break}
			"ENT" 	{$LicenseEditionType = "Enterprise Edition"; Break}
			"PLT" 	{$LicenseEditionType = "Platinum Edition"; Break}
			"STD" 	{$LicenseEditionType = "VDI Edition"; Break}
			Default {$LicenseEditionType = "License edition could not be determined: $($Script:CCSite2.ProductEdition)"; Break}
		}
	}
	ElseIf($Script:CCSite2.ProductCode -eq "VADS" -or $Script:CCSite2.ProductCode -eq "VAS")
	{
		Switch ($Script:CCSite2.ProductEdition)
		{
			"Advanced" 	{$LicenseEditionType = "Advanced Edition"; Break}
			"Premium" 	{$LicenseEditionType = "Premium Edition"; Break}
			Default {$LicenseEditionType = "License edition could not be determined: $($Script:CCSite2.ProductEdition)"; Break}
		}
	}
	ElseIf($Script:CCSite2.ProductCode -eq "VDS")
	{
		Switch ($Script:CCSite2.ProductEdition)
		{
			"Premium" 	{$LicenseEditionType = "Premium Edition"; Break}
			Default {$LicenseEditionType = "License edition could not be determined: $($Script:CCSite2.ProductEdition)"; Break}
		}
	}
	ElseIf($Script:CCSite2.ProductCode -eq "CVADS")
	{
		Switch ($Script:CCSite2.ProductEdition)
		{
			"AzureVdi"				{$LicenseEditionType = "Azure VDI"; Break}
			"DaaS"					{$LicenseEditionType = "Desktops as a Service"; Break}
			"ExpressAdmin"			{$LicenseEditionType = "Express Admin"; Break}
			"Full"					{$LicenseEditionType = "Citrix Cloud Full"; Break}
			"FullTrial"				{$LicenseEditionType = "Citrix Cloud Full Trial"; Break}
			"MultitenantCustomer"	{$LicenseEditionType = "Multitenant Customer"; Break}
			"XAOnly"				{$LicenseEditionType = "XenApp Only"; Break}
			"XDOnly"				{$LicenseEditionType = "XenDesktop Only"; Break}
			Default					{$LicenseEditionType = "License edition could not be determined: $($Script:CCSite2.ProductEdition)"; Break}
		}
	}

	If($Script:CCSite1.LicenseModel -eq "UserDevice")
	{
		$LicenseModelType = "User/Device"
	}
	Else
	{
		$LicenseModelType = $Script:CCSite1.LicenseModel.ToString()
	}
	$tmpdate = '{0:yyyy\.MMdd}' -f $Script:CCSite1.LicensingBurnInDate
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Licensing"
		WriteWordLine 2 0 "Licensing Overview"

		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Product"; Value = $LicensedProduct; }) > $Null
		$ScriptInformation.Add(@{Data = "Edition"; Value = $LicenseEditionType; }) > $Null
		$ScriptInformation.Add(@{Data = "License model"; Value = $LicenseModelType; }) > $Null
		$ScriptInformation.Add(@{Data = "Required SA date"; Value = $tmpdate; }) > $Null
		$ScriptInformation.Add(@{Data = "Licensed sessions active"; Value = $Script:CCSite1.LicensedSessionsActive.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Peak Licensed CCU"; Value = $Script:CCSite1.PeakConcurrentLicenseUsers.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Peak Licensed CCD"; Value = $Script:CCSite1.PeakConcurrentLicensedDevices.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Total unique licensed users"; Value = $Script:CCSite1.TotalUniqueLicenseUsers.ToString(); }) > $Null
		
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($Text)
	{
		Line 0 "Licensing"
		Line 0 "Licensing Overview"
		Line 0 ""
		Line 0 "Product`t`t`t`t: " $LicensedProduct
		Line 0 "Edition`t`t`t`t: " $LicenseEditionType
		Line 0 "License model`t`t`t: " $LicenseModelType
		Line 0 "Required SA date`t`t: " $tmpdate
		Line 0 "Licensed sessons active`t`t: " $Script:CCSite1.LicensedSessionsActive.ToString()
		Line 0 "Peak Licensed CCU`t`t: " $Script:CCSite1.PeakConcurrentLicenseUsers.ToString()
		Line 0 "Peak Licensed CCD`t`t: " $Script:CCSite1.PeakConcurrentLicensedDevices.ToString()
		Line 0 "Total unique licensed users`t: " $Script:CCSite1.TotalUniqueLicenseUsers.ToString()
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Licensing"
		WriteHTMLLine 2 0 "Licensing Overview"
		$rowdata = @()
		$columnHeaders = @("Product",($global:htmlsb),$LicensedProduct,$htmlwhite)
		$rowdata += @(,('Edition',($global:htmlsb),$LicenseEditionType,$htmlwhite))
		$rowdata += @(,('License model',($global:htmlsb),$LicenseModelType,$htmlwhite))
		$rowdata += @(,('Required SA date',($global:htmlsb),$tmpdate,$htmlwhite))
		$rowdata += @(,('Licensed sessons active',($global:htmlsb),$Script:CCSite1.LicensedSessionsActive.ToString(),$htmlwhite))
		$rowdata += @(,("Peak Licensed CCU",($global:htmlsb),$Script:CCSite1.PeakConcurrentLicenseUsers.ToString(),$htmlwhite))
		$rowdata += @(,("Peak Licensed CCD",($global:htmlsb),$Script:CCSite1.PeakConcurrentLicensedDevices.ToString(),$htmlwhite))
		$rowdata += @(,("Total unique licensed users",($global:htmlsb),$Script:CCSite1.TotalUniqueLicenseUsers.ToString(),$htmlwhite))

		$msg = ""
		$columnWidths = @("150","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
	}
}
#endregion

#region StoreFront functions
Function ProcessStoreFront
{
	Write-Verbose "$(Get-Date -Format G): Processing StoreFront"
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "StoreFront"
	}
	If($Text)
	{
		Line 0 "StoreFront"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "StoreFront"
	}
	
	Write-Verbose "$(Get-Date -Format G): `tRetrieving StoreFront information"
	$SFInfos = Get-BrokerMachineConfiguration -EA 0 -Name rs* -SortBy LeafName
	If($? -and ($Null -ne $SFInfos))
	{
		$First = $True
		ForEach($SFInfo in $SFInfos)
		{
			$Script:TotalStoreFrontServers++

			$SFByteArray = $SFInfo.Policy
			Write-Verbose "$(Get-Date -Format G): `t`tRetrieving StoreFront server information for $($SFInfo.LeafName)"
			## GRL add Try/Catch
            Try
            {
			    $SFServer = Get-SFStoreFrontAddress -ByteArray $SFByteArray 4>$Null
            }
            Catch
            {
                $SFServer = $null
            }
			If($? -and ($Null -ne $SFServer))
			{
				If($MSWord -or $PDF)
				{
					If(!$First)
					{
						$Selection.InsertNewPage()
					}
					$First = $False
				}
				OutputStoreFront $SFServer $SFInfo
				If($StoreFront)
				{
					If($SFInfo.DesktopGroupUids.Count -gt 0)
					{
						OutputStoreFrontDeliveryGroups $SFInfo
					}
					
					Write-Verbose "$(Get-Date -Format G): `t`tProcessing administrators for StoreFront server $($SFServer.Name)"
					$Admins = GetAdmins "Storefront"
					
					If($? -and ($Null -ne $Admins))
					{
						OutputAdminsForDetails $Admins
					}
					ElseIf($? -and ($Null -eq $Admins))
					{
						$txt = "There are no administrators for StoreFront server $($SFServer.Name)"
						OutputNotice $txt
					}
					Else
					{
						$txt = "Unable to retrieve administrators for StoreFront server $($SFServer.Name)"
						OutputWarning $txt
					}
				}
			}
			ElseIf($? -and ($Null -eq $SFServer))
			{
				$txt = "There was no StoreFront Server found for $($SFInfo.LeafName)"
				OutputWarning $txt
			}
			Else
			{
				$txt = "Unable to retrieve StoreFront Server for $($SFInfo.LeafName)"
				OutputWarning $txt
			}
		}
	}
	ElseIf($? -and ($Null -eq $SFInfos))
	{
		$txt = "StoreFront is not configured for this Site"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve StoreFront configuration"
		OutputWarning $txt
	}
	
	Write-Verbose "$(Get-Date -Format G): "
}

Function OutputStoreFront
{
	Param([object]$SFServer, [object] $SFInfo)
	
	$DGCnt = $SFInfo.DesktopGroupUids.Count
	
	Write-Verbose "$(Get-Date -Format G): `t`t`tOutput StoreFront server $($SFServer.Name)"
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Server: " $SFServer.Name
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "StoreFront Server"; Value = $SFServer.Name; }) > $Null
		$ScriptInformation.Add(@{Data = "Used by # Delivery Groups"; Value = $DGCnt.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "URL"; Value = $SFServer.Url; }) > $Null
		$ScriptInformation.Add(@{Data = "Description"; Value = $SFServer.Description; }) > $Null
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 300;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "Server"
		Line 0 "StoreFront Server`t`t: " $SFServer.Name
		Line 0 "Used by # Delivery Groups`t: " $DGCnt.ToString()
		Line 0 "URL`t`t`t`t: " $SFServer.Url
		Line 0 "Description`t`t`t: " $SFServer.Description
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Server: " $SFServer.Name
		$rowdata = @()
		$columnHeaders = @("StoreFront Server",($global:htmlsb),$SFServer.Name,$htmlwhite)
		$rowdata += @(,('Used by # Delivery Groups',($global:htmlsb),$DGCnt.ToString(),$htmlwhite))
		$rowdata += @(,('URL',($global:htmlsb),$SFServer.Url,$htmlwhite))
		$rowdata += @(,('Description',($global:htmlsb),$SFServer.Description,$htmlwhite))

		$msg = ""
		$columnWidths = @("150","300")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
	}
}

Function OutputStoreFrontDeliveryGroups
{
	Param([object] $SFInfo)
	
	$DeliveryGroups = @()
	ForEach($DGUid in $SFInfo.DesktopGroupUids)
	{
		$Results = Get-BrokerDesktopGroup -EA 0 -Uid $DGUid
		If($? -and $Null -ne $Results)
		{
			$DeliveryGroups += $Results.Name
		}
	}

	$DeliveryGroups = $DeliveryGroups | Sort-Object Name
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Delivery Groups"
		$DGWordTable = @()
	}
	If($Text)
	{
		Line 0 "Delivery Groups"
		Line 0 ""
		$cnt = -1
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Delivery Groups"
		$rowdata = @()
	}
	
	ForEach($Group in $DeliveryGroups)
	{
		If($MSWord -or $PDF)
		{
			$DGWordTable += @{DGName = $Group;}
		}
		If($Text)
		{
			$cnt++
			If($cnt -eq 0)
			{
				Line 1 "Delivery Group: " $Group
			}
			Else
			{
				Line 3 "" $Group
			}
		}
		If($HTML)
		{
			$rowdata += @(,($Group,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $DGWordTable `
		-Columns DGName `
		-Headers "Delivery Group" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		$columnHeaders = @(
		'Delivery Group',($global:htmlsb))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
}
#endregion

#region AppV functions
Function ProcessAppV
{
	Write-Verbose "$(Get-Date -Format G): Processing App-V"
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "App-V Publishing"
	}
	If($Text)
	{
		Line 0 "App-V Publishing"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "App-V Publishing"
	}
	
	Write-Verbose "$(Get-Date -Format G): `tRetrieving App-V configuration"
	$AppVConfigs = Get-BrokerMachineConfiguration -EA 0 -Name appv* 4>$Null
	
	If($? -and $Null -ne $AppVConfigs)
	{
		Write-Verbose "$(Get-Date -Format G): `t`tRetrieving App-V server information"
		
		$AppVs = New-Object System.Collections.ArrayList
		ForEach($AppVConfig in $AppVConfigs)
		{
			$AppV = Get-CtxAppVServer -ByteArray $AppVConfig.Policy -EA 0 4>$Null
			If($? -and ($Null -ne $AppV))
			{
				$obj = [PSCustomObject] @{
					MgmtServer = $AppV.ManagementServer				
					PubServer  = $AppV.PublishingServer				
				}
				$null = $AppVs.Add($obj)
			}
			ElseIf($? -and ($Null -eq $AppV))
			{
				$txt = "There was no App-V server information found for $($AppVConfig.Policy)"
				OutputWarning $txt
			}
			Else
			{
				$txt = "Unable to retrieve App-V information for $($AppVConfig.Policy)"
				OutputWarning $txt
			}
		}
		
		$AppVs = $AppVs | Sort-Object MgmtServer
		
		OutputAppV $AppVs
	}
	ElseIf($? -and $Null -eq $AppVConfigs)
	{
		$txt = "App-V is not configured for this Site"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve App-V configuration"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date -Format G): "
}

Function OutputAppV
{
	Param([object]$AppVs)
	
	Write-Verbose "$(Get-Date -Format G): `t`t`tOutput App-V server information"
	If($MSWord -or $PDF)
	{
		$AppVWordTable = @()
	}
	If($HTML)
	{
		$rowdata = @()
	}
	
	ForEach($AppV in $AppVs)
	{
		Write-Verbose "$(Get-Date -Format G): `t`t`tAdding AppV Server $($AppV.MgmtServer)"

		If($MSWord -or $PDF)
		{
			$AppVWordTable += @{
			MgmtServer = $AppV.MgmtServer; 
			PubServer = $AppV.PubServer; 
			}
		}
		If($Text)
		{
			Line 1 "App-V management server`t: " $AppV.MgmtServer
			Line 1 "App-V publishing server`t: " $AppV.PubServer
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,(
			$AppV.MgmtServer,$htmlwhite,
			$AppV.PubServer,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $AppVWordTable `
		-Columns  MgmtServer,PubServer `
		-Headers  "App-V management server","App-V publishing server" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($HTML)
	{
		$columnHeaders = @(
		'App-V management server',($global:htmlsb),
		'App-V publishing server',($global:htmlsb))

		$msg = ""
		$columnWidths = @("250","250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
	}
}
#endregion

#region zones
Function ProcessZones
{
	Write-Verbose "$(Get-Date -Format G): Processing Zones"
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Zones"
	}
	If($Text)
	{
		Line 0 "Zones"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Zones"
	}
	
	#get all zone names
	Write-Verbose "$(Get-Date -Format G): `tRetrieving All Zones"
	$Zones = Get-ConfigZone -EA 0 -Filter {Name -ne "Initial Zone" -and Name -ne "00000000-0000-0000-0000-000000000000"} -SortBy Name
	$ZoneMembers = New-Object System.Collections.ArrayList
	
	ForEach($Zone in $Zones)
	{
		$Script:TotalZones++
		Write-Verbose "$(Get-Date -Format G): `t`tRetrieving Machine Catalogs for Zone $($Zone.Name)"
		$ZoneCatalogs = Get-BrokerCatalog @CCParams2 -ZoneUid $Zone.Uid
		ForEach($ZoneCatalog in $ZoneCatalogs)
		{
			$obj = [PSCustomObject] @{
				MemName = $ZoneCatalog.Name			
				MemDesc = $ZoneCatalog.Description			
				MemType = "Machine Catalog"			
				MemZone = $Zone.Name			
			}
			$null = $ZoneMembers.Add($obj)
		}
		
		Write-Verbose "$(Get-Date -Format G): `t`tRetrieving Host Connections for Zone $($Zone.Name)"
		$ZoneHosts = Get-ChildItem -EA 0 -path 'xdhyp:\connections' 4>$Null | Where-Object{$_.ZoneUid -eq $Zone.Uid}
		ForEach($ZoneHost in $ZoneHosts)
		{
			$obj = [PSCustomObject] @{
				MemName = $ZoneHost.HypervisorConnectionName			
				MemDesc = ""			
				MemType = "Host Connection"		
				MemZone = $Zone.Name			
			}
			$null = $ZoneMembers.Add($obj)
		}

		Write-Verbose "$(Get-Date -Format G): `t`tRetrieving Cloud Connectors for Zone $($Zone.Name)"
		$ZoneCCs = Get-ConfigEdgeServer -ZoneUid $Zone.Uid
		ForEach($ZoneCC in $ZoneCCs)
		{
			$obj = [PSCustomObject] @{
				MemName = $ZoneCC.MachineAddress
				MemDesc = ""		
				MemType = "Citrix Cloud Connector"		
				MemZone = $Zone.Name			
			}
			$null = $ZoneMembers.Add($obj)
		}
	}
	
	OutputZoneSiteView $ZoneMembers
	
	OutputPerZoneView $ZoneMembers $Zones
}

Function OutputZoneSiteView
{
	Param([array]$ZoneMembers)
	
	Write-Verbose "$(Get-Date -Format G): `tOutput Zone Site View"
	$ZoneMembers = $ZoneMembers | Sort-Object MemType, MemName
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Site View"
		$ZoneWordTable = @()

		ForEach($ZoneMember in $ZoneMembers)
		{
			$ZoneWordTable += @{ 
			xName = $ZoneMember.MemName;
			xDesc = $ZoneMember.MemDesc;
			xType = $ZoneMember.MemType;
			xZone = $ZoneMember.MemZone;
			}
		}

		$Table = AddWordTable -Hashtable $ZoneWordTable `
		-Columns xName, xDesc, xType, xZone `
		-Headers "Name", "Description", "Type", "Zone" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 175;
		$Table.Columns.Item(2).Width = 100;
		$Table.Columns.Item(3).Width = 125;
		$Table.Columns.Item(4).Width = 100;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($Text)
	{
		Line 0 "Site View"
		Line 0 ""
		ForEach($ZoneMember in $ZoneMembers)
		{
			Line 1 "Name`t`t: " $ZoneMember.MemName
			Line 1 "Description`t: " $ZoneMember.MemDesc
			Line 1 "Type`t`t: " $ZoneMember.MemType
			Line 1 "Zone`t`t: " $ZoneMember.MemZone
			Line 0 ""
		}
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Site View"
		$rowdata = @()
		ForEach($ZoneMember in $ZoneMembers)
		{
			$rowdata += @(,(
			$ZoneMember.MemName,$htmlwhite,
			$ZoneMember.MemDesc,$htmlwhite,
			$ZoneMember.MemType,$htmlwhite,
			$ZoneMember.MemZone,$htmlwhite))
		}
		
		$columnHeaders = @(
		'Name',($global:htmlsb),
		'Description',($global:htmlsb),
		'Type',($global:htmlsb),
		'Zone',($global:htmlsb))

		$msg = ""
		$columnWidths = @("150","200","150","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "650"
	}
}

Function OutputPerZoneView
{
	Param([array]$ZoneMembers, [object]$Zones)
	
	Write-Verbose "$(Get-Date -Format G): `tOutput Per Zone View"
	$ZoneMembers = $ZoneMembers | Sort-Object MemZone, MemType, MemName

	ForEach($Zone in $Zones)
	{
		$TmpZoneMembers = $ZoneMembers | Where-Object{$_.MemZone -eq $Zone.Name}
		
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 $Zone.Name
			If($TmpZoneMembers -isnot [array])
			{
				WriteWordLine 0 0 "There are no zone members for Zone " $Zone.Name
			}
			Else
			{
				$ZoneWordTable = @()

				ForEach($ZoneMember in $TmpZoneMembers)
				{
					$ZoneWordTable += @{ 
					xName = $ZoneMember.MemName;
					xDesc = $ZoneMember.MemDesc;
					xType = $ZoneMember.MemType;
					}
				}

				$Table = AddWordTable -Hashtable $ZoneWordTable `
				-Columns xName, xDesc, xType `
				-Headers "Name", "Description", "Type" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 175;
				$Table.Columns.Item(2).Width = 100;
				$Table.Columns.Item(3).Width = 125;
				
				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 $Zone.Name
			Line 0 ""
			
			If($TmpZoneMembers -isnot [array])
			{
				Line 1 "There are no zone members for Zone " $Zone.Name
				Line 0 ""
			}
			Else
			{
				ForEach($ZoneMember in $TmpZoneMembers)
				{
					Line 1 "Name`t`t: " $ZoneMember.MemName
					Line 1 "Description`t: " $ZoneMember.MemDesc
					Line 1 "Type`t`t: " $ZoneMember.MemType
					Line 0 ""
				}
			}
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 $Zone.Name
			If($TmpZoneMembers -isnot [array])
			{
				WriteHTMLLine 0 0 "There are no zone members for Zone " $Zone.Name
			}
			Else
			{
				$rowdata = @()
				ForEach($ZoneMember in $TmpZoneMembers)
				{
					$rowdata += @(,(
					$ZoneMember.MemName,$htmlwhite,
					$ZoneMember.MemDesc,$htmlwhite,
					$ZoneMember.MemType,$htmlwhite))
				}
				
				$columnHeaders = @(
				'Name',($global:htmlsb),
				'Description',($global:htmlsb),
				'Type',($global:htmlsb))

				$msg = ""
				$columnWidths = @("150","200","150")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
			}
		}
	}
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region summary page
Function OutputSummaryPage
{
	#summary page
	Write-Verbose "$(Get-Date -Format G): Create Summary Page"
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Citrix Cloud $($CCSiteName) Summary Page"
	}
	If($Text)
	{
		Line 0 ""
		Line 0 "Citrix Cloud $($CCSiteName) Summary Page"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Citrix Cloud $($CCSiteName) Summary Page"
	}

	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		WriteWordLine 4 0 "Machine Catalogs"
		$ScriptInformation.Add(@{Data = "Total Multi-session OS Catalogs"; Value = $Script:TotalServerOSCatalogs.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = 'Total Single-session OS Catalogs'; Value = $Script:TotalDesktopOSCatalogs.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = 'Total RemotePC Catalogs'; Value = $Script:TotalRemotePCCatalogs.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = '     Total Machine Catalogs'; Value = ($Script:TotalServerOSCatalogs+$Script:TotalDesktopOSCatalogs+$Script:TotalRemotePCCatalogs).ToString(); }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 50;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
		
		$ScriptInformation = New-Object System.Collections.ArrayList
		WriteWordLine 4 0 "Delivery Groups"
		$ScriptInformation.Add(@{Data = "Total Application Groups"; Value = $Script:TotalApplicationGroups.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = 'Total Desktop Groups'; Value = $Script:TotalDesktopGroups.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = 'Total Apps & Desktop Groups'; Value = $Script:TotalAppsAndDesktopGroups.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = '     Total Delivery Groups'; Value = ($Script:TotalApplicationGroups+$Script:TotalDesktopGroups+$Script:TotalAppsAndDesktopGroups).ToString(); }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 50;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		$ScriptInformation = New-Object System.Collections.ArrayList
		WriteWordLine 4 0 "Applications"
		$ScriptInformation.Add(@{Data = "Total Published Applications"; Value = $Script:TotalPublishedApplications.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = 'Total App-V Applications'; Value = $Script:TotalAppvApplications.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = '     Total Applications'; Value = ($Script:TotalPublishedApplications + $Script:TotalAppvApplications).ToString(); }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 50;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
		
		If($Policies -eq $True)
		{
			$ScriptInformation = New-Object System.Collections.ArrayList
			WriteWordLine 4 0 "Policies"
			$ScriptInformation.Add(@{Data = "Total Computer Policies"; Value = $Script:TotalComputerPolicies.ToString(); }) > $Null
			$ScriptInformation.Add(@{Data = 'Total User Policies'; Value = $Script:TotalUserPolicies.ToString(); }) > $Null
			$ScriptInformation.Add(@{Data = '     Total Policies'; Value = $Script:TotalPolicies.ToString(); }) > $Null
			$ScriptInformation.Add(@{Data = ''; Value = ""; }) > $Null
			$ScriptInformation.Add(@{Data = 'Site Policies'; Value = $Script:TotalSitePolicies.ToString(); }) > $Null
			If($NoADPolicies -eq $False)
			{
				$ScriptInformation.Add(@{Data = "Citrix AD Policies Processed "; Value = "$($Script:TotalADPolicies) *"; }) > $Null
				$ScriptInformation.Add(@{Data = 'Citrix AD Policies not Processed'; Value = $Script:TotalADPoliciesNotProcessed.ToString(); }) > $Null
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 250;
			$Table.Columns.Item(2).Width = 50;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 '* (AD Policies can contain multiple Citrix policies)' "" $Null 8 $False $True
			WriteWordLine 0 0 ""
		}
		
		$ScriptInformation = New-Object System.Collections.ArrayList
		WriteWordLine 4 0 "Administrators"
		$ScriptInformation.Add(@{Data = "Total Cloud Admins"; Value = $Script:TotalCloudAdmins.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Total Delivery Group Admins"; Value = $Script:TotalDeliveryGroupAdmins.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = 'Total Full Admins'; Value = $Script:TotalFullAdmins.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = 'Total Full Monitor Admins'; Value = $Script:TotalFullMonitorAdmins.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = 'Total Help Desk Admins'; Value = $Script:TotalHelpDeskAdmins.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = 'Total Host Admins'; Value = $Script:TotalHostAdmins.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = 'Total Machine Catalog Admins'; Value = $Script:TotalMachineCatalogAdmins.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Total Probe Agent Admins"; Value = $Script:TotalProbeAdmins.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = 'Total Read Only Admins'; Value = $Script:TotalReadOnlyAdmins.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Total Session Admins"; Value = $Script:TotalSessionAdmins.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = 'Total Custom Admins'; Value = $Script:TotalCustomAdmins.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = '     Total Administrators'; Value = (
		$Script:TotalCloudAdmins+
		$Script:TotalDeliveryGroupAdmins+
		$Script:TotalFullAdmins+
		$Script:TotalFullMonitorAdmins+
		$Script:TotalHelpDeskAdmins+
		$Script:TotalHostAdmins+
		$Script:TotalMachineCatalogAdmins+
		$Script:TotalProbeAdmins+
		$Script:TotalReadOnlyAdmins+
		$Script:TotalSessionAdmins+
		$Script:TotalCustomAdmins
		).ToString(); }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 50;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		$ScriptInformation = New-Object System.Collections.ArrayList
		WriteWordLine 4 0 "Hosting Connections"
		$ScriptInformation.Add(@{Data = "     Total Hosting Connections"; Value = $Script:TotalHostingConnections.ToString(); }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 50;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		$ScriptInformation = New-Object System.Collections.ArrayList
		WriteWordLine 4 0 "StoreFront"
		$ScriptInformation.Add(@{Data = "     Total StoreFront Servers"; Value = $Script:TotalStoreFrontServers.ToString(); }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 50;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		$ScriptInformation = New-Object System.Collections.ArrayList
		WriteWordLine 4 0 "Zones"
		$ScriptInformation.Add(@{Data = "     Total Zones"; Value = $Script:TotalZones.ToString(); }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 50;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "Machine Catalogs"
		Line 1 "Total Multi-session OS Catalogs`t: " $Script:TotalServerOSCatalogs
		Line 1 "Total Single-session OS Catalogs: " $Script:TotalDesktopOSCatalogs
		Line 1 "Total RemotePC Catalogs`t`t: " $Script:TotalRemotePCCatalogs
		Line 2 "Total Machine Catalogs`t: " ($Script:TotalServerOSCatalogs+$Script:TotalDesktopOSCatalogs+$Script:TotalRemotePCCatalogs)
		Line 0 ""
		Line 0 "Delivery Groups"
		Line 1 "Total Application Groups`t: " $Script:TotalApplicationGroups
		Line 1 "Total Desktop Groups`t`t: " $Script:TotalDesktopGroups
		Line 1 "Total Apps & Desktop Groups`t: " $Script:TotalAppsAndDesktopGroups
		Line 2 "Total Delivery Groups`t: " ($Script:TotalApplicationGroups+$Script:TotalDesktopGroups+$Script:TotalAppsAndDesktopGroups)
		Line 0 ""
		Line 0 "Applications"
		Line 1 "Total Published Applications`t: " $Script:TotalPublishedApplications
		Line 1 "Total App-V Applications`t: " $Script:TotalAppvApplications
		Line 2 "Total Applications`t: " ($Script:TotalPublishedApplications + $Script:TotalAppvApplications)
		Line 0 ""
		
		If($Policies -eq $True)
		{
			Line 0 "Policies"
			Line 1 "Total Computer Policies`t`t: " $Script:TotalComputerPolicies
			Line 1 "Total User Policies`t`t: " $Script:TotalUserPolicies
			Line 2 "Total Policies`t`t: " $Script:TotalPolicies
			Line 0 ""
			Line 1 "Site Policies`t`t`t: " $Script:TotalSitePolicies
			
			If($NoADPolicies -eq $False)
			{
				Line 1 "Citrix AD Policies Processed`t: $($Script:TotalADPolicies)`t (AD Policies can contain multiple Citrix policies)"
				Line 1 "Citrix AD Policies not Processed: " $Script:TotalADPoliciesNotProcessed
			}
			Line 0 ""
		}
		
		Line 0 "Administrators"
		Line 1 "Total Cloud Admins`t`t: " $Script:TotalCloudAdmins
		Line 1 "Total Delivery Group Admins`t: " $Script:TotalDeliveryGroupAdmins
		Line 1 "Total Full Admins`t`t: " $Script:TotalFullAdmins
		Line 1 "Total Full Monitor Admins`t: " $Script:TotalFullMonitorAdmins
		Line 1 "Total Help Desk Admins`t`t: " $Script:TotalHelpDeskAdmins
		Line 1 "Total Host Admins`t`t: " $Script:TotalHostAdmins
		Line 1 "Total Machine Catalog Admins`t: " $Script:TotalMachineCatalogAdmins
		Line 1 "Total Probe Agent Admins`t: " $Script:TotalProbeAdmins
		Line 1 "Total Read Only Admins`t`t: " $Script:TotalReadOnlyAdmins
		Line 1 "Total Session Admins`t`t: " $Script:TotalSessionAdmins
		Line 1 "Total Custom Admins`t`t: " $Script:TotalCustomAdmins
		Line 2 "Total Administrators`t: " (
		$Script:TotalCloudAdmins+
		$Script:TotalDeliveryGroupAdmins+
		$Script:TotalFullAdmins+
		$Script:TotalFullMonitorAdmins+
		$Script:TotalHelpDeskAdmins+
		$Script:TotalHostAdmins+
		$Script:TotalMachineCatalogAdmins+
		$Script:TotalProbeAdmins+
		$Script:TotalReadOnlyAdmins+
		$Script:TotalSessionAdmins+
		$Script:TotalCustomAdmins
		)
		Line 0 ""
		Line 0 "Hosting Connections"
		Line 1 "Total Hosting Connections`t: " $Script:TotalHostingConnections
		Line 0 ""
		Line 0 "StoreFront"
		Line 1 "Total StoreFront Servers`t: " $Script:TotalStoreFrontServers
		Line 0 ""
		Line 0 "Zones"
		Line 1 "Total Zones`t`t`t: " $Script:TotalZones
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Total Multi-session OS Catalogs",($global:htmlsb),$Script:TotalServerOSCatalogs.ToString(),$htmlwhite)
		$rowdata += @(,('Total Single-session OS Catalogs',($global:htmlsb),$Script:TotalDesktopOSCatalogs.ToString(),$htmlwhite))
		$rowdata += @(,('Total RemotePC Catalogs',($global:htmlsb),$Script:TotalRemotePCCatalogs.ToString(),$htmlwhite))
		$rowdata += @(,('     Total Machine Catalogs',($global:htmlsb),($Script:TotalServerOSCatalogs+$Script:TotalDesktopOSCatalogs+$Script:TotalRemotePCCatalogs).ToString(),$htmlwhite))

		$msg = "Machine Catalogs"
		$columnWidths = @("250","50")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		
		$rowdata = @()
		$columnHeaders = @("Total Application Groups",($global:htmlsb),$Script:TotalApplicationGroups.ToString(),$htmlwhite)
		$rowdata += @(,('Total Desktop Groups',($global:htmlsb),$Script:TotalDesktopGroups.ToString(),$htmlwhite))
		$rowdata += @(,('Total Apps & Desktop Groups',($global:htmlsb),$Script:TotalAppsAndDesktopGroups.ToString(),$htmlwhite))
		$rowdata += @(,('     Total Delivery Groups',($global:htmlsb),($Script:TotalApplicationGroups+$Script:TotalDesktopGroups+$Script:TotalAppsAndDesktopGroups).ToString(),$htmlwhite))

		$msg = "Delivery Groups"
		$columnWidths = @("250","50")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths

		$rowdata = @()
		$columnHeaders = @("Total Published Applications",($global:htmlsb),$Script:TotalPublishedApplications.ToString(),$htmlwhite)
		$rowdata += @(,('Total App-V Applications',($global:htmlsb),$Script:TotalAppvApplications.ToString(),$htmlwhite))
		$rowdata += @(,('     Total Applications',($global:htmlsb),($Script:TotalPublishedApplications + $Script:TotalAppvApplications).ToString(),$htmlwhite))

		$msg = "Applications"
		$columnWidths = @("250","50")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		
		If($Policies -eq $True)
		{
			$rowdata = @()
			$columnHeaders = @("Total Computer Policies",($global:htmlsb),$Script:TotalComputerPolicies.ToString(),$htmlwhite)
			$rowdata += @(,('Total User Policies',($global:htmlsb),$Script:TotalUserPolicies.ToString(),$htmlwhite))
			$rowdata += @(,('     Total Policies',($global:htmlsb),$Script:TotalPolicies.ToString(),$htmlwhite))
			$rowdata += @(,('',($global:htmlsb),"",$htmlwhite))
			$rowdata += @(,('Site Policies',($global:htmlsb),$Script:TotalSitePolicies.ToString(),$htmlwhite))
			If($NoADPolicies -eq $False)
			{
				$rowdata += @(,("Citrix AD Policies Processed ",($global:htmlsb),"$($Script:TotalADPolicies) *",$htmlwhite))
				$rowdata += @(,('Citrix AD Policies not Processed',($global:htmlsb),$Script:TotalADPoliciesNotProcessed.ToString(),$htmlwhite))
			}

			$msg = "Policies"
			$columnWidths = @("250","50")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 '* (AD Policies can contain multiple Citrix policies)' "" "Calibri" 1 $htmlbold
		}
		
		$rowdata = @()
		$columnHeaders = @("Total Cloud Admins",($global:htmlsb),$Script:TotalCloudAdmins.ToString(),$htmlwhite)
		$rowdata += @(,('Total Delivery Group Admins',($global:htmlsb),$Script:TotalDeliveryGroupAdmins.ToString(),$htmlwhite))
		$rowdata += @(,('Total Full Admins',($global:htmlsb),$Script:TotalFullAdmins.ToString(),$htmlwhite))
		$rowdata += @(,('Total Full Monitor Admins',($global:htmlsb),$Script:TotalFullMonitorAdmins.ToString(),$htmlwhite))
		$rowdata += @(,('Total Help Desk Admins',($global:htmlsb),$Script:TotalHelpDeskAdmins.ToString(),$htmlwhite))
		$rowdata += @(,('Total Host Admins',($global:htmlsb),$Script:TotalHostAdmins.ToString(),$htmlwhite))
		$rowdata += @(,('Total Machine Catalog Admins',($global:htmlsb),$Script:TotalMachineCatalogAdmins.ToString(),$htmlwhite))
		$rowdata += @(,('Total Probe Agent Admins',($global:htmlsb),$Script:TotalProbeAdmins.ToString(),$htmlwhite))
		$rowdata += @(,('Total Read Only Admins',($global:htmlsb),$Script:TotalReadOnlyAdmins.ToString(),$htmlwhite))
		$rowdata += @(,('Total Session Admins',($global:htmlsb),$Script:TotalSessionAdmins.ToString(),$htmlwhite))
		$rowdata += @(,('Total Custom Admins',($global:htmlsb),$Script:TotalCustomAdmins.ToString(),$htmlwhite))
		$rowdata += @(,('     Total Administrators',($global:htmlsb),(
		$Script:TotalCloudAdmins+
		$Script:TotalDeliveryGroupAdmins+
		$Script:TotalFullAdmins+
		$Script:TotalFullMonitorAdmins+
		$Script:TotalHelpDeskAdmins+
		$Script:TotalHostAdmins+
		$Script:TotalMachineCatalogAdmins+
		$Script:TotalProbeAdmins+
		$Script:TotalReadOnlyAdmins+
		$Script:TotalSessionAdmins+
		$Script:TotalCustomAdmins
		).ToString(),$htmlwhite))

		$msg = "Administrators"
		$columnWidths = @("250","50")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths

		$rowdata = @()
		$columnHeaders = @("     Total Hosting Connections",($global:htmlsb),$Script:TotalHostingConnections.ToString(),$htmlwhite)

		$msg = "Hosting Connections"
		$columnWidths = @("250","50")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths

		$rowdata = @()
		$columnHeaders = @("     Total Zones",($global:htmlsb),$Script:TotalZones.ToString(),$htmlwhite)

		$msg = "Zones"
		$columnWidths = @("250","50")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
	}

	Write-Verbose "$(Get-Date -Format G): Finished Create Summary Page"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region script setup function
Function ProcessScriptSetup
{
	$script:startTime = Get-Date

	#borrowed from Martic Zugec
	Write-Verbose "$(Get-Date -Format G): Testing required PowerShell modules for Citrix Cloud"
	If ($(Get-Module Citrix.PoshSdkProxy.Commands -ListAvailable) -isnot [System.Management.Automation.PSModuleInfo]) 
	{
		Write-Host "Citrix Virtual Apps and Desktops Remote PowerShell SDK is required for Citrix Cloud connection." -ForegroundColor Red
		Write-Host "Script will abort. You can try again after installing the required SDK from https://download.apps.cloud.com/CitrixPoshSdk.exe" -ForegroundColor Red

		Write-Error "
	`n`r
	CVADS Remote Powershell SDK is not available
	`n`n
	If you are running XA/XD 7.0 through 7.7, please use: 
	https://carlwebster.com/downloads/download-info/xenappxendesktop-7-x-documentation-script/
	`n`n
	If you are running XA/XD 7.8 through CVAD 2006, please use: 
	https://carlwebster.com/downloads/download-info/xenappxendesktop-7-8/
	`n`n
	If you are running CVAD 2006 and later, please use:
	https://carlwebster.com/downloads/download-info/citrix-virtual-apps-and-desktops-v3-script/
	`n`n
	Script will now close.
	`n
		"
		Exit
	} 

	Write-Verbose "$(Get-Date -Format G): Loading Citrix.Common.GroupPolicy PSSnapin"
	If(!(Check-NeededPSSnapins "Citrix.Common.GroupPolicy"))
	{
		#We're missing Citrix Snapins that we need
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
	`n`n
	Missing Citrix PowerShell Snap-ins Detected. 
	`n`n
	Install the Citrix Group Policy Management Console from the CVAD 2006 or later installation media. 
	Note: This is required by the StoreFront and Citrix Policy cmdlets and functions.
	x:\x64\Citrix Policy\CitrixGroupPolicyManagement_x64.msi
	Installing this console installs the required Citrix.Common.GroupPolicy PowerShell snapin.
	`n`n
	If you are running XA/XD 7.0 through 7.7, please use: 
	https://carlwebster.com/downloads/download-info/xenappxendesktop-7-x-documentation-script/
	`n`n
	If you are running XA/XD 7.8 through CVAD 2006, please use: 
	https://carlwebster.com/downloads/download-info/xenappxendesktop-7-8/
	`n`n
	If you are running CVAD 2006 and later, please use:
	https://carlwebster.com/downloads/download-info/citrix-virtual-apps-and-desktops-v3-script/
	`n`n
	Script will now close.
	`n
		"
		Exit
	}

	Write-Verbose "$(Get-Date -Format G): Importing required Citrix PowerShell modules"
	Write-Verbose "$(Get-Date -Format G): `tCitrix.ADIdentity.Commands"
	Import-Module "Citrix.ADIdentity.Commands" 4>$Null
	If(!$?)
	{
		Write-Error "Unable to import the Citrix.ADIdentity.Commands module. Script cannot continue."
		Exit
	}
	
	Write-Verbose "$(Get-Date -Format G): `tCitrix.Analytics.Commands"
	Import-Module "Citrix.Analytics.Commands" 4>$Null
	If(!$?)
	{
		Write-Error "Unable to import the Citrix.Analytics.Commands module. Script cannot continue."
		Exit
	}
	
	Write-Verbose "$(Get-Date -Format G): `tCitrix.AppLibrary.Commands"
	Import-Module "Citrix.AppLibrary.Commands" 4>$Null
	If(!$?)
	{
		Write-Error "Unable to import the Citrix.AppLibrary.Commands module. Script cannot continue."
		Exit
	}
	
	Write-Verbose "$(Get-Date -Format G): `tCitrix.Broker.Commands"
	Import-Module "Citrix.Broker.Commands" 4>$Null
	If(!$?)
	{
		Write-Error "Unable to import the Citrix.Broker.Commands module. Script cannot continue."
		Exit
	}
	
	Write-Verbose "$(Get-Date -Format G): `tCitrix.Common.Commands"
	Import-Module "Citrix.Common.Commands" 4>$Null
	If(!$?)
	{
		Write-Error "Unable to import the Citrix.Common.Commands module. Script cannot continue."
		Exit
	}
	
	Write-Verbose "$(Get-Date -Format G): `tCitrix.Configuration.Commands"
	Import-Module "Citrix.Configuration.Commands" 4>$Null
	If(!$?)
	{
		Write-Error "Unable to import the Citrix.Configuration.Commands module. Script cannot continue."
		Exit
	}
	
	Write-Verbose "$(Get-Date -Format G): `tCitrix.ConfigurationLogging.Commands"
	Import-Module "Citrix.ConfigurationLogging.Commands" 4>$Null
	If(!$?)
	{
		Write-Error "Unable to import the Citrix.ConfigurationLogging.Commands module. Script cannot continue."
		Exit
	}
	
	Write-Verbose "$(Get-Date -Format G): `tCitrix.DelegatedAdmin.Commands"
	Import-Module "Citrix.DelegatedAdmin.Commands" 4>$Null
	If(!$?)
	{
		Write-Error "Unable to import the Citrix.DelegatedAdmin.Commands module. Script cannot continue."
		Exit
	}
	
	Write-Verbose "$(Get-Date -Format G): `tCitrix.Host.Commands"
	Import-Module "Citrix.Host.Commands" 4>$Null
	If(!$?)
	{
		Write-Error "Unable to import the Citrix.Host.Commands module. Script cannot continue."
		Exit
	}
	
	Write-Verbose "$(Get-Date -Format G): `tCitrix.MachineCreation.Commands"
	Import-Module "Citrix.MachineCreation.Commands" 4>$Null
	If(!$?)
	{
		Write-Error "Unable to import the  Citrix.MachineCreation.Commandsmodule. Script cannot continue."
		Exit
	}
	
	Write-Verbose "$(Get-Date -Format G): `tCitrix.Monitor.Commands"
	Import-Module "Citrix.Monitor.Commands" 4>$Null
	If(!$?)
	{
		Write-Error "Unable to import the Citrix.Monitor.Commands module. Script cannot continue."
		Exit
	}
	
	Write-Verbose "$(Get-Date -Format G): `tCitrix.Storefront.Commands"
	Import-Module "Citrix.Storefront.Commands" 4>$Null
	If(!$?)
	{
		Write-Error "Unable to import the Citrix.Storefront.Commands module. Script cannot continue."
		Exit
	}
	
	$Script:DoPolicies = $False
	If($NoPolicies)
	{
		Write-Verbose "$(Get-Date -Format G): NoPolicies was specified so do not search for Citrix.GroupPolicy.Commands.psm1"
		$Script:DoPolicies = $False
	}

	If($Policies -eq $True)
	{
		If(Test-Path "$([Environment]::GetFolderPath( [Environment+SpecialFolder]::System))\WindowsPowerShell\v1.0\Modules\Citrix.GroupPolicy.Commands\Citrix.GroupPolicy.Commands.psm1")
		{
			Write-Verbose "Importing module: $([Environment]::GetFolderPath( [Environment+SpecialFolder]::System))\WindowsPowerShell\v1.0\Modules\Citrix.GroupPolicy.Commands\Citrix.GroupPolicy.Commands.psm1"

			Import-Module "$([Environment]::GetFolderPath( [Environment+SpecialFolder]::System))\WindowsPowerShell\v1.0\Modules\Citrix.GroupPolicy.Commands\Citrix.GroupPolicy.Commands.psm1" 4>$Null
			If(!$?)
			{
				Write-Warning "
		`n
		Citrix Group Policy module not loaded:
		$([Environment]::GetFolderPath( [Environment+SpecialFolder]::System))\WindowsPowerShell\v1.0\Modules\Citrix.GroupPolicy.Commands\Citrix.GroupPolicy.Commands.psm1 could not be loaded 
		`n
		Please see the Prerequisites section in the ReadMe file: https://carlwebster.sharefile.com/d-sb4e144f9ecc48e7b.
		`n
		Because the Policies parameter was used, Policies will not be processed.
		"
				Write-Verbose "$(Get-Date -Format G): "
				$Script:DoPolicies = $False
			}
			Else
			{
				$Script:DoPolicies = $True
			}
		}
		Else
		{
			Write-Warning "
		`n
		Citrix Group Policy module was not found:
		$([Environment]::GetFolderPath( [Environment+SpecialFolder]::System))\WindowsPowerShell\v1.0\Modules\Citrix.GroupPolicy.Commands.psm1
		`n
		Please see the Prerequisites section in the ReadMe file: https://carlwebster.sharefile.com/d-sb4e144f9ecc48e7b.
		`n
		Because the Policies parameter was used, Policies will not be processed.
		"
			Write-Verbose "$(Get-Date -Format G): "
			$Script:DoPolicies = $False
		}
	}
	
	If($Policies -eq $False -and $NoPolicies -eq $False -and $NoADPolicies -eq $False)
	{
		#script defaults, so don't process policies
		$Script:DoPolicies = $False
	}

	If($NoPolicies -eq $True)
	{
		#don't process policies
		$Script:DoPolicies = $False
	}
	
	#set value for MaxRecordCount
	$Script:MaxRecordCount = [int]::MaxValue 

	$Script:CCParams2 = @{
	EA = 0;
	MaxRecordCount = $Script:MaxRecordCount;
	}

	# Get Site information
	Write-Verbose "$(Get-Date -Format G): Gathering initial Site data"

	$Script:CCSite1 = Get-BrokerSite -EA 0

#Do not indent the following write-error lines. Doing so will mess up the console formatting of the error message.
	If( !($?) -or $Null -eq $Script:CCSite1)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Warning "CC Site1 information could not be retrieved. Script cannot continue"
		Write-Error "
	`n`n
Unable to connect to your Citrix Cloud Account $($Script:CustomerId)
	`n`n
cmdlet failed: $($error[ 0 ].ToString())
	`n`n
		"
		AbortScript
	}

	$Script:CCSite2 = Get-ConfigSite -EA 0

	If( !($?) -or $Null -eq $Script:CCSite2)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Warning "CC Site2 information could not be retrieved. Script cannot continue"
		Write-Error "
	`n`n
cmdlet failed $($error[ 0 ].ToString())
	`n`n
		"
		AbortScript
	}

	$Script:CCSiteVersion = $Script:CCSite2.ProductVersion
	$tmp = $Script:CCSiteVersion.Split(".")
	[int]$MajorVersion = $tmp[0]
	[int]$MinorVersion = $tmp[1]

	If($MajorVersion -eq 7)
	{
		#this is a CVAD 7.x Site, now test to see if it is less than 7.26 (CVAD 2006)
		If($MinorVersion -lt 26)
		{
			Write-Warning "You are running version $Script:CCSiteVersion"
			Write-Warning "Are the PowerShell Snapins/Modules or Studio installed?"
			Write-Error "
	`n`n
This script is designed for CVADS.
	`n`n
Script cannot continue
	`n`n
			"
			AbortScript
		}
	}
	ElseIf($MajorVersion -eq 0 -and $MinorVersion -eq 0)
	{
		#something is wrong, we shouldn't be here
		Write-Error "
	`n`n
Something bad happened. We shouldn't be here. Could not find the version information.
	`n`n
Script cannot continue
	`n`n
		"
		AbortScript
	}
	Else
	{
		#this is not a CVADS Site, script cannot proceed
		Write-Warning "You are running version $Script:CCSiteVersion"
		Write-Warning "Are the PowerShell Snapins/Modules or Studio installed?"
		Write-Error "
	`n`n
This script is designed for CVADS.
	`n`n
Script cannot continue
	`n`n
		"
		AbortScript
	}
	
	$tmp = $Script:CCSiteVersion
	Switch ($tmp)
	{
		"7.27"	{$Script:CCSiteVersion = "2009"; Break}
		"7.26"	{$Script:CCSiteVersion = "2006"; Break}
		"7.25"	{$Script:CCSiteVersion = "2003"; Break}
		"7.24"	{$Script:CCSiteVersion = "1912"; Break}
		"7.23"	{$Script:CCSiteVersion = "1909"; Break}
		"7.22"	{$Script:CCSiteVersion = "1906"; Break}
		"7.21"	{$Script:CCSiteVersion = "1903"; Break}
		"7.20"	{$Script:CCSiteVersion = "1811"; Break}
		"7.19"	{$Script:CCSiteVersion = "1808"; Break}
		Default	{$Script:CCSiteVersion = $tmp; Break}
	}
	Write-Verbose "$(Get-Date -Format G): You are running version $Script:CCSiteVersion"

	If($SiteName -ne "")
	{
		[string]$Script:CCSiteName = $SiteName
	}
	Else
	{
		[string]$Script:CCSiteName = $Script:CCSite2.SiteName
	}

	Switch ($Section)
	{
		"Admins"		{[string]$Script:Title = "Inventory Report for the $($Script:CCSiteName) Site (Administrators Only)"; Break}
		"Apps"			{[string]$Script:Title = "Inventory Report for the $($Script:CCSiteName) Site (Applications Only)"; Break}
		"AppV"			{[string]$Script:Title = "Inventory Report for the $($Script:CCSiteName) Site (App-V Only"; Break}
		"Catalogs"		{[string]$Script:Title = "Inventory Report for the $($Script:CCSiteName) Site (Machine Catalogs Only)"; Break}
		"Config"		{[string]$Script:Title = "Inventory Report for the $($Script:CCSiteName) Site (Configuration Only)"; Break}
		"Groups" 		{[string]$Script:Title = "Inventory Report for the $($Script:CCSiteName) Site (Delivery Groups Only)"; Break}
		"Hosting"		{[string]$Script:Title = "Inventory Report for the $($Script:CCSiteName) Site (Hosting Only)"; Break}
		"Licensing"		{[string]$Script:Title = "Inventory Report for the $($Script:CCSiteName) Site (Licensing Only)"; Break}
		"Policies"		{[string]$Script:Title = "Inventory Report for the $($Script:CCSiteName) Site (Policies Only)"; Break}
		"StoreFront"	{[string]$Script:Title = "Inventory Report for the $($Script:CCSiteName) Site (StoreFront Only)"; Break}
		"Zones"			{[string]$Script:Title = "Inventory Report for the $($Script:CCSiteName) Site (Zones Only)"; Break}
		"All"			{[string]$Script:Title = "Inventory Report for the $($Script:CCSiteName) Site"; Break}
	}
	Write-Verbose "$(Get-Date -Format G): Initial Site data has been gathered"
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date -Format G): Script has completed"
	Write-Verbose "$(Get-Date -Format G): "

	Write-Verbose "$(Get-Date -Format G): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date -Format G): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date -Format G): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$Script:pwdpath\CCInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime       : $($AddDateTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Administrators     : $($Administrators)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Applications       : $($Applications)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name       : $($Script:CoName)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Address    : $($CompanyAddress)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Email      : $($CompanyEmail)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Fax        : $($CompanyFax)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Phone      : $($CompanyPhone)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page         : $($CoverPage)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "CSV                : $($CSV)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Customer ID        : $($Script:CustomerID)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "DeliveryGroups     : $($DeliveryGroups)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev                : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile       : $($Script:DevErrorFile)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "DGUtilization      : $($DeliveryGroupsUtilization)" 4>$Null
		If($HTML)
		{
			Out-File -FilePath $SIFile -Append -InputObject "HTMLFilename       : $($Script:HTMLFileName)" 4>$Null
		}
		If($MSWord)
		{
			Out-File -FilePath $SIFile -Append -InputObject "WordFilename       : $($Script:WordFileName)" 4>$Null
		}
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "PDFFilename        : $($Script:PDFFileName)" 4>$Null
		}
		If($Text)
		{
			Out-File -FilePath $SIFile -Append -InputObject "TextFilename       : $($Script:TextFileName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder             : $($Folder)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From               : $($From)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Hosting            : $($Hosting)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                : $($Log)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Logging            : $($Logging)" 4>$Null
		If($Logging)
		{
			Out-File -FilePath $SIFile -Append -InputObject "   Start Date      : $($StartDate)" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "   End Date        : $($EndDate)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "MachineCatalogs    : $($MachineCatalogs)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "MaxDetails         : $($MaxDetails)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "NoADPolicies       : $($NoADPolicies)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "NoPolicies         : $($NoPolicies)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Policies           : $($Policies)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML       : $($HTML)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF        : $($PDF)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT       : $($TEXT)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD       : $($MSWORD)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info        : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Section            : $($Section)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Site Name          : $($CCSiteName)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port          : $($SmtpPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server        : $($SmtpServer)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title              : $($Script:Title)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To                 : $($To)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL            : $($UseSSL)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name          : $($UserName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "VDA Registry Keys  : $($VDARegistryKeys)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "CC Version         : $($Script:CCSiteVersion)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected        : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version       : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture          : $($PSCulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture        : $($PSUICulture)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language      : $($Script:WordLanguageValue)" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version       : $($Script:WordProduct)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start       : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time       : $($Str)" 4>$Null
	}

	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $true) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date -Format G): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date -Format G): Transcript/log stop failed"
			}
		}
	}
	$ErrorActionPreference = $SaveEAPreference
}
#endregion

#region AppendixA
Function OutputAppendixA
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix A VDA Registry Items"

	If($CSV)
	{
		$CSVDone = $False
		$File = "$($Script:pwdpath)\$($CCSiteName)_Documentation_AppendixA_VDARegistryItems.csv"
		If($MSWord -or $PDF)
		{
			$Script:WordALLVDARegistryItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File *> $Null
			$CSVDone = $True
		}
		If($Text -and $CSVDone -eq $False)
		{
			$Script:TextALLVDARegistryItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File *> $Null
			$CSVDone = $True
		}
		If($HTML -and $CSVDone -eq $False)
		{
			$Script:HTMLALLVDARegistryItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File *> $Null
			$CSVDone = $True
		}
	}
	
	If($MSWord -or $PDF)
	{
		#sort the array by regkey, regvalue and servername
		$Script:WordALLVDARegistryItems = $Script:WordALLVDARegistryItems | Sort-Object RegKey, RegValue, VDAType, ComputerName
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix A - VDA Registry Items"
		WriteWordLine 0 0 "Miscellaneous Registry Items That May or May Not Exist on VDAs"
		WriteWordLine 0 0 "Linux VDAs are excluded"
		WriteWordLine 0 0 "These items may or may not be needed"
		WriteWordLine 0 0 "This Appendix is for VDA comparison only"
		WriteWordLine 0 0 ""
		
		$Save = ""
		$First = $True
		If($Script:WordAllVDARegistryItems)
		{
			$AppendixWordTable = @()
			ForEach($Item in $Script:WordALLVDARegistryItems)
			{
				If(!$First -and $Save -ne "$($Item.RegKey.ToString())$($Item.RegValue.ToString())$($Item.VDAType)")
				{
					$AppendixWordTable += @{ 
					RegKey = "";
					RegValue = "";
					RegData = "";
					VDAType = "";
					ComputerName = "";
					}
				}

				$AppendixWordTable += @{ 
				RegKey = $Item.RegKey;
				RegValue = $Item.RegValue;
				RegData = $Item.Value;
				VDAType = $Item.VDAType.Substring(0,1);
				ComputerName = $Item.ComputerName;
				}
				$Save = "$($Item.RegKey.ToString())$($Item.RegValue.ToString())$($Item.VDAType)"
				If($First)
				{
					$First = $False
				}
			}
			$Table = AddWordTable -Hashtable $AppendixWordTable `
			-Columns RegKey, RegValue, RegData, VDAType, ComputerName `
			-Headers "Registry Key", "Registry Value", "Data", "VDA", "Computer Name" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
	}
	If($Text)
	{
		#sort the array by regkey, regvalue and servername
		$Script:TextALLVDARegistryItems = $Script:TextALLVDARegistryItems | Sort-Object RegKey, RegValue, VDAType, ComputerName
		Line 0 "Appendix A - VDA Registry Items"
		Line 0 "Miscellaneous Registry Items That May or May Not Exist on VDAs"
		Line 0 "Linux VDAs are excluded"
		Line 0 "These items may or may not be needed"
		Line 0 "This Appendix is for VDA comparison only"
		Line 0 ""
		Line 1 "Registry Key                                                                                    Registry Value                                     Data                                                                                       VDA Type Computer Name                 " 
		Line 1 "====================================================================================================================================================================================================================================================================================="
		#       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345S12345678901234567890123456789012345678901234567890S123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890S12345678S123456789012345678901234567890
		
		$Save = ""
		$First = $True
		If($Script:TextAllVDARegistryItems)
		{
			ForEach($Item in $Script:TextALLVDARegistryItems)
			{
				If(!$First -and $Save -ne "$($Item.RegKey.ToString())$($Item.RegValue.ToString())$($Item.VDAType)")
				{
					Line 0 ""
				}

				Line 1 ( "{0,-95} {1,-50} {2,-90} {3,-8} {4,-30}" -f `
				$Item.RegKey, $Item.RegValue, $Item.Value, $Item.VDAType, $Item.ComputerName )
				
				$Save = "$($Item.RegKey.ToString())$($Item.RegValue.ToString())$($Item.VDAType)"
				If($First)
				{
					$First = $False
				}
			}
		}
		Else
		{
			Line 1 "<None found>"
		}
		Line 0 ""
	}
	If($HTML)
	{
		#sort the array by regkey, regvalue and servername
		$Script:HTMLALLVDARegistryItems = $Script:HTMLALLVDARegistryItems | Sort-Object RegKey, RegValue, VDAType, ComputerName
		WriteHTMLLine 1 0 "Appendix A - VDA Registry Items"
		WriteHTMLLine 0 0 "Miscellaneous Registry Items That May or May Not Exist on VDAs"
		WriteHTMLLine 0 0 "Linux VDAs are excluded"
		WriteHTMLLine 0 0 "These items may or may not be needed"
		WriteHTMLLine 0 0 "This Appendix is for VDA comparison only"
		
		$Save = ""
		$First = $True
		If($Script:HTMLAllVDARegistryItems)
		{
			$rowdata = @()
			ForEach($Item in $Script:HTMLAllVDARegistryItems)
			{
				If(!$First -and $Save -ne "$($Item.RegKey.ToString())$($Item.RegValue.ToString())$($Item.VDAType)")
				{
					$rowdata += @(,(
					"",$htmlwhite))
				}

				$rowdata += @(,(
				$Item.RegKey,$htmlwhite,
				$Item.RegValue,$htmlwhite,
				$Item.Value,$htmlwhite,
				$Item.VDAType,$htmlwhite,
				$Item.ComputerName,$htmlwhite))
				$Save = "$($Item.RegKey.ToString())$($Item.RegValue.ToString())$($Item.VDAType)"
				If($First)
				{
					$First = $False
				}
			}
			$columnHeaders = @(
			'Registry Key',($global:htmlsb),
			'Registry Key',($global:htmlsb),
			'Data',($global:htmlsb),
			'VDA Type',($global:htmlsb),
			'Computer Name',($global:htmlsb))

			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		}
		Else
		{
			WriteHTMLLine 0 1 "None found"
		}
		WriteHTMLLine 0 0 ""
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix A VDA Registry Items"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region script core
#Script begins

ProcessScriptSetup

SetFileNames "$($Script:CCSiteName)"

If($Section -eq "All" -or $Section -eq "Catalogs")
{
	ProcessMachineCatalogs
}

If($Section -eq "All" -or $Section -eq "Groups")
{
	ProcessDeliveryGroups
}

If($Section -eq "All" -or $Section -eq "Apps")
{
	ProcessApplications
	ProcessApplicationGroupDetails
}

If($Section -eq "All" -or $Section -eq "Policies")
{
	If($NoPolicies -or $Script:DoPolicies -eq $False)
	{
		#don't process policies
	}
	Else
	{
		ProcessPolicies
	}
}

If($Section -eq "All" -or $Section -eq "Logging")
{
	ProcessConfigLogging
}

If($Section -eq "All" -or $Section -eq "Config")
{
	ProcessConfiguration
}

If($Section -eq "All" -or $Section -eq "Admins")
{
	ProcessAdministrators
	ProcessScopes
	ProcessRoles
}

If($Section -eq "All" -or $Section -eq "Hosting")
{
	ProcessHosting
}

If($Section -eq "All" -or $Section -eq "Licensing")
{
	ProcessLicensing
}

If($Section -eq "All" -or $Section -eq "StoreFront")
{
	ProcessStoreFront
}

If($Section -eq "All" -or $Section -eq "AppV")
{
	ProcessAppV
}

If($Section -eq "All" -or $Section -eq "Zones")
{
	ProcessZones
}

If($Section -eq "All")
{
	OutputSummaryPage
}

If($VDARegistryKeys)
{
	OutputAppendixA
}
#endregion

#region finish script
Write-Verbose "$(Get-Date -Format G): Finishing up document"
#end of document processing

$AbstractTitle = "Citrix Cloud $($Script:CCSiteVersion) Inventory"
$SubjectTitle = "Citrix Cloud $($Script:CCSiteVersion) Site Inventory"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

ProcessScriptEnd
#endregion