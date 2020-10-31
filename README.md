# CitrixCloud
Citrix Cloud or Virtual Apps and Desktop service
	Creates an inventory of a Citrix Cloud Site using Microsoft PowerShell, Word, plain 
	text, or HTML.
	
	This Script requires at least PowerShell version 5.
	
	This script must run from an elevated PowerShell Session.

	Default output is HTML.
	
	Run this script on a computer with the Remote SDK installed.
	
	https://download.apps.cloud.com/CitrixPoshSdk.exe
		
	This script was developed and run from two Windows 10 VMs. One was domain-joined and 
    the other was in a Workgroup.
	
	This script supports only Citrix Cloud, not the on-premises CVAD products.
	
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
		
