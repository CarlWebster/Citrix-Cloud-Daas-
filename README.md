# Citrix Cloud (DaaS)
Creates an inventory of a Citrix Cloud (now DaaS) Site using Microsoft PowerShell, Word, plain text, or HTML.
	
This script requires at least PowerShell version 5.
	
The default output is HTML.
	
Run this script on a computer with the Remote SDK installed.
	
https://download.apps.cloud.com/CitrixPoshSdk.exe
	
This script was developed and run from two Windows 10 VMs. One was domain-joined and the other was in a Workgroup.
	
This script supports only Citrix Cloud (now DaaS), not the on-premises CVAD products.
	
If you are running XA/XD 7.0 through 7.7, please use: 
https://carlwebster.com/downloads/download-info/xenappxendesktop-7-x-documentation-script/

If you are running XA/XD 7.8 through CVAD 2006, please use:
https://carlwebster.com/downloads/download-info/xenappxendesktop-7-8/

If you are running CVAD 2006 and later, please use:
https://carlwebster.com/downloads/download-info/citrix-virtual-apps-and-desktops-v3-script/

To prevent multiple Citrix Cloud (now DaaS) authentication prompts, follow the instructions in the Authentication section of the ReadMe file to create a profile named Default.
	
ReadMe file: https://carlwebster.sharefile.com/d-s1ef10b6883eb473fa2f4eef00be83799
	
By default, the script only gives summary information for:
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

Using the MachineCatalogs parameter can cause the report to take a very long time to complete and can generate an extremely long report.
	
Using the DeliveryGroups parameter can cause the report to take a very long time to complete and can generate an extremely long report.

Using both the MachineCatalogs and DeliveryGroups parameters can cause the report to take an extremely long time to complete and generate an exceptionally long report.

Creates an output file named after the Citrix Cloud(now DaaS) Site (which by default is cloudxdsite) unless you use the SiteName parameter.
	
Word and PDF Document includes a Cover Page, Table of Contents, and Footer.
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
 
Prerequisites
Let us ensure we have the requirements before using PowerShell to document anything in a Citrix DaaS Site.
1.	Verify script execution is Enabled on the computer running this script. By default, PowerShell script execution is blocked.
a.	Start, Run, gpedit.msc
b.	Computer Configuration/Administrative Templates/Windows Components/Windows PowerShell/Turn on Script Execution: Enabled
c.	Execution Policy: select the option that follows your recommended security practices.
2.	The script can be run from a non-domain or domain-joined computer.
3.	Install the Remote PowerShell SDK.
a.	https://download.apps.cloud.com/CitrixPoshSdk.exe
b.	Do not install the Remote PowerShell SDK on a computer running the Cloud Connector.
c.	Do not install the Remote PowerShell SDK on a computer running the on-premises version of Studio, or that has the on-premises PowerShell snap-ins or modules installed.
d.	You must restart the computer after the installation.
4.	Verify that the Visual C++ 2015 Runtime is installed.
a.	Option 1: Install from the CVAD 2106 or later installation media 
i.	x:\Support\VcRedist_2017\vcredist_x86.exe (install first)
ii.	x:\Support\VcRedist_2017\vcredist_x64.exe
b.	Option2: Install from the Microsoft download page
i.	https://aka.ms/vs/17/release/vc_redist.x86.exe (install first)
ii.	https://aka.ms/vs/17/release/vc_redist.x64.exe
5.	Install the Citrix Group Policy Management Console
a.	Option 1: Install from the CVAD 2106 or later installation media. Note: This is required by the StoreFront and Citrix Policy cmdlets and functions.
i.	x:\x64\Citrix Policy\CitrixGroupPolicyManagement_x64.msi
ii.	Installing this console installs the required Citrix.Common.GroupPolicy PowerShell snapin.
b.	Option 2: Install from https://www.citrix.com/downloads/citrix-cloud/product-software/xenapp-and-xendesktop-service.html. Note: This is required by the StoreFront and Citrix Policy cmdlets and functions.
i.	This downloads the latest CitrixGroupPolicyManagement_x64.msi.
ii.	Installing this console installs the required Citrix.Common.GroupPolicy PowerShell snapin.

Script Usage
How to use this script?
1.	Save the script as CC_Inventory_V1.ps1 in your PowerShell scripts folder.
2.	Follow the process in the previous Authentication section to create an authentication profile.
3.	Change to your PowerShell scripts folder from a new elevated PowerShell prompt.
a.	Note: The script author recommends you run the script from a NEW elevated PowerShell session each time.
4.	From the PowerShell prompt, type in:
a.	.\CC_Inventory_V1.ps1 -ProfileName "Example_1" and press Enter.
5.	By default, the script creates an HTML document named after the CVADS Site, which is cloudxdsite. Use the -SiteName parameter to specify a name for the report file.
6.	If you use the –MSWord option, the script creates a Microsoft Word file named after the CVADS Site, which is cloudxdsite. Use the -SiteName parameter to specify a name for the report file.
7.	If you use the –PDF option, the script creates a PDF file named after the CVADS Site, which is cloudxdsite. Use the -SiteName parameter to specify a name for the report file.
8.	If you use the –Text option, the script creates a Text file named after the CVADS Site, which is cloudxdsite. Use the -SiteName parameter to specify a name for the report file.
To run the script:
.\CC_Inventory_V1.ps1 -ProfileName "Example_1" and press Enter.
Full help text is available.
Get-Help .\CC_Inventory_V1.ps1 –full

The help text explains all the parameters the script accepts.
