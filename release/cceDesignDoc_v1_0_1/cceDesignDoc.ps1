<#PSScriptInfo

.VERSION 1.0.1

.GUID 52f620ce-6560-42d9-afac-d1124aa65d1c

.AUTHOR Shane Hoey

.COMPANYNAME 

.COPYRIGHT 2016-2018 Shane Hoey

.TAGS Skype for Business, Skype for Business Online, Microsoft Office, Office

.LICENSEURI https://shanehoey.github.io/ccedesigndoc/license

.PROJECTURI https://shanehoey.github.io/ccedesigndoc

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES
https://shanehoey.github.io/ccedesigndoc

#> 

#Requires -Module WordDoc






<# 

.DESCRIPTION 
Create a Design Document from Cloud Connector

#> 

<#
MIT License

Copyright (c) 2016-2018 Shane Hoey

Permission is hereby granted, free of charge, to any person obtaining a copy 
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

#>
[cmdletbinding()]
Param(  

    [ValidateNotNullOrEmpty()]  
    [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".ini")})]  
    [Parameter(ValueFromPipeline=$false,Mandatory=$false)]  
    [string]$CloudConnectorFile,
  
    [ValidateNotNullOrEmpty()]  
    [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -like ".do*x")})]  
    [Parameter(ValueFromPipeline=$false,Mandatory=$false)]  
    [string]$TemplateFile,  

    [ValidateNotNullOrEmpty()]  
    [ValidateScript({ (Test-Path ([System.IO.Path]::GetDirectoryName($_)) -pathtype container ) -and ([System.IO.Path]::GetExtension($_) -eq ".docx") })]  
    [Parameter(ValueFromPipeline=$false,Mandatory=$false)]  
    [string]$SaveAsFile

)

#region File Variables
    if ($PSBoundParameters.ContainsKey('CloudConnectorFile')) 
        {
            $CloudConnectorFile = (get-item -path $CloudConnectorFile).fullname
        }    
    else
        { 
            Add-Type -AssemblyName System.windows.forms
            $OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.initialDirectory =  [Environment]::GetFolderPath('MyDocuments')
            $OpenFileDialog.filter = 'CloudConnector.ini (*.ini)|*.ini'
            $OpenFileDialog.title = 'Select cloudconnector.ini to import'
            $OpenFileDialog.ShowHelp = $True
            [void]$OpenFileDialog.ShowDialog()
            $CloudConnectorFile = $OpenFileDialog.filename
            Remove-Variable -Name OpenFileDialog
        }
    Write-Verbose -message $CloudConnectorFile

    if ($PSBoundParameters.ContainsKey('TemplateFile')) 
        {
            $Template = (get-item -path $TemplateFile).fullname
        }    
    else
        { 
            Add-Type -AssemblyName System.windows.forms 
            $OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.initialDirectory =  [Environment]::GetFolderPath('MyDocuments')
            $OpenFileDialog.filter = 'Word Document (*.docx)|*.docx|Word Template (*.dotx)|*.dotx'
            $OpenFileDialog.title = 'Select Word Template to import'
            [void]$OpenFileDialog.ShowDialog()
            $TemplateFile = $OpenFileDialog.filename
            Remove-Variable -Name OpenFileDialog
        }
    Write-Verbose -message $TemplateFile

    if ($PSBoundParameters.ContainsKey('SaveAsFile')) 
        {
            $SaveAsFile = [System.IO.Path]::GetFullPath($SaveAsFile)
        }
    else    
        {     
            $SaveFileDialog = New-Object -TypeName System.Windows.Forms.saveFileDialog
            $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
            $SaveFileDialog.filter = 'Word Document (*.docx)|*.docx|HTML File (*.html)|*.html|PDF (*.pdf)|*.pdf'
            $SaveFileDialog.title = "Save As:"
            [void]$SaveFileDialog.ShowDialog()
            $SaveAsFile = $SaveFileDialog.FileName
            remove-variable -Name SaveFileDialog
        }
    Write-Verbose -message $SaveAsFile
#endregion

#region Version Control & 14 Day usage stats
# Please do not remove this section,
# It is only used for version Control and unique users via github 
# I only see number of unique users over 14 days period, 
# Collecting the stats gives me an indication how often this script is used to determine if I should continue developing it, or concentrate on other projects
# If you want to silence the notice set notify to $false rather than deleting the section
# thank in advance
$notify = $true #$false if you do not want to be notified of updates. 
$thisversion = "d5a1c799-b160-47a1-bfcf-2dbe5562dd84"
$Version = (Invoke-WebRequest -Uri https://raw.githubusercontent.com/shanehoey/versions/master/ccedesigndoc.json -UserAgent cceDesignDoc -Method Get -DisableKeepAlive -TimeoutSec 2).content | convertfrom-json
if (($thisversion -ne $version.release) -and ($thisversion -ne $version.dev)) {
    Write-Verbose -message "cceDesignDoc has been updated" -Verbose
    if($notify) { 
        Write-Host -object "**********************`nCCE has been Updated`n**********************`nMore details available at $($version.link)"
        start-sleep -Seconds 5 
    }
}
#endregion

#region textOverview
$textOverview = @" 
Skype for Business Online is Microsoft’s cloud-based version of Skype for Business Server and is part of the Office 365 cloud offering from Microsoft

Phone System (formerly known as Cloud PBX) in Office 365 enables external connectivity either through Microsoft's Calling Plans (formerly known as PSTN Calling) add-on — available in selected countries only or via your existing PSTN circuits using Cloud Connector Edition. 

As an Official Cloud Connector appliance, the AudioCodes Mediant Cloud Connector Appliances has been tested for performance and suitability to meet load and quality requirements of PSTN Calling. This hybrid appliances that consists of a set of packaged Virtual Machines (VMs), and a Mediant Gateway/SBC, that together implement on-premises PSTN connectivity. Your users are homed in the Skype for Business online, and receive PBX services from the cloud, but PSTN connectivity is provided through the AudioCodes Mediant Appliance. 

To ensure a seamless transition to Skype for Business Online and Phone System, AudioCodes' One Voice for Microsoft 365 offering includes a complete portfolio of products and services design to smooth the transition to cloud communications.
"@
#endregion

#region objectRelatedDocuments
$objectRelatedDocuments = (@"
[
    {
        "DocumentName":  "Plan for Skype for Business Cloud Connector Edition",
        "DocumentLink":  "https://technet.microsoft.com/en-us/library/mt605227.aspx "
    },
    {
        "DocumentName":  "Configure Skype for Business Cloud Connector Edition",
        "DocumentLink":  "https://technet.microsoft.com/en-us/library/mt605228.aspx"
    },
    {
        "DocumentName":  "LTRT-28088 Mediant Appliance for Skype for Business CCE Installation",
        "DocumentLink":  "https://www.AudioCodes.com/media/12399/ltrt-28088-mediant-appliance-for Microsoft Skype for Business cce-installation-manual-ver-210.pdf"
    },
    {
        "DocumentName":  "LTRT-0318 Product Notice Release of AudioCodes Mediant CCE Appliance Software",
        "DocumentLink":  "https://www.AudioCodes.com/media/8714/0318-product-notice-release-of-AudioCodes-mediant-cce-appliance-software-version-210.pdf"
    },
    {
        "DocumentName":  "LTRT-09943-400hd Series IP Phone for Microsoft Skype for Business administrators manual",
        "DocumentLink":  "https://www.AudioCodes.com/media/9548/ltrt-09943-400hd-series-ip-phone-for-microsoft-skype-for-business-administrators-manual-ver-301.pdf"
    }
]
"@ | convertfrom-json)
#todo update typedata 
#endregion

#region textDesign
$textDesign = @" 
CCE appliances combined with a Gateway/SBC and can be deployed in multiple scenarios. Deploying multiple CCEs and Gateways/SBCs allows you to enhance your site capacity and implement High Availability (HA) configuration for the CCE instance and for Gateway/SBC redundancy configuration. 

CCE appliances are a Physical Host running Windows Server and Hyper-V, that hosts a set of packaged VMs that contain a minimal Skype for Business Server topology—consisting of an Edge component, Mediation component, a Central Management Store (CMS) role, and a Domain Controller (Each CCE Appliance has its own Forest). Additionally, the addition of a Session Border Controllers (SBC) (Physical or Virtual) offers direct SIP connectivity between existing enterprise voice infrastructure, Skype for Business, the PSTN and SIP trunk services.

IMPORTANT: This Design Document only covers a Single CCE Appliance and single/multiple Gateways/SBC, however this design can be adapted to include Multiple CCE appliances configured in a High Available configuration.

"@
#endregion

#region listDesign
$listDesign = @()
$listDesign += "A dedicated account should be used for management that has <b>Skype for Business</B> Tenant administrator Rights." 
$listDesign += "Cloud Connector cannot co-exist with Lync or Skype for Business on-premises servers." 
$listDesign += "Dial-in conferencing can be configured with PSTN Conferencing." 
#endregion

#region objectHardware
$objectHardware = (@"
[
    {
        "Manufacture" : "AudioCodes",
        "HardwareAppliance":  "AudioCodes Mediant 800 CCE Appliance",
        "Hardware":  "Mediant 800 with OSN server",
        "Processor":  "64-bit processor, four core (4 real cores), 2.70 Gigahertz (GHz)",
        "Ram":  "32 gigabytes (GB) RAM",
        "Disk":  "Single SSD 512 GB",
        "Network":  "Two NICs 1Gbps + one 100Mbps",
        "Link": ""
    },
    {
        "Manufacture" : "AudioCodes",
        "HardwareAppliance":  "AudioCodes Mediant Server CCE Appliance (Gen8)",
        "Hardware":  "HP ProLiant server",
        "Processor":  "64-bit processor, six core (12 real cores), 2.10 Gigahertz (GHz)",
        "Ram":  "64 gigabytes (GB) RAM",
        "Disk":  "Four 600gb 10k RPM 512M Cache SAS 6Gbps disk, configured in a RAID 5 configuration",
        "Network":  "Four 1Gbps",
        "Link": ""
    },
    {
        "Manufacture" : "AudioCodes",
        "HardwareAppliance":  "AudioCodes Mediant Server CCE Appliance (Gen9)",
        "Hardware":  "HP ProLiant server",
        "Processor":  "64-bit processor, six core (16 real cores), 2.10 Gigahertz (GHz)",
        "Ram":  "64 gigabytes (GB) RAM",
        "Disk":  "Four 600gb 10k RPM 2G Cache SAS 6Gbps disk, configured in a RAID 5 configuration",
        "Network":  "Eight 1Gbps",
        "Link": ""
    }
]
"@ | convertfrom-json )
#endregion

#region objectIPPhone
$objectIPPhone = (@"
[
    {
        "Manufacture" : "AudioCodes", 
        "Model" :  "450HD",
        "Description" :  "AudioCodes 450HD IP Phone, An easy-to-use executive high-end business phone with a large color touch screen and full UC integration.",
        "Link" : "https://www.AudioCodes.com/solutions-products/products/ip-phones/450hd-ip-phone"
    },
    {
        "Manufacture" : "AudioCodes",
        "Model" :  "445HD",
        "Description" : "AudioCodes 445HD IP Phone, An advanced high-end business phone with a color screen and integrated sidecar for speed dial contacts and presence monitoring.",
        "Link" : "https://www.AudioCodes.com/solutions-products/products/ip-phones/445hd-ip-phone"
    },
    {
        "Manufacture" : "AudioCodes",
        "Model" :  "440HD",
        "Description" :  "AudioCodes 440HD IP Phone,  An advanced mid-range business phone with an integrated sidecar for speed dial contacts and presence monitoring.",
        "Link" : "https://www.AudioCodes.com/solutions-products/products/ip-phones/440hd-ip-phone"
    },
    {
        "Manufacture" : "AudioCodes",
        "Model" :  "420HD",
        "Description" :  "AudioCodes 420HD IP Phone, An entry-level phone with high voice quality at an affordable price.",
        "Link" : "https://www.AudioCodes.com/solutions-products/products/ip-phones/420hd-ip-phone"
    },
    {
        "Manufacture" : "AudioCodes",
        "Model" :  "405HD",
        "Description" :  "AudioCodes 405HD IP Phone, A cost-effective phone packed with essential features.",
        "Link" : "https://www.AudioCodes.com/solutions-products/products/ip-phones/405hd-ip-phone"
    }
]
"@ | convertfrom-json )
#endregion

#region textCloudConnectorAppliance
$textCloudConnectorAppliance = @" 
Cloud Connector Edition is a set of packaged VMs that contain a minimal Skype for Business Server topology—consisting of an Edge component, Mediation component, and a Central Management Store (CMS) role.  It also includes a domain controller for management of the VM’s only and has no dependency or trust requirements on your existing domain.
"@
#endregion

#region textSipDomains
$textSipDomains = @"
Sip Domains must be registered in Office 365, and you should include all SIP domains. Also you cannot use the default onmicrosoft.com domain
"@
#endregion

#region textWSUSSettings
$textWSUSSettings = @"
OPTIONAL: The address of the Windows Server Update Services (WSUS) used
"@
#endregion

#region textOtherSettings
$textOtherSettings = @"
The Federation FQDN in most cases will be kept at the default “sipfed.online.lync.com” refer to TechNet documentation
OPTIONAL:  the Base VM is generally not required, refer to TechNet documentation
"@
#endregion

#region testSiteDetails
$textPSTNSite = @" 
A PSTN site contains Cloud Connector appliances and SBC/gateways that are generally connected at the same location. A PSTN Site can be configured with a maximum of 16 appliances and 16 SBC/Gateways, with potential of handling up to 8000 simultaneous calls dependant on chosen hardware. When multiple CCE appliances in a single PSTN site are deployed calls are distributed on random order between CCE appliances.

The site name must exist in office 365 if it does not it will be created when the site is registered.
"@
#endregion

#region textNetwork
$textNetwork = @" 
TODO - Write Details on Network Requirements
"@
#endregion

#region textCorpNetwork
$textCorpNetwork = @" 
Cloud Connector Edition Requires a network for internal communication between Cloud Connector components.
"@
#endregion

#region textInternetNetwork
$textInternetNetwork = @" 
Cloud Connector Edition Requires a network for External communication between Cloud Connector Edge server and the Internet. 
"@
#endregion

#region textManagementNetwork
$textManagementNetwork = @" 
Management subnet is a temporary subnet that is created automatically during the deployment. Generally, there is no requirement to modify this network.
"@
#endregion

#region textServers
$textServers = @" 
Cloud Connector Edition is a set of packaged VMs that contain a minimal Skype for Business Server topology—consisting of an Edge component, Mediation component, and a Central Management Store (CMS) role.  It also includes a domain controller for management of the VM’s only and has no dependency or trust requirements on your existing domain.
"@
#endregion

#region textDomainController
$textDomainController = @" 
Active Directory services is used to store global settings required to deploy clod connector components. The AD Domain name (including NETBIOS name) must be different from any production names. There is one forest per Cloud Connector, however the name should be the same across all Cloud Connector Appliances.  
"@
#endregion

#region textCMS
$textCMS = @" 
The Primary CMS is used for the  configuration store,  including CMS File Transfer, and synchronising global CMS data.
"@
#endregion

#region textMediation
$textMediation = @" 
The mediation server is used for SIP and Media gateway mapping protocol between Skype for Business and PSTN gateways.
"@
#endregion

#region textEdgeServer
$textEdgeServer = @" 
The Edge Server is used for communication between the CCE topology and Office 365
"@
#endregion

#region textGateways
$textGateways = @" 
The Gateway is used to connect the PSTN to the Skype for Business Mediation Server.
"@
#endregion

#region textTrunkConfiguration
$textTrunkConfiguration = @" 
Trunk Configuration should be configured the same across all cce appliances in a PSTN site.
"@
#endregion

#region textHybridVoiceRoutes
$textHybridVoiceRoutes = @" 
Voice route can be modified to restrict the outbound call numbers
"@
#endregion

#region textFirewalls
$textFirewalls = @" 
CCE must be deployed in a perimeter network with an External DMZ that is routable to the internet and an Internal DMZ that is routable via the internal network.  Any Client coming from the internet will connect via the edge server.  Clients coming from the internal network will connect via the mediation server.
"@
#endregion

#region textCertificates
$textCertificates = @" 
TODO - Write Details on Certificate Requirements
"@
#endregion

#region textExternalDNSresolution 
$textExternalDNSresolution = @" 
The Edge component needs to resolve the external names of Office 365 services and the internal names of other Cloud Connector components.

TODO - Write more details on External DNS Resolution
"@
#endregion


#region import cloudconnector.ini
#Credit Oliver Lipkau
#https://blogs.technet.microsoft.com/heyscriptingguy/2011/08/20/use-powershell-to-work-with-any-ini-file/
$CloudConnector = @{}
switch -regex -file $CloudConnectorFile
{
  "^\[(.+)\]" # Section
  {
    $section = $matches[1]
    $CloudConnector[$section] = @{}
    $CommentCount = 0
  }

  "(.+?)\s*=(.*)" # Key
  {
    $name,$value = $matches[1..2]
    $CloudConnector[$section][$name] = $value
  }
}
#endregion

$CCESettings = [pscustomobject]@{
    "Sip Domains" = $CloudConnector.Common.SIPDomains.split(" ")
    "Active Directory Domain FQDN" = $CloudConnector.Common.VirtualMachineDomain 
    "Active Directory Domain Netbios" = ($CloudConnector.Common.VirtualMachineDomain).Split('.')[0]
    "Federation FQDN" = $CloudConnector.Common.OnlineSipFederationFqdn 
    "Base VM" = $CloudConnector.Common.BaseVMIP 
    "WSUS Server" = $CloudConnector.Common.WSUSServer 
    "WSUS Status" = $CloudConnector.Common.WSUSStatusServer 
    }

$CCESite = [pscustomobject]@{
    "Site name" = $CloudConnector.Common.SiteName 
    "City" = $CloudConnector.Common.City
    "State" = $CloudConnector.Common.State
    "Country" = $CloudConnector.Common.CountryCode 
   }

$Corpnet = [pscustomobject]@{
    "Corpnet VN Switch Name" = $CloudConnector.Network.CorpnetSwitchName
    "Corpnet Default Gateway" = $CloudConnector.Network.CorpnetDefaultGateway
    "Corpnet Subnet" = $CloudConnector.Network.CorpnetIPPrefixLength
    "Corpnet DNS Forwarder" =   $CloudConnector.Network.CorpnetDNSIPAddress
    }

$DMZNet = [pscustomobject]@{
    "Internet VN Switch Name" = $CloudConnector.Network.InternetSwitchName
    "Internet Default Gateway" = $CloudConnector.Network.InternetDefaultGateway
    "Internet Subnet" = $CloudConnector.Network.InternetIPPrefixLength
    "Internet DNS Forwarder" =   $CloudConnector.Network.InternetDNSIPAddress
    }

$ManagementNet = [pscustomobject]@{
    "Management VN Switch Name" = $CloudConnector.Network.ManagementSwitchName
    "Management Network" = $CloudConnector.Network.ManagementIPPrefix
    "Management Subnet" = $CloudConnector.Network.ManagementIPPrefixLength
    }

$DomainController = [pscustomobject]@{
    "Domain Controller Name" = $CloudConnector.Common.ServerName
    "AD Domain" =  $CloudConnector.Common.VirtualMachineDomain
    "IP Address" = $CloudConnector.Common.IP
    "Subnet" = $CloudConnector.Network.CorpnetIPPrefixLength
    "Default Gateway" = $CloudConnector.Network.CorpnetDefaultGateway
    "DNS Forwarder" =   $CloudConnector.Network.CorpnetDNSIPAddress
    }

$CMS = [pscustomobject]@{
    "CMS Servername" = $CloudConnector.Common.ServerName
    "AD Domain" =  $CloudConnector.Common.VirtualMachineDomain
    "IP Address" = $CloudConnector.PrimaryCMS.IP
    "Subnet" = $CloudConnector.Network.CorpnetIPPrefixLength
    "Default Gateway" = $CloudConnector.Network.CorpnetDefaultGateway
    "DNS Server" =   $CloudConnector.Common.IP
    "CMS Share" =  "\\$($CloudConnector.Common.ServerName)\$($cloudconnector.PrimaryCMS.ShareName)"
    }

$Mediation = [pscustomobject]@{
    "Mediation Pool name" = $cloudconnector.MediationServer.PoolName
    "Mediation Server name" = $cloudconnector.MediationServer.ServerName
    "AD Domain" =  $CloudConnector.Common.VirtualMachineDomain
    "IP Address" =  $cloudconnector.MediationServer.IP
    "Subnet" = $CloudConnector.Network.CorpnetIPPrefixLength
    "Default Gateway" = $CloudConnector.Network.CorpnetDefaultGateway
    "DNS Server" =   $CloudConnector.Common.IP
    "CMS Share" =  "\\$($CloudConnector.Common.ServerName)\$($cloudconnector.PrimaryCMS.ShareName)"
    }




$Edge =[pscustomobject]@{
    "Edge Pool name" = "$($cloudconnector.EdgeServer.InternalPoolName).$($CloudConnector.Common.VirtualMachineDomain)"
    "Edge Server name" = $cloudconnector.EdgeServer.InternalServerName
    "AD Domain" =  $CloudConnector.Common.VirtualMachineDomain
    "IP Address" =  $cloudconnector.EdgeServer.InternalServerIPs
    "Subnet" = $CloudConnector.Network.CorpnetIPPrefixLength
    "Default Gateway" = $CloudConnector.Network.CorpnetDefaultGateway
    "DNS Server" =   $CloudConnector.Common.IP
    "SIP Pool Name" =   "$($cloudconnector.EdgeServer.ExternalSIPPoolName).$($CCESettings.'Sip Domains'[0])"
    "SIP IP Address" =   $cloudconnector.EdgeServer.ExternalSIPIPs
    "Media Relay Pool Name" =   "$($cloudconnector.EdgeServer.ExternalMRFQDNPoolName).$($CCESettings.'Sip Domains'[0])"
    "Media Relay IP Address" =   $cloudconnector.EdgeServer.ExternalMRIPs
    "Media Relay Public IP Address" =   $cloudconnector.EdgeServer.ExternalMRPublicIPs
    "Media Relay Port Range" =   $cloudconnector.EdgeServer.ExternalMRPortRange
    }

$gateways = @()
for ($i = 1; $i -lt 16; $i++)
{ 
    if ($cloudconnector."Gateway$i".fqdn -ne $null) {     
        $gateways += [pscustomobject]@{
            "Gateway FQDN" = $cloudconnector.Gateway1.fqdn
            "IP Address" = $cloudconnector.Gateway1.ip
            "Port" =  $cloudconnector.Gateway1.port
            "Protocol" =  $cloudconnector.Gateway1.protocol
            "Voice Routes" = $cloudconnector.Gateway1.voiceroutes
        }
    }
}
    
$TrunkConfiguration = [pscustomobject]@{
    "Forward Call History" = $CloudConnector.TrunkConfiguration.ForwardCallHistory
    "Enable Fast Failover timer " =  $CloudConnector.TrunkConfiguration.EnableFastFailoverTimer
    "Enable Refer Support" = $CloudConnector.TrunkConfiguration.EnableReferSupport
    "Forward PAI" = $CloudConnector.TrunkConfiguration.ForwardPAI
   }

$HybridVoiceRoutes = @()
foreach ($routes in $cloudconnector.HybridVoiceRoutes.GetEnumerator()) {
    $HybridVoiceRoutes += [pscustomobject]@{
       'Voice Route Name' = $routes.Name
       'Voice Route' = $routes.Value
    }
}

$Firewall = @()
foreach($gateway in $gateways) 
{
  $Firewall += [pscustomobject]@{ 'Firewall' = 'Internal'; 'ruleset' = '1'; 'Source IP' = $Mediation.'IP Address' ; 'Source Port' = 'Any' ; 'Destination IP' = $gateway.'IP Address' ; 'Destination Port' =  "$($gateway.Protocol) $($gateway.Port)" }
  $Firewall += [pscustomobject]@{ 'Firewall' = 'Internal'; 'ruleset' = '2'; 'Source IP' = $gateway.'IP Address' ; 'Source Port' = 'Any' ; 'Destination IP' =  $Mediation.'IP Address' ; 'Destination Port' = 'TCP 5068' }
  $Firewall += [pscustomobject]@{ 'Firewall' = 'Internal'; 'ruleset' = '3'; 'Source IP' = $gateway.'IP Address' ; 'Source Port' = 'Any' ; 'Destination IP' =  $Mediation.'IP Address' ; 'Destination Port' =  'TLS 5067' }
  $Firewall += [pscustomobject]@{ 'Firewall' = 'Internal'; 'ruleset' = '4'; 'Source IP' = $Mediation.'IP Address' ; 'Source Port' = 'UDP 49,152 - 57,500' ; 'Destination IP' =  $gateway.'IP Address' ; 'Destination Port' =  'Any' }
  $Firewall += [pscustomobject]@{ 'Firewall' = 'Internal'; 'ruleset' = '5'; 'Source IP' = $gateway.'IP Address' ; 'Source Port' = 'Any' ; 'Destination IP' =  $Mediation.'IP Address' ; 'Destination Port' =  'UDP 49,152 - 57500' }
  $Firewall += [pscustomobject]@{ 'Firewall' = 'Internal'; 'ruleset' = '6'; 'Source IP' = $Mediation.'IP Address' ; 'Source Port' =  'TCP 49,152 - 57,500' ; 'Destination IP' =  'Internal Clients' ; 'Destination Port' = 'TCP 50,000-50,019' }
  $Firewall += [pscustomobject]@{ 'Firewall' = 'Internal'; 'ruleset' = '7'; 'Source IP' = $Mediation.'IP Address' ; 'Source Port' =  'UDP 49,152 - 57,500' ; 'Destination IP' =  'Internal Clients' ; 'Destination Port' = 'UDP 50,000-50,019' }
  $Firewall += [pscustomobject]@{ 'Firewall' = 'Internal'; 'ruleset' = '8'; 'Source IP' = 'Internal Clients' ; 'Source Port' = 'TCP 50,000 - 50,019' ; 'Destination IP' =  $Mediation.'IP Address' ; 'Destination Port' = 'TCP 49,152-57,500' }
  $Firewall += [pscustomobject]@{ 'Firewall' = 'Internal'; 'ruleset' = '9'; 'Source IP' = 'Internal Clients' ; 'Source Port' = 'UDP 50,000 - 50,019' ; 'Destination IP' =  $Mediation.'IP Address'; 'Destination Port' =  'UDP 49,152-57,500' }
}
 
#Edge Server 
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Minimum'; 'ruleset' = '10'; 'Source IP' = 'Any' ; 'Source Port' =   'Any' ; 'Destination IP' =  $Edge.'Media Relay IP Address'; 'Destination Port' = 'TCP 5061' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Minimum'; 'ruleset' = '10'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'Any' ; 'Destination IP' =  'Any'; 'Destination Port' = 'TCP 5061' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Minimum'; 'ruleset' = '10'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'Any' ; 'Destination IP' =  'Any'; 'Destination Port' = 'TCP 80' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Minimum'; 'ruleset' = '10'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'Any' ; 'Destination IP' =  'Any'; 'Destination Port' = 'UDP 53' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Minimum'; 'ruleset' = '10'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'Any' ; 'Destination IP' =  'Any'; 'Destination Port' = 'TCP 53' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Minimum'; 'ruleset' = '10'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'UDP 3478' ; 'Destination IP' =  'Any'; 'Destination Port' = 'UDP 3478' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Minimum'; 'ruleset' = '10'; 'Source IP' = 'Any' ; 'Source Port' =   'TCP 50,000 - 59,999' ; 'Destination IP' =  $Edge.'Media Relay IP Address'; 'Destination Port' = 'TCP 443' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Minimum'; 'ruleset' = '10'; 'Source IP' = 'Any' ; 'Source Port' =   'UDP 3478' ; 'Destination IP' =  $Edge.'Media Relay IP Address'; 'Destination Port' = 'UDP 3478' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Minimum'; 'ruleset' = '10'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'TCP 50,000-59,999' ; 'Destination IP' =  'Any'; 'Destination Port' = 'TCP 443' }

#Edge Server  Recommended
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Recommended'; 'ruleset' = '11'; 'Source IP' = 'Any' ; 'Source Port' =   'Any' ; 'Destination IP' =  $Edge.'Media Relay IP Address'; 'Destination Port' = 'TCP 5061' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Recommended'; 'ruleset' = '11'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'Any' ; 'Destination IP' =  'Any'; 'Destination Port' = 'TCP 5061' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Recommended'; 'ruleset' = '11'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'Any' ; 'Destination IP' =  'Any'; 'Destination Port' = 'TCP 80' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Recommended'; 'ruleset' = '11'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'Any' ; 'Destination IP' =  'Any'; 'Destination Port' = 'UDP 53' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Recommended'; 'ruleset' = '11'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'Any' ; 'Destination IP' =  'Any'; 'Destination Port' = 'TCP 53' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Recommended'; 'ruleset' = '11'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'TCP 50,000-59,999' ; 'Destination IP' =  'Any'; 'Destination Port' = 'Any' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Recommended'; 'ruleset' = '11'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'UDP 3478' ; 'Destination IP' =  'Any'; 'Destination Port' = 'Any' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Recommended'; 'ruleset' = '11'; 'Source IP' = $Edge.'Media Relay IP Address' ; 'Source Port' =   'UDP 50,000-59,999' ; 'Destination IP' =  'Any'; 'Destination Port' = 'Any' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Recommended'; 'ruleset' = '11'; 'Source IP' = 'Any' ; 'Source Port' =   'Any' ; 'Destination IP' =  $Edge.'Media Relay IP Address'; 'Destination Port' = 'TCP 443' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Recommended'; 'ruleset' = '11'; 'Source IP' = 'Any' ; 'Source Port' =   'Any' ; 'Destination IP' =  $Edge.'Media Relay IP Address'; 'Destination Port' = 'TCP 50,000 - 59,999' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Recommended'; 'ruleset' = '11'; 'Source IP' = 'Any' ; 'Source Port' =   'Any' ; 'Destination IP' =  $Edge.'Media Relay IP Address'; 'Destination Port' = 'UDP 3478' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'External-Recommended'; 'ruleset' = '11'; 'Source IP' = 'Any' ; 'Source Port' =   'Any' ; 'Destination IP' =  $Edge.'Media Relay IP Address'; 'Destination Port' = 'UDP 50,000 - 59,999' }

#Edge Server  Recommended
$Firewall += [pscustomobject]@{ 'Firewall' = 'Host'; 'ruleset' = '12'; 'Source IP' = 'TBA -HostIP' ; 'Source Port' =   'Any' ; 'Destination IP' =  'any'; 'Destination Port' = 'TCP 53' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'Host'; 'ruleset' = '12'; 'Source IP' = 'TBA -HostIP' ; 'Source Port' =   'Any' ; 'Destination IP' =  'any'; 'Destination Port' = 'UDP 53' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'Host'; 'ruleset' = '12'; 'Source IP' = 'TBA -HostIP' ; 'Source Port' =   'Any' ; 'Destination IP' =  'any'; 'Destination Port' = 'TCP 80' }
$Firewall += [pscustomobject]@{ 'Firewall' = 'Host'; 'ruleset' = '12'; 'Source IP' = 'TBA -HostIP' ; 'Source Port' =   'Any' ; 'Destination IP' =  'any'; 'Destination Port' = 'TCP 443' }



$certificates = @()
$certificates += [pscustomobject]@{ 
    "cert" = "Option1"
    "SN" = "$($edge.'SIP Pool Name')"
    "SAN" = ("$(($CCESettings.'sip domains').foreach({'sip.' + $_})) $($edge.'SIP Pool Name')").split(" ")
}
$certificates += [pscustomobject]@{ 
    "cert" = "Option2"
    "SN" = "$($edge.'SIP Pool Name')"
    "SAN" = ("$(($CCESettings.'sip domains').foreach({'sip.' + $_})) *.$(($CCESettings.'sip domains')[0])").split(" ")
}
$certificates += [pscustomobject]@{ 
    "cert" = "Option3"
    "SN" = "$($edge.'SIP Pool Name')"
    "SAN" = ("$(($CCESettings.'sip domains').foreach({'sip.' + $_})) $($edge.'SIP Pool Name') AllDeployedEdgeServers").split(" ")
}

#region Create Word Document 
    New-WordInstance
    New-WordDocument
#endregion 

#region Apply Template or CoverPage
    if ($TemplateFile -eq "") 
    { 
        Add-WordCoverPage -CoverPage Facet
        Add-WordBreak -breaktype NewPage
    }
    else 
    {
        Add-WordTemplate -filename $TemplateFile
        Add-WordBreak -breaktype Paragraph
        Add-WordBreak -breaktype Paragraph
        Add-WordBreak -breaktype Paragraph
        Add-WordText -text "Cloud Connector Implementation Design " -WDBuiltinStyle wdStyleTitle
        Add-WordBreak -breaktype Paragraph
        Add-WordBreak -breaktype Paragraph
        Add-WordBreak -breaktype Paragraph
        Add-WordText -text "for" -WDBuiltinStyle wdStyleTitle
        Add-WordBreak -breaktype Paragraph
        Add-WordBreak -breaktype Paragraph
        Add-WordBreak -breaktype Paragraph
        Add-WordText -text "CustomerName" -WDBuiltinStyle wdStyleTitle
        Add-WordBreak -breaktype NewPage
    }
#endregion

#region update document settings
    Set-WordBuiltInProperty -WdBuiltInProperty wdPropertytitle "cceDesignDoc"
    Set-WordBuiltInProperty -WdBuiltInProperty wdPropertySubject "Cloud Connector Implementation"
    Set-WordBuiltInProperty -WdBuiltInProperty wdPropertyAuthor $([Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('cwBoAGEAbgBlAGgAbwBlAHkA')))
    Set-WordBuiltInProperty -WdBuiltInProperty wdPropertyComments $([Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('aAB0AHQAcABzADoALwAvAHMAaABhAG4AZQBoAG8AZQB5AC4AZwBpAHQAaAB1AGIALgBpAG8ALwB3AG8AcgBkAGQAbwBjAC8AYwBjAGUARABlAHMAaQBnAG4ARABvAGMA')))
    Set-WordBuiltInProperty -WdBuiltInProperty wdPropertyManager $([Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('cwBoAGEAbgBlAGgAbwBlAHkA')))
#endregion

#region License
$license = @"
MIT License
Copyright (c) 2016-2018 Shane Hoey

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
"@
Add-WordBreak -breaktype Paragraph
Add-WordText -text "This document has been created with cceDesignDoc which has been distributed under the MIT license. For more information visit http://shanehoey.github.io/worddoc/ccedesigndoc" -WDBuiltinStyle wdStyleBookTitle
Add-WordBreak -breaktype Paragraph
#bug with bold/italic in worddoc module
$selection = (Get-WordDocument).application.selection
$selection.font.Bold = $false
$selection.ParagraphFormat.Alignment = 3
Add-WordText -text $license -WDBuiltinStyle wdStyleNormal
Add-WordBreak -breaktype Paragraph
Add-WordText -text "Are you using this commercially? Show your appreciation and encourage more development of this script at https://paypal.me/shanehoey" -WDBuiltinStyle wdStyleIntenseQuote
Add-WordBreak -breaktype NewPage
#endregion

#region Add TOC
    Add-WordBreak -breaktype Paragraph
    Add-WordText -text 'Table of Contents' -WDBuiltinStyle wdStyleTitle
    Add-WordTOC 
    Add-WordBreak -breaktype NewPage 
#endregion 

#region Overview
    Add-WordText -text "Overview" -WDBuiltinStyle wdStyleHeading1 
    Add-WordBreak -breaktype Paragraph  
    Add-WordText -text $textOverview -WDBuiltinStyle wdStyleBodyText 
    Add-WordBreak -breaktype Paragraph 
#endregion

#region Supporting Documents
    Add-WordText -text "Supporting Documentation" -WDBuiltinStyle  wdStyleHeading1
    Add-WordText -text "The following supporting documentation should be used as a reference when reading this high Level this Design Document" -WDBuiltinStyle wdStyleBodyText
    Add-WordBreak -breaktype Paragraph
    foreach ($object in $objectRelatedDocuments)
    { 
        [void](Get-WordDocument).Hyperlinks.Add((Get-WordDocument).application.selection.range,$object.DocumentLink,$null,$null,$object.DocumentName)
        Add-WordBreak -breaktype Paragraph
    }
    Add-WordBreak -breaktype NewPage
#endregion

#region Cloud Connector Design 
    Add-WordText -text "Cloud Connector Design " -WDBuiltinStyle wdStyleHeading1 
    Add-WordBreak -breaktype Paragraph
    Add-WordText -text $textDesign
    #Bullets not yet done so need to manually create it 
    $selection = (Get-WordDocument).application.selection
    $selection.range.ListFormat.ApplyBulletDefault()
    foreach ($object in $Listdesign) 
    {
        $selection.typetext($object)
        $Selection.TypeParagraph()
    }
    $selection.range.ListFormat.ApplyBulletDefault()
#endregion 

#region Cloud Connector PlatForm 
Add-WordBreak -breaktype NewPage
Add-WordText -text "Recommended CCE Appliances" -WDBuiltinStyle wdStyleHeading2
$object = $objecthardware | Select-Object manufacture,@{name="Hardware Appliance";Expression={$_.hardwareappliance }},hardware,Processor,Ram,Disk,Network,link
foreach ($o in $object) { 
    Add-WordTable -Object $o -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $FALSE
    }
Add-WordBreak -breaktype NewPage
Add-WordText -text "Recommended IP Phones" -WDBuiltinStyle wdStyleHeading2
$object = $objectIPPhone
foreach ($o in $object) { 
    Add-WordTable -Object $o -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $FALSE
    }
#endregion 

#region Cloud Connector Appliance
Add-WordBreak -breaktype NewPage
Add-WordText -text 'Cloud Connector Appliance' -WDBuiltinStyle wdStyleHeading1 
Add-WordBreak -breaktype Paragraph
Add-WordText -text $textCloudConnectorAppliance 
Add-WordBreak -breaktype Paragraph
Add-WordText -text 'Sip domains' -WDBuiltinStyle wdStyleHeading2
Add-WordText -text $textSipDomains
#BUG IN add-word tables means I need to add additional property
$object = $CCESettings.'sip domains'.foreach({[pscustomobject]@{"Sip Domain"=$_;"Default"=$null}})
foreach($o in $object | select 'sip domain') { 
    Add-WordTable -Object $o -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5'  -NoParagraph -FirstColumn $false -BandedRow $false -RemoveProperties
}
Add-WordBreak -breaktype Paragraph
Add-WordText -text 'WSUS Settings' -WDBuiltinStyle wdStyleHeading2
Add-WordText -text $textWSUSSettings
$object = $CCESettings | select-object "WSUS Server","WSUS Status"
Add-WordTable -Object $object  -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
Add-WordText -text 'Other Settings' -WDBuiltinStyle wdStyleHeading2
$object = $CCESettings | select-object "Federation FQDN","Base VM"
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion

#region CCESite Details 
Add-WordText -text 'PSTN Site' -WDBuiltinStyle wdStyleHeading2 
Add-WordText -text $textPSTNSite -WDBuiltinStyle wdStyleBodyText 
Add-WordTable -Object $CCESite -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion 

#region Network Details
Add-WordBreak -breaktype NewPage
Add-WordText -text 'Network Details' -WDBuiltinStyle wdStyleHeading1 
Add-WordBreak -breaktype Paragraph  
Add-WordText -text $textNetwork -WDBuiltinStyle wdStyleBodyText -WdColor wdColorRed
Add-WordBreak -breaktype Paragraph  
Add-WordText -text 'Corporate Network' -WDBuiltinStyle wdStyleHeading2
Add-WordText -text $textCorpNetwork -WDBuiltinStyle wdStyleBodyText
Add-WordTable -Object $Corpnet -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
Add-WordText -text 'Internet Network' -WDBuiltinStyle wdStyleHeading2 
Add-WordText -text $textInternetNetwork -WDBuiltinStyle wdStyleBodyText
Add-WordTable -Object $DMZNet -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
Add-WordText -text 'Management Network' -WDBuiltinStyle wdStyleHeading2
Add-WordText -text $textManagementNetwork -WDBuiltinStyle wdStyleBodyText
Add-WordTable -Object $ManagementNet -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion 

#region Servers
Add-WordBreak -breaktype NewPage  
Add-WordText -text 'Servers' -WDBuiltinStyle wdStyleHeading1 
Add-WordBreak -breaktype Paragraph  
Add-WordText -text $textServers -WDBuiltinStyle wdStyleBodyText 
Add-WordBreak -breaktype Paragraph  
#endregion 

#region DomainController
Add-WordText -text 'AD Domain Controller' -WDBuiltinStyle wdStyleHeading2 
Add-WordText -text $textDomainController -WDBuiltinStyle wdStyleBodyText
$object = $CCESettings | select-object "Active Directory Domain FQDN","Active Directory Domain Netbios"
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true
Add-WordTable -Object $DomainController -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion

#region PrimaryCMS
Add-WordText -text 'Primary CMS' -WDBuiltinStyle wdStyleHeading2 
Add-WordText -text $textCMS -WDBuiltinStyle wdStyleBodyText
Add-WordTable -Object $cms -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion

#region MediationServer
Add-WordText -text 'Mediation Server' -WDBuiltinStyle wdStyleHeading2 
Add-WordText -text $textMediation -WDBuiltinStyle wdStyleBodyText 
Add-WordTable -Object $Mediation -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion

#region EdgeServer
Add-WordText -text 'Edge Server' -WDBuiltinStyle wdStyleHeading2
Add-WordText -text $textEdgeServer -WDBuiltinStyle wdStyleBodyText
Add-WordText -text 'Internal Interface' -WDBuiltinStyle wdStyleHeading4 
$object = $Edge | select-object "Edge Pool name","Edge Server name" ,"AD Domain","IP Address","Subnet","Default Gateway","DNS Server"
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
Add-WordText -text 'DMZ Interface' -WDBuiltinStyle wdStyleHeading4 
$object = $Edge | select-object "SIP Pool Name","SIP IP Address","Media Relay Pool Name","Media Relay IP Address","Media Relay Public IP Address","Media Relay Port Range"
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion

#region Gateway
Add-WordBreak -breaktype NewPage
Add-WordText -text 'Gateways' -WDBuiltinStyle wdStyleHeading1
Add-WordBreak -breaktype Paragraph
Add-WordBreak -breaktype Paragraph
Add-WordText -text $textGateways -WDBuiltinStyle wdStyleBodyText 
foreach ($gateway in $gateways) { 
  Add-WordText  -text "Gateway $($Gateway.'gateway FQDN')" -WDBuiltinStyle wdStyleHeading2 
  Add-WordTable -Object $gateway -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false  -BandedRow $false -FirstColumn $true 
}
#endregion

#region TRUNK CONFIGURATION 
Add-WordText -text 'Trunk Configuration' -WDBuiltinStyle wdStyleHeading2 
Add-wordtext -text $textTrunkConfiguration  -WDBuiltinStyle wdStyleBodyText 
Add-WordTable -Object $TrunkConfiguration -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion

#region HybridVoices
Add-WordText -text 'Voice Routes' -WDBuiltinStyle wdStyleHeading2 
Add-wordtext -text $textHybridVoiceRoutes -WDBuiltinStyle wdStyleBodyText 
Add-WordTable -Object $HybridVoiceRoutes -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false 
#endregion

#region Firewalls
Add-WordBreak -breaktype NewPage  
Add-WordText -text 'Firewall' -WDBuiltinStyle wdStyleHeading1 
Add-WordBreak -breaktype Paragraph
Add-WordText -text $textFirewalls -WDBuiltinStyle wdStyleBodyText
Add-WordBreak -breaktype Paragraph
Add-WordText -text 'Internal Firewall' -WDBuiltinStyle wdStyleHeading2
Add-WordTable -Object ($firewall | where {$_.firewall -eq "Internal"} | sort ruleset | select 'Source IP', 'Source Port','Destination IP','Destination Port') -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false 
Add-WordText -text 'External Firewall - Minimum Configuration' -WDBuiltinStyle wdStyleHeading2
Add-WordTable -Object ($firewall | where {$_.firewall -eq "External-Minimum"}| select 'Source IP', 'Source Port','Destination IP','Destination Port') -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false 
Add-WordText -text 'External Firewall - Recommended Configuration' -WDBuiltinStyle wdStyleHeading2
Add-WordTable -Object ($firewall | where {$_.firewall -eq "External-Recommended"}| select 'Source IP', 'Source Port','Destination IP','Destination Port') -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false 
Add-WordText -text 'Appliance Internet Firewall' -WDBuiltinStyle wdStyleHeading2
Add-WordTable -Object ($firewall | where {$_.firewall -eq "Host"}| select 'Source IP', 'Source Port','Destination IP','Destination Port') -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false 
#endregion 

#region Certificates
Add-WordBreak -breaktype NewPage  
Add-WordText -text 'Certificates' -WDBuiltinStyle wdStyleHeading1 
Add-WordBreak -breaktype Paragraph  
Add-WordText -text $textCertificates -WDBuiltinStyle wdStyleBodyText  -WdColor wdColorRed
Add-WordBreak -breaktype Paragraph  
Add-WordText -text 'Option 1' -WDBuiltinStyle wdStyleHeading2 
Add-WordTable -Object ($certificates | where {$_.cert -eq 'option1'}) -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false 
Add-WordText -text 'Option 2' -WDBuiltinStyle wdStyleHeading2 
Add-WordTable -Object  ($certificates | where {$_.cert -eq 'option2'}) -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false 
Add-WordText -text 'Option 3' -WDBuiltinStyle wdStyleHeading2 
Add-WordTable -Object  ($certificates | where {$_.cert -eq 'option3'}) -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false 
Update-WordTOC
#endregion 

#region External DNS Resolution 
Add-WordBreak -breaktype NewPage  
Add-WordText -text 'External DNS Resolution' -WDBuiltinStyle wdStyleHeading1 
Add-WordBreak -breaktype Paragraph  
Add-WordText -text $textExternalDNSresolution -WDBuiltinStyle wdStyleBodyText  -WdColor wdColorRed
Add-WordBreak -breaktype Paragraph  
#endregion 

#region Close Document
Save-WordDocument -filename $SaveAsFile -wordsaveformat wdFormatDocument
Close-WordDocument
Close-WordInstance
#endregion

# SIG # Begin signature block
# MIINCgYJKoZIhvcNAQcCoIIM+zCCDPcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU5r5Y5Q0cq1rTgtzhsMmUtvJs
# Km2gggpMMIIFFDCCA/ygAwIBAgIQDq/cAHxKXBt+xmIx8FoOkTANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE4MDEwMzAwMDAwMFoXDTE5MDEw
# ODEyMDAwMFowUTELMAkGA1UEBhMCQVUxGDAWBgNVBAcTD1JvY2hlZGFsZSBTb3V0
# aDETMBEGA1UEChMKU2hhbmUgSG9leTETMBEGA1UEAxMKU2hhbmUgSG9leTCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBANAI9q03Pl+EpWcVZ7PQ3AOJ17k6
# OoS9SCIbZprs7NhyRIg7mKzxdcHMnjKwUe/7NDlt5mYzXT2yY/0MeUkyspiEs1+t
# eiHJ6IIs9llWgPGOkV4Ro5fZzlutqeeaomEW/ulH7mVjihVCR6mP/O09YSNo0Dv4
# AltYmVXqhXTB64NdwupL2G8fmTmVUJsww9abtGxy3mhL/l2W3VBcozZbCZVw363p
# 9mjeR9WUz5AxZji042xldKB/97cNHd/2YyWuJ8eMlYfRqz1nVgmmpuU+SuApRult
# hy6wNEngVmJBVhH/a8AH29dEZNL9pzhJGRwGBFi+m/vIr5SFhQVFZYJy79kCAwEA
# AaOCAcUwggHBMB8GA1UdIwQYMBaAFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB0GA1Ud
# DgQWBBROEIC6bKfPIk2DtUTZh7HSa5ajqDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0l
# BAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1oDOgMYYvaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwNaAzoDGGL2h0dHA6Ly9j
# cmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMEwGA1UdIARF
# MEMwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2lj
# ZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGEBggrBgEFBQcBAQR4MHYwJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBOBggrBgEFBQcwAoZCaHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRENvZGVT
# aWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggEBAIly
# KESC2V2sBAl6sIQiHRRgQ9oQdtQamES3fVBNHwmsXl76DdjDURDNi6ptwve3FALo
# ROZHkrjTU+5r6GaOIopKwE4IXkboVoPBP0wJ4jcVm7kcfKJqllSBGZfpnSUjlaRp
# EE5k1XdVAGEoz+m0GG+tmb9gGblHUiCAnGWLw9bmRoGbJ20a0IQ8jZsiEq+91Ft3
# 1vJSBO2RRBgqHTama5GD16OyE3Aps5ypaKYXuq0cnNZCaCasRtDJPolSP4KQ+NVg
# Z/W/rDiO8LNOTDwGcZ2bYScAT88A5KX42wiKnKldmyXnd4ffrwWk8fPngR5sVhus
# Arv6TbwR8dRMGwXwQqMwggUwMIIEGKADAgECAhAECRgbX9W7ZnVTQ7VvlVAIMA0G
# CSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0
# IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMzEwMjIxMjAwMDBaFw0yODEwMjIxMjAw
# MDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNV
# BAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNz
# dXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZz9D7RZmxOttE9X/lqJ3bMtdx6nadBS63
# j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnkoOn7p0WfTxvspJ8fTeyOU5JEjlpB3gvmh
# hCNmElQzUHSxKCa7JGnCwlLyFGeKiUXULaGj6YgsIJWuHEqHCN8M9eJNYBi+qsSy
# rnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q8grkV7tKtel05iv+bMt+dDk2DZDv5LVO
# pKnqagqrhPOsZ061xPeM0SAlI+sIZD5SlsHyDxL0xY4PwaLoLFH3c7y9hbFig3NB
# ggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPRAgMBAAGjggHNMIIByTASBgNVHRMBAf8E
# CDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDAzB5
# BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0
# LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0
# cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# bDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDBPBgNVHSAESDBGMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEF
# BQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAKBghghkgBhv1sAzAd
# BgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHwYDVR0jBBgwFoAUReuir/SS
# y4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQELBQADggEBAD7sDVoks/Mi0RXILHwl
# KXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L/e8q3yBVN7Dh9tGSdQ9RtG6ljlriXiSB
# ThCk7j9xjmMOE0ut119EefM2FAaK95xGTlz/kLEbBw6RFfu6r7VRwo0kriTGxycq
# oSkoGjpxKAI8LpGjwCUR4pwUR6F6aGivm6dcIFzZcbEMj7uo+MUSaJ/PQMtARKUT
# 8OZkDCUIQjKyNookAv4vcn4c10lFluhZHen6dGRrsutmQ9qzsIzV6Q3d9gEgzpkx
# Yz0IGhizgZtPxpMQBvwHgfqL2vmCSfdibqFT+hKUGIUukpHqaGxEMrJmoecYpJpk
# Ue8xggIoMIICJAIBATCBhjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNl
# cnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdp
# Q2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBAhAOr9wAfEpcG37G
# YjHwWg6RMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkG
# CSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEE
# AYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQtgT+3NHF4gqD0db9jce/xmBlkdDANBgkq
# hkiG9w0BAQEFAASCAQCJF9LNHorW+JdJy+40325hy547k/g2ig0tFKO3T2S+TSEO
# N2TGWZZQn5ZB6xoQ62T6EaEpgYfNFu0D+lEHq+k2pcqAFHYC8gaixGWrE9IxPIRJ
# 9hDNhd6KdPrENX0rJVe+Fq0OP1rIwSW1FXX631QzVNgN3YRpMr30wVV7yJeWfKHP
# ce+YdE2NgbeQBC0DzchMCxe2F1s0aXFt2D130VRbyuExytRUTWBAC6EeWigFbIw1
# ApWXOEtvvrRrZID3/Ezu6S1E+iO13tlpR4lg+KIi66m5+MYY/Zm2ookx/wxG0NRt
# vMkpa9nm2gjLRPrHTK/MuR52m0RTWkgnIZVXHSXL
# SIG # End signature block
