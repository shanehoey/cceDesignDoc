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