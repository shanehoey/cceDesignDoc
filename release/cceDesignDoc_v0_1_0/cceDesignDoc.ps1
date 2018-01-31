#requires -version 5.0
#requires -module WordDoc

<#
    Requires WordDoc module download from 
    https://gallery.technet.microsoft.com/WordDoc-Create-Word-75739cf9

    Full Documentation aT
    https://bitbucket.org/shanehoey/ccedesigndoc

    Author
    Shane Hoey

#>

Param(  
  [ValidateNotNullOrEmpty()]  
  [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".ini")})]  
  [Parameter(ValueFromPipeline=$True,Mandatory=$True)]  
  [string]$FilePath
)

Import-Module Worddoc -force

#region update this part with your
#$ Lipsum text generated from http://www.randomtext.me/#/lorem/p-4/20-42
$lipsum1 = "Lorem ipsum tempus inceptos urna torquent eros tempor vitae, ut fringilla nullam torquent sociosqu ornare sem eu vel, rutrum habitant quis habitant erat euismod turpis condimentum curabitur rhoncus mauris iaculis ornare fermentum libero dolor cursus lacinia at phasellus."
$lipsum2 = "Posuere netus nec odio elit eros imperdiet adipiscing dolor sem praesent himenaeos elit vestibulum ornare, porttitor nullam sapien malesuada nullam potenti non vulputate imperdiet mauris tempus rutrum integer nunc cubilia sed lectus varius cubilia porttitor taciti sollicitudin odio massa dictum."
$lipsum3 = "Imperdiet mollis a habitant tincidunt iaculis praesent nunc ornare hac feugiat, class torquent elementum venenatis luctus cras vivamus sociosqu etiam sem sit augue sem netus habitasse curabitur netus class nullam ipsum."
$lipsum4 = "Consectetur pellentesque vitae arcu netus morbi vel a sem feugiat vitae, nibh curae non sit hendrerit amet sodales ut mattis ut ad quisque primis rhoncus conubia iaculis tempor erat habitant ante phasellus morbi phasellus."

$textOverview = @" 
$lipsum1
$lipsum2
$lipsum3
$lipsum4
"@

$textDesign = @" 
$lipsum1
"@

$textSiteDetails = @" 
$lipsum1
"@

$textCommonSettings = @" 
$lipsum1
"@

$textNetwork = @" 
$lipsum1
"@

$textCorpNetwork = @" 
$lipsum1
"@

$textInternetNetwork = @" 
$lipsum1
"@

$textManagementNetwork = @" 
$lipsum1
"@

$textServers = @" 
$lipsum1
"@
$textDomainController = @" 
$lipsum1
"@
$textCMS = @" 
$lipsum1
"@
$textMediation = @" 
$lipsum1
"@
$textEdgeServer = @" 
$lipsum1
"@
$textGateways = @" 
$lipsum1
"@

$textFirewalls = @" 
$lipsum1
"@

$textCertificates = @" 
$lipsum1
"@
#endregion 

#region Import Cloud Connector INI

#Credit Oliver Lipkau
#https://blogs.technet.microsoft.com/heyscriptingguy/2011/08/20/use-powershell-to-work-with-any-ini-file/
$CloudConnector = @{}
switch -regex -file $FilePath
{
  “^\[(.+)\]” # Section
  {
    $section = $matches[1]
    $CloudConnector[$section] = @{}
    $CommentCount = 0
  }

  “(.+?)\s*=(.*)” # Key
  {
    $name,$value = $matches[1..2]
    $CloudConnector[$section][$name] = $value
  }
}

#endregion 

#region Create Word Document 
$word = Invoke-Word
$worddoc = New-WordDocument -Word $word
#endregion 

#region Create Cover & TOC
Add-WordCoverPage -CoverPage 'Slice (Dark)' -word $word -WordDoc $worddoc
Add-WordBreak -breaktype NewPage -word $word -WordDoc $worddoc 
Add-WordText -text 'Table of Contents' -WDBuiltinStyle wdStyleTitle -WordDoc $worddoc
Add-WordTOC -word $word -worddoc $worddoc 
#endregion 

#region Overview
Add-WordBreak -breaktype NewPage -word $word -WordDoc $worddoc
Add-WordText -text "Overview" -WDBuiltinStyle wdStyleHeading1 -WordDoc $worddoc
Add-WordBreak -breaktype Paragraph -Word $Word -WordDoc $worddoc
Add-WordText -text $textOverview -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
#endregion

#region Cloud Connector Design 
Add-WordBreak -breaktype NewPage -Word $Word -WordDoc $worddoc
Add-WordText -text "Cloud Connector Design " -WDBuiltinStyle wdStyleHeading1 -WordDoc $worddoc
Add-WordText -text $textDesign -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
Add-WordBreak -breaktype Paragraph -Word $Word -WordDoc $worddoc

#region CCESite Details 
class CCESite{
  [String]$SiteName
  [String]$CountryCode
  [String]$City
  [String]$State

  CCESite(
    [String]$SiteName,
    [String]$CountryCode,
    [String]$City,
    [String]$State
  )
  {
    $this.SiteName = $SiteName
    $this.CountryCode = $CountryCode
    $this.City = $City
    $this.State = $State
  }
  
}

$CCESite = [CCESite]::new(
  $CloudConnector.Common.SiteName,
  $CloudConnector.Common.CountryCode,
  $CloudConnector.Common.City,
  $CloudConnector.Common.State
)

$Object = $CCESite  | select-object -Property @{name = "Site name";Expression = {$_.SiteName} },
@{name = "City";Expression = {$_.City} },
@{name = "State";Expression = {$_.State} },
@{name = "Country";Expression = {$_.CountryCode} }


Add-WordText -text 'Site Details' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
Add-WordText -text $textSiteDetails -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
Add-WordTable -Object $Object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true -WordDoc $worddoc

#endregion 

#region CCESettings
class CCESettings{
  [String]$SipDomains
  [String]$ADDomain
  [String]$ADDomainNetbios
  [String]$FederationFQDN
  [String]$BaseVM
  [String]$WSUSServer
  [String]$WSUSstatus

  CCESettings(
    [String]$SipDomains,
    [String]$ADDomain,
    [String]$ADDomainNetbios,
    [String]$FederationFQDN,
    [String]$BaseVM,
    [String]$WSUSServer,
    [String]$WSUSstatus
  )
  {
    $this.SipDomains = $SipDomains
    $this.ADDomain = $ADDomain
    $this.ADDomainNetbios = $ADDomainNetbios
    $this.FederationFQDN = $FederationFQDN
    $this.BaseVM = $BaseVM
    $this.WSUSServer = $WSUSServer
    $this.WSUSstatus = $WSUSstatus
  }
  
}

$CCESettings = [CCESettings]::new( 
  $CloudConnector.Common.SIPDomains,
  $CloudConnector.Common.VirtualMachineDomain,
  ($CloudConnector.Common.VirtualMachineDomain).Split('.')[0],
  $CloudConnector.Common.OnlineSipFederationFqdn,
  $CloudConnector.Common.BaseVMIP,
  $CloudConnector.Common.WSUSServer,
  $CloudConnector.Common.WSUSStatusServer
)

$object = $CCESettings | Select-Object -Property @{Name = "Sip Domains"; Expression = { $_.SipDomains } },
@{Name = "Federation FQDN"; Expression = { $_.FederationFQDN } },
@{Name = "Active Directory Domain FQDN"; Expression = { $_.ADDomain } },
@{Name = "Active Directory Domain Netbios"; Expression = { $_.ADDomainNetbios } },
@{Name = "Base VM"; Expression = { $_.BaseVM } },
@{Name = "WSUS Server"; Expression = { $_.WSUSServer } },
@{Name = "WSUS Server Status"; Expression = { $_.WSUSstatus } }

Add-WordText -text 'Common Settings' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
Add-WordText -text $textCommonSettings -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
Add-WordTable -Object $Object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true -WordDoc $worddoc

#endregion

#endregion

#region Network 

Add-WordBreak -breaktype NewPage -Word $Word -WordDoc $worddoc
Add-WordText -text 'Network Design' -WDBuiltinStyle wdStyleHeading1 -WordDoc $worddoc
Add-WordText -text $textNetwork -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
Add-WordBreak -breaktype Paragraph -Word $Word -WordDoc $worddoc

class cceNetwork{
  [String]$CorpnetSwitchName
  [String]$CorpnetDefaultGateway
  [String]$CorpnetIPPrefixLength
  [String]$CorpnetDNSIPAddress
  [String]$InternetSwitchName
  [String]$InternetDefaultGateway
  [String]$InternetIPPrefixLength
  [String]$InternetDNSIPAddress
  [String]$ManagementSwitchName
  [String]$ManagementIPPrefix  
  [String]$ManagementIPPrefixLength

  cceNetwork(
    [String]$CorpnetSwitchName,
    [String]$CorpnetDefaultGateway,
    [String]$CorpnetIPPrefixLength,
    [String]$CorpnetDNSIPAddress,
    [String]$InternetSwitchName,
    [String]$InternetDefaultGateway,
    [String]$InternetIPPrefixLength,
    [String]$InternetDNSIPAddress,
    [String]$ManagementSwitchName,
    [String]$ManagementIPPrefix,
    [String]$ManagementIPPrefixLength
  )
  {
    $this.CorpnetSwitchName = $CorpnetSwitchName
    $this.CorpnetDefaultGateway = $CorpnetDefaultGateway
    $this.CorpnetIPPrefixLength = $CorpnetIPPrefixLength
    $this.CorpnetDNSIPAddress = $CorpnetDNSIPAddress
    $this.InternetSwitchName = $InternetSwitchName
    $this.InternetDefaultGateway = $InternetDefaultGateway
    $this.InternetIPPrefixLength = $InternetIPPrefixLength
    $this.InternetDNSIPAddress = $InternetDNSIPAddress
    $this.ManagementSwitchName = $ManagementSwitchName
    $this.ManagementIPPrefix = $ManagementIPPrefix
    $this.ManagementIPPrefixLength = $ManagementIPPrefixLength
  }
  
}

$cceNetwork = [cceNetwork]::new(
  $CloudConnector.Network.CorpnetSwitchName,
  $CloudConnector.Network.CorpnetDefaultGateway,
  $CloudConnector.Network.CorpnetIPPrefixLength,
  $CloudConnector.Network.CorpnetDNSIPAddress,
  $CloudConnector.Network.InternetSwitchName,
  $CloudConnector.Network.InternetDefaultGateway,
  $CloudConnector.Network.InternetIPPrefixLength,
  $CloudConnector.Network.InternetDNSIPAddress,
  $CloudConnector.Network.ManagementSwitchName,
  $CloudConnector.Network.ManagementIPPrefix,
  $CloudConnector.Network.ManagementIPPrefixLength
)

$object = $cceNetwork | Select-Object -Property @{Name = "Hyper-V Switch"; Expression = { $_.CorpnetSwitchName } },
@{Name = "Gateway"; Expression = { $_.CorpnetDefaultGateway } },
@{Name = "Subnet"; Expression = { $_.CorpnetIPPrefixLength } },
@{Name = "DNS forwarder"; Expression = { $_.CorpnetDNSIPAddress } }

Add-WordText -text 'Corporate Network' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
Add-WordText -text $textCorpNetwork -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
Add-WordTable -Object $Object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true -WordDoc $worddoc

$object = $cceNetwork | Select-Object -Property @{Name = "Hyper-V switch"; Expression = { $_.InternetSwitchName } },
@{Name = "Gateway"; Expression = { $_.InternetDefaultGateway } },
@{Name = "Subnet"; Expression = { $_.InternetIPPrefixLength } },
@{Name = "DNS forwarder"; Expression = { $_.InternetDNSIPAddress } }

Add-WordText -text 'Internet Network' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
Add-WordText -text $textInternetNetwork -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
Add-WordTable -Object $Object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true -WordDoc $worddoc

$object = $cceNetwork | Select-Object -Property @{Name = "Hyper-V Switch"; Expression = { $_.ManagementSwitchName } },
@{Name = "Network"; Expression = { $_.ManagementIPPrefix } },
@{Name = "Subnet"; Expression = { $_.ManagementIPPrefixLength } }

Add-WordText -text 'Management Network' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
Add-WordText -text $textManagementNetwork -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
Add-WordTable -Object $Object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true -WordDoc $worddoc

#endregion 

#region Servers
Add-WordBreak -breaktype NewPage -Word $Word -WordDoc $worddoc
Add-WordText -text 'Servers' -WDBuiltinStyle wdStyleHeading1 -WordDoc $worddoc
Add-WordText -text $textServers -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
Add-WordBreak -breaktype Paragraph -Word $Word -WordDoc $worddoc


#region DomainController
class DomainController{
  [String]$Servername
  [String]$Domain
  [String]$IP
  [String]$Subnet
  [String]$Gateway
  [String]$DNS
    
  DomainController(
    [String]$Servername,
    [String]$Domain,
    [String]$IP,
    [String]$Subnet,
    [String]$Gateway,
    [String]$DNS
  )
  {
    $this.Servername =  $Servername
    $this.Domain     =  $Domain
    $this.IP         =  $IP
    $this.Subnet     =  $Subnet
    $this.Gateway    =  $Gateway
    $this.DNS        =  $DNS
  }
}

$DomainController = [DomainController]::new( 
  $CloudConnector.Common.ServerName,
  $CCESettings.ADDomain,
  $CloudConnector.Common.IP,
  $cceNetwork.CorpnetIPPrefixLength,
  $cceNetwork.CorpnetDefaultGateway,
  $CloudConnector.Common.IP
)

$Object = $DomainController
Add-WordText -text 'Domain Controller' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
Add-WordText -text $textDomainController -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
Add-WordTable -Object $Object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true -WordDoc $worddoc

#endregion

#region PrimaryCMS
class PrimaryCMS{
  [String]$Servername
  [String]$Domain
  [String]$IP
  [String]$Subnet
  [String]$Gateway
  [String]$DNS
  [String]$ShareName
  
  PrimaryCMS(
    [String]$Servername,
    [String]$Domain,
    [String]$IP,
    [String]$Subnet,
    [String]$Gateway,
    [String]$DNS,
    [String]$ShareName
  )

  {
    $this.Servername  =  $Servername
    $this.Domain      =  $Domain
    $this.IP          =  $IP
    $this.Subnet      =  $Subnet
    $this.Gateway     =  $Gateway
    $this.DNS         =  $DNS
    $this.ShareName   =  $ShareName
  }
  
}

$PrimaryCMS=[PrimaryCMS]::new( 
  $cloudconnector.PrimaryCMS.ServerName,
  $CCESettings.ADDomain,
  $cloudconnector.PrimaryCMS.IP,
  $cceNetwork.CorpnetIPPrefixLength,
  $cceNetwork.CorpnetDefaultGateway,
  $CloudConnector.Common.IP,
  $cloudconnector.PrimaryCMS.ShareName
)


Add-WordText -text 'Primary CMS' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
Add-WordText -text $textCMS -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc

$object = $PrimaryCMS | Select-object @{Name='Server Name';Expression={ $_.Servername}},
domain,
ip,
subnet,
gateway,
DNS

Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true -WordDoc $worddoc

$object = $PrimaryCMS | Select-object @{Name='ShareName';Expression={ "$($_.Servername)\$($_.sharename)"}}
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true -WordDoc $worddoc

#endregion

#region MediationServer
class MediationServer{
  [string]$Poolname
  [String]$Servername
  [String]$Domain
  [String]$IP
  [String]$Subnet
  [String]$Gateway
  [String]$DNS
  
  MediationServer(
    [string]$Poolname,
    [String]$Servername,
    [String]$Domain,
    [String]$IP,
    [String]$Subnet,
    [String]$Gateway,
    [String]$DNS
  )

  {
    $this.Poolname = $Poolname
    $this.Servername = $Servername
    $this.Domain = $DOMAIN
    $this.IP = $IP
    $this.Subnet = $Subnet
    $this.Gateway = $Gateway
    $this.DNS = $DNS
  }
  
}

$MediationServer=[MediationServer]::new( 
  $cloudconnector.MediationServer.PoolName,
  $cloudconnector.MediationServer.ServerName,
  $CCESettings.ADDomain,
  $cceNetwork.CorpnetIPPrefixLength,
  $cceNetwork.CorpnetDefaultGateway,
  $CloudConnector.Common.IP,
  $cloudconnector.MediationServer.IP
)

$Object = $MediationServer  | select-object PoolName,
ServerName,
Domain, 
IP,
Subnet,
Gateway,
DNS

Add-WordText -text 'Mediation Server' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
Add-WordText -text $textMediation -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
Add-WordTable -Object $Object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true -WordDoc $worddoc

#endregion

#region EdgeServer
class EdgeServer{
  [String]$InternalServerName
  [String]$InternalPoolName
  [String]$InternalServerIPs
  [String]$ExternalSIPPoolName
  [String]$ExternalSIPIPs
  [String]$ExternalMRFQDNPoolName
  [String]$ExternalMRIPs
  [String]$ExternalMRPublicIPs
  [String]$ExternalMRPortRange
  
  
  EdgeServer(
    [String]$InternalServerName,
    [String]$InternalPoolName,
    [String]$InternalServerIPs,
    [String]$ExternalSIPPoolName,
    [String]$ExternalSIPIPs,
    [String]$ExternalMRFQDNPoolName,
    [String]$ExternalMRIPs,
    [String]$ExternalMRPublicIPs,
    [String]$ExternalMRPortRange
  )

  {
    $this.InternalServerName = $InternalServerName
    $this.InternalPoolName = $InternalPoolName
    $this.InternalServerIPs = $InternalServerIPs
    $this.ExternalSIPPoolName = $ExternalSIPPoolName
    $this.ExternalSIPIPs = $ExternalSIPIPs
    $this.ExternalMRFQDNPoolName = $ExternalMRFQDNPoolName
    $this.ExternalMRIPs = $ExternalMRIPs
    $this.ExternalMRPublicIPs = $ExternalMRPublicIPs
    $this.ExternalMRPortRange = $ExternalMRPortRange
  }
  
}

$EdgeServer=[EdgeServer]::new( 
  $cloudconnector.EdgeServer.InternalServerName,
  $cloudconnector.EdgeServer.InternalPoolName,
  $cloudconnector.EdgeServer.InternalServerIPs,
  $cloudconnector.EdgeServer.ExternalSIPPoolName,
  $cloudconnector.EdgeServer.ExternalSIPIPs,
  $cloudconnector.EdgeServer.ExternalMRFQDNPoolName,
  $cloudconnector.EdgeServer.ExternalMRIPs,
  $cloudconnector.EdgeServer.ExternalMRPublicIPs,
  $cloudconnector.EdgeServer.ExternalMRPortRange  
)

$object = $EdgeServer
Add-WordText -text 'Edge Server' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
Add-WordText -text $textEdgeServer -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc


$object = $EdgeServer | Select-Object @{Name="Pool Name";Expression={$_.InternalPoolName}},
@{Name="Server Name";Expression={$_.InternalServerName}},
@{Name="IP Address";Expression={$_.InternalServerIPs}}

Add-WordText -text 'Internal' -WDBuiltinStyle wdStyleHeading4 -WordDoc $worddoc
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true -WordDoc $worddoc

$object = $EdgeServer | Select-Object @{Name="Sip Poolname";Expression={$_.ExternalSIPPoolName}},
@{Name="IP Address";Expression={$_.ExternalSIPIPs}},
@{Name="Media Poolname";Expression={$_.ExternalMRFQDNPoolName}},
@{Name="Media IP Address";Expression={$_.ExternalMRIPs}},
@{Name="Media Public IP Address";Expression={$_.ExternalMRPublicIPs}},
@{Name="Media Port Range";Expression={$_.ExternalMRPortRange}}

Add-WordText -text 'External' -WDBuiltinStyle wdStyleHeading4 -WordDoc $worddoc
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true -WordDoc $worddoc


#endregion


#endregion

#region Gateway
Add-WordBreak -breaktype NewPage -Word $Word -WordDoc $worddoc
Add-WordText -text 'Gateways' -WDBuiltinStyle wdStyleHeading1 -WordDoc $worddoc
Add-WordText -text $textGateways -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
Add-WordBreak -breaktype Paragraph -Word $Word -WordDoc $worddoc

class Gateway
{
  [string]$FQDN
  [string]$IP
  [string]$Port
  [string]$Protocol
  [string]$VoiceRoutes
  
  Gateway ([string]$FQDN,[string]$IP,[string]$Port,[string]$Protocol,[string]$VoiceRoutes){
 
    $this.FQDN =$FQDN
    $this.IP = $IP
    $this.Port = $Port
    $this.Protocol = $Protocol
    $this.VoiceRoutes = $VoiceRoutes
  }
}

$gateways = @()
if ($cloudconnector.Gateway1.fqdn -ne $null) { 
  $gateways += [gateway]::new($cloudconnector.Gateway1.fqdn,$cloudconnector.Gateway1.ip,$cloudconnector.Gateway1.port,$cloudconnector.Gateway1.protocol,$cloudconnector.Gateway1.voiceroutes)
}
if ($cloudconnector.Gateway2.fqdn -ne $null) { 
  $gateways += [gateway]::new($cloudconnector.Gateway2.fqdn,$cloudconnector.Gateway2.ip,$cloudconnector.Gateway2.port,$cloudconnector.Gateway2.protocol,$cloudconnector.Gateway2.voiceroutes)
}
if ($cloudconnector.Gateway3.fqdn -ne $null) { 
  $gateways += [gateway]::new($cloudconnector.Gateway3.fqdn,$cloudconnector.Gateway3.ip,$cloudconnector.Gateway3.port,$cloudconnector.Gateway3.protocol,$cloudconnector.Gateway3.voiceroutes)
}
if ($cloudconnector.Gateway4.fqdn -ne $null) { 
  $gateways += [gateway]::new($cloudconnector.Gateway4.fqdn,$cloudconnector.Gateway4.ip,$cloudconnector.Gateway4.port,$cloudconnector.Gateway4.protocol,$cloudconnector.Gateway4.voiceroutes)
}

foreach ($gateway in $gateways) { 
  Add-WordText  -text "Gateway $($Gateway.FQDN)" -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
  Add-WordTable -Object $gateways -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false  -BandedRow $false -FirstColumn $true -WordDoc $worddoc
}

#endregion Gateway

#region TRUNK CONFIGURATION 
Add-WordText -text 'Trunk Configuration' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc

class TrunkConfiguration {
  [bool]$EnableReferSupport
  [bool]$ForwardPAI

  TrunkConfiguration ([bool]$EnableReferSupport,[bool]$ForwardPAI) {
    $this.EnableReferSupport = $EnableReferSupport
    $this.ForwardPAI         = $ForwardPAI
  }
}
$TrunkConfiguration = [TrunkConfiguration]::new($cloudconnector.TrunkConfiguration.EnableReferSupport,$cloudconnector.TrunkConfiguration.ForwardPAI)

Add-WordTable -Object $TrunkConfiguration -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true -WordDoc $worddoc

#endregion

#region HybridVoices
class HybridVoiceRoutes {
  [string]$VoiceRouteName
  [string]$VoiceRoute

  HybridVoiceRoutes ([string]$VoiceRouteName,[string]$VoiceRoute) 
  {
    $this.VoiceRouteName = $VoiceRouteName
    $this.VoiceRoute     = $VoiceRoute
  }
}

$HybridVoiceRoutes = @()
foreach ($routes in $cloudconnector.HybridVoiceRoutes.GetEnumerator()) {$HybridVoiceRoutes += [HybridVoiceRoutes]::new($routes.Name,$routes.value)}
$object = $HybridVoiceRoutes | Select-Object -Property @{
  Name       = 'Voice Route Name'
  Expression = {$_.voiceroutename}
}, @{
  Name       = 'Voice Route'
  Expression = {$_.voiceroute}
}
Add-WordText -text 'Voice Routes' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false -WordDoc $worddoc
#endregion

#region Firewall
class Firewall {
  [String]$Firewall
  [string]$SourceIP
  [string]$SourcePort
  [string]$DestinationIP
  [string]$DestinationPort
  
  Firewall (
    [String]$Firewall,
    [string]$SourceIP,
    [string]$SourcePort,
    [string]$DestinationIP,
    [string]$DestinationPort
  ) 
  {
    $this.Firewall = $Firewall
    $this.SourceIP = $SourceIP
    $this.SourcePort = $SourcePort
    $this.DestinationIP = $DestinationIP
    $this.DestinationPort = $DestinationPort
  }
}

$Firewall = @()

#Gateway
foreach($gateway in $gateways) 
{
  $Firewall += [firewall]::new('Internal', $MediationServer.IP, 'Any', $gateway.IP, "$($gateway.Protocol) $($gateway.Port)")
  $Firewall += [firewall]::new('Internal', $gateway.IP, 'Any', $MediationServer.IP, 'TCP 5068')
  $Firewall += [firewall]::new('Internal', $gateway.IP, 'Any', $MediationServer.IP, 'TLS 5067')
  $Firewall += [firewall]::new('Internal', $MediationServer.IP, 'UDP 49,152 - 57,500', $gateway.IP, 'Any')
  $Firewall += [firewall]::new('Internal', $gateway.IP, 'Any', $MediationServer.IP, 'UDP 49,152 - 57500')
  
  $Firewall += [firewall]::new('Internal', $MediationServer.IP, 'TCP 49,152 - 57,500', 'Internal Clients', 'TCP 50,000-50,019')
  $Firewall += [firewall]::new('Internal', $MediationServer.IP, 'UDP 49,152 - 57,500', 'Internal Clients', 'UDP 50,000-50,019')
  $Firewall += [firewall]::new('Internal', 'Internal Clients', 'TCP 50,000 - 50,019', $MediationServer.IP, 'TCP 49,152-57,500')
  $Firewall += [firewall]::new('Internal', 'Internal Clients', 'UDP 50,000 - 50,019', $MediationServer.IP,  'UDP 49,152-57,500')
}
 
#Edge Server 
$Firewall += [firewall]::new('External', 'Any', 'Any',$EdgeServer.ExternalMRIPs,'TCP 5061')
$Firewall += [firewall]::new('External', $EdgeServer.ExternalMRIPs, 'Any','Any','TCP 5061')
$Firewall += [firewall]::new('External', $EdgeServer.ExternalMRIPs, 'Any','Any','TCP 80')
$Firewall += [firewall]::new('External', $EdgeServer.ExternalMRIPs, 'Any','Any','UDP 53')
$Firewall += [firewall]::new('External', $EdgeServer.ExternalMRIPs, 'Any','Any','TCP 53')
$Firewall += [firewall]::new('External', $EdgeServer.ExternalMRIPs, 'TCP 50,000-59,999','Any','Any')
$Firewall += [firewall]::new('External', $EdgeServer.ExternalMRIPs, 'UDP 3478','Any','Any')
$Firewall += [firewall]::new('External', $EdgeServer.ExternalMRIPs, 'UDP 50,000-59,999','Any','Any')
$Firewall += [firewall]::new('External', 'Any', 'Any',$EdgeServer.ExternalMRIPs,'TCP 443')
$Firewall += [firewall]::new('External', 'Any', 'Any',$EdgeServer.ExternalMRIPs,'TCP 50,000 - 59,999')
$Firewall += [firewall]::new('External', 'Any', 'Any',$EdgeServer.ExternalMRIPs,'UDP 3478')
$Firewall += [firewall]::new('External', 'Any', 'Any',$EdgeServer.ExternalMRIPs,'UDP 50,000 - 59,999')

Add-WordBreak -breaktype NewPage -Word $Word -WordDoc $worddoc
Add-WordText -text 'Firewall' -WDBuiltinStyle wdStyleHeading1 -WordDoc $worddoc
Add-WordText -text $textFirewalls -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc


$object = $Firewall | Where-Object -FilterScript {$_.firewall -eq 'Internal'} | Select-Object -Property @{ Name = 'Source IP'; Expression = {$_.sourceIP }},
@{ Name = 'Source Port'; Expression = {$_.sourcePort }},
@{ Name = 'Destination IP'; Expression = {$_.DestinationIP }},
@{ Name = 'Destination Port'; Expression = {$_.DestinationPort }}
Add-WordText -text 'Internal Firewall' -WDBuiltinStyle wdStyleHeading1 -WordDoc $worddoc
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false -WordDoc $worddoc

$object = $Firewall | Where-Object -FilterScript {$_.firewall -eq 'External'} | Select-Object -Property @{ Name = 'Source IP'; Expression = {$_.sourceIP }},
@{ Name = 'Source Port'; Expression = {$_.sourcePort }},
@{ Name = 'Destination IP'; Expression = {$_.DestinationIP }},
@{ Name       = 'Destination Port'; Expression = {$_.DestinationPort }}

Add-WordText -text 'External Firewall' -WDBuiltinStyle wdStyleHeading1 -WordDoc $worddoc
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false -WordDoc $worddoc

#endregion 

#region Certificates
class Certificates {
  [String]$SubjectName
  [string]$SubjectAlternativeName
  
  Certificates (
    [String]$SubjectName,
    [string]$SubjectAlternativeName
  ) 
  {
    $this.SubjectName = $SubjectName
    $this.SubjectAlternativeName = $SubjectAlternativeName
  }
}


Add-WordBreak -breaktype NewPage -Word $Word -WordDoc $worddoc
Add-WordText -text 'Certificates' -WDBuiltinStyle wdStyleHeading1 -WordDoc $worddoc
Add-WordText -text $textCertificates -WDBuiltinStyle wdStyleBodyText -WordDoc $worddoc
Add-WordBreak -breaktype Paragraph -Word $Word -WordDoc $worddoc

$certificates = [Certificates]::new("$($Edgeserver.ExternalSIPPoolName).$($CCESettings.ADDomain)","sip.$($CCESettings.SIPDomains),$($Edgeserver.ExternalSIPPoolName).$($CCESettings.SIPDomains),{All Additional Edge Servers}")
$object = $certificates | select-object @{ Name = 'Subject Name'; Expression = {$_.SubjectName }},@{ Name = 'Subject Alternative Name'; Expression = {($_.SubjectAlternativeName).split(',') }}

Add-WordText -text 'Option 1' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false -WordDoc $worddoc

$certificates = [Certificates]::new("$($Edgeserver.ExternalSIPPoolName).$($CCESettings.SIPDomains)","sip.$($CCESettings.SIPDomains),*.$($CCESettings.SIPDomains)")
$object = $certificates | select-object @{ Name = 'Subject Name'; Expression = {$_.SubjectName }},@{ Name = 'Subject Alternative Name'; Expression = {($_.SubjectAlternativeName).split(',') }}

Add-WordText -text 'Option 2' -WDBuiltinStyle wdStyleHeading2 -WordDoc $worddoc
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false -WordDoc $worddoc

#endregion 

Set-WordBuiltInProperty -WdBuiltInProperty wdPropertyTitle -text 'Cloud Connector Design' -WordDoc $worddoc
Set-WordBuiltInProperty -WdBuiltInProperty wdPropertySubject -text 'CCE Design Document' -WordDoc $worddoc
Set-WordBuiltInProperty -WdBuiltInProperty wdPropertyCompany -text 'CompanyName' -WordDoc $worddoc
Set-WordBuiltInProperty -WdBuiltInProperty wdPropertyAuthor -text 'Shane Hoey' -WordDoc $worddoc

Update-WordTOC -WordDoc $worddoc