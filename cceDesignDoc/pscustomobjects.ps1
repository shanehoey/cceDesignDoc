
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