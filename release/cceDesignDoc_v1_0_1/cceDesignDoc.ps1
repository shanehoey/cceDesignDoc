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
    try {
        if ($PSBoundParameters.ContainsKey('CloudConnectorFile')) 
            {
                $CloudConnectorContent = get-content (get-item -path $CloudConnectorFile).fullname
            }    
        else
            { 
                switch (($host.ui.PromptForChoice("","Do you want to use an existing cloudconnector.ini or download a sample ??", [System.Management.Automation.Host.ChoiceDescription[]]((New-Object System.Management.Automation.Host.ChoiceDescription "&Existing"), (New-Object System.Management.Automation.Host.ChoiceDescription "&Download","Download a cloudconnect.ini from shanehoey.com")), 1)))
                {
                    0   {  
                            Write-warning -Message "Due to a bug the open file dialog box may be behind other windows"
                            Add-Type -AssemblyName System.Windows.Forms
                            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
                            $OpenFileDialog.initialDirectory =  [Environment]::GetFolderPath('MyDocuments')
                            $OpenFileDialog.filter = 'CloudConnector.ini (*.ini)|*.ini'
                            $OpenFileDialog.title = 'Select cloudconnector.ini to import'
                            $result = $OpenFileDialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
                            if ($result -eq [Windows.Forms.DialogResult]::OK)
                            {
                                $CloudConnectorFile = $OpenFileDialog.filename
                            }
                            else 
                            {
                               Write-Verbose "No file selected" -VERBOSE
                               throw "No File Selected"
                            } 
                            Remove-Variable -Name OpenFileDialog
                        }
                    1   {  
                            Write-verbose -Message "Downloading CloudConnectorSample.ini from shanehoey.com/cloudconnectorsample.ini" -Verbose
                            $CloudConnectorContent = (invoke-WebRequest -uri "https://shanehoey.com/cloudconnectorsample.ini" -ContentType "text/plain").content -split '\n'
                            Write-verbose -Message "Downloading CloudConnectorSample.ini Complete" -Verbose
                        }
                }
            }
        }
    catch 
        {
            Write-warning "Sorry unable to get cloudconnector.ini file, please try again"
            Break
        }

    if ($PSBoundParameters.ContainsKey('TemplateFile')) 
        {
            $TemplateFile = (get-item -path $TemplateFile).fullname
        }    
    else
        { 
            switch (($host.ui.PromptForChoice("","Do you want to use an existing word Document as a Template ??", [System.Management.Automation.Host.ChoiceDescription[]]((New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"), (New-Object System.Management.Automation.Host.ChoiceDescription "&No")), 1)))
            {
                0   {  
                            Write-warning -Message "Due to a bug the open file dialog box may be behind other windows"
                            Add-Type -AssemblyName System.Windows.Forms
                            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
                            $OpenFileDialog.initialDirectory =  [Environment]::GetFolderPath('MyDocuments')
                            $OpenFileDialog.filter = 'Word Document (*.docx)|*.docx|Word Template (*.dotx)|*.dotx'
                            $OpenFileDialog.title = 'Select Word Template to import'
                            $result = $OpenFileDialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
                            if ($result -eq [Windows.Forms.DialogResult]::OK)
                            {
                                $TemplateFile = $OpenFileDialog.filename
                            }
                            else 
                            {
                               Write-Verbose "No file selected" -VERBOSE
                            } 
                            Remove-Variable -Name OpenFileDialog
                    }
                1   {  }
            }
        }

    if ($PSBoundParameters.ContainsKey('SaveAsFile')) 
        {
            $SaveAsFile = [System.IO.Path]::GetFullPath($SaveAsFile)
        }
    else    
        {   
           # redundant no longer needed
           # $SaveFileDialog = New-Object -TypeName System.Windows.Forms.saveFileDialog
           #$SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
           # $SaveFileDialog.filter = 'Word Document (*.docx)|*.docx|HTML File (*.html)|*.html|PDF (*.pdf)|*.pdf'
           # $SaveFileDialog.title = "Save As:"
           # [void]$SaveFileDialog.ShowDialog()
           # $SaveAsFile = $SaveFileDialog.FileName
           # remove-variable -Name SaveFileDialog
        }
    
    
    try 
    {
        switch (($host.ui.PromptForChoice("","Do you want to download a basic design template ??", [System.Management.Automation.Host.ChoiceDescription[]]((New-Object System.Management.Automation.Host.ChoiceDescription "&No"), (New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Download a standard design text example")), 1)))
        {
            0   {  
                $includetext = $false
                }
            1   {  
                    Write-verbose -Message "Downloading design document from shanehoey.com/cloudconnector.json" -Verbose
                    $CloudConnectorText = ((invoke-WebRequest -uri "https://shanehoey.com/cloudconnector.json" -ContentType "text/plain").content -split '\n') | convertfrom-json
                    Write-verbose -Message "Downloading CloudConnector.json Complete" -Verbose
                }
        }
    } 
    catch
    {
        write-warning "Unable to download cloudconnector design template, Defaulting to no template"
        Remove-Variable CloudConnectorText -ErrorAction SilentlyContinue
    }


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


#region import cloudconnector.ini
#Credit Oliver Lipkau
#https://blogs.technet.microsoft.com/heyscriptingguy/2011/08/20/use-powershell-to-work-with-any-ini-file/
$cloudconnector = @{}
switch -regex ($CloudConnectorContent)
{
    "^\[([^\\].+)\]$" # Section
    {
        Write-Verbose "SECTION ---> $_" 
        $section = $matches[1].replace(" ","")
        $cloudconnector[$section] = @{}
        $CommentCount = 0
    }
    "^(.+?)\s*=(.*)$" # Key
    {
        Write-Verbose "KEY    ---> $_" 
        $name,$value = $matches[1..2]
        $cloudconnector[$section][$name] = $value
    }
    default { Write-verbose "IGNORE  ---> $_" }
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

Write-Verbose "Creating Word Documents" -Verbose

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
$license = "MIT License`nCopyright (c) 2016-2018 Shane Hoey`rPermission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the 'Software'), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:`nThe above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.`nTHE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."
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
    if($CloudConnectorText) 
        { 
            Add-WordBreak -breaktype Paragraph  
            Add-WordText -text $CloudConnectorText.textOverview -WDBuiltinStyle wdStyleBodyText 
        }
    Add-WordBreak -breaktype Paragraph 
#endregion

#region Supporting Documents
    Add-WordText -text "Supporting Documentation" -WDBuiltinStyle  wdStyleHeading1
    Add-WordText -text "The following supporting documentation should be used as a reference when reading this high Level this Design Document" -WDBuiltinStyle wdStyleBodyText
    Add-WordBreak -breaktype Paragraph
    if($CloudConnectorText) 
        { 
            foreach ($object in $CloudConnectorText.objectRelatedDocuments)
            { 
                    [void](Get-WordDocument).Hyperlinks.Add((Get-WordDocument).application.selection.range,$object.DocumentLink,$null,$null,$object.DocumentName)
                    Add-WordBreak -breaktype Paragraph
            }
         }
    Add-WordBreak -breaktype NewPage
#endregion

#region Cloud Connector Design 
    Add-WordText -text "Cloud Connector Design " -WDBuiltinStyle wdStyleHeading1 
    Add-WordBreak -breaktype Paragraph
    if($CloudConnectorText) 
    { 
        Add-WordText -text $CloudConnectorText.textDesign
        #Bullets not yet done so need to manually create it 
        $selection = (Get-WordDocument).application.selection
        $selection.range.ListFormat.ApplyBulletDefault()
        foreach ($object in $CloudConnectorText.listDesign) 
        {
            $selection.typetext($object)
            $Selection.TypeParagraph()
        }
        $selection.range.ListFormat.ApplyBulletDefault()
    }
#endregion 

#region Cloud Connector PlatForm 
Add-WordBreak -breaktype NewPage
Add-WordText -text "Recommended CCE Appliances" -WDBuiltinStyle wdStyleHeading2
if($CloudConnectorText) 
{ 
    $object = $CloudConnectorText.objecthardware | Select-Object manufacture,@{name="Hardware Appliance";Expression={$_.hardwareappliance }},hardware,Processor,Ram,Disk,Network,link
    foreach ($o in $object) { 
        Add-WordTable -Object $o -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $FALSE
        }
    Add-WordBreak -breaktype NewPage
}
Add-WordText -text "Recommended IP Phones" -WDBuiltinStyle wdStyleHeading2
if($CloudConnectorText) 
{ 
    foreach ($o in $CloudConnectorText.objectIPPhone) 
    { 
        Add-WordTable -Object $o -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $FALSE
    }
}
#endregion 

#region Cloud Connector Appliance
Add-WordBreak -breaktype NewPage
Add-WordText -text 'Cloud Connector Appliance' -WDBuiltinStyle wdStyleHeading1 
Add-WordBreak -breaktype Paragraph
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textCloudConnectorAppliance 
    Add-WordBreak -breaktype Paragraph
}
Add-WordText -text 'Sip domains' -WDBuiltinStyle wdStyleHeading2
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textSipDomains
}
#BUG IN add-word tables means I need to add additional property
$object = $CCESettings.'sip domains'.foreach({[pscustomobject]@{"Sip Domain"=$_;"Default"=$null}})
foreach($o in $object | select 'sip domain') { 
        Add-WordTable -Object $o -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5'  -NoParagraph -FirstColumn $false -BandedRow $false -RemoveProperties
}
Add-WordBreak -breaktype Paragraph
Add-WordText -text 'WSUS Settings' -WDBuiltinStyle wdStyleHeading2
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textWSUSSettings
}
$object = $CCESettings | select-object "WSUS Server","WSUS Status"
Add-WordTable -Object $object  -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
Add-WordText -text 'Other Settings' -WDBuiltinStyle wdStyleHeading2
$object = $CCESettings | select-object "Federation FQDN","Base VM"
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion

#region CCESite Details 
Add-WordText -text 'PSTN Site' -WDBuiltinStyle wdStyleHeading2 
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textPSTNSite -WDBuiltinStyle wdStyleBodyText 
}
Add-WordTable -Object $CCESite -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion 

#region Network Details
Add-WordBreak -breaktype NewPage
Add-WordText -text 'Network Details' -WDBuiltinStyle wdStyleHeading1 
Add-WordBreak -breaktype Paragraph 
if($CloudConnectorText) 
{  
    Add-WordText -text $CloudConnectorText.textNetwork -WDBuiltinStyle wdStyleBodyText -WdColor wdColorRed
    Add-WordBreak -breaktype Paragraph  
}
Add-WordText -text 'Corporate Network' -WDBuiltinStyle wdStyleHeading2
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textCorpNetwork -WDBuiltinStyle wdStyleBodyText
}
Add-WordTable -Object $Corpnet -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
Add-WordText -text 'Internet Network' -WDBuiltinStyle wdStyleHeading2 
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textInternetNetwork -WDBuiltinStyle wdStyleBodyText
}
Add-WordTable -Object $DMZNet -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
Add-WordText -text 'Management Network' -WDBuiltinStyle wdStyleHeading2
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textManagementNetwork -WDBuiltinStyle wdStyleBodyText
}
Add-WordTable -Object $ManagementNet -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion 

#region Servers
Add-WordBreak -breaktype NewPage  
Add-WordText -text 'Servers' -WDBuiltinStyle wdStyleHeading1 
Add-WordBreak -breaktype Paragraph  
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textServers -WDBuiltinStyle wdStyleBodyText 
    Add-WordBreak -breaktype Paragraph  
}
#endregion 

#region DomainController
Add-WordText -text 'AD Domain Controller' -WDBuiltinStyle wdStyleHeading2 
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textDomainController -WDBuiltinStyle wdStyleBodyText
}
$object = $CCESettings | select-object "Active Directory Domain FQDN","Active Directory Domain Netbios"
Add-WordTable -Object $object -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true
Add-WordTable -Object $DomainController -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion

#region PrimaryCMS
Add-WordText -text 'Primary CMS' -WDBuiltinStyle wdStyleHeading2 
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textCMS -WDBuiltinStyle wdStyleBodyText
}
Add-WordTable -Object $cms -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion

#region MediationServer
Add-WordText -text 'Mediation Server' -WDBuiltinStyle wdStyleHeading2 
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textMediation -WDBuiltinStyle wdStyleBodyText 
}
Add-WordTable -Object $Mediation -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion

#region EdgeServer
Add-WordText -text 'Edge Server' -WDBuiltinStyle wdStyleHeading2
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textEdgeServer -WDBuiltinStyle wdStyleBodyText
}
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
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textGateways -WDBuiltinStyle wdStyleBodyText 
    Add-WordBreak -breaktype Paragraph
}
foreach ($gateway in $gateways) { 
  Add-WordText  -text "Gateway $($Gateway.'gateway FQDN')" -WDBuiltinStyle wdStyleHeading2 
  Add-WordTable -Object $gateway -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false  -BandedRow $false -FirstColumn $true 
}
#endregion

#region TRUNK CONFIGURATION 
Add-WordText -text 'Trunk Configuration' -WDBuiltinStyle wdStyleHeading2 
if($CloudConnectorText) 
{ 
    Add-wordtext -text $CloudConnectorText.textTrunkConfiguration  -WDBuiltinStyle wdStyleBodyText 
}
Add-WordTable -Object $TrunkConfiguration -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -BandedRow $false -FirstColumn $true 
#endregion

#region HybridVoices
Add-WordText -text 'Voice Routes' -WDBuiltinStyle wdStyleHeading2 
if($CloudConnectorText) 
{ 
    Add-wordtext -text $CloudConnectorText.textHybridVoiceRoutes -WDBuiltinStyle wdStyleBodyText 
}
Add-WordTable -Object $HybridVoiceRoutes -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false 
#endregion

#region Firewalls
Add-WordBreak -breaktype NewPage  
Add-WordText -text 'Firewall' -WDBuiltinStyle wdStyleHeading1 
Add-WordBreak -breaktype Paragraph
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textFirewalls -WDBuiltinStyle wdStyleBodyText
    Add-WordBreak -breaktype Paragraph
}
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
if($CloudConnectorText) 
{ 
    Add-WordText -text $CloudConnectorText.textCertificates -WDBuiltinStyle wdStyleBodyText  -WdColor wdColorRed
    Add-WordBreak -breaktype Paragraph  
}
Add-WordText -text 'Option 1' -WDBuiltinStyle wdStyleHeading2 
Add-WordTable -Object ($certificates | where {$_.cert -eq 'option1'}) -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false 
Add-WordText -text 'Option 2' -WDBuiltinStyle wdStyleHeading2 
Add-WordTable -Object  ($certificates | where {$_.cert -eq 'option2'}) -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false 
Add-WordText -text 'Option 3' -WDBuiltinStyle wdStyleHeading2 
Add-WordTable -Object  ($certificates | where {$_.cert -eq 'option3'}) -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -BandedRow $false -FirstColumn $false 

#endregion 

#region External DNS Resolution 
Add-WordBreak -breaktype NewPage  
Add-WordText -text 'External DNS Resolution' -WDBuiltinStyle wdStyleHeading1 
Add-WordBreak -breaktype Paragraph  
if($CloudConnectorText) 
{ 
Add-WordText -text $CloudConnectorText.textExternalDNSresolution -WDBuiltinStyle wdStyleBodyText  -WdColor wdColorRed
Add-WordBreak -breaktype Paragraph  
}
#endregion 

#region Close Document
Update-WordTOC 
IF($SaveAsFile) { Save-WordDocument -filename $SaveAsFile -wordsaveformat wdFormatDocument} ELSE {Save-WordDocument}

Close-WordDocument
Close-WordInstance
#endregion

# SIG # Begin signature block
# MIINCgYJKoZIhvcNAQcCoIIM+zCCDPcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUM68T4F0USm7y+JNoj0fSgXz8
# JC6gggpMMIIFFDCCA/ygAwIBAgIQDq/cAHxKXBt+xmIx8FoOkTANBgkqhkiG9w0B
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
# AYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQls29nDGdJRvdclNth1QJ4dwhfPzANBgkq
# hkiG9w0BAQEFAASCAQBqTYppVpWPF5RX+J7h2fwc+uSTghQiRNYnp1eP1S0aVBoB
# iTJAHy2J4y6hSUSIN6CmM/AtYQBziouhWBB/FsP6frxfUOl9XUCYJCnUAQTzL51w
# GU8Utmi/zuZWXUvpogVVCnk5fr3SYc2iNOhsw9t4im6cjnlcWbL+h8b7pfty4RNS
# 1gsh7xNC6adke1qmqtDic6S5QF3UNbRNdPfjvdFl67HF5V5lOOZDHT5r47zZGuFZ
# 9Gvu4qmhC7So04pKZqzZTsrIfjr+yh7UtuzW3RAzAbmj7XlWJe0HU/mzwxOimdUd
# D361NkAwKiES6ftp9f/jVWE8IbzNIyPqLQSyaUb9
# SIG # End signature block
