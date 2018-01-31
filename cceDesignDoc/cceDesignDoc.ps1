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

. .\customtext.ps1

. .\pscustomobjects.ps1

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
