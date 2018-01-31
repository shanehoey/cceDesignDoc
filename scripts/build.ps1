exit

set-location $env:USERPROFILE\github\ccedesigndoc

$NuGetApiKey
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert
$cert | format-table subject,issuer

$version = "1.0.1"

Update-ScriptFileInfo -Path ".\ccedesigndoc\ccedesigndoc.ps1" -Version $version -Guid "52f620ce-6560-42d9-afac-d1124aa65d1c" -Author "Shane Hoey" -Copyright "2016-2018 Shane Hoey" `
                        -RequiredModules WordDoc -ProjectUri https://shanehoey.github.io/ccedesigndoc -ReleaseNotes https://shanehoey.github.io/ccedesigndoc `
                        -LicenseUri https://shanehoey.github.io/ccedesigndoc/license -Tags "Skype for Business, Skype for Business Online, Microsoft Office, Office" -Description "Create a Design Document from Cloud Connector"

Import-Module -name PowerShellProTools
$script =  ".\cceDesigndoc\cceDesignDoc.ps1"
$bundle =  ".\release\cceDesignDoc_v$($version.replace(".","_"))\"

Merge-Script -Script $script -OutputPath $bundle -Bundle

copy ..\cceDesignDoc\license $bundle

Set-AuthenticodeSignature -filepath "$($bundle)cceDesignDoc.ps1" -Certificate $cert
(Get-AuthenticodeSignature -FilePath "$($bundle)cceDesignDoc.ps1").Status

set-location $bundle

.\cceDesignDoc.ps1  -TemplateFile 'C:\users\shane\OneDrive\Documents\Custom Office Templates\AudioCodes.dotx' `
                      -CloudConnectorFile 'C:\users\shane\OneDrive\Documents\CloudConnector.Sample.ini' `
                      -SaveAsFile 'C:\users\shane\OneDrive\Documents\cceDesignDoc.docx'


### IMPORTANT ONLY RUN AFTER ALL ABOVE IS COMPLETED
pause
Publish-Script -path .\cceDesignDoc.ps1 -NuGetApiKey $NuGetApiKey