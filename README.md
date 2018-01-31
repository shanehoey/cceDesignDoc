# About CCE Design Doc #

Quickly and effortless create a Skype for Business Cloud Connector Eddition (CCE) Design Document or As Built Document using cloudconnector.ini and Powershell.

**Highlights Include:**

* Generate a full Design or As Built document from the cloudconnector.ini file 
* Full List of Servers Created
* Firewall Requirements
* Certificate Requirements

### Important Prerequisites ###
Before attempting to use this module make sure you install the WordDoc Module

[WordDoc](https://shanehoey.github.io/worddoc/)

### Installation ###
Option 1 Install from PowerShell Gallery 

```
Install-Module WordDoc -Scope CurrentUser
Install-Script cceDesignDoc  -Scope CurrentUser
```

Option 2 Download Current release from TechNet Gallery
```
https://gallery.technet.microsoft.com/office/CCE-Design-Doc-Document-a29166ab
```

### Example Usage ###

When you run cceDesignDoc it will prompt you for the following 
 * cloudconnector.ini file 
 * Word Template,  (use your own template to brand the document automatically) click cancle to use blamk document
 * Name of File to save document as 
 
```
 .\cceDesignDoc.ps1
```


### Who do I talk to? ###

* Shane Hoey - ShaneHoey.com
* https://github.com/shanehoey/cceDesignDoc/

### Credit and Thanks ###
Massive thanks to the following people for blog there own scripts that have helped me write this module

* [Oliver Lipkau](https://blogs.technet.microsoft.com/heyscriptingguy/2011/08/20/use-powershell-to-work-with-any-ini-file/)

