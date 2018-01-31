# About CCE Design Doc #
This project is hosted on bitbucket [https://bitbucket.org/shanehoey/cceDesignDoc](https://bitbucket.org/shanehoey/cceDesignDoc)


Quickly and effortless create a Skype for Business Cloud Connector Eddition (CCE) Design Document or As Built Document using cloudconnector.ini and Powershell.

**Highlights Include:**

* Generate a full Design or As Built document from the cloudconnectpr.ini file 
* Full List of Servers Created
* Firewall Requirements
* Certificate Requirements

### Important Prerequisites ###
Before attempting to use this module make sure you download a copy of the WordDoc Module from either

* Option 1 - [https://gallery.technet.microsoft.com/WordDoc-Create-Word-75739cf9](https://gallery.technet.microsoft.com/WordDoc-Create-Word-75739cf9)
* Option 2 - [https://bitbucket.org/shanehoey/worddoc](https://bitbucket.org/shanehoey/worddoc)

### Example Usage ###

```
#!powershell
 .\cceDesignDoc.ps1 -filepath .\cloudconnector.ini

```


### Who do I talk to? ###

* Shane Hoey - ShaneHoey.com
* https://bitbucket.org/shanehoey/ccedesigndoc

### Credit and Thanks ###
Massive thanks to the following people for blog there own scripts that have helped me write this module

* Oliver Lipkau - https://blogs.technet.microsoft.com/heyscriptingguy/2011/08/20/use-powershell-to-work-with-any-ini-file/

### License ###
refer to license.txt