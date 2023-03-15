# Get-PinInfo
##### This is a script to send mail to the user who Lync dialin PIN will expire
---
### Changelog
  
#### v1.1.1 - 19.04.2016
* Added additional checks AD Group and loaded module availability
  
#### v1.1.0 - 16.04.2016
* Reduced console output
* created summary
* added progress bars
* added event log switch and entries
  
#### v1.0.0 - 13.04.2016
* Original Script (PCA)
  
### How to use it
    .\Get-Pininfo.ps1 -CSGroup Groupname
> The summary output will be returned to the console

    .\Get-Pininfo.ps1 -CSGroup Groupname -ToEvents
> The summary will be saved in event logs

[Visit my website](https://train2play.eu)
