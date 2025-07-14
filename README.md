# Get-MacInfo
Module for looking up OUI/MAC address vendors/organistaions. Works offline once OUI is downloaded once.

### Install module
`install-Module Get-MAC`

### Download offline MAC-address database
`Update-MACdatabase`
Content from IEEE is parsed from text into a hashtable, and then saved to file as a hashtable using Export-CliXML

### Search the offline database
`Get-MAC 08:EA:44`
Offline OUI-file is stored as a hashtable. By default it gets saved to PSProfile dir inside a folder names "Lookups"

### To-do
1. More logging and error handling (Add Write-Console function)
2. Suppress console messages unless -VerboseLogging is used (not added yet)
