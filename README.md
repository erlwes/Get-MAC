![PowerShell](https://img.shields.io/badge/PowerShell-5+-blue)
![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/Get-MacInfo)

# Get-MacInfo
Module for looking up OUI/MAC address vendors/organistaions. Works offline once OUI is downloaded once.

## 📦 Installation
```powershell
Install-Script -Name Get-MAC
```

---

## 🚀 Quickstart

```powershell
Update-MACdatabase -VerboseLogging
Get-MAC 08:EA:44
```

---

## 📌 Features
 - Lookup MAC-vendor/company
 - No API key, no limiting or throttling
 - Works offline, after running `Update-MACdatabase` once

---

## 💻 Cmdlets

### `Update-MACdatabase`
Content from IEEE is parsed from text into a hashtable, and then saved to file as a hashtable using Export-CliXML

### `Get-MAC`
Lookup of MAC vendor/company information

---
