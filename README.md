![PowerShell](https://img.shields.io/badge/PowerShell-5+-blue)
![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/Get-MAC)

# Get-MAC
Module for looking up OUI/MAC address vendors/organistaions. Works offline once OUI is downloaded once.

## 📦 Installation
```powershell
Install-Module -Name Get-MAC
```

---

## 🚀 Quickstart

### Create local DB + CLI lookup

```powershell
Update-MACdatabase -VerboseLogging
Get-MAC '08:EA:44'
```
<img width="826" height="201" alt="image" src="https://github.com/user-attachments/assets/ff39e4da-30d1-488f-8f6f-462aeffa3955" />

### Multi-lookup from pipeline

```powershell
$MAC = @('00.10.20','00-11-21','00:12:22','001323','00-14-24','00:15:25')
$MAC | Get-Mac
```
<img width="898" height="240" alt="image" src="https://github.com/user-attachments/assets/e9beeb20-e743-48c1-bb95-a627d4e6c3b4" />


### GUI lookup

```powershell
Get-MACGui
```
![Record](https://github.com/user-attachments/assets/524b1c1c-3a11-4852-940b-6618111d68ec)

---

## 📌 Features
 - Lookup MAC-vendor/company
 - No API key, no limiting or throttling
 - Works offline, after running `Update-MACdatabase` once
 - GUI for live response
 - Multiple inputs from pipeline

---

## 🔍 Sources
| URL                                                         | Type |
|-------------------------------------------------------------|------|
| https://standards-oui.ieee.org/oui/oui.csv                  | MA-L |
| https://standards-oui.ieee.org/oui28/mam.csv                | MA-M |
| https://standards-oui.ieee.org/oui36/oui36.csv              | MA-S |
| https://standards-oui.ieee.org/iab/iab.csv                  | IAB  |
---

## 💻 Cmdlets

### `Update-MACDatabase`

Downloads the latest IEEE OUI database and saves it for offline use.

#### Parameters

| Name             | Type     | Required | Description                                                       |
|------------------|----------|----------|-------------------------------------------------------------------|
| `MacDBFolder`    | `String` | No       | Folder path to save the database (default: user's profile path)   |
| `VerboseLogging` | `Switch` | No       | Enables verbose console output for debugging                      |

---

### `Get-MAC`

Searches the offline MAC database for a vendor/organization by MAC or OUI.

#### Parameters

| Name             | Type     | Required | Description                                                       |
|------------------|----------|----------|-------------------------------------------------------------------|
| `OUI`            | `String` | Yes      | Full or partial MAC address or OUI to search                      |
| `MacDBFolder`    | `String` | No       | Folder path to load the database (default: user's profile path)   |
| `VerboseLogging` | `Switch` | No       | Enables verbose console output for debugging                      |

---

### `Get-MACGui`

Opens a graphical user interface for MAC/OUI lookup.

#### Parameters

| Name             | Type     | Required | Description                                                       |
|------------------|----------|----------|-------------------------------------------------------------------|
| `MacDBFolder`    | `String` | No       | Folder path to load the database (default: user's profile path)   |
| `VerboseLogging` | `Switch` | No       | Enables verbose console output for debugging                      |

---
