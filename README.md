![PowerShell](https://img.shields.io/badge/PowerShell-5+-blue)
![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/Get-MacInfo)

# Get-MAC
Module for looking up OUI/MAC address vendors/organistaions. Works offline once OUI is downloaded once.

## 📦 Installation
```powershell
Install-Script -Name Get-MAC
```

---

## 🚀 Quickstart

### Create local DB + CLI lookup

```powershell
Update-MACdatabase -VerboseLogging
Get-MAC 08:EA:44
```
<img width="826" height="201" alt="image" src="https://github.com/user-attachments/assets/ff39e4da-30d1-488f-8f6f-462aeffa3955" />

### GUI lookup

```powershell
Get-MACGui
```
![GUI version](https://github.com/user-attachments/assets/cd759099-d251-4fff-b0d8-d5f3dccb9066)

---

## 📌 Features
 - Lookup MAC-vendor/company
 - No API key, no limiting or throttling
 - Works offline, after running `Update-MACdatabase` once

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

---
