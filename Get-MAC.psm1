<#
.SYNOPSIS
    Module for looking up OUI/MAC address vendors/organistaions. Works offline once OUI is downloaded once.

.DESCRIPTION
    Module for looking up OUI/MAC address vendors/organistaions. Works offline once OUI is downloaded once.

.NOTES
    https://github.com/erlwes/Get-MacInfo

.EXAMPLE
    To download offline copy of OUI MAC-database
    Update-MACDatabase

.EXAMPLE
    To search offline MAC-database
    Get-MAC -OUI '08:EA:44'    

.EXAMPLE
    To open a graphical interface for searching MAC-addresses
    Get-MACGui
    
#>
function Test-MACOui {
    param([string]$InputString)

    # Accpets full MAC (6 bytes) or OUI (3 bytes)
    # Allows colon and hyphen as separator, or no separators

    if ($InputString -match '^([0-9A-Fa-f]{2}([-:]?)([0-9A-Fa-f]{2}\2){1,4}[0-9A-Fa-f]{2})$') {
        # Normalize: Remove all separators, uppercase
        $NormalizedInputString = (($InputString -replace '[-:]', '').ToUpper())[0..5] -join ''        
        return $NormalizedInputString
    }
    else {
        return $false
    }
}

function New-DirectoryIfNotExist {
    param([string]$Path)

    # Removes filename from path
    if ([System.IO.Path]::HasExtension($Path)) {
        $Path = [System.IO.Path]::GetDirectoryName($Path)
    }

    # Create directory if it doesn't exist
    if (!(Test-Path $Path)) {
        try {
            New-Item -ItemType Directory -Path $Path -ErrorAction Stop | Out-Null
            #Success
        }
        catch {
            #Fail
        }
    }
}

function Update-MACDatabase {
    Param([string]$MacDBFolder = "$(($profile | Split-Path))\Lookups")
    $wr = (Invoke-WebRequest -Uri https://standards-oui.ieee.org/)
    $lines = (($wr | Select-Object -ExpandProperty Content) -split "`n") | Select-Object -Skip 3

    $OuiMap = @{}
    $Block = @()
    $OuiFile = "$MacDBFolder\ouiMap.xml" -replace "\\\\", '\'
    Write-Host "Working vs. '$OuiFile"

    New-DirectoryIfNotExist -Path $OuiFile

    function ParseBlock {
        param([string[]]$block)

        if ($block.Count -lt 2) { return $null }
        
        $block = $block | ForEach-Object { $_.Trim() }

        # 1st line: hex
        if ($block[0] -match "^([0-9A-Fa-f\-]+)\s+\(hex\)\s+(.*)$") {
            $ouiHex = $matches[1].ToUpper()
            $org = $matches[2].Trim()
        } else {
            return $null
        }

        # 2nd line: base16
        if ($block[1] -match "^([0-9A-Fa-f]+)\s+\(base 16\)\s+(.*)$") {
            $ouiBase16 = $matches[1].ToUpper()
        } else {
            return $null
        }

        # Then address
        $addressLines = $block[2..($block.Count - 1)] | Where-Object { $_ -ne "" }
        $address = $addressLines -join ", "

        return [PSCustomObject]@{
            OuiHex      = $ouiHex
            OuiBase16   = $ouiBase16
            Org         = $org
            Address     = $address
        }
    }

    foreach ($line in $lines) {
        # A new record starts when we see a line with "(hex)" — save previous block
        if ($line -match "^\s*[0-9A-Fa-f\-]+\s+\(hex\)\s+") {
            if ($block.Count -gt 0) {
                # Process the previous block
                $parsed = ParseBlock -block $block
                if ($parsed) {
                    $ouiMap[$parsed.ouiBase16] = $parsed
                }
                $block = @()
            }
        }
        $block += $line
    }

    # Handle final block
    if ($block.Count -gt 0) {
        $parsed = ParseBlock -block $block
        if ($parsed) {
            $ouiMap[$parsed.ouiBase16] = $parsed
        }
    }
    Write-Host "Parsed entries: $($ouiMap.Count)"

    try {
        $ouiMap | Export-Clixml -Path "$OuiFile" -Force -ErrorAction Stop
        Write-Host "Export success" -ForegroundColor Green
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

function Search-OUIFile {
    Param([string]$lookupKey)
    if ($ouiMap.ContainsKey($lookupKey)) {
        $data = $ouiMap[$lookupKey]
        return [pscustomobject]$data

    } else {        
        Return $null
    }
}

function Get-MAC {
    Param(        
        [string]$OUI,
        [string]$MacDBFolder = "$(($profile | Split-Path))\Lookups"
    )

    $OuiFile = "$MacDBFolder\ouiMap.xml" -replace "\\\\", '\'
    Write-Host "Looking for '$OuiFile'."

    if (Test-Path $OuiFile) {
        Write-Host "Found '$OuiFile'."

        $NormalizedOUI = Test-MacOui $OUI

        if ($NormalizedOUI) {
            if(!$ouiMap) {
                try {                
                    $ouiMap = Import-Clixml -Path $OuiFile -ErrorAction Stop
                    $result = Search-OUIFile -lookupKey $NormalizedOUI
                    return $result
                }
                catch {
                    #
                }
            }
            else {
                $result = Search-OUIFile -lookupKey $NormalizedOUI
                return $result
            }
        }
        else {
            #
        }
    }
    else {
        Write-Host "Not able to find '$OuiFile'."
    }    
}
function Get-MACGui {
    param(
        [string]$MacDBFolder = "$(($profile | Split-Path))\Lookups"
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Load the OUI hashtable from CLIXML    
    $OuiFile = "$MacDBFolder\ouiMap.xml" -replace "\\\\", '\'

    if (-not (Test-Path $OuiFile)) {
        [System.Windows.Forms.MessageBox]::Show("Could not find OUI data at $OuiFile")
        return
    }

    $ouiMap = Import-Clixml $OuiFile
   
    # Create Form
    $form = New-Object Windows.Forms.Form
    $form.Text = "MAC Address Lookup"
    $form.Size = New-Object Drawing.Size(700, 200)
    $form.StartPosition = "CenterScreen"

    # Search Label
    $label = New-Object Windows.Forms.Label
    $label.Text = "Enter MAC or OUI:"
    $label.Location = New-Object Drawing.Point(10, 20)
    $label.AutoSize = $true
    $form.Controls.Add($label)

    # Textbox for MAC input
    $textBox = New-Object Windows.Forms.TextBox
    $textBox.Location = New-Object Drawing.Point(130, 16)
    $textBox.Width = 120
    $form.Controls.Add($textBox)

    # Result label
    $resultLabel = New-Object Windows.Forms.Label
    $resultLabel.Text = "Please start typing in input field above :)"
    $resultLabel.Location = New-Object Drawing.Point(10, 60)
    $resultLabel.Size = New-Object Drawing.Size(560, 80)
    $resultLabel.AutoSize = $false
    $resultLabel.TextAlign = "TopLeft"
    $form.Controls.Add($resultLabel)

    # Lookup logic triggered on each keypress
    $textBox.Add_TextChanged({        
        $Normalised = Test-MacOui -InputString ($textBox.Text)

        if ($Normalised) {
            $Result = Search-OUIFile -lookupKey $Normalised
            if ($Result) {
                $resultLabel.Text = "OuiHex: $($Result.OuiHex)`nOrganization: $($Result.Org)`nAddress: $($Result.Address)"
            }
            else {
                $resultLabel.Text = "Org/vendor not found"
            }
        }
        else {
            $resultLabel.Text = "Invalid/incompete input"
        }
    })

    # Run the form
    [void]$form.ShowDialog()
}
