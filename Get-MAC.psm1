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

.EXAMPLE
    Get vendor for all local networkadapters
    Get-NetAdapter | select -ExpandProperty MacAddress | Get-MAC


.EXAMPLE
    Get vendor for all MAC-addresses in ARP
    Get-NetNeighbor | select -ExpandProperty LinkLayerAddress -Unique | ? {$_} | Get-MAC

.EXAMPLE
    Get vendor for all MAC in array
    $MAC = @('00.10.20','00-11-21','00:12:22','001323','00-14-24','00:15:25')
    $MAC | Get-Mac
#>

function Write-Console {
    param(
        [ValidateSet(0, 1, 2, 3, 4)]
        [int]$Level,

        [Parameter(Mandatory=$true)]
        [string]$Message
    )
    $Message = $Message.Replace("`r",'').Replace("`n",' ')
    switch ($Level) {
        0 { $Status = 'Info'        ;$FGColor = 'White'   }
        1 { $Status = 'Success'     ;$FGColor = 'Green'   }
        2 { $Status = 'Warning'     ;$FGColor = 'Yellow'  }
        3 { $Status = 'Error'       ;$FGColor = 'Red'     }
        4 { $Status = 'Highlight'   ;$FGColor = 'Gray'    }
        Default { $Status = ''      ;$FGColor = 'Black'   }
    }
    if ($VerboseLogging) {
        Write-Host "$((Get-Date).ToString()) " -ForegroundColor 'DarkGray' -NoNewline
        Write-Host "$Status" -ForegroundColor $FGColor -NoNewline

        if ($level -eq 4) {
            Write-Host ("`t " + $Message) -ForegroundColor 'Cyan'
        }
        else {
            Write-Host ("`t " + $Message) -ForegroundColor 'White'
        }
    }
    if ($Level -eq 3) {
        $LogErrors += $Message
    }
}

function Test-MACOui {
    param([string]$InputString)

    # Accpets full MAC (6 bytes) or OUI (3 bytes)
    # Allows colon and hyphen as separator, or no separators

    if ($InputString -match '^([0-9A-Fa-f]{2}([-:.]?)([0-9A-Fa-f]{2}\2){1,4}[0-9A-Fa-f]{2})$') {
        # Normalize: Remove all separators, uppercase
        $NormalizedInputString = (($InputString -replace '[-:.]', '').ToUpper())[0..5] -join ''        
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
            Write-Console -Level 1 -Message "New-Item - Directory created: '$Path'"
        }
        catch {
            Write-Console -Level 3 -Message "New-Item - Failed to create new directory '$Path'. Error: $($_.Exception.Message)"
        }
    }
}

function Update-MACDatabase {
    Param(
        [string]$MacDBFolder = "$(($profile | Split-Path))\Lookups",
        [switch]$VerboseLogging = $false
    )
    try {
        $wr = (Invoke-WebRequest -Uri https://standards-oui.ieee.org/ -ErrorAction Stop)
        Write-Console -Level 1 -Message "Invoke-WebRequest - Downloaded OUI from IEEE. Statuscode: $($wr.StatusCode)"        
    }
    catch {
        Write-Console -Level 3 -Message "Invoke-WebRequest - Failed to download OUI from IEEE. Statuscode: $($wr.StatusCode). Error: $($_.Exception.Message)"
        Break
    }

    $lines = (($wr | Select-Object -ExpandProperty Content) -split "`n") | Select-Object -Skip 3
    Write-Console -Level 0 -Message "The downloaded OUI-file contains $($lines.count) lines"

    $OuiMap = @{}
    $Block = @()
    $OuiFile = "$MacDBFolder\ouiMap.xml" -replace "\\\\", '\'
    Write-Console -Level 0 -Message "Local DB - Using following path '$OuiFile"

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
    Write-Console -Level 0 -Message "Local DB - Parsed entries from OUI-file: $($ouiMap.Count)"

    try {
        $ouiMap | Export-Clixml -Path "$OuiFile" -Force -ErrorAction Stop
        Write-Console -Level 1 -Message "Local DB - Saved hashtable for offline lookups (Export-CliXml)"
    }
    catch {
        Write-Console -Level 3 -Message "Local DB - Failed to save hashtable. Error: $($_.Exception.Message)"
    }
}

function Search-OUIFile {
    Param([string]$lookupKey)
    if ($ouiMap.ContainsKey($lookupKey)) {
        $data = $ouiMap[$lookupKey]
        Write-Console -Level 1 -Message "Local DB - Search-OUIFile: Found results for OUI '$lookupKey' in file"
        return [pscustomobject]$data

    } else {       
        Write-Console -Level 0 -Message "Local DB - Search-OUIFile: No results for OUI '$lookupKey' in file"
        Return $null
    }
}

function Get-MAC {
    Param(        
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)][string]$OUI,
        [string]$MacDBFolder = "$(($profile | Split-Path))\Lookups",        
        [switch]$VerboseLogging = $false
    )

    Begin {
        $OuiFile = "$MacDBFolder\ouiMap.xml" -replace "\\\\", '\'
        Write-Console -Level 0 -Message "Local DB - Get-MAC: Using following path '$OuiFile"

        if (Test-Path $OuiFile) {
            Write-Console -Level 1 -Message "Local DB - Get-MAC: File found"            
            if(!$ouiMap) {
                try {                
                    $ouiMap = Import-Clixml -Path $OuiFile -ErrorAction Stop
                    Write-Console -Level 1 -Message "Local DB - Import-Clixml: File imported as hashtable"                        
                }
                catch {
                    Write-Console -Level 3 -Message "Local DB - Import-Clixml: Failed to imported file. Error: $($_.Exception.Message)"
                    Break
                }
            }
            else {
                # OK :)
            }            
        }
        else {        
            Write-Console -Level 2 -Message "Local DB - Get-MAC: File not found. Please run 'Update-MACDatabase'"
            Write-Warning "Local DB - Get-MAC: File not found. Please run 'Update-MACDatabase'"
        }

        $NormalizedOUIs = @()

    }
    Process {        
        $OUI | ForEach-Object {
            $NormalizedOUIs += Test-MacOui $OUI
        }
    }
    End {
        $Results = @()
        Foreach ($NOUI in $NormalizedOUIs) {
            $Results += Search-OUIFile -lookupKey $NOUI
        }
        return $Results
    }
     
}

function Get-MACGui {
    param(
        [string]$MacDBFolder = "$(($profile | Split-Path))\Lookups",
        [switch]$VerboseLogging = $false
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
    $form.StartPosition   = 'CenterScreen'
    $form.Size = New-Object Drawing.Size(400, 200)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'None'
    $form.TopMost         = $true    
    $form.BackColor       = [System.Drawing.Color]::Fuchsia
    $form.TransparencyKey = [System.Drawing.Color]::Fuchsia
    $form.Opacity         = 0.88

    $panel           = New-Object System.Windows.Forms.Panel
    $panel.Dock      = 'Fill'
    $panel.BackColor = [System.Drawing.Color]::FromArgb(32,32,32)  # dark
    $panel.Padding   = New-Object System.Windows.Forms.Padding(16) # <-- fixed
    $form.Controls.Add($panel)

    # Title
    $title = New-Object System.Windows.Forms.Label
    $title.Text      = "Get-Mac"
    $title.AutoSize  = $true
    $title.Font      = New-Object System.Drawing.Font('Segoe UI', 12)
    $title.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
    $title.Location  = New-Object System.Drawing.Point(16, 16)
    $panel.Controls.Add($title)

    # Search Label
    $label = New-Object Windows.Forms.Label
    $label.Text = "Search MAC or OUI:"
    $label.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
    $label.Location = New-Object Drawing.Point(16, 60)
    $label.AutoSize = $true
    $panel.Controls.Add($label)

    # Textbox for MAC input
    $textBox = New-Object Windows.Forms.TextBox
    $textBox.Location = New-Object Drawing.Point(150, 60)
    $textBox.Width = 120
    $panel.Controls.Add($textBox)

    # Result label
    $resultLabel = New-Object Windows.Forms.Label
    $resultLabel.Text = "Please start typing in input field above :)"
    $resultLabel.Location = New-Object Drawing.Point(16, 100)
    $resultLabel.Size = New-Object Drawing.Size(560, 80)
    $resultLabel.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
    $resultLabel.AutoSize = $false
    $resultLabel.TextAlign = "TopLeft"
    $panel.Controls.Add($resultLabel)

    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "✕"
    $closeBtn.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $closeBtn.Size = New-Object System.Drawing.Size(34, 30)
    $closeBtn.FlatStyle = 'Flat'
    $closeBtn.FlatAppearance.BorderSize = 0
    $closeBtn.BackColor = [System.Drawing.Color]::FromArgb(55,55,55)
    $closeBtn.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
    $closeBtn.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 50), 16)
    $closeBtn.Anchor = 'Top,Right'
    $closeBtn.Add_MouseEnter({ $closeBtn.BackColor = [System.Drawing.Color]::FromArgb(75,75,75) })
    $closeBtn.Add_MouseLeave({ $closeBtn.BackColor = [System.Drawing.Color]::FromArgb(55,55,55) })
    $closeBtn.Add_Click({ $form.Close() })
    $panel.Controls.Add($closeBtn)

    # Lookup logic triggered on each keypress
    $textBox.Add_TextChanged({        
        $Normalised = Test-MacOui -InputString ($textBox.Text -replace "(\.|-|:)\d$" -replace "(\.|-|:)$")

        if ($Normalised) {
            $Result = Search-OUIFile -lookupKey $Normalised
            if ($Result) {
                
                $Width = ($Result.Address).length * 8
                if ($Width -gt 400) {
                    $form.Size = New-Object Drawing.Size($Width, 200)
                    $resultLabel.Size = New-Object Drawing.Size($Width, 80)
                }
                
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

    # Make form draggable
    $mouseDown = $false
    $offset    = [System.Drawing.Point]::Empty
    $startDrag = {
        $script:mouseDown = $true
        $script:offset = [System.Drawing.Point]::Subtract([System.Windows.Forms.Cursor]::Position, $form.Location)
    }
    $doDrag = {
        if ($script:mouseDown) {
            $form.Location = [System.Drawing.Point]::Subtract([System.Windows.Forms.Cursor]::Position, $script:offset)
        }
    }
    $stopDrag = { $script:mouseDown = $false }
    $form.Add_MouseDown($startDrag);  $form.Add_MouseMove($doDrag);  $form.Add_MouseUp($stopDrag)
    $panel.Add_MouseDown($startDrag); $panel.Add_MouseMove($doDrag); $panel.Add_MouseUp($stopDrag)

    # Run the form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [void]$form.ShowDialog()
}

Export-ModuleMember -Function 'Update-MACDatabase', 'Get-MAC', 'Get-MACGui'