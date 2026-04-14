# This script simplifies running Auto-Editor with a modern GUI including sliders and textboxes for parameter adjustments
# If you do not already have Auto-Editor installed, make sure you have Python installed then run: pip install auto-editor
# Check the website for more info: https://auto-editor.com/

# Instructions for making this script act as an app (searchable in the Windows search bar) are included in the original script comments

## Useful Commands ##
# Check Version: auto-editor --version
# Install: pip install auto-editor==24.13.1
# Uninstall: pip uninstall auto-editor
# Upgrade: pip install auto-editor --upgrade

# Retrieve the installed version of auto-editor
try {
    $autoEditorVersion = & auto-editor --version
    if (-not $autoEditorVersion) {
        $autoEditorVersion = "Unknown (auto-editor not installed or inaccessible)"
    }
} catch {
    $autoEditorVersion = "Error retrieving version"
}

# Retrieve the most recent release version from GitHub
$latestReleaseUrl = "https://api.github.com/repos/WyattBlue/auto-editor/releases/latest"
$headers = @{ "User-Agent" = "PowerShell Script" }
try {
    $latestRelease = Invoke-RestMethod -Uri $latestReleaseUrl -Headers $headers -Method Get
    $latestVersion = $latestRelease.tag_name
    Write-Output "Latest version of Auto-Editor: $latestVersion"
} catch {
    Write-Output "Error retrieving the latest version from GitHub: $_"
}

# Bypass the execution policy for the current session
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Default values
$defaultConfig = @{
    audioLevelValue = 1.5
    beforeValue = 0.05
    afterValue = 0.05
    videoEditorValue = "resolve"
}

# Function to load settings from config file
function Load-Settings {
    $configPath = Join-Path $PSScriptRoot "auto-editor-config.json"
    if (Test-Path $configPath) {
        $config = Get-Content $configPath | ConvertFrom-Json
        return @{
            audioLevelValue = $config.audioLevelValue
            beforeValue = $config.beforeValue
            afterValue = $config.afterValue
            videoEditorValue = $config.videoEditorValue
        }
    }
    return $defaultConfig
}

# Function to save settings to config file
function Save-Settings {
    param (
        $audioLevel,
        $before,
        $after,
        $videoEditorType
    )
    
    $config = @{
        audioLevelValue = [double]$audioLevel
        beforeValue = [double]$before
        afterValue = [double]$after
        videoEditorValue = $videoEditorType
    }
    
    $configPath = Join-Path $PSScriptRoot "auto-editor-config.json"
    $config | ConvertTo-Json | Set-Content $configPath
}

# Load saved settings or use defaults
$settings = Load-Settings

# Initialize form with modern styling
$form = New-Object System.Windows.Forms.Form
$form.Text = "Auto-Editor - v$autoEditorVersion (Latest: $latestVersion)"
$form.Size = New-Object System.Drawing.Size(550, 550)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Create Update button
$updateButton = New-Object System.Windows.Forms.Button
$updateButton.Text = "Update"
$updateButton.Location = New-Object System.Drawing.Point(400, 20)
$updateButton.Size = New-Object System.Drawing.Size(120, 30)
$updateButton.FlatStyle = "Flat"
$updateButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$updateButton.ForeColor = [System.Drawing.Color]::White
if ($latestVersion -and $autoEditorVersion -ne $latestVersion) {
    $updateButton.Enabled = $true
} else {
    $updateButton.Enabled = $false
    $updateButton.Text = "Up-to-date"
    $updateButton.BackColor = [System.Drawing.Color]::Gray
}
$form.Controls.Add($updateButton)

$updateButton.Add_Click({
    try {
        Log-Message "Updating Auto-Editor to the latest version..."
        $command = "pip install auto-editor --upgrade"
        $updateProcess = Start-Process powershell -ArgumentList "-Command & { $command }" -PassThru -Wait -RedirectStandardError "error.log" -RedirectStandardOutput "output.log"

        if ($updateProcess.ExitCode -ne 0) {
            $errorLog = Get-Content "error.log" | Where-Object { $_ -notmatch "^\[notice\]" }
            Show-Error "Update failed with the following error(s):`r`n$($errorLog -join "`r`n")"
        } else {
            Log-Message "Auto-Editor successfully updated to the latest version."
            $updateButton.Text = "Up-to-date"
            $updateButton.Enabled = $false
            $updateButton.BackColor = [System.Drawing.Color]::Gray
        }
    } catch {
        Show-Error "An error occurred during the update: $_"
    }
})

# Create labels, sliders, and textboxes for audio threshold, margin before, and margin after
$audioThresholdLabel = New-Object System.Windows.Forms.Label
$audioThresholdLabel.Text = "Audio Threshold:"
$audioThresholdLabel.Location = New-Object System.Drawing.Point(20, 60)
$audioThresholdLabel.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($audioThresholdLabel)

$audioThresholdSlider = New-Object System.Windows.Forms.TrackBar
$audioThresholdSlider.Minimum = 0
$audioThresholdSlider.Maximum = 500  # 0 to 5 with 2 decimal places
$audioThresholdSlider.Value = [int]($settings.audioLevelValue * 100)
$audioThresholdSlider.TickFrequency = 10
$audioThresholdSlider.Location = New-Object System.Drawing.Point(150, 60)
$audioThresholdSlider.Size = New-Object System.Drawing.Size(250, 45)
$form.Controls.Add($audioThresholdSlider)

$audioThresholdTextBox = New-Object System.Windows.Forms.TextBox
$audioThresholdTextBox.Text = $settings.audioLevelValue.ToString("F2")
$audioThresholdTextBox.Location = New-Object System.Drawing.Point(410, 60)
$audioThresholdTextBox.Size = New-Object System.Drawing.Size(60, 20)
$audioThresholdTextBox.Add_TextChanged({
    try {
        $value = [double]$audioThresholdTextBox.Text
        if ($value -ge 0 -and $value -le 5) {
            $audioThresholdSlider.Value = [int]($value * 100)
        } else {
            $audioThresholdTextBox.Text = ($audioThresholdSlider.Value / 100).ToString("F2")
        }
    } catch {
        $audioThresholdTextBox.Text = ($audioThresholdSlider.Value / 100).ToString("F2")
    }
})
$form.Controls.Add($audioThresholdTextBox)

$audioThresholdSlider.Add_ValueChanged({
    $value = $audioThresholdSlider.Value / 100
    $audioThresholdTextBox.Text = $value.ToString("F2")
})

$marginBeforeLabel = New-Object System.Windows.Forms.Label
$marginBeforeLabel.Text = "Margin Before:"
$marginBeforeLabel.Location = New-Object System.Drawing.Point(20, 100)
$marginBeforeLabel.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($marginBeforeLabel)

$marginBeforeSlider = New-Object System.Windows.Forms.TrackBar
$marginBeforeSlider.Minimum = 0
$marginBeforeSlider.Maximum = 50  # 0.00 to 0.50 with 2 decimal places
$marginBeforeSlider.Value = [int]($settings.beforeValue * 100)
$marginBeforeSlider.TickFrequency = 1
$marginBeforeSlider.Location = New-Object System.Drawing.Point(150, 100)
$marginBeforeSlider.Size = New-Object System.Drawing.Size(250, 45)
$form.Controls.Add($marginBeforeSlider)

$marginBeforeTextBox = New-Object System.Windows.Forms.TextBox
$marginBeforeTextBox.Text = $settings.beforeValue.ToString("F2")
$marginBeforeTextBox.Location = New-Object System.Drawing.Point(410, 100)
$marginBeforeTextBox.Size = New-Object System.Drawing.Size(60, 20)
$marginBeforeTextBox.Add_TextChanged({
    try {
        $value = [double]$marginBeforeTextBox.Text
        if ($value -ge 0 -and $value -le 0.50) {
            $marginBeforeSlider.Value = [int]($value * 100)
        } else {
            $marginBeforeTextBox.Text = ($marginBeforeSlider.Value / 100).ToString("F2")
        }
    } catch {
        $marginBeforeTextBox.Text = ($marginBeforeSlider.Value / 100).ToString("F2")
    }
})
$form.Controls.Add($marginBeforeTextBox)

$marginBeforeSlider.Add_ValueChanged({
    $value = $marginBeforeSlider.Value / 100
    $marginBeforeTextBox.Text = $value.ToString("F2")
})

$marginAfterLabel = New-Object System.Windows.Forms.Label
$marginAfterLabel.Text = "Margin After:"
$marginAfterLabel.Location = New-Object System.Drawing.Point(20, 140)
$marginAfterLabel.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($marginAfterLabel)

$marginAfterSlider = New-Object System.Windows.Forms.TrackBar
$marginAfterSlider.Minimum = 0
$marginAfterSlider.Maximum = 50  # 0.00 to 0.50 with 2 decimal places
$marginAfterSlider.Value = [int]($settings.afterValue * 100)
$marginAfterSlider.TickFrequency = 1
$marginAfterSlider.Location = New-Object System.Drawing.Point(150, 140)
$marginAfterSlider.Size = New-Object System.Drawing.Size(250, 45)
$form.Controls.Add($marginAfterSlider)

$marginAfterTextBox = New-Object System.Windows.Forms.TextBox
$marginAfterTextBox.Text = $settings.afterValue.ToString("F2")
$marginAfterTextBox.Location = New-Object System.Drawing.Point(410, 140)
$marginAfterTextBox.Size = New-Object System.Drawing.Size(60, 20)
$marginAfterTextBox.Add_TextChanged({
    try {
        $value = [double]$marginAfterTextBox.Text
        if ($value -ge 0 -and $value -le 0.50) {
            $marginAfterSlider.Value = [int]($value * 100)
        } else {
            $marginAfterTextBox.Text = ($marginAfterSlider.Value / 100).ToString("F2")
        }
    } catch {
        $marginAfterTextBox.Text = ($marginAfterSlider.Value / 100).ToString("F2")
    }
})
$form.Controls.Add($marginAfterTextBox)

$marginAfterSlider.Add_ValueChanged({
    $value = $marginAfterSlider.Value / 100
    $marginAfterTextBox.Text = $value.ToString("F2")
})

# Create a dropdown for video editor
$videoEditorLabel = New-Object System.Windows.Forms.Label
$videoEditorLabel.Text = "Video Editor:"
$videoEditorLabel.Location = New-Object System.Drawing.Point(20, 180)
$videoEditorLabel.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($videoEditorLabel)

$videoEditor = New-Object System.Windows.Forms.ComboBox
$videoEditor.Items.AddRange(@("resolve", "final-cut-pro", "shotcut"))
$videoEditor.SelectedItem = $settings.videoEditorValue
$videoEditor.Location = New-Object System.Drawing.Point(150, 180)
$videoEditor.Size = New-Object System.Drawing.Size(250, 20)
$videoEditor.DropDownStyle = "DropDownList"
$form.Controls.Add($videoEditor)

# Create Save Settings button
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save Settings"
$saveButton.Location = New-Object System.Drawing.Point(400, 180)
$saveButton.Size = New-Object System.Drawing.Size(120, 30)
$saveButton.FlatStyle = "Flat"
$saveButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$saveButton.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($saveButton)

$saveButton.Add_Click({
    try {
        $audioLevel = $audioThresholdSlider.Value / 100
        $before = $marginBeforeSlider.Value / 100
        $after = $marginAfterSlider.Value / 100
        Save-Settings -audioLevel $audioLevel -before $before -after $after -videoEditorType $videoEditor.SelectedItem
        Log-Message "Settings saved successfully!"
    }
    catch {
        Show-Error "Failed to save settings: $_"
    }
})

# Create Pin to Start Menu button
$pinButton = New-Object System.Windows.Forms.Button
$pinButton.Text = "Pin to Start"
$pinButton.Location = New-Object System.Drawing.Point(400, 220)
$pinButton.Size = New-Object System.Drawing.Size(120, 30)
$pinButton.FlatStyle = "Flat"
$pinButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$pinButton.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($pinButton)

$pinButton.Add_Click({
    try {
        $scriptDir   = $PSScriptRoot
        $scriptPath  = Join-Path $scriptDir "AutoEditorGui.ps1"
        $iconPath    = Join-Path $scriptDir "AutoEditorScript.ico"
        $shortcutPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Auto-Editor.lnk"

        $shell    = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath       = "powershell.exe"
        $shortcut.Arguments        = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
        $shortcut.WorkingDirectory = $scriptDir
        $shortcut.IconLocation     = $iconPath
        $shortcut.Save()

        Log-Message "Shortcut created. Search 'Auto-Editor' in the Start bar."
    } catch {
        Show-Error "Failed to create shortcut: $_"
    }
})

# Create a label and box for dragging files
$dropLabel = New-Object System.Windows.Forms.Label
$dropLabel.Text = "Drag files here:"
$dropLabel.Location = New-Object System.Drawing.Point(20, 265)
$dropLabel.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($dropLabel)

$dropBox = New-Object System.Windows.Forms.Panel
$dropBox.BorderStyle = 'FixedSingle'
$dropBox.AllowDrop = $true
$dropBox.Size = New-Object System.Drawing.Size(350, 50)
$dropBox.Location = New-Object System.Drawing.Point(150, 265)
$dropBox.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$form.Controls.Add($dropBox)

$dropBoxLabel = New-Object System.Windows.Forms.Label
$dropBoxLabel.Text = "Drop files here"
$dropBoxLabel.TextAlign = 'MiddleCenter'
$dropBoxLabel.Dock = 'Fill'
$dropBoxLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
$dropBox.Controls.Add($dropBoxLabel)

# Create a textbox to show the log with a monospaced font
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ReadOnly = $true
$logBox.ScrollBars = 'Vertical'
$logBox.Size = New-Object System.Drawing.Size(500, 185)
$logBox.Location = New-Object System.Drawing.Point(20, 335)
$logBox.Font = New-Object System.Drawing.Font('Consolas', 10)
$logBox.BackColor = [System.Drawing.Color]::White
$logBox.BorderStyle = 'FixedSingle'
$form.Controls.Add($logBox)

# Function to log actions in the logBox
function Log-Message {
    param ($message)
    $logBox.AppendText("$message`r`n")
    $logBox.ScrollToCaret()
}

# Function to show an error message box
function Show-Error {
    param ($message)
    [System.Windows.Forms.MessageBox]::Show($message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Drawing.MessageBoxIcon]::Error)
}

# Handle file drag and drop
$dropBox.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [Windows.Forms.DragDropEffects]::Copy
        $dropBox.BackColor = [System.Drawing.Color]::FromArgb(200, 230, 201)
    }
})

$dropBox.Add_DragLeave({
    $dropBox.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
})

$dropBox.Add_DragDrop({
    $dropBox.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    foreach ($file in $files) {
        try {
            $filePath = "`"$file`""
            $marginBeforeValue = $marginBeforeSlider.Value / 100
            $marginAfterValue = $marginAfterSlider.Value / 100
            $audioThresholdValue = $audioThresholdSlider.Value / 100
            $videoEditorValue = $videoEditor.SelectedItem
            
            Log-Message "Processing file: $filePath"
            Log-Message "Audio Threshold: $audioThresholdValue"
            Log-Message "Margin Before: $marginBeforeValue"
            Log-Message "Margin After: $marginAfterValue"
            Log-Message "Video Editor: $videoEditorValue"
            
            $command = "auto-editor '$filePath' --margin $marginBeforeValue's',$marginAfterValue'sec' --export $videoEditorValue --edit audio:threshold=$audioThresholdValue% -sn"
            Log-Message "Running command: $command"
            
            $process = Start-Process powershell -ArgumentList "-Command & { $command }" -PassThru -Wait -RedirectStandardError "error.log" -RedirectStandardOutput "output.log"

            $errorLog = Get-Content "error.log"
            if ($errorLog) {
                Show-Error "The command failed with the following error(s):`r`n$errorLog"
            }
        } catch {
            Show-Error "An error occurred: $_"
        }
    }
})

# Show the form
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()