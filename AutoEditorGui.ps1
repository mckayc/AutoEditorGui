# Auto-Editor GUI
# Requires auto-editor binary on PATH: https://github.com/wyattblue/auto-editor/releases/latest

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------------------------------------
# Version detection
# ---------------------------------------------------------------------------

try {
    $autoEditorVersion = & auto-editor --version 2>$null
    if (-not $autoEditorVersion) { $autoEditorVersion = "Not installed" }
} catch {
    $autoEditorVersion = "Not installed"
}

try {
    $latestRelease = Invoke-RestMethod `
        -Uri "https://api.github.com/repos/WyattBlue/auto-editor/releases/latest" `
        -Headers @{ "User-Agent" = "PowerShell Script" }
    $latestVersion = $latestRelease.tag_name
} catch {
    $latestVersion = $null
}

# ---------------------------------------------------------------------------
# Settings helpers
# ---------------------------------------------------------------------------

$defaultConfig = @{
    audioLevelValue  = 1.5
    beforeValue      = 0.05
    afterValue       = 0.05
    videoEditorValue = "resolve"
}

function Load-Settings {
    $path = Join-Path $PSScriptRoot "auto-editor-config.json"
    if (Test-Path $path) {
        $c = Get-Content $path | ConvertFrom-Json
        return @{
            audioLevelValue  = $c.audioLevelValue
            beforeValue      = $c.beforeValue
            afterValue       = $c.afterValue
            videoEditorValue = $c.videoEditorValue
        }
    }
    return $defaultConfig
}

function Save-Settings {
    param($audioLevel, $before, $after, $videoEditorType)
    $path = Join-Path $PSScriptRoot "auto-editor-config.json"
    @{
        audioLevelValue  = [double]$audioLevel
        beforeValue      = [double]$before
        afterValue       = [double]$after
        videoEditorValue = $videoEditorType
    } | ConvertTo-Json | Set-Content $path
}

$settings = Load-Settings

# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------

$clrBg       = [System.Drawing.Color]::FromArgb(28,  28,  28 )  # form background
$clrPanel    = [System.Drawing.Color]::FromArgb(44,  44,  46 )  # input / panel background
$clrText     = [System.Drawing.Color]::FromArgb(229, 229, 229)  # primary text
$clrBlue     = [System.Drawing.Color]::FromArgb(0,   120, 215)  # accent blue
$clrGray     = [System.Drawing.Color]::FromArgb(72,  72,  74 )  # disabled / secondary button
$clrLogBg    = [System.Drawing.Color]::FromArgb(15,  15,  15 )  # log background
$clrLogText  = [System.Drawing.Color]::FromArgb(204, 204, 204)  # log text
$clrDrop     = [System.Drawing.Color]::FromArgb(44,  44,  46 )  # drop zone idle
$clrDropHov  = [System.Drawing.Color]::FromArgb(20,  60,  35 )  # drop zone hover

# ---------------------------------------------------------------------------
# Form
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text            = "Auto-Editor  |  Installed: v$autoEditorVersion  |  Latest: $(if ($latestVersion) { $latestVersion } else { 'unknown' })"
$form.Size            = New-Object System.Drawing.Size(555, 590)
$form.MinimumSize     = New-Object System.Drawing.Size(555, 590)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $clrBg
$form.ForeColor       = $clrText
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 10)

# ---------------------------------------------------------------------------
# Helper: styled button
# ---------------------------------------------------------------------------

function New-Button {
    param([string]$text, [int]$x, [int]$y, [int]$w, [int]$h,
          [System.Drawing.Color]$bg)
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $text
    $b.Location  = New-Object System.Drawing.Point($x, $y)
    $b.Size      = New-Object System.Drawing.Size($w, $h)
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderSize = 0
    $b.BackColor = $bg
    $b.ForeColor = [System.Drawing.Color]::White
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $b
}

# ---------------------------------------------------------------------------
# Sliders + textboxes
# ---------------------------------------------------------------------------

function New-SliderRow {
    param([string]$label, [int]$y, [int]$sliderMax, [double]$initValue,
          [double]$textMax, [int]$tickFreq)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $label
    $lbl.Location  = New-Object System.Drawing.Point(20, ($y + 4))
    $lbl.Size      = New-Object System.Drawing.Size(135, 20)
    $lbl.ForeColor = $clrText
    $form.Controls.Add($lbl)

    $slider = New-Object System.Windows.Forms.TrackBar
    $slider.Minimum       = 0
    $slider.Maximum       = $sliderMax
    $slider.Value         = [int]($initValue * 100)
    $slider.TickFrequency = $tickFreq
    $slider.Location      = New-Object System.Drawing.Point(160, $y)
    $slider.Size          = New-Object System.Drawing.Size(280, 45)
    $slider.BackColor     = $clrBg
    $form.Controls.Add($slider)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Text      = $initValue.ToString("F2")
    $tb.Location  = New-Object System.Drawing.Point(450, ($y + 8))
    $tb.Size      = New-Object System.Drawing.Size(75, 24)
    $tb.BackColor = $clrPanel
    $tb.ForeColor = $clrText
    $tb.BorderStyle = "FixedSingle"
    $form.Controls.Add($tb)

    return @{ Slider = $slider; TextBox = $tb; TextMax = $textMax }
}

$audioRow  = New-SliderRow "Audio Threshold:" 18  500 $settings.audioLevelValue 5.0  10
$beforeRow = New-SliderRow "Margin Before:"   65  50  $settings.beforeValue      0.50 1
$afterRow  = New-SliderRow "Margin After:"    112 50  $settings.afterValue       0.50 1

$audioSlider  = $audioRow.Slider;  $audioTB  = $audioRow.TextBox
$beforeSlider = $beforeRow.Slider; $beforeTB = $beforeRow.TextBox
$afterSlider  = $afterRow.Slider;  $afterTB  = $afterRow.TextBox

# Wire slider <-> textbox (explicit per-control to avoid closure-in-loop capture bug)
$audioSlider.Add_ValueChanged({ $audioTB.Text = ($audioSlider.Value / 100).ToString("F2") })
$audioTB.Add_TextChanged({
    try {
        $v = [double]$audioTB.Text
        if ($v -ge 0 -and $v -le 5.0) { $audioSlider.Value = [int]($v * 100) }
        else { $audioTB.Text = ($audioSlider.Value / 100).ToString("F2") }
    } catch { $audioTB.Text = ($audioSlider.Value / 100).ToString("F2") }
})

$beforeSlider.Add_ValueChanged({ $beforeTB.Text = ($beforeSlider.Value / 100).ToString("F2") })
$beforeTB.Add_TextChanged({
    try {
        $v = [double]$beforeTB.Text
        if ($v -ge 0 -and $v -le 0.50) { $beforeSlider.Value = [int]($v * 100) }
        else { $beforeTB.Text = ($beforeSlider.Value / 100).ToString("F2") }
    } catch { $beforeTB.Text = ($beforeSlider.Value / 100).ToString("F2") }
})

$afterSlider.Add_ValueChanged({ $afterTB.Text = ($afterSlider.Value / 100).ToString("F2") })
$afterTB.Add_TextChanged({
    try {
        $v = [double]$afterTB.Text
        if ($v -ge 0 -and $v -le 0.50) { $afterSlider.Value = [int]($v * 100) }
        else { $afterTB.Text = ($afterSlider.Value / 100).ToString("F2") }
    } catch { $afterTB.Text = ($afterSlider.Value / 100).ToString("F2") }
})

# ---------------------------------------------------------------------------
# Tooltips
# ---------------------------------------------------------------------------

$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.SetToolTip($audioSlider,  "Volume level (%) that counts as 'active' audio. Higher = only louder sounds are kept.")
$tooltip.SetToolTip($beforeSlider, "Seconds of audio to keep before each active section (prevents hard cuts).")
$tooltip.SetToolTip($afterSlider,  "Seconds of audio to keep after each active section (prevents hard cuts).")

# ---------------------------------------------------------------------------
# Video editor dropdown + Save button
# ---------------------------------------------------------------------------

$videoEditorLabel = New-Object System.Windows.Forms.Label
$videoEditorLabel.Text     = "Video Editor:"
$videoEditorLabel.Location  = New-Object System.Drawing.Point(20, 164)
$videoEditorLabel.Size      = New-Object System.Drawing.Size(135, 20)
$videoEditorLabel.ForeColor = $clrText
$form.Controls.Add($videoEditorLabel)

$videoEditor = New-Object System.Windows.Forms.ComboBox
$videoEditor.Items.AddRange(@("resolve", "final-cut-pro", "shotcut"))
$videoEditor.SelectedItem  = $settings.videoEditorValue
$videoEditor.Location      = New-Object System.Drawing.Point(160, 160)
$videoEditor.Size          = New-Object System.Drawing.Size(235, 26)
$videoEditor.DropDownStyle = "DropDownList"
$videoEditor.BackColor     = $clrPanel
$videoEditor.ForeColor     = $clrText
$form.Controls.Add($videoEditor)
$tooltip.SetToolTip($videoEditor, "Output format - choose the video editor you will import the result into.")

$saveButton = New-Button "Save Settings" 405 158 125 30 $clrBlue
$form.Controls.Add($saveButton)

$saveButton.Add_Click({
    try {
        Save-Settings `
            -audioLevel    ($audioSlider.Value  / 100) `
            -before        ($beforeSlider.Value / 100) `
            -after         ($afterSlider.Value  / 100) `
            -videoEditorType $videoEditor.SelectedItem
        Log-Message "Settings saved."
    } catch {
        Show-Error "Failed to save settings: $_"
    }
})

# ---------------------------------------------------------------------------
# Separator
# ---------------------------------------------------------------------------

$separator = New-Object System.Windows.Forms.Panel
$separator.Location  = New-Object System.Drawing.Point(20, 198)
$separator.Size      = New-Object System.Drawing.Size(510, 1)
$separator.BackColor = [System.Drawing.Color]::FromArgb(63, 63, 70)
$form.Controls.Add($separator)

# ---------------------------------------------------------------------------
# Drop zone (full width)
# ---------------------------------------------------------------------------

$dropBox = New-Object System.Windows.Forms.Panel
$dropBox.BorderStyle = "FixedSingle"
$dropBox.AllowDrop   = $true
$dropBox.Size        = New-Object System.Drawing.Size(510, 50)
$dropBox.Location    = New-Object System.Drawing.Point(20, 207)
$dropBox.BackColor   = $clrDrop
$form.Controls.Add($dropBox)

$dropBoxLabel = New-Object System.Windows.Forms.Label
$dropBoxLabel.Text      = "Drop video files here"
$dropBoxLabel.TextAlign = "MiddleCenter"
$dropBoxLabel.Dock      = "Fill"
$dropBoxLabel.ForeColor = [System.Drawing.Color]::FromArgb(140, 140, 140)
$dropBoxLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
$dropBox.Controls.Add($dropBoxLabel)

# ---------------------------------------------------------------------------
# Browse / GitHub Releases / Clear Log buttons (evenly spaced)
# ---------------------------------------------------------------------------

$browseButton  = New-Button "Browse Files"    20  267 160 28 $clrBlue
$releasesButton = New-Button "GitHub Releases" 190 267 160 28 $clrBlue
$clearButton   = New-Button "Clear Log"       360 267 160 28 $clrGray
$clearButton.FlatAppearance.BorderSize  = 1
$clearButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(90, 90, 92)
$form.Controls.Add($browseButton)
$form.Controls.Add($releasesButton)
$form.Controls.Add($clearButton)

# ---------------------------------------------------------------------------
# Log box
# ---------------------------------------------------------------------------

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline   = $true
$logBox.ReadOnly    = $true
$logBox.ScrollBars  = "Vertical"
$logBox.Size        = New-Object System.Drawing.Size(510, 255)
$logBox.Location    = New-Object System.Drawing.Point(20, 303)
$logBox.Font        = New-Object System.Drawing.Font("Consolas", 9)
$logBox.BackColor   = $clrLogBg
$logBox.ForeColor   = $clrLogText
$logBox.BorderStyle = "FixedSingle"
$form.Controls.Add($logBox)

# ---------------------------------------------------------------------------
# Helper functions (reference UI controls - defined after controls exist)
# ---------------------------------------------------------------------------

function Log-Message {
    param([string]$message)
    $logBox.AppendText("$message`r`n")
    $logBox.ScrollToCaret()
}

function Show-Error {
    param([string]$message)
    [System.Windows.Forms.MessageBox]::Show(
        $message, "Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
}

function Invoke-ProcessFile {
    param([string]$file)
    try {
        $threshold = $audioSlider.Value  / 100
        $before    = $beforeSlider.Value / 100
        $after     = $afterSlider.Value  / 100
        $editor    = $videoEditor.SelectedItem

        Log-Message "--- $file"
        Log-Message "    threshold=${threshold}%  before=${before}s  after=${after}s  export=$editor"

        # Use & with separate arguments so PowerShell handles quoting correctly.
        # This avoids the path-with-spaces bug that occurs when building a command
        # string and passing it through Start-Process powershell -Command.
        $output = & auto-editor $file `
            --margin "${before}sec,${after}sec" `
            --export $editor `
            --edit "audio:threshold=${threshold}%" 2>&1

        $exitCode = $LASTEXITCODE

        foreach ($line in $output) {
            $text = "$line".Trim()
            if ($text) { Log-Message "    $text" }
        }

        if ($exitCode -ne 0) {
            Show-Error "auto-editor failed (exit code $exitCode). See log for details."
        } else {
            Log-Message "    Done."
        }
    } catch {
        Show-Error "Error processing file: $_"
    }
}

# ---------------------------------------------------------------------------
# Event handlers
# ---------------------------------------------------------------------------

$clearButton.Add_Click({ $logBox.Clear() })

$releasesButton.Add_Click({
    Start-Process "https://github.com/WyattBlue/auto-editor/releases"
})

$browseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Multiselect = $true
    $dialog.Filter      = "Video files|*.mp4;*.mov;*.avi;*.mkv;*.wmv;*.webm;*.flv|All files|*.*"
    $dialog.Title       = "Select video files to process"
    if ($dialog.ShowDialog() -eq "OK") {
        foreach ($file in $dialog.FileNames) {
            Invoke-ProcessFile $file
        }
    }
})

$dropBox.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [Windows.Forms.DragDropEffects]::Copy
        $dropBox.BackColor = $clrDropHov
    }
})

$dropBox.Add_DragLeave({ $dropBox.BackColor = $clrDrop })

$dropBox.Add_DragDrop({
    $dropBox.BackColor = $clrDrop
    foreach ($file in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) {
        Invoke-ProcessFile $file
    }
})

# ---------------------------------------------------------------------------
# Show
# ---------------------------------------------------------------------------

$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
