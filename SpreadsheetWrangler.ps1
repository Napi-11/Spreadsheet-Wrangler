# SpreadsheetWrangler.ps1
# GUI application for spreadsheet operations and folder backups

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Spreadsheet Wrangler"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Create a table layout panel for the main layout
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = "Fill"
$mainLayout.RowCount = 1
$mainLayout.ColumnCount = 2
$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 70)))
$form.Controls.Add($mainLayout)

#region Left Panel
$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Dock = "Fill"
$leftPanel.Padding = New-Object System.Windows.Forms.Padding(10)
$mainLayout.Controls.Add($leftPanel, 0, 0)

# Create a table layout for the left panel
$leftLayout = New-Object System.Windows.Forms.TableLayoutPanel
$leftLayout.Dock = "Fill"
$leftLayout.RowCount = 3
$leftLayout.ColumnCount = 1
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
$leftPanel.Controls.Add($leftLayout)

# Backup Locations Section
$backupPanel = New-Object System.Windows.Forms.GroupBox
$backupPanel.Text = "Backup Locations"
$backupPanel.Dock = "Fill"
$leftLayout.Controls.Add($backupPanel, 0, 0)

$backupLayout = New-Object System.Windows.Forms.TableLayoutPanel
$backupLayout.Dock = "Fill"
$backupLayout.RowCount = 2
$backupLayout.ColumnCount = 1
$backupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 90)))
$backupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
$backupPanel.Controls.Add($backupLayout)

# List of backup locations
$backupLocations = New-Object System.Windows.Forms.ListView
$backupLocations.View = "Details"
$backupLocations.FullRowSelect = $true
$backupLocations.Columns.Add("Folder Path", -2)
$backupLocations.Dock = "Fill"
$backupLayout.Controls.Add($backupLocations, 0, 0)

# Add button for backup locations
$addBackupBtn = New-Object System.Windows.Forms.Button
$addBackupBtn.Text = "+"
$addBackupBtn.Dock = "Right"
$addBackupBtn.Width = 40
$addBackupBtn.FlatStyle = "Flat"
$backupLayout.Controls.Add($addBackupBtn, 0, 1)

# Spreadsheet Folders Section
$spreadsheetPanel = New-Object System.Windows.Forms.GroupBox
$spreadsheetPanel.Text = "Spreadsheet Folder Locations"
$spreadsheetPanel.Dock = "Fill"
$leftLayout.Controls.Add($spreadsheetPanel, 0, 1)

$spreadsheetLayout = New-Object System.Windows.Forms.TableLayoutPanel
$spreadsheetLayout.Dock = "Fill"
$spreadsheetLayout.RowCount = 2
$spreadsheetLayout.ColumnCount = 1
$spreadsheetLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 90)))
$spreadsheetLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
$spreadsheetPanel.Controls.Add($spreadsheetLayout)

# List of spreadsheet folder locations
$spreadsheetLocations = New-Object System.Windows.Forms.ListView
$spreadsheetLocations.View = "Details"
$spreadsheetLocations.FullRowSelect = $true
$spreadsheetLocations.Columns.Add("Folder Path", -2)
$spreadsheetLocations.Dock = "Fill"
$spreadsheetLayout.Controls.Add($spreadsheetLocations, 0, 0)

# Add button for spreadsheet locations
$addSpreadsheetBtn = New-Object System.Windows.Forms.Button
$addSpreadsheetBtn.Text = "+"
$addSpreadsheetBtn.Dock = "Right"
$addSpreadsheetBtn.Width = 40
$addSpreadsheetBtn.FlatStyle = "Flat"
$spreadsheetLayout.Controls.Add($addSpreadsheetBtn, 0, 1)

# Run Button
$runBtn = New-Object System.Windows.Forms.Button
$runBtn.Text = "Run"
$runBtn.Dock = "Fill"
$runBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$runBtn.ForeColor = [System.Drawing.Color]::White
$runBtn.FlatStyle = "Flat"
$runBtn.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$leftLayout.Controls.Add($runBtn, 0, 2)
#endregion

#region Right Panel
$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.Dock = "Fill"
$rightPanel.Padding = New-Object System.Windows.Forms.Padding(10)
$mainLayout.Controls.Add($rightPanel, 1, 0)

# Create a table layout for the right panel
$rightLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rightLayout.Dock = "Fill"
$rightLayout.RowCount = 3
$rightLayout.ColumnCount = 1
$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 70)))
$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
$rightPanel.Controls.Add($rightLayout)

# Options Panel
$optionsPanel = New-Object System.Windows.Forms.GroupBox
$optionsPanel.Text = "Options"
$optionsPanel.Dock = "Fill"
$rightLayout.Controls.Add($optionsPanel, 0, 0)

$optionsLayout = New-Object System.Windows.Forms.FlowLayoutPanel
$optionsLayout.Dock = "Fill"
$optionsLayout.FlowDirection = "LeftToRight"
$optionsLayout.WrapContents = $true
$optionsPanel.Controls.Add($optionsLayout)

# Create 9 checkboxes for options
for ($i = 1; $i -le 9; $i++) {
    $checkbox = New-Object System.Windows.Forms.CheckBox
    $checkbox.Text = "Option $i"
    $checkbox.Width = 100
    $checkbox.Height = 30
    $checkbox.Margin = New-Object System.Windows.Forms.Padding(5)
    $optionsLayout.Controls.Add($checkbox)
}

# Output Panel
$outputPanel = New-Object System.Windows.Forms.GroupBox
$outputPanel.Text = "Terminal Output"
$outputPanel.Dock = "Fill"
$rightLayout.Controls.Add($outputPanel, 0, 1)

# Terminal output textbox
$outputTextbox = New-Object System.Windows.Forms.RichTextBox
$outputTextbox.Dock = "Fill"
$outputTextbox.ReadOnly = $true
$outputTextbox.BackColor = [System.Drawing.Color]::Black
$outputTextbox.ForeColor = [System.Drawing.Color]::LightGreen
$outputTextbox.Font = New-Object System.Drawing.Font("Consolas", 10)
$outputPanel.Controls.Add($outputTextbox)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock = "Fill"
$progressBar.Style = "Continuous"
$progressBar.Value = 0
$rightLayout.Controls.Add($progressBar, 0, 2)
#endregion

# Initialize the form with some sample data for visualization
$outputTextbox.AppendText("Application initialized and ready to run.`r`n")
$outputTextbox.AppendText("Please add backup and spreadsheet folder locations.`r`n")

# Show the form
$form.ShowDialog()
