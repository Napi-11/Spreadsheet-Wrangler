# SpreadsheetWrangler.ps1
# GUI application for spreadsheet operations and folder backups

#region Helper Functions

# Function to create a timestamp string for folder naming
function Get-TimeStampString {
    return Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
}

# Function to extract number from filename
function Get-FileNumber {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FileName
    )
    
    # Look for patterns like "1", "(1)", "_1", etc.
    if ($FileName -match "[\(\[_\s-](\d+)[\)\]\s]*$") {
        return $matches[1]
    }
    elseif ($FileName -match "(\d+)[\)\]\s]*$") {
        return $matches[1]
    }
    
    # If no number found, return null
    return $null
}

# Function to log messages to the output textbox
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$Color = "LightGreen"
    )
    
    # Ensure we're on the UI thread
    if ($outputTextbox.InvokeRequired) {
        $outputTextbox.Invoke([Action[string, string]]{ param($msg, $clr) 
            $outputTextbox.SelectionColor = [System.Drawing.Color]::$clr
            $outputTextbox.AppendText("$msg`r`n")
            $outputTextbox.ScrollToCaret()
        }, $Message, $Color)
    } else {
        $outputTextbox.SelectionColor = [System.Drawing.Color]::$Color
        $outputTextbox.AppendText("$Message`r`n")
        $outputTextbox.ScrollToCaret()
    }
}

# Function to update the progress bar
function Update-ProgressBar {
    param (
        [Parameter(Mandatory=$true)]
        [int]$Value
    )
    
    # Ensure we're on the UI thread
    if ($progressBar.InvokeRequired) {
        $progressBar.Invoke([Action[int]]{ param($val) 
            $progressBar.Value = $val
        }, $Value)
    } else {
        $progressBar.Value = $Value
    }
}

# Function to create backup of a folder
function Backup-Folder {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SourcePath
    )
    
    try {
        # Create backup directory if it doesn't exist
        $backupRootDir = Join-Path -Path $PSScriptRoot -ChildPath ".backup"
        if (-not (Test-Path -Path $backupRootDir)) {
            New-Item -Path $backupRootDir -ItemType Directory -Force | Out-Null
            Write-Log "Created backup directory: $backupRootDir"
        }
        
        # Get folder name from source path
        $folderName = Split-Path -Path $SourcePath -Leaf
        
        # Create timestamped backup folder
        $timestamp = Get-TimeStampString
        $backupFolderName = "$folderName-$timestamp"
        $backupPath = Join-Path -Path $backupRootDir -ChildPath $backupFolderName
        
        # Create the backup folder
        New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
        Write-Log "Created backup folder: $backupPath"
        
        # Copy all items from source to backup
        Write-Log "Starting backup of $SourcePath..."
        Copy-Item -Path "$SourcePath\*" -Destination $backupPath -Recurse -Force
        Write-Log "Backup completed successfully!" "Yellow"
        
        return $true
    }
    catch {
        Write-Log "Error during backup: $_" "Red"
        return $false
    }
}

# Function to perform backup of all selected folders
function Start-BackupProcess {
    if ($backupLocations.Items.Count -eq 0) {
        Write-Log "No backup locations selected." "Yellow"
        return
    }
    
    Write-Log "Starting backup process..." "Cyan"
    Update-ProgressBar 0
    
    $totalFolders = $backupLocations.Items.Count
    $completedFolders = 0
    
    foreach ($item in $backupLocations.Items) {
        $folderPath = $item.Text
        Write-Log "Processing backup for: $folderPath" "White"
        
        $success = Backup-Folder -SourcePath $folderPath
        $completedFolders++
        
        # Update progress
        $progressPercentage = [int](($completedFolders / $totalFolders) * 100)
        Update-ProgressBar $progressPercentage
    }
    
    Write-Log "Backup process completed." "Cyan"
    Update-ProgressBar 100
}

# Function to combine spreadsheets
function Combine-Spreadsheets {
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.ArrayList]$FolderPaths,
        
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath,
        
        [Parameter(Mandatory=$false)]
        [string]$FileExtension = "*.xlsx",
        
        [Parameter(Mandatory=$false)]
        [bool]$ExcludeHeaders = $false,
        
        [Parameter(Mandatory=$false)]
        [bool]$DuplicateQuantityTwoRows = $false,
        
        [Parameter(Mandatory=$false)]
        [bool]$NormalizeQuantities = $false
    )
    
    try {
        # Load Excel COM object
        Write-Log "Initializing Excel..." "White"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # Create a dictionary to store spreadsheets by number
        $spreadsheetGroups = @{}
        
        # First, scan all folders and group spreadsheets by their number
        Write-Log "Scanning folders for spreadsheets..." "White"
        
        foreach ($folderPath in $FolderPaths) {
            Write-Log "Scanning folder: $folderPath" "White"
            
            # Handle different file formats if All Formats option is selected
            if ($FileExtension -eq "*.*") {
                $files = @()
                $files += Get-ChildItem -Path $folderPath -Filter "*.xlsx" -File
                $files += Get-ChildItem -Path $folderPath -Filter "*.xls" -File
                $files += Get-ChildItem -Path $folderPath -Filter "*.csv" -File
            } else {
                $files = Get-ChildItem -Path $folderPath -Filter $FileExtension -File
            }
            
            foreach ($file in $files) {
                $fileNumber = Get-FileNumber -FileName $file.Name
                
                if ($fileNumber) {
                    if (-not $spreadsheetGroups.ContainsKey($fileNumber)) {
                        $spreadsheetGroups[$fileNumber] = New-Object System.Collections.ArrayList
                    }
                    
                    $null = $spreadsheetGroups[$fileNumber].Add($file.FullName)
                    Write-Log "  Found spreadsheet: $($file.Name) (Group: $fileNumber)" "White"
                }
            }
        }
        
        # Check if we found any spreadsheets
        if ($spreadsheetGroups.Count -eq 0) {
            Write-Log "No spreadsheets with matching numbers found in the selected folders." "Yellow"
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            return $false
        }
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path -Path $DestinationPath)) {
            New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
            Write-Log "Created destination directory: $DestinationPath" "White"
        }
        
        # Process each group of spreadsheets
        $totalGroups = $spreadsheetGroups.Count
        $completedGroups = 0
        
        foreach ($groupNumber in $spreadsheetGroups.Keys) {
            $files = $spreadsheetGroups[$groupNumber]
            
            if ($files.Count -lt 2) {
                Write-Log "Skipping group $groupNumber - only one spreadsheet found." "Yellow"
                continue
            }
            
            Write-Log "Processing spreadsheet group $groupNumber..." "Cyan"
            
            # Create a new workbook for the combined data
            $combinedWorkbook = $excel.Workbooks.Add()
            $combinedWorksheet = $combinedWorkbook.Worksheets.Item(1)
            $combinedWorksheet.Name = "Combined"
            
            $rowIndex = 1
            $isFirstFile = $true
            
            # Process each file in the group
            foreach ($file in $files) {
                Write-Log "  Combining: $file" "White"
                
                # Open the source workbook
                $sourceWorkbook = $excel.Workbooks.Open($file)
                $sourceWorksheet = $sourceWorkbook.Worksheets.Item(1)
                
                # Get used range
                $usedRange = $sourceWorksheet.UsedRange
                $lastRow = $usedRange.Rows.Count
                $lastColumn = $usedRange.Columns.Count
                
                # Handle headers based on the ExcludeHeaders option
                if ($isFirstFile -and -not $ExcludeHeaders) {
                    # Copy header row
                    $headerRange = $sourceWorksheet.Range($sourceWorksheet.Cells(1, 1), $sourceWorksheet.Cells(1, $lastColumn))
                    $headerRange.Copy() | Out-Null
                    $combinedWorksheet.Range($combinedWorksheet.Cells(1, 1), $combinedWorksheet.Cells(1, $lastColumn)).PasteSpecial(-4163) | Out-Null
                    $rowIndex = 2
                    $isFirstFile = $false
                } else {
                    # For subsequent files or if excluding headers, start from row 2 (skip header)
                    if ($isFirstFile) {
                        $isFirstFile = $false
                    }
                    $startRow = 2
                }
                
                # Copy data (excluding header for all files except first when headers are included)
                $startRow = if ($isFirstFile) { 1 } else { 2 }
                if ($lastRow -ge $startRow) {
                    $dataRange = $sourceWorksheet.Range($sourceWorksheet.Cells($startRow, 1), $sourceWorksheet.Cells($lastRow, $lastColumn))
                    $dataRange.Copy() | Out-Null
                    $combinedWorksheet.Range($combinedWorksheet.Cells($rowIndex, 1), $combinedWorksheet.Cells($rowIndex + $lastRow - $startRow, $lastColumn)).PasteSpecial(-4163) | Out-Null
                    $rowIndex += $lastRow - $startRow + 1
                }
                
                # Close the source workbook without saving
                $sourceWorkbook.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sourceWorkbook) | Out-Null
            }
            
            # Process special options for 'Add to Quantity' column if needed
            if ($DuplicateQuantityTwoRows -or $NormalizeQuantities) {
                # Find the 'Add to Quantity' column if it exists
                $addToQuantityColIndex = $null
                
                # Re-open the workbook to process it
                $combinedFilePath = Join-Path -Path $DestinationPath -ChildPath "Combined_Spreadsheet_$groupNumber.xlsx"
                $tempWorkbook = $excel.Workbooks.Add()
                $tempWorksheet = $tempWorkbook.Worksheets.Item(1)
                $tempWorksheet.Name = "Combined"
                
                # Copy all data from combined worksheet to temp worksheet
                $combinedWorksheet.UsedRange.Copy() | Out-Null
                $tempWorksheet.Range("A1").PasteSpecial(-4163) | Out-Null
                
                # Find the 'Add to Quantity' column if it exists
                $lastColumn = $tempWorksheet.UsedRange.Columns.Count
                $lastRow = $tempWorksheet.UsedRange.Rows.Count
                
                for ($col = 1; $col -le $lastColumn; $col++) {
                    $headerText = $tempWorksheet.Cells(1, $col).Text
                    if ($headerText -eq "Add to Quantity") {
                        $addToQuantityColIndex = $col
                        Write-Log "  Found 'Add to Quantity' column at index $col" "White"
                        break
                    }
                }
                
                # Process the column if found
                if ($addToQuantityColIndex) {
                    # First duplicate rows with '2' in the 'Add to Quantity' column if option is enabled
                    if ($DuplicateQuantityTwoRows) {
                        Write-Log "  Processing 'Duplicate Qty=2' option..." "White"
                        
                        # We need to process from bottom to top to avoid index shifting issues
                        $rowsToInsert = @()
                        
                        for ($row = $lastRow; $row -ge 2; $row--) {
                            $cellValue = $tempWorksheet.Cells($row, $addToQuantityColIndex).Text
                            if ($cellValue -eq "2") {
                                Write-Log "    Found row $row with quantity 2, duplicating..." "White"
                                $rowsToInsert += $row
                            }
                        }
                        
                        foreach ($row in $rowsToInsert) {
                            # Insert a new row
                            $range = $tempWorksheet.Rows($row)
                            $range.Copy() | Out-Null
                            $tempWorksheet.Rows($row).Insert(-4121) | Out-Null  # -4121 is xlShiftDown
                            $lastRow++
                        }
                        
                        Write-Log "  Duplicated $($rowsToInsert.Count) rows with quantity 2" "Green"
                    }
                    
                    # Then normalize all quantities to '1' if option is enabled
                    if ($NormalizeQuantities) {
                        Write-Log "  Processing 'Normalize Qty to 1' option..." "White"
                        $changedCount = 0
                        
                        for ($row = 2; $row -le $lastRow; $row++) {
                            $cellValue = $tempWorksheet.Cells($row, $addToQuantityColIndex).Text
                            if ($cellValue -ne "1") {
                                $tempWorksheet.Cells($row, $addToQuantityColIndex).Value = "1"
                                $changedCount++
                            }
                        }
                        
                        Write-Log "  Normalized $changedCount cells to quantity 1" "Green"
                    }
                } else {
                    Write-Log "  'Add to Quantity' column not found, skipping quantity processing" "Yellow"
                }
                
                # Save the processed workbook
                $tempWorkbook.SaveAs($combinedFilePath)
                $tempWorkbook.Close($true)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($tempWorkbook) | Out-Null
            } else {
                # Save the combined workbook normally if no special processing needed
                $combinedFilePath = Join-Path -Path $DestinationPath -ChildPath "Combined_Spreadsheet_$groupNumber.xlsx"
                $combinedWorkbook.SaveAs($combinedFilePath)
                $combinedWorkbook.Close($true)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($combinedWorkbook) | Out-Null
            }
            
            Write-Log "  Saved combined spreadsheet: $combinedFilePath" "Green"
            
            $completedGroups++
            $progressPercentage = [int](($completedGroups / $totalGroups) * 100)
            Update-ProgressBar $progressPercentage
        }
        
        # Clean up Excel
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Log "Spreadsheet combining process completed." "Cyan"
        Update-ProgressBar 100
        
        return $true
    }
    catch {
        Write-Log "Error during spreadsheet combining: $_" "Red"
        
        # Try to clean up Excel if an error occurs
        if ($excel) {
            try {
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            } catch {}
        }
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        return $false
    }
}

# Function to start the spreadsheet combining process
function Start-SpreadsheetCombiningProcess {
    if ($spreadsheetLocations.Items.Count -lt 2) {
        Write-Log "At least two spreadsheet folder locations are required." "Yellow"
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($destinationLocation.Text)) {
        Write-Log "Please select a combined destination location." "Yellow"
        return
    }
    
    Write-Log "Starting spreadsheet combining process..." "Cyan"
    Update-ProgressBar 0
    
    # Get all folder paths
    $folderPaths = New-Object System.Collections.ArrayList
    foreach ($item in $spreadsheetLocations.Items) {
        $null = $folderPaths.Add($item.Text)
    }
    
    # Get options from checkboxes
    $fileExtension = if ($optionCheckboxes[1].Checked) { "*.*" } else { "*.xlsx" }
    $excludeHeaders = $optionCheckboxes[2].Checked
    $duplicateQuantityTwoRows = $optionCheckboxes[3].Checked
    $normalizeQuantities = $optionCheckboxes[4].Checked
    
    # Start combining spreadsheets with selected options
    $success = Combine-Spreadsheets `
        -FolderPaths $folderPaths `
        -DestinationPath $destinationLocation.Text `
        -FileExtension $fileExtension `
        -ExcludeHeaders $excludeHeaders `
        -DuplicateQuantityTwoRows $duplicateQuantityTwoRows `
        -NormalizeQuantities $normalizeQuantities
    
    if ($success) {
        Write-Log "Spreadsheet combining completed successfully." "Green"
    } else {
        Write-Log "Spreadsheet combining completed with errors." "Red"
    }
}

# Function to browse for a folder
function Select-FolderDialog {
    # Use the standard Windows folder browser dialog
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select a folder"
    $folderBrowser.SelectedPath = $PSScriptRoot  # Start in the application directory
    $folderBrowser.ShowNewFolderButton = $true   # Allow creating new folders
    
    $result = $folderBrowser.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    
    return $null
}

#endregion

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Spreadsheet Wrangler"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)
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
$leftLayout.RowCount = 4
$leftLayout.ColumnCount = 1
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 30)))
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 30)))
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
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
$backupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 85)))
$backupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
$backupPanel.Controls.Add($backupLayout)

# List of backup locations
$backupLocations = New-Object System.Windows.Forms.ListView
$backupLocations.View = "Details"
$backupLocations.FullRowSelect = $true
$backupLocations.Columns.Add("Folder Path", -2)
$backupLocations.Dock = "Fill"
$backupLocations.ToolTip = "List of folders to back up. Select an item and press Delete or use the minus button to remove it."
$backupLocations.Add_KeyDown({
    param($sender, $e)
    # Delete selected item when Delete key is pressed
    if ($e.KeyCode -eq 'Delete' -and $backupLocations.SelectedItems.Count -gt 0) {
        foreach ($item in $backupLocations.SelectedItems) {
            Write-Log "Removed backup location: $($item.Text)" "Yellow"
            $backupLocations.Items.Remove($item)
        }
    }
})
$backupLayout.Controls.Add($backupLocations, 0, 0)

# Button panel for backup locations
$backupButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$backupButtonPanel.Dock = "Fill"
$backupButtonPanel.FlowDirection = "RightToLeft"
$backupButtonPanel.WrapContents = $false
$backupButtonPanel.Padding = New-Object System.Windows.Forms.Padding(5, 2, 5, 2)
$backupLayout.Controls.Add($backupButtonPanel, 0, 1)

# Remove button for backup locations
$removeBackupBtn = New-Object System.Windows.Forms.Button
$removeBackupBtn.Text = "-"
$removeBackupBtn.Width = 40
$removeBackupBtn.Height = 25
$removeBackupBtn.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$removeBackupBtn.FlatStyle = "Flat"
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($removeBackupBtn, "Remove selected backup location")
$removeBackupBtn.Add_Click({
    if ($backupLocations.SelectedItems.Count -gt 0) {
        foreach ($item in $backupLocations.SelectedItems) {
            Write-Log "Removed backup location: $($item.Text)" "Yellow"
            $backupLocations.Items.Remove($item)
        }
    } else {
        Write-Log "Please select a backup location to remove" "Yellow"
    }
})
$backupButtonPanel.Controls.Add($removeBackupBtn)

# Add button for backup locations
$addBackupBtn = New-Object System.Windows.Forms.Button
$addBackupBtn.Text = "+"
$addBackupBtn.Width = 40
$addBackupBtn.Height = 25
$addBackupBtn.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$addBackupBtn.FlatStyle = "Flat"
$toolTip.SetToolTip($addBackupBtn, "Add a new folder to back up")
$addBackupBtn.Add_Click({
    $folderPath = Select-FolderDialog
    if ($folderPath) {
        $item = New-Object System.Windows.Forms.ListViewItem($folderPath)
        $backupLocations.Items.Add($item)
        Write-Log "Added backup location: $folderPath"
    }
})
$backupButtonPanel.Controls.Add($addBackupBtn)

# Spreadsheet Folders Section
$spreadsheetPanel = New-Object System.Windows.Forms.GroupBox
$spreadsheetPanel.Text = "Spreadsheet Folder Locations"
$spreadsheetPanel.Dock = "Fill"
$leftLayout.Controls.Add($spreadsheetPanel, 0, 1)

$spreadsheetLayout = New-Object System.Windows.Forms.TableLayoutPanel
$spreadsheetLayout.Dock = "Fill"
$spreadsheetLayout.RowCount = 2
$spreadsheetLayout.ColumnCount = 1
$spreadsheetLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 85)))
$spreadsheetLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
$spreadsheetPanel.Controls.Add($spreadsheetLayout)

# List of spreadsheet folder locations
$spreadsheetLocations = New-Object System.Windows.Forms.ListView
$spreadsheetLocations.View = "Details"
$spreadsheetLocations.FullRowSelect = $true
$spreadsheetLocations.Columns.Add("Folder Path", -2)
$spreadsheetLocations.Dock = "Fill"
$spreadsheetLocations.ToolTip = "List of folders containing spreadsheets to combine. Select an item and press Delete or use the minus button to remove it."
$spreadsheetLocations.Add_KeyDown({
    param($sender, $e)
    # Delete selected item when Delete key is pressed
    if ($e.KeyCode -eq 'Delete' -and $spreadsheetLocations.SelectedItems.Count -gt 0) {
        foreach ($item in $spreadsheetLocations.SelectedItems) {
            Write-Log "Removed spreadsheet folder location: $($item.Text)" "Yellow"
            $spreadsheetLocations.Items.Remove($item)
        }
    }
})
$spreadsheetLayout.Controls.Add($spreadsheetLocations, 0, 0)

# Button panel for spreadsheet locations
$spreadsheetButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$spreadsheetButtonPanel.Dock = "Fill"
$spreadsheetButtonPanel.FlowDirection = "RightToLeft"
$spreadsheetButtonPanel.WrapContents = $false
$spreadsheetButtonPanel.Padding = New-Object System.Windows.Forms.Padding(5, 2, 5, 2)
$spreadsheetLayout.Controls.Add($spreadsheetButtonPanel, 0, 1)

# Remove button for spreadsheet locations
$removeSpreadsheetBtn = New-Object System.Windows.Forms.Button
$removeSpreadsheetBtn.Text = "-"
$removeSpreadsheetBtn.Width = 40
$removeSpreadsheetBtn.Height = 25
$removeSpreadsheetBtn.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$removeSpreadsheetBtn.FlatStyle = "Flat"
$toolTip.SetToolTip($removeSpreadsheetBtn, "Remove selected spreadsheet folder location")
$removeSpreadsheetBtn.Add_Click({
    if ($spreadsheetLocations.SelectedItems.Count -gt 0) {
        foreach ($item in $spreadsheetLocations.SelectedItems) {
            Write-Log "Removed spreadsheet folder location: $($item.Text)" "Yellow"
            $spreadsheetLocations.Items.Remove($item)
        }
    } else {
        Write-Log "Please select a spreadsheet folder location to remove" "Yellow"
    }
})
$spreadsheetButtonPanel.Controls.Add($removeSpreadsheetBtn)

# Add button for spreadsheet locations
$addSpreadsheetBtn = New-Object System.Windows.Forms.Button
$addSpreadsheetBtn.Text = "+"
$addSpreadsheetBtn.Width = 40
$addSpreadsheetBtn.Height = 25
$addSpreadsheetBtn.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$addSpreadsheetBtn.FlatStyle = "Flat"
$toolTip.SetToolTip($addSpreadsheetBtn, "Add a new folder containing spreadsheets to combine")
$addSpreadsheetBtn.Add_Click({
    $folderPath = Select-FolderDialog
    if ($folderPath) {
        $item = New-Object System.Windows.Forms.ListViewItem($folderPath)
        $spreadsheetLocations.Items.Add($item)
        Write-Log "Added spreadsheet folder location: $folderPath"
    }
})
$spreadsheetButtonPanel.Controls.Add($addSpreadsheetBtn)

# Combined Destination Location Section
$destinationPanel = New-Object System.Windows.Forms.GroupBox
$destinationPanel.Text = "Combined Destination Location"
$destinationPanel.Dock = "Fill"
$leftLayout.Controls.Add($destinationPanel, 0, 2)

$destinationLayout = New-Object System.Windows.Forms.TableLayoutPanel
$destinationLayout.Dock = "Fill"
$destinationLayout.RowCount = 2
$destinationLayout.ColumnCount = 1
$destinationLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 85)))
$destinationLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
$destinationPanel.Controls.Add($destinationLayout)

# Destination location display
$destinationLocation = New-Object System.Windows.Forms.TextBox
$destinationLocation.ReadOnly = $true
$destinationLocation.Dock = "Fill"
$destinationLocation.BackColor = [System.Drawing.Color]::White
$toolTip.SetToolTip($destinationLocation, "Location where combined spreadsheets will be saved")
$destinationLayout.Controls.Add($destinationLocation, 0, 0)

# Browse button for destination location
$browseDestinationBtn = New-Object System.Windows.Forms.Button
$browseDestinationBtn.Text = "Browse..."
$browseDestinationBtn.Dock = "Right"
$browseDestinationBtn.Width = 80
$browseDestinationBtn.Height = 25
$browseDestinationBtn.Margin = New-Object System.Windows.Forms.Padding(3, 2, 3, 2)
$browseDestinationBtn.FlatStyle = "Flat"
$toolTip.SetToolTip($browseDestinationBtn, "Select a folder where combined spreadsheets will be saved")
$browseDestinationBtn.Add_Click({
    $folderPath = Select-FolderDialog
    if ($folderPath) {
        $destinationLocation.Text = $folderPath
        Write-Log "Set combined destination location: $folderPath"
    }
})
$destinationLayout.Controls.Add($browseDestinationBtn, 0, 1)

# Run Button
$runBtn = New-Object System.Windows.Forms.Button
$runBtn.Text = "Run"
$runBtn.Dock = "Fill"
$runBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$runBtn.ForeColor = [System.Drawing.Color]::White
$runBtn.FlatStyle = "Flat"
$runBtn.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$toolTip.SetToolTip($runBtn, "Start the backup and spreadsheet combining process")
$runBtn.Add_Click({
    # Clear previous output
    $outputTextbox.Clear()
    Write-Log "Starting operations..." "Cyan"
    
    # Start backup process if not skipped
    if (-not $optionCheckboxes[0].Checked) {
        Start-BackupProcess
    } else {
        Write-Log "Backup process skipped due to 'Skip Backup' option." "Yellow"
    }
    
    # Start spreadsheet combining process
    Start-SpreadsheetCombiningProcess
    
    Write-Log "All operations completed." "Cyan"
})
$leftLayout.Controls.Add($runBtn, 0, 3)
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

$optionsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$optionsLayout.Dock = "Fill"
$optionsLayout.RowCount = 3
$optionsLayout.ColumnCount = 3
$optionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$optionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$optionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$optionsPanel.Controls.Add($optionsLayout)

# Create checkboxes for options with specific functionality
$optionCheckboxes = @()

# Option 1: Skip backup process
$optionCheckboxes += $checkbox1 = New-Object System.Windows.Forms.CheckBox
$checkbox1.Text = "Skip Backup"
$checkbox1.Dock = "Fill"
$checkbox1.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox1, "Skip the backup process and only combine spreadsheets")
$optionsLayout.Controls.Add($checkbox1, 0, 0)

# Option 2: Support multiple file formats
$optionCheckboxes += $checkbox2 = New-Object System.Windows.Forms.CheckBox
$checkbox2.Text = "All Formats"
$checkbox2.Dock = "Fill"
$checkbox2.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox2, "Process all spreadsheet formats (.xlsx, .xls, .csv)")
$optionsLayout.Controls.Add($checkbox2, 1, 0)

# Option 3: Exclude headers
$optionCheckboxes += $checkbox3 = New-Object System.Windows.Forms.CheckBox
$checkbox3.Text = "No Headers"
$checkbox3.Dock = "Fill"
$checkbox3.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox3, "Exclude headers when combining spreadsheets")
$optionsLayout.Controls.Add($checkbox3, 2, 0)

# Option 4: Duplicate rows with '2' in 'Add to Quantity' column
$optionCheckboxes += $checkbox4 = New-Object System.Windows.Forms.CheckBox
$checkbox4.Text = "Duplicate Qty=2"
$checkbox4.Dock = "Fill"
$checkbox4.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox4, "Duplicate rows with '2' in the 'Add to Quantity' column")
$optionsLayout.Controls.Add($checkbox4, 0, 1)

# Option 5: Normalize all quantities to '1'
$optionCheckboxes += $checkbox5 = New-Object System.Windows.Forms.CheckBox
$checkbox5.Text = "Normalize Qty to 1"
$checkbox5.Dock = "Fill"
$checkbox5.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox5, "Change all values in 'Add to Quantity' column to '1' (runs after duplication)")
$optionsLayout.Controls.Add($checkbox5, 1, 1)

# Option 6: Placeholder
$optionCheckboxes += $checkbox6 = New-Object System.Windows.Forms.CheckBox
$checkbox6.Text = "Option 6"
$checkbox6.Dock = "Fill"
$checkbox6.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox6, "Reserved for future functionality")
$optionsLayout.Controls.Add($checkbox6, 2, 1)

# Option 7: Placeholder
$optionCheckboxes += $checkbox7 = New-Object System.Windows.Forms.CheckBox
$checkbox7.Text = "Option 7"
$checkbox7.Dock = "Fill"
$checkbox7.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox7, "Reserved for future functionality")
$optionsLayout.Controls.Add($checkbox7, 0, 2)

# Option 8: Placeholder
$optionCheckboxes += $checkbox8 = New-Object System.Windows.Forms.CheckBox
$checkbox8.Text = "Option 8"
$checkbox8.Dock = "Fill"
$checkbox8.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox8, "Reserved for future functionality")
$optionsLayout.Controls.Add($checkbox8, 1, 2)

# Option 9: Placeholder
$optionCheckboxes += $checkbox9 = New-Object System.Windows.Forms.CheckBox
$checkbox9.Text = "Option 9"
$checkbox9.Dock = "Fill"
$checkbox9.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox9, "Reserved for future functionality")
$optionsLayout.Controls.Add($checkbox9, 2, 2)

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
$toolTip.SetToolTip($outputTextbox, "Displays real-time progress and status information")
$outputPanel.Controls.Add($outputTextbox)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock = "Fill"
$progressBar.Style = "Continuous"
$progressBar.Value = 0
$toolTip.SetToolTip($progressBar, "Shows overall progress of the current operation")
$rightLayout.Controls.Add($progressBar, 0, 2)
#endregion

# Initialize the form with some sample data for visualization
Write-Log "Application initialized and ready to run." "Cyan"
Write-Log "Please add backup, spreadsheet, and combined destination folder locations." "White"
Write-Log "Tip: You can remove locations by selecting them and pressing Delete." "Yellow"

# Show the form
$form.ShowDialog()
