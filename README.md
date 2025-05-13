# Spreadsheet Wrangler

A PowerShell GUI application for folder backups and spreadsheet combining operations.

<img src="https://github.com/user-attachments/assets/c78e7d1f-e388-4f1a-841d-0bfc4900e2b8" width=75% height=75%>

## Features

### Backup Functionality
- Create timestamped backups of selected folders
- Support for multiple backup locations
- Automatic ".backup" folder creation
- Option to skip backup process

### Spreadsheet Combining
- Combine spreadsheets with similar numbering across folders
- Support for multiple spreadsheet formats (.xlsx, .xls, .csv)
- Maintain headers from the first spreadsheet (optional)
- Save combined spreadsheets to a user-selected destination

### Advanced Spreadsheet Options
- **No Headers**: Exclude headers when combining spreadsheets
- **Duplicate Qty=2**: Duplicate rows with '2' in the "Add to Quantity" column
- **Normalize Qty to 1**: Change all values in "Add to Quantity" column to '1'
- **Log to File**: Save terminal output to a log file for future reference

### Configuration Management
- Save/load application settings to/from XML files
- Menu system with keyboard shortcuts
- Persistent settings across sessions

## Requirements

- Windows operating system
- PowerShell 5.1 or higher
- Microsoft Excel (for spreadsheet operations)

## Installation

1. Clone this repository or download the latest release
2. Extract the files to your preferred location
3. Run `SpreadsheetWrangler.ps1` with PowerShell:
```powershell
powershell -ExecutionPolicy Bypass -File .\SpreadsheetWrangler.ps1
```

## Usage

### Backup Process
1. Add folder locations to back up using the "+" button
2. Select "Skip Backup" option if you only want to combine spreadsheets

### Spreadsheet Combining
1. Add folder locations containing spreadsheets using the "+" button
2. Set the destination folder for combined spreadsheets
3. Select desired options for the combining process
4. Click "Run" to start the process

### Configuration
- **File → New Configuration**: Reset all settings to default
- **File → Open Configuration**: Load settings from an XML file
- **File → Save Configuration**: Save settings to the current file
- **File → Save Configuration As**: Save settings to a new file

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

Created by Bryant Welch

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request
