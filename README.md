# KrakenImport

**Author:** Designed and developed from scratch to automate laboratory plate data processing, integrating database access, Excel parsing, and XML generation for LIMS.

## Background
In our laboratory workflow, generating import files for the Kraken system required manual Excel formatting and data matching with our LIMS (SampleManager). This process was time-consuming and error-prone.
KrakenImport automates this, ensuring consistent well mappings and file formats, reducing human error and speeding up plate processing.


    
## Impact

Reduced import preparation time from ~30 minutes to <5 minutes per plate

Eliminated mismatches between Kraken imports and LIMS naming conventions

## Tech Stack

.NET 8.0 (C#)

SQL Server integration

Excel parsing (ClosedXML)

XML generation for LIMS

Interactive CLI

**KrakenImport** is a .NET 8.0 console application for importing, processing, and exporting laboratory plate data, including support for different Excel formats, database integration, and XML export for LIMS systems.


## Features

- Import sample data from Excel or SQL databases
- Supports multiple plate formats (Intertek, EIB, etc.)
- Generates master plate XML files for LIMS
- Handles routine and verification workflows
- Customizable well mapping and plate layouts
- Command-line interface with interactive prompts

## Requirements

- [.NET 8.0 SDK](https://dotnet.microsoft.com/download)
- Windows OS (tested)
- Access to the relevant SQL database (if using database features)

## Getting Started

### 1. Clone the repository

```sh
git clone https://github.com/yourusername/KrakenImport.git
cd KrakenImport
```

### 2. Set up the database connection

Set the environment variable `KRAKEN_DB_CONNECTION` to your database connection string:

**PowerShell:**
```powershell
$env:KRAKEN_DB_CONNECTION="Server=YOUR_SERVER;Database=YOUR_DB;Integrated Security=True;"
```

**Command Prompt:**
```cmd
set KRAKEN_DB_CONNECTION=Server=YOUR_SERVER;Database=YOUR_DB;Integrated Security=True;
```

### 3. Build and publish

To build a single-file executable:

```sh
dotnet publish -c Release -r win-x64 --self-contained true
```

The executable will be in `bin/Release/net8.0/win-x64/publish/`.

### 4. Run the application

```sh
.\bin\Release\net8.0\win-x64\publish\Kraken.exe
```

Follow the interactive prompts to import and export your data.

## Usage

- Enter the order ID and output path as prompted.
- Choose the data source and format.
- The application will process the data and generate the required XML files.

## Logs

Application logs (if enabled) are stored in the `logs/` directory.

## License

This project is licensed under the MIT License.

---

**Note:**  
Do **not** commit real database credentials or sensitive data to this repository. Use environment variables for secrets.
