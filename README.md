# Intune ODC Log Collector

**GUI tool to collect Intune One Data Collector (ODC) logs for Microsoft Support.**

## What It Does

This tool automates the collection of Intune diagnostic logs using Microsoft's official ODC scripts. These logs are used by Microsoft Support to troubleshoot Intune enrollment, policy, and compliance issues.

## Process

The tool performs these steps:
1. Creates directory: `C:\IntuneODCLogs`
2. Downloads Microsoft diagnostic scripts from `aka.ms/intunexml` and `aka.ms/intuneps1`
3. Runs the Intune ODC collection script (~10 minutes)
4. Creates a compressed ZIP file with all diagnostic data

## Requirements

- Windows 10/11
- Administrator privileges (required for log collection)
- Internet connection (to download Microsoft scripts)
- Python 3.7+ (if running from source)

## Usage

### Option 1: Run with Python

```bash
git clone https://github.com/kbaker827/intune-odc-collector.git
cd intune-odc-collector
python main.py
```

### Option 2: Download Executable

Download `IntuneODCCollector.exe` from [Releases](../../releases) and run as Administrator.

## Instructions

1. **Run as Administrator** (required)
2. Click **"📥 Start Collection"**
3. Wait ~10 minutes for collection to complete
4. Click **"📂 Open Log Folder"** to access the ZIP file
5. Send the ZIP file to Microsoft Support

## Output Location

All collected logs are saved to:
```
C:\IntuneODCLogs\
```

Look for a ZIP file with a name like:
```
IntuneODC_YYYYMMDD_HHMMSS.zip
```

## Building Executable

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "IntuneODCCollector" main.py
```

## Safety

- ✅ Uses official Microsoft scripts from `aka.ms` links
- ✅ Only collects diagnostic information
- ✅ No data is uploaded - everything stays local
- ✅ Creates logs in standard `C:\IntuneODCLogs` location

## License

MIT License
