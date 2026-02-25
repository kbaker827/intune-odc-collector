<#
    Stand-alone implementation of One Data Collector
#>

# Copyright © 2016, Microsoft Corporation. All rights reserved.
# :: ======================================================= ::

#region Fields
$Global:ResultRootDirectory = [System.IO.Path]::Combine(($env:TEMP), 'CollectedData')
$fileTime = Get-Date ([datetime]::UtcNow) -UFormat "%m_%d_%Y_%H_%M_UTC%Z"
$CompressedResultFileName = "$($env:COMPUTERNAME)_CollectedData_$fileTime.ZIP"
$global:LogName = "$env:systemroot\temp\stdout.log"
$ODCversion = "2026.1.30"
#endregion

# Write-Log function
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [string]$Message = "",
        [Parameter()]
        [ValidateSet('Information', 'Warning', 'Error', 'Verbose')]
        [string]$Level = 'Information',
        [Parameter()]
        [switch]$WriteStdOut,
        [Parameter()]
        [string]$LogName = $global:LogName
    )
    BEGIN {
        if (($null -eq $LogName) -or ($LogName -eq "")) { Write-Error "Please set variable \`$global\`:LogName." }
    }
    PROCESS {
        if (($Level -eq "Verbose") -and (-not ($debugMode))) {
        } else {
            [pscustomobject]@{
                Time = (Get-Date -f u)
                Line = "`[$($MyInvocation.ScriptLineNumber)`]"
                Level = $Level
                Message = $Message
            } | Export-Csv -Path $LogName -Append -Force -NoTypeInformation -Encoding Unicode
            if ($WriteStdOut -or (($Level -eq "Verbose") -and $debugMode)) { Write-Output $Message }
        }
    }
    END {}
}

# Test-IsAdmin function
Function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal] `
      [Security.Principal.WindowsIdentity]::GetCurrent() `
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Create-ZipFromDirectory function
function Create-ZipFromDirectory {
    PARAM
    (
        [Alias('S')]
        [Parameter(Position = 1, Mandatory = $true)]
        [ValidateScript({Test-Path -Path $_})]
        [string]$SourceDirectory,
        [Alias('O')]
        [parameter(Position = 2, Mandatory = $false)]
        [string]$ZipFileName,
        [Alias('Force')]
        [Parameter(Mandatory = $false)]
        [switch]$Overwrite
    )
    PROCESS {
        $ZipFileName = ("{0}.zip" -f $ZipFileName), $ZipFileName)[$ZipFileName.EndsWith('.zip', [System.StringComparison]::OrdinalIgnoreCase)]
        if(![System.IO.Path]::IsPathRooted($ZipFileName)) {
            $ZipFileName = ("{0}\{1}" -f (Get-Location), $ZipFileName)
        }
        if($Overwrite) {
           if(Test-Path($ZipFileName)){ Remove-Item $ZipFileName -Force -ErrorAction SilentlyContinue }
        }
        $source = Get-Item $SourceDirectory
        if ($source.PSIsContainer) {
            try {
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                [System.IO.Compression.ZipFile]::CreateFromDirectory($source.FullName, $ZipFileName, [System.IO.Compression.CompressionLevel]::Optimal, $false)
            }
            catch {
                $timeElapsed = 0
                while (Get-Process -Name msinfo32 -ErrorAction SilentlyContinue) {
                    sleep 20
                    $timeElapsed += 20
                    "Waiting for msinfo32 to exit. Elapsed time: $timeElapsed seconds."
                }
                [System.IO.Compression.ZipFile]::CreateFromDirectory($source.FullName, $ZipFileName, [System.IO.Compression.CompressionLevel]::Optimal, $false)
            }
        }
    }
}

# Initialize
if (-not (Test-IsAdmin)) {
    Return "Please run PowerShell elevated (run as administrator) and run the script again."
    Break
}

$ResultRootDirectory = [System.IO.Path]::Combine(($env:TEMP), 'CollectedData')
if (Test-Path $ResultRootDirectory) {
    Remove-Item -Recurse -Path $ResultRootDirectory -Force
}

Write-Output "Starting Intune ODC ver. $ODCversion"
$xmlPath = Join-Path (Get-Location) 'Intune.XML'
Write-Output "XML path is $xmlPath"
Write-Output "Working folder is $(Get-Location)"

# Download XML if not present
if (-not (Test-Path $xmlPath)) {
    try {
        $downloadLocation = "https://raw.githubusercontent.com/markstan/IntuneOneDataCollector/master/Intune.xml"
        Invoke-WebRequest -UseBasicParsing -Uri $downloadLocation -OutFile .\Intune.XML
    }
    catch {
        Write-Output "Unable to download Intune.XML"
        exit 1
    }
}

# Verify XML
if ((Get-Content .\Intune.xml -TotalCount 1) -ne '<?xml version="1.0" encoding="utf-8"?>') {
    Write-Output "Warning: Intune.xml does not contain valid header"
}

Write-Output "Loading $xmlPath"

# Process XML and collect data
[xml]$xml = Get-Content $xmlPath
$packages = $xml.DataPoints.Package

foreach ($package in $packages) {
    $packageID = $package.ID
    Write-Output "Processing package: $packageID"
    
    # Collect Files
    if ($package.Files.File) {
        foreach ($file in $package.Files.File) {
            $filePath = [System.Environment]::ExpandEnvironmentVariables($file.Value)
            $filePath = $filePath -replace '"', ""
            if (Test-Path $filePath) {
                try {
                    $resolvedFiles = Get-ChildItem -Path $filePath
                    foreach ($resolvedFile in $resolvedFiles) {
                        $teamName = if ($file.Team) { $file.Team } else { 'General' }
                        $dstFileName = $env:COMPUTERNAME + "_" + $resolvedFile.Name
                        $destDir = Join-Path $ResultRootDirectory $packageID "Files" $teamName
                        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                        $destFile = Join-Path $destDir $dstFileName
                        if ($resolvedFile.Length -ne 0) {
                            Copy-Item -Path $resolvedFile.FullName -Destination $destFile -Force
                            Write-Output "Collected: $($resolvedFile.Name)"
                        }
                    }
                }
                catch {
                    Write-Output "Error collecting file: $filePath"
                }
            }
        }
    }
    
    # Collect Registry Keys
    if ($package.Registries.Registry) {
        foreach ($reg in $package.Registries.Registry) {
            $regKey = $reg.Value.Replace('\*', '')
            $teamName = if ($reg.Team) { $reg.Team } else { 'General' }
            $outputFile = if ($reg.OutputFileName) { $reg.OutputFileName } else { ($regKey -replace '\\', '_') }
            $outputFile = $env:COMPUTERNAME + "_" + $outputFile + ".txt"
            $destDir = Join-Path $ResultRootDirectory $packageID "RegistryKeys" $teamName
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            $outputPath = Join-Path $destDir $outputFile
            try {
                REG.exe EXPORT $regKey $outputPath /y /reg:64 2>&1 | Out-Null
                Write-Output "Collected registry: $regKey"
            }
            catch {
                Write-Output "Error collecting registry: $regKey"
            }
        }
    }
    
    # Collect Event Logs
    if ($package.EventLogs.EventLog) {
        foreach ($evt in $package.EventLogs.EventLog) {
            $sourceLog = [System.Environment]::ExpandEnvironmentVariables($evt.Value)
            if (Test-Path $sourceLog) {
                try {
                    $resolvedLogs = Get-ChildItem -Path $sourceLog
                    foreach ($resolvedLog in $resolvedLogs) {
                        $teamName = if ($evt.Team) { $evt.Team } else { 'General' }
                        $dstLog = $env:COMPUTERNAME + "_" + $resolvedLog.Name
                        $destDir = Join-Path $ResultRootDirectory $packageID "EventLogs" $teamName
                        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                        $destFile = Join-Path $destDir $dstLog
                        Copy-Item -Path $resolvedLog.FullName -Destination $destFile -Force
                        Write-Output "Collected event log: $($resolvedLog.Name)"
                    }
                }
                catch {
                    Write-Output "Error collecting event log: $sourceLog"
                }
            }
        }
    }
    
    # Collect Commands
    if ($package.Commands.Command) {
        foreach ($cmd in $package.Commands.Command) {
            $teamName = if ($cmd.Team) { $cmd.Team } else { 'General' }
            $outputFile = if ($cmd.OutputFileName -and $cmd.OutputFileName -ne "NA") { 
                $cmd.OutputFileName 
            } else { 
                [System.IO.Path]::GetRandomFileName() + ".txt"
            }
            $outputFile = $env:COMPUTERNAME + "_" + $outputFile
            $destDir = Join-Path $ResultRootDirectory $packageID "Commands" $teamName
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            $outputPath = Join-Path $destDir ([System.IO.Path]::GetFileNameWithoutExtension($outputFile) + ".txt")
            
            try {
                if ($cmd.Type -eq "PS") {
                    (Invoke-Expression $cmd.Value) | Out-File $outputPath
                    Write-Output "Collected command output: $($cmd.Value)"
                }
                elseif ($cmd.Type -eq "CMD") {
                    (CMD.exe /c $cmd.Value 2>&1) | Out-File $outputPath
                    Write-Output "Collected command output: $($cmd.Value)"
                }
            }
            catch {
                Write-Output "Error running command: $($cmd.Value)"
            }
        }
    }
}

# Compress results
Write-Output "Compressing collected data..."
if (Test-Path $ResultRootDirectory) {
    try {
        Create-ZipFromDirectory -Source $ResultRootDirectory -ZipFileName $CompressedResultFileName -Force
        Copy-Item -Path (Join-Path (Get-Location) $CompressedResultFileName) -Destination (Get-Location) -Force
        Write-Output "Created: $CompressedResultFileName"
    }
    catch {
        Start-Sleep -Seconds 15
        Create-ZipFromDirectory -Source $ResultRootDirectory -ZipFileName $CompressedResultFileName -Force
    }
    finally {
        Remove-Item -Path $ResultRootDirectory -Force -Recurse -ErrorAction SilentlyContinue
    }
}

Write-Output "Collection complete!"
Write-Output "Output file: $CompressedResultFileName"
