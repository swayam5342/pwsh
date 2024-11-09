# Get All Shell Folder Shortcuts Script
# https://gist.github.com/ThioJoe/16eac0ea7d586c4edba41b454b58b225

# How to Use:
# 1. Open powershell, and navigate to the path with the script using 'cd' command
# 2. Run the following command to allow running scripts for the current session:
#        Set-ExecutionPolicy -ExecutionPolicy unrestricted -Scope Process
# 3. Without closing the powershell window, run the script by typing the name of the script file starting with .\  for example:
#        .\Get_All_Shell_Folder_Shortcuts.ps1
# 4. Wait for it to finish, then look in the "Shell Folder Shortcuts" folder for the results

# Create a main folder for all shell folder shortcuts in the current script directory
$mainShortcutsFolder = Join-Path $PSScriptRoot "Shell Folder Shortcuts"
New-Item -Path $mainShortcutsFolder -ItemType Directory -Force | Out-Null

# Create a subfolder for CLSID-based shortcuts
$CLSIDshortcutsOutputFolder = Join-Path $mainShortcutsFolder "CLSID Shell Folder Shortcuts"
New-Item -Path $CLSIDshortcutsOutputFolder -ItemType Directory -Force | Out-Null

# Create a subfolder for named special folders shortcuts
$namedShortcutsOutputFolder = Join-Path $mainShortcutsFolder "Named Shell Folder Shortcuts"
New-Item -Path $namedShortcutsOutputFolder -ItemType Directory -Force | Out-Null

# Add necessary type for string localization
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    public class Windows
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);
    }
"@


if (-not ([System.Management.Automation.PSTypeName]'Win32').Type) {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    public class Win32 {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr LoadLibrary(string lpFileName);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);
    }
"@
}

function Get-LocalizedString {
    param ( [string]$StringReference )
    if ($StringReference -match '@(.+),-(\d+)') {
        $dllPath = [Environment]::ExpandEnvironmentVariables($Matches[1])
        $resourceId = [uint32]$Matches[2]
        $hModule = [Win32]::LoadLibrary($dllPath)
        if ($hModule -eq [IntPtr]::Zero) {
            Write-Error "Failed to load library: $dllPath"
            return $null
        }
        $stringBuilder = New-Object System.Text.StringBuilder 1024
        $result = [Win32]::LoadString($hModule, $resourceId, $stringBuilder, $stringBuilder.Capacity)
        if ($result -ne 0) {
            return $stringBuilder.ToString()
        } else {
            Write-Error "Failed to load string resource: $resourceId from $dllPath"
            return $null
        }
    } else {
        Write-Error "Invalid string reference format: $StringReference"
        return $null
    }
}

function Get-FolderName {
    param (
        [string]$clsid
    )

    $Global:NameSource = "Unknown"

    Write-Host "Attempting to get folder name for CLSID: $clsid"

    # Check the default value in HKEY_CLASSES_ROOT\CLSID\
    $defaultPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid"
    Write-Host "Checking default value at: $defaultPath"
    $defaultName = (Get-ItemProperty -Path $defaultPath -ErrorAction SilentlyContinue).'(default)'
    if ($defaultName) { 
        Write-Host "Found default name: $defaultName"
        if ($defaultName -match '@.+,-\d+') {
            Write-Host "Default name is a localized string reference"
            $resolvedName = Get-LocalizedString $defaultName
            if ($resolvedName) {
                $Global:NameSource = "Localized String"
                Write-Host "Resolved default name to: $resolvedName"
                return $resolvedName
            }
            else {
                Write-Host "Failed to resolve default name, using original value"
            }
        }
        $Global:NameSource = "Default Value"
        return $defaultName 
    }
    else {
        Write-Host "No default name found"
    }

    # Check for TargetKnownFolder
    $initPropertyBagPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\Instance\InitPropertyBag"
    Write-Host "Checking for TargetKnownFolder at: $initPropertyBagPath"
    $targetKnownFolder = (Get-ItemProperty -Path $initPropertyBagPath -ErrorAction SilentlyContinue).TargetKnownFolder
    if ($targetKnownFolder) {
        Write-Host "Found TargetKnownFolder: $targetKnownFolder"
        $folderDescriptionsPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\$targetKnownFolder"
        Write-Host "Checking for folder name at: $folderDescriptionsPath"
        $folderName = (Get-ItemProperty -Path $folderDescriptionsPath -ErrorAction SilentlyContinue).Name
        if ($folderName) { 
            $Global:NameSource = "Known Folder ID"
            Write-Host "Found folder name: $folderName"
            return $folderName 
        }
        else {
            Write-Host "No folder name found in FolderDescriptions"
        }
    }
    else {
        Write-Host "No TargetKnownFolder found"
    }

    # Check for LocalizedString
    $localizedStringPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid"
    Write-Host "Checking for LocalizedString at: $localizedStringPath"
    $localizedString = (Get-ItemProperty -Path $localizedStringPath -ErrorAction SilentlyContinue).LocalizedString
    if ($localizedString) {
        Write-Host "Found LocalizedString: $localizedString"
        $resolvedString = Get-LocalizedString $localizedString
        if ($resolvedString) {
            $Global:NameSource = "Localized String"
            Write-Host "Resolved LocalizedString to: $resolvedString"
            return $resolvedString
        }
        else {
            Write-Host "Failed to resolve LocalizedString"
        }
    }
    else {
        Write-Host "No LocalizedString found"
    }

    # Check Desktop\NameSpace registry key as a fallback
    $namespacePath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$clsid"
    Write-Host "Checking Desktop\NameSpace at: $namespacePath"
    $namespaceName = (Get-ItemProperty -Path $namespacePath -ErrorAction SilentlyContinue).'(default)'
    if ($namespaceName) {
        $Global:NameSource = "Desktop Namespace"
        Write-Host "Found name in Desktop\NameSpace: $namespaceName"
        return $namespaceName
    }
    else {
        Write-Host "No name found in Desktop\NameSpace"
    }

    # If all else fails, return the CLSID
    Write-Host "Returning CLSID as folder name"
    return $clsid
}

function Create-Shortcut {
    param (
        [string]$clsid,
        [string]$name,
        [string]$shortcutPath
    )

    try {
        Write-Host "Creating shortcut for $name at $shortcutPath"
        $shell = New-Object -ComObject WScript.Shell

        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = "explorer.exe"
        $shortcut.Arguments = "shell:::$clsid"

        # Find the icon
        $iconPath = (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\DefaultIcon" -ErrorAction SilentlyContinue).'(default)'
        if ($iconPath) {
            Write-Host "Setting custom icon: $iconPath"
            $shortcut.IconLocation = $iconPath
        }
        else {
            Write-Host "No custom icon found. Setting default folder icon."
            $shortcut.IconLocation = "%SystemRoot%\System32\shell32.dll,3"
        }

        $shortcut.Save()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        return $true
    }
    catch {
        Write-Host "Error creating shortcut for $name`: $($_.Exception.Message)"
        return $false
    }
}

# Get all CLSIDs with ShellFolder subkey
$shellFolders = Get-ChildItem -Path 'Registry::HKEY_CLASSES_ROOT\CLSID' | 
    Where-Object {$_.GetSubKeyNames() -contains "ShellFolder"} | 
    Select-Object PSChildName

Write-Host "Found $($shellFolders.Count) shell folders"

# Create an array to store CLSID information
$clsidInfo = @()

# Modify the main loop for processing CLSIDs
foreach ($folder in $shellFolders) {
    $clsid = $folder.PSChildName
    Write-Host "`nProcessing CLSID: $clsid"
    $Global:NameSource = "Unknown"
    $name = Get-FolderName -clsid $clsid

    $sanitizedName = $name -replace '[\\/:*?"<>|]', '_'
    $shortcutPath = Join-Path $CLSIDshortcutsOutputFolder "$sanitizedName.lnk"

    Write-Host "Attempting to create shortcut: $shortcutPath"
    $success = Create-Shortcut -clsid $clsid -name $name -shortcutPath $shortcutPath

    if ($success) {
        Write-Host "Successfully created shortcut for $name"
    }
    else {
        Write-Host "Failed to create shortcut for $name"
    }

    # Store the CLSID information
    $iconPath = (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\DefaultIcon" -ErrorAction SilentlyContinue).'(default)'
    $clsidInfo += [PSCustomObject]@{
        CLSID = $clsid
        Name = $name
        NameSource = $Global:NameSource
        IconPath = $iconPath
    }
}

Write-Host "`nAll shortcuts have been processed. Check the Shortcuts folder at $CLSIDshortcutsOutputFolder"



function Create-NamedShortcut {
    param (
        [string]$name,
        [string]$shortcutPath,
        [string]$iconPath
    )

    try {
        Write-Host "Creating named shortcut for $name at $shortcutPath"
        $shell = New-Object -ComObject WScript.Shell

        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = "explorer.exe"
        $shortcut.Arguments = "shell:$name"

        # Set icon
        if ($iconPath) {
            Write-Host "Setting custom icon: $iconPath"
            $shortcut.IconLocation = $iconPath
        }
        else {
            Write-Host "Setting default folder icon"
            $shortcut.IconLocation = "%SystemRoot%\System32\shell32.dll,3"
        }

        $shortcut.Save()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        return $true
    }
    catch {
        Write-Host "Error creating shortcut for $name`: $($_.Exception.Message)"
        return $false
    }
}

# -------------------------------------------------------------------------------------------------------------


# Get all named special folders
$namedFolders = Get-ChildItem -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions"

Write-Host "Found $($namedFolders.Count) named special folders"

# Create shortcuts for each named special folder
foreach ($folder in $namedFolders) {
    $folderProperties = Get-ItemProperty -Path $folder.PSPath
    $folderName = $folderProperties.Name
    $iconPath = $folderProperties.Icon

    if ($folderName) {
        Write-Host "`nProcessing named folder: $folderName"

        $sanitizedName = $folderName -replace '[\\/:*?"<>|]', '_'
        $shortcutPath = Join-Path $namedShortcutsOutputFolder "$sanitizedName.lnk"

        Write-Host "Attempting to create shortcut: $shortcutPath"
        $success = Create-NamedShortcut -name $folderName -shortcutPath $shortcutPath -iconPath $iconPath

        if ($success) {
            Write-Host "Successfully created shortcut for $folderName"
        }
        else {
            Write-Host "Failed to create shortcut for $folderName"
        }
    }
    else {
        Write-Host "Skipping folder with no name: $($folder.PSChildName)"
    }
}

Write-Host "`nAll named shortcuts have been processed. Check the Shortcuts2 folder at $namedShortcutsOutputFolder"



# ---------------------------------------------------------------------

function Create-CLSIDCsvFile {
    param (
        [string]$outputPath,
        [array]$clsidData
    )

    $csvContent = "CLSID,ExplorerCommand,Name,NameSource,CustomIcon`n"

    foreach ($item in $clsidData) {
        $explorerCommand = "explorer shell:::$($item.CLSID)"
        $iconPath = if ($item.IconPath) { 
            "`"$($item.IconPath -replace '"', '""')`""
        } else { 
            "None" 
        }

        # Escape any double quotes in the name
        $escapedName = $item.Name -replace '"', '""'

        $csvContent += "$($item.CLSID),`"$explorerCommand`",`"$escapedName`",$($item.NameSource),$iconPath`n"
    }

    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

function Create-NamedFoldersCsvFile {
    param (
        [string]$outputPath
    )

    $csvContent = "Name,ExplorerCommand,RelativePath,ParentFolder`n"

    $namedFolders = Get-ChildItem -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions"

    foreach ($folder in $namedFolders) {
        $folderProperties = Get-ItemProperty -Path $folder.PSPath
        $name = $folderProperties.Name
        if ($name) {
            $explorerCommand = "explorer shell:$name"
            $relativePath = $folderProperties.RelativePath -replace ',', '","'
            $parentFolderGuid = $folderProperties.ParentFolder
            $parentFolderName = "None"
            if ($parentFolderGuid) {
                $parentFolderPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\$parentFolderGuid"
                $parentFolderName = (Get-ItemProperty -Path $parentFolderPath -ErrorAction SilentlyContinue).Name
            }

            $csvContent += "`"$name`",`"$explorerCommand`",`"$relativePath`",`"$parentFolderName`"`n"
        }
    }

    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

# Create the CSV file using the stored data
$clsidCsvPath = Join-Path $mainShortcutsFolder "CLSID_Shell_Folders.csv"
Create-CLSIDCsvFile -outputPath $clsidCsvPath -clsidData $clsidInfo

$namedFoldersCsvPath = Join-Path $mainShortcutsFolder "Named_Shell_Folders.csv"
Create-NamedFoldersCsvFile -outputPath $namedFoldersCsvPath