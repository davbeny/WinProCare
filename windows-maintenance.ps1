# Complete system maintenance script, browser and application updates

# Global variables
$patchmypcPath = "$env:USERPROFILE\PatchMyPC.exe"
$logFilePath = "$env:USERPROFILE\PatchMyPC_Update_Log.txt"
$downloadsPath = "$env:USERPROFILE\Downloads"

# Create restore point
function Manage-RestorePoints {
    try {
        Checkpoint-Computer -Description "Before cleaning and updates" -RestorePointType "MODIFY_SETTINGS"
        Write-Host "[*] New restore point created: 'Before cleaning and updates'" -ForegroundColor Green
    }
    catch {
        Write-Host "[!] Unable to create a restore point. Error: $_" -ForegroundColor Red
    }
}

# Remove .exe files from downloads folder
function Remove-ExeFiles {
    Write-Host "[*] Deleting all .exe files in the Downloads folder..." -ForegroundColor Cyan
    $deletedFiles = @()
    try {
        $exeFiles = Get-ChildItem -Path $downloadsPath -Filter "*.exe"
        foreach ($file in $exeFiles) {
            Remove-Item -Path $file.FullName -Force
            $deletedFiles += $file.Name
        }
        
        if ($deletedFiles.Count -gt 0) {
            Write-Host "[+] Deleted exe files:" -ForegroundColor Green
            foreach ($fileName in $deletedFiles) {
                Write-Host "   - $fileName" -ForegroundColor Green
            }
            Write-Host "[*] A total of $($deletedFiles.Count) exe files were deleted" -ForegroundColor Green
        }
        else {
            Write-Host "[*] No .exe files found to delete" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "[!] Error deleting .exe files. Error: $_" -ForegroundColor Red
    }
}

function Remove-IsoFiles {
    Write-Host "[*] Deleting all .iso files in the Downloads folder..." -ForegroundColor Cyan
    $deletedFiles = @()
    try {
        $exeFiles = Get-ChildItem -Path $downloadsPath -Filter "*.iso"
        foreach ($file in $exeFiles) {
            Remove-Item -Path $file.FullName -Force
            $deletedFiles += $file.Name
        }
        
        if ($deletedFiles.Count -gt 0) {
            Write-Host "[+] Deleted iso files:" -ForegroundColor Green
            foreach ($fileName in $deletedFiles) {
                Write-Host "   - $fileName" -ForegroundColor Green
            }
            Write-Host "[*] A total of $($deletedFiles.Count) iso files were deleted" -ForegroundColor Green
        }
        else {
            Write-Host "[*] No .iso files found to delete" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "[!] Error deleting .iso files. Error: $_" -ForegroundColor Red
    }
}

# List of browser processes to close
$browsers = @(
    "chrome", 
    "firefox", 
    "msedge"
)

# Close all browsers
function Close-Browsers {
    Write-Host "Closing all browsers..."
    foreach ($browser in $browsers) {
        try {
            Get-Process -Name $browser -ErrorAction SilentlyContinue | Stop-Process -Force
            Write-Host "Browser $browser closed successfully."
        }
        catch {
            Write-Host "Unable to close $browser."
        }
    }
}

# Clear browsing history and data
function Clear-BrowserData {
    Write-Host "Clearing browsing data..."
    
    $chromePaths = @(
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cookies",
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\History*"
    )

    $firefoxPaths = @(
        "$env:APPDATA\Mozilla\Firefox\Profiles\*.default\cache2",
        "$env:APPDATA\Mozilla\Firefox\Profiles\*.default\cookies.sqlite",
        "$env:APPDATA\Mozilla\Firefox\Profiles\*.default\places.sqlite"
    )

    $edgePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cookies",
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\History"
    )

    foreach ($path in ($chromePaths + $firefoxPaths + $edgePaths)) {
        try {
            Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
        }
        catch {
            Write-Host "Unable to delete $path"
        }
    }

    Write-Host "Browsing data cleared."
}

# Check and install only Windows security updates
function Update-WindowsSecurity {
    Write-Host "[*] Checking for Windows security updates..." -ForegroundColor Cyan
    
    Import-Module PSWindowsUpdate -ErrorAction Stop

    try {
        $securityUpdates = Get-WindowsUpdate -Category "Security Updates"
        
        if ($securityUpdates) {
            Write-Host "[*] Security updates found:" -ForegroundColor Green
            foreach ($update in $securityUpdates) {
                Write-Host "   - $($update.Title)" -ForegroundColor Green
            }
            
            $installedUpdates = Get-WindowsUpdate -Category "Security Updates" -Install -AcceptAll -AutoReboot
            
            Write-Host "[+] Security updates installed:" -ForegroundColor Green
            foreach ($update in $installedUpdates) {
                Write-Host "   - $($update.Title)" -ForegroundColor Green
            }
            Write-Host "[*] $($installedUpdates.Count) security updates were installed" -ForegroundColor Green
        }
        else {
            Write-Host "[*] No security updates are available" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "[!] Security update installation error. Please check that the PSWindowsUpdate module is installed." -ForegroundColor Red
    }
}

function Clear-RecycleBinSafely {
    try {
        # Built-in command to empty the recycle bin
        Clear-RecycleBin -Confirm:$false
        Write-Host "[+] Recycle bin emptied successfully"
    }
    catch {
        Write-Host "[!] Unable to empty recycle bin"
    }
}

function Clean-Disk {
    # Disk cleaning
    Write-Output "Disk cleaning in progress..."

    try {
        # Deleting the temporary files in Windows
        Get-ChildItem -Path "C:\Windows\Temp\*" -Recurse -Force | ForEach-Object {
            try {
                Remove-Item $_.FullName -Force -ErrorAction Stop
                Write-Output "[+] Deleted : $($_.FullName)"
            }
            catch {
                Write-Output "[!] Unable to delete : $($_.FullName). The file may be in use."
            }
        }

        # Deleting the temporary files for the user
        Get-ChildItem -Path "C:\Users\$env:USERNAME\AppData\Local\Temp\*" -Recurse -Force | ForEach-Object {
            try {
                Remove-Item $_.FullName -Force -ErrorAction Stop
                Write-Output "[+] Deleted : $($_.FullName)"
            }
            catch {
                Write-Output "[!] Unable to delete : $($_.FullName). The file may be in use."
            }
        }

        Write-Output "Disk cleaning completed."
    }
    catch {
        Write-Output "[!] A general error occurred during the cleaning : $_"
    }
}

# Scan disk
function Scan {
    try {
        sfc /scannow 
    }
    catch {
        Write-Host "[!] Unable to scan disk"
    }
}

function Quick-DefenderScan { 
    Write-Output "Performing a quick antivirus scan..." 
    try { 
        # Using MpCmdRun.exe to perform a quick antivirus scan
        & "C:\Program Files\Windows Defender\MpCmdRun.exe" -Scan -ScanType 1 
        Write-Output "[+] Quick antivirus scan completed." 
        } 
    catch 
    { 
        Write-Output "[!] An error occurred during the antivirus scanning process : $_"
    } 
}

Write-Output "All the maintenance tasks are finished."

# Main script - Execute functions in order
Manage-RestorePoints
Remove-ExeFiles
Remove-IsoFiles
Close-Browsers
Clear-BrowserData
Clear-RecycleBinSafely
Clean-Disk
Update-WindowsSecurity
Scan
Quick-DefenderScan
