# Enhanced WinCare Maintenance Tool
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create form with improved styling
$form = New-Object System.Windows.Forms.Form
$form.Text = "WinCare : The Advanced Windows Maintenance Tool"  # Set the title of the form
$form.Size = New-Object System.Drawing.Size(900, 850)  # Define the size of the form
$form.StartPosition = "CenterScreen"  # Center the form on the screen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle  # Set a fixed border style
$form.MaximizeBox = $false  # Disable the maximize button
$form.BackColor = [System.Drawing.Color]::FromArgb(18, 18, 18)  # Set the background color
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")  # Use PowerShell icon

# Create main panel with improved layout
$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Size = New-Object System.Drawing.Size(860, 780)  # Set the size of the main panel
$mainPanel.Location = New-Object System.Drawing.Point(20, 20)  # Position the main panel within the form
$mainPanel.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)  # Set the background color of the main panel

# Add title label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "System Maintenance Optimization"  # Set the text of the label
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)  # Set the font style and size
$titleLabel.ForeColor = [System.Drawing.Color]::White  # Set the font color
$titleLabel.AutoSize = $true  # Allow the label to resize automatically
$titleLabel.Location = New-Object System.Drawing.Point(20, 20)  # Position the label within the panel
$mainPanel.Controls.Add($titleLabel)  # Add the label to the main panel

# Load and display the PNG image
$imagePath = ".\hero2.png"  # Define the path to the image
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Image = [System.Drawing.Image]::FromFile($imagePath)  # Load the image
$pictureBox.Size = New-Object System.Drawing.Size(560, 600)  # Set the size of the picture box
$pictureBox.Location = New-Object System.Drawing.Point(280, 60)  # Position the picture box within the panel
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage  # Stretch the image to fit
$mainPanel.Controls.Add($pictureBox)  # Add the picture box to the main panel

# Create checkboxes panel with categories
$checkboxesPanel = New-Object System.Windows.Forms.Panel
$checkboxesPanel.Size = New-Object System.Drawing.Size(820, 600)  # Set the size of the checkboxes panel
$checkboxesPanel.Location = New-Object System.Drawing.Point(20, 60)  # Position the checkboxes panel within the main panel
$checkboxesPanel.BackColor = [System.Drawing.Color]::FromArgb(35, 35, 35)  # Set the background color of the checkboxes panel

# Custom font and colors
$categoryFont = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)  # Define the font for category labels
$controlFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)  # Define the font for controls
$checkboxControls = @{}  # Initialize a hashtable for checkbox controls
$yPos = 20  # Initialize the vertical position for checkboxes


# Define categories and checkboxes
$categories = @{
    "Security" = @{
        "UpdateWindows" = "Install Security Updates"
        "AntivirusScan" = "Run Antivirus Scan"
        "RestorePoint" = "Create System Restore Point"
    }
    "Performance" = @{
        "AnalyseDiskUsage" = "Analyze Disk Space Usage"
        "DefragDisk" = "Smart Disk Defragmentation"
        "AnalyzeResourceUsage" = "Monitor System Performance"
    }
    "Maintenance" = @{
        "ScanDisk" = "Run System File Check"
        "CheckServices" = "Verify Critical Services"
        "UnusedPrograms" = "Remove Unused Programs"
    }
    "Cleanup" = @{
        "CloseBrowsers" = "Clear Browser Data"
        "CleanupTemp" = "Remove Temporary Files"
    }
}

$tooltips = @{
    "Security" = @{
        "UpdateWindows" = "Installs the latest Windows security updates to protect your system against known vulnerabilities. Requires an Internet connection."
        "AntivirusScan" = "Runs a quick scan with Windows Defender to detect potential threats and malware on your system."
        "RestorePoint" = "Creates a system restore point to revert to a stable state in case of issues after modifications."
    }
    "Performance" = @{
        "AnalyseDiskUsage" = "Analyzes disk space usage and identifies large files and folders to free up space."
        "DefragDisk" = "Optimizes file organization on the hard drive to improve performance (automatically ignores SSDs)."
        "AnalyzeResourceUsage" = "Monitors CPU, memory, and I/O usage to identify resource-hungry processes."
    }
    "Maintenance" = @{
        "ScanDisk" = "Checks the integrity of Windows system files and repairs corrupted files (sfc /scannow)."
        "CheckServices" = "Verifies and starts essential Windows services if necessary for optimal system operation."
        "UnusedPrograms" = "Identifies rarely used programs and suggests uninstalling them to free up space."
    }
    "Cleanup" = @{
        "CloseBrowsers" = "Closes browsers and clears temporary browsing data (history, cache, cookies)."
        "CleanupTemp" = "Removes temporary system and application files to free up disk space."
    }
}

# Define icons for sections
$icons = @{
    "Security" = ".\shield.ico"       # Shield icon for Security section
    "Performance" = ".\perf.ico"      # Icon for Performance section
    "Maintenance" = ".\handy.ico"     # Handy icon for Maintenance section
    "Cleanup" = ".\cleani.ico"        # Broom icon for Cleanup section
}

foreach ($category in $categories.Keys) {
    # Add icon for the category
    $iconPath = $icons[$category]  # Get the icon path for the current category
    $iconBox = New-Object System.Windows.Forms.PictureBox
    $iconBox.Image = [System.Drawing.Image]::FromFile($iconPath)  # Load the icon image
    $iconBox.Size = New-Object System.Drawing.Size(24, 24)  # Set the size of the icon
    $iconBox.Location = New-Object System.Drawing.Point(5, $yPos)  # Position the icon
    $iconBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage  # Set the image size mode
    $checkboxesPanel.Controls.Add($iconBox)  # Add the icon to the checkboxes panel

    # Create category label
    $categoryLabel = New-Object System.Windows.Forms.Label
    $categoryLabel.Text = $category  # Set the text of the category label
    $categoryLabel.Font = $categoryFont  # Set the font of the category label
    $categoryLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)  # Set the font color
    $categoryLabel.Location = New-Object System.Drawing.Point(40, $yPos)  # Position the category label
    $categoryLabel.AutoSize = $true  # Allow the label to resize automatically
    $checkboxesPanel.Controls.Add($categoryLabel)  # Add the category label to the checkboxes panel
    
    $yPos += 40  # Adjust vertical position for the next control
    
    # Create checkboxes for this category with tooltips
    foreach ($key in $categories[$category].Keys) {
        $checkbox = New-Object System.Windows.Forms.CheckBox
        $checkbox.Text = $categories[$category][$key]  # Set the text of the checkbox
        $checkbox.Location = New-Object System.Drawing.Point(40, $yPos)  # Position the checkbox
        $checkbox.Font = $controlFont  # Set the font of the checkbox
        $checkbox.AutoSize = $true  # Allow the checkbox to resize automatically
        $checkbox.ForeColor = [System.Drawing.Color]::White  # Set the font color
        
        # Create and configure tooltip for this checkbox
        $tooltip = New-Object System.Windows.Forms.ToolTip
        $tooltip.InitialDelay = 500  # Set the initial delay before the tooltip appears
        $tooltip.ReshowDelay = 100  # Set the delay before the tooltip reappears
        $tooltip.AutoPopDelay = 10000  # Set the duration the tooltip remains visible
        $tooltip.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)  # Set the background color of the tooltip
        $tooltip.ForeColor = [System.Drawing.Color]::White  # Set the font color of the tooltip
        $tooltip.SetToolTip($checkbox, $tooltips[$category][$key])  # Assign the tooltip to the checkbox
        
        if ($key -eq "RestorePoint") {
            $checkbox.Checked = $true  # Set the checkbox as checked if it is the "RestorePoint" checkbox
        }
        
        $checkboxesPanel.Controls.Add($checkbox)  # Add the checkbox to the checkboxes panel
        $checkboxControls[$key] = $checkbox  # Add the checkbox to the hashtable of controls
        $yPos += 30  # Adjust vertical position for the next checkbox
    }
    
    $yPos += 20  # Add extra space before the next category
}

# Functions

function Remove-TempFiles {
    try {
        # Check if the script is being run with administrative privileges
        if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
            return
        }

        # Remove temporary files with the necessary privileges
        Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue  # Delete temp files from user's temp directory
        Remove-Item -Path "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue  # Delete temp files from Windows temp directory

        return @{Success=$true; Message="Temporary files cleaned"}  # Return success message if files are cleaned
    } catch {
        return @{Success=$false; Message="Error cleaning temp files: $_"}  # Return error message if an exception occurs
    }
}

function Show-DiskUsage {

    # Analyze disk usage
    function Analyze-DiskUsage {
        param (
            [string]$Path = "C:\",  # Default path to analyze
            [int]$SizeThresholdMB = 100  # Default size threshold in MB
        )

        $sizeThresholdBytes = $SizeThresholdMB * 1MB  # Convert MB to bytes

        $largeItems = @()
        try {
            # Get all child items recursively
            Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
                if ($_.PSIsContainer) {
                    # Calculate folder size
                    $folderSize = (Get-ChildItem -Path $_.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                    if ($folderSize -gt $sizeThresholdBytes) {
                        $largeItems += [PSCustomObject]@{
                            Path = $_.FullName
                            SizeMB = [math]::Round($folderSize / 1MB, 2)
                            Type = "Directory"
                        }
                    }
                } elseif ($_.Length -gt $sizeThresholdBytes) {
                    # Add large files to the list
                    $largeItems += [PSCustomObject]@{
                        Path = $_.FullName
                        SizeMB = [math]::Round($_.Length / 1MB, 2)
                        Type = "File"
                    }
                }
            }
        } catch {
            Write-Error "An error occurred while processing items in the path: $_"
        }

        return $largeItems
    }

    # Create the user interface form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Disk Usage Analyzer"  # Set the form title
    $form.Size = New-Object System.Drawing.Size(800, 600)  # Set form size
    $form.StartPosition = "CenterScreen"  # Center the form on the screen

    # Create the list view to display results
    $listView = New-Object System.Windows.Forms.ListView
    $listView.Size = New-Object System.Drawing.Size(760, 480)  # Set list view size
    $listView.Location = New-Object System.Drawing.Point(20, 20)  # Set list view position
    $listView.View = [System.Windows.Forms.View]::Details  # Set list view to display details
    $listView.FullRowSelect = $true  # Enable full row selection
    $listView.CheckBoxes = $true  # Enable checkboxes

    # Add columns to the list view
    $listView.Columns.Add("Type", 100)
    $listView.Columns.Add("Path", 400)
    $listView.Columns.Add("Size (MB)", 100)

    # Add column sorting event
    $listView.add_ColumnClick({
        param($sender, $e)

        # Get the clicked column
        $columnIndex = $e.Column
        $listViewItems = $listView.Items
        $sortOrder = $sender.Tag

        # Determine sort order
        if ($sortOrder -eq "Ascending") {
            $sender.Tag = "Descending"
            $ascending = $false
        } else {
            $sender.Tag = "Ascending"
            $ascending = $true
        }

        # Sort list items
        $sortedItems = $listViewItems | Sort-Object -Property { $_.SubItems[$columnIndex].Text } -Descending:$ascending

        # Reset list items
        $listView.Items.Clear()
        $sortedItems | ForEach-Object { $listView.Items.Add($_) }
    })

    # Add Analyze button
    $btnAnalyze = New-Object System.Windows.Forms.Button
    $btnAnalyze.Text = "Analyze"  # Set button text
    $btnAnalyze.Size = New-Object System.Drawing.Size(100, 30)  # Set button size
    $btnAnalyze.Location = New-Object System.Drawing.Point(20, 520)  # Set button position
    $btnAnalyze.Add_Click({
        # Show folder browser dialog
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPath = $folderDialog.SelectedPath
            $listView.Items.Clear()
            $largeItems = Analyze-DiskUsage -Path $selectedPath -SizeThresholdMB 100
            foreach ($item in $largeItems) {
                $listItem = New-Object System.Windows.Forms.ListViewItem($item.Type)
                $listItem.SubItems.Add($item.Path)
                $listItem.SubItems.Add($item.SizeMB.ToString())
                $listView.Items.Add($listItem)
            }
        }
    })

    # Add View Selected button
    $btnView = New-Object System.Windows.Forms.Button
    $btnView.Text = "View Selected"  # Set button text
    $btnView.Size = New-Object System.Drawing.Size(100, 30)  # Set button size
    $btnView.Location = New-Object System.Drawing.Point(260, 520)  # Set button position
    $btnView.Add_Click({
        # Get selected items and open them
        $selectedItems = $listView.CheckedItems
        foreach ($listItem in $selectedItems) {
            $itemPath = $listItem.SubItems[1].Text
            try {
                if (Test-Path $itemPath -PathType Container) {
                    Start-Process explorer.exe -ArgumentList $itemPath  # Open folder
                } else {
                    Start-Process -FilePath $itemPath  # Open file
                }
                Write-Host "Opened: $itemPath" -ForegroundColor Blue
            } catch {
                Write-Error "Failed to open $($listItem.Text): $itemPath. Error: $_"
            }
        }
    })

    # Add Delete Selected button
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = "Delete Selected"  # Set button text
    $btnDelete.Size = New-Object System.Drawing.Size(100, 30)  # Set button size
    $btnDelete.Location = New-Object System.Drawing.Point(140, 520)  # Set button position
    $btnDelete.Add_Click({
        # Get selected items and delete them
        $selectedItems = $listView.CheckedItems
        $deletedItems = @()
        foreach ($listItem in $selectedItems) {
            $itemPath = $listItem.SubItems[1].Text
            try {
                if ($listItem.Text -eq "Directory") {
                    Remove-Item -Path $itemPath -Recurse -Force  # Delete directory
                } else {
                    Remove-Item -Path $itemPath -Force  # Delete file
                }
                $deletedItems += $itemPath
                Write-Host "$($listItem.Text) $itemPath has been deleted." -ForegroundColor Green
            } catch {
                Write-Error "Failed to delete $($listItem.Text): $itemPath. Error: $_"
            }
        }
        
        if ($deletedItems.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show("The following items have been deleted:`n" + ($deletedItems -join "`n"), "Items Deleted", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }

        $listView.Items.Clear()
    })

    $form.Controls.Add($listView)
    $form.Controls.Add($btnAnalyze)
    $form.Controls.Add($btnView)
    $form.Controls.Add($btnDelete)

    $form.ShowDialog()  # Show the form
}

function Manage-RestorePoints {
    try {
        $date = Get-Date -Format "yyyy-MM-dd"  # Get the current date in "yyyy-MM-dd" format
        $description = "Maintenance $date"  # Create a description for the restore point using the current date
        Checkpoint-Computer -Description $description -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop  # Create a system restore point
        return @{Success=$true; Message="Restore point created successfully"}  # Return a success message if the restore point is created
    } catch {
        return @{Success=$false; Message="Failed to create restore point: $_"}  # Return an error message if an exception occurs
    }
}

function Close-And-Clear-Browsers {
    try {
        $browsers = @("chrome", "firefox", "msedge")  # List of browsers to close
        $closedBrowsers = @()  # List to track closed browsers
        $clearedData = @()  # List to track cleared browser data
        
        # Close browsers
        foreach ($browser in $browsers) {
            $processes = Get-Process -Name $browser -ErrorAction SilentlyContinue  # Get browser processes
            if ($processes) {
                Stop-Process -Name $browser -Force -ErrorAction SilentlyContinue  # Forcefully stop browser processes
                $closedBrowsers += $browser  # Add closed browser to the list
            }
        }

        # Clear browser data
        $paths = @{
            'Chrome' = @(
                "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache*",  # Chrome cache
                "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cookies",  # Chrome cookies
                "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\History*"  # Chrome history
            )
            'Firefox' = @(
                "$env:APPDATA\Mozilla\Firefox\Profiles\*.default\cache2",  # Firefox cache
                "$env:APPDATA\Mozilla\Firefox\Profiles\*.default\cookies.sqlite"  # Firefox cookies
            )
            'Edge' = @(
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache*",  # Edge cache
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cookies"  # Edge cookies
            )
        }

        foreach ($browser in $paths.Keys) {
            $cleared = $false  # Flag to track cleared data
            foreach ($path in $paths[$browser]) {
                if (Test-Path $path) {
                    Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue  # Remove browser data
                    $cleared = $true
                }
            }
            if ($cleared) {
                $clearedData += $browser  # Add browser to the cleared data list
            }
        }

        $message = ""
        if ($closedBrowsers.Count -gt 0) {
            $message += "Closed browsers: $($closedBrowsers -join ', '). `n"  # List closed browsers
        }
        if ($clearedData.Count -gt 0) {
            $message += "Cleared data for: $($clearedData -join ', ')."  # List cleared data
        }
        if ($message -eq "") {
            $message = "No browsers found to clean."  # Message if no browsers found
        }

        return @{Success=$true; Message=$message}  # Return success message
    } catch {
        return @{Success=$false; Message="Browser cleanup error: $_"}  # Return error message if an exception occurs
    }
}

function Update-WindowsSecurity {
    # Check for administrator privileges
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin) {
        Write-Host "[!] This function requires administrator privileges" -ForegroundColor Red
        return @{Success=$false; Message="Administrator privileges required"}
    }

    # Save current execution policy
    $originalExecutionPolicy = Get-ExecutionPolicy
    
    try {
        # Temporarily set execution policy to Bypass for this session
        Write-Host "[*] Configuring temporary execution policy..." -ForegroundColor Cyan
        Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
        
        Write-Host "[*] Checking for Windows security updates..." -ForegroundColor Cyan
        
        # Check if PSWindowsUpdate module is installed
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-Host "[*] Installing PSWindowsUpdate module..." -ForegroundColor Yellow
            Install-Module PSWindowsUpdate -Force -AllowClobber
        }
        
        Import-Module PSWindowsUpdate -ErrorAction Stop  # Import the PSWindowsUpdate module
        
        # Search for security updates
        $securityUpdates = Get-WindowsUpdate -Category "Security Updates" -MicrosoftUpdate
        
        if ($securityUpdates) {
            Write-Host "[*] Security updates found:" -ForegroundColor Green
            foreach ($update in $securityUpdates) {
                Write-Host "   - $($update.Title)" -ForegroundColor Green  # List found security updates
            }
            
            # Install security updates
            $installedUpdates = Get-WindowsUpdate -Install -AcceptAll -AutoReboot
            
            Write-Host "[+] Security updates installed:" -ForegroundColor Green
            foreach ($update in $installedUpdates) {
                Write-Host "   - $($update.Title)" -ForegroundColor Green  # List installed security updates
            }
            Write-Host "[*] $($installedUpdates.Count) security updates have been installed" -ForegroundColor Green
            return @{Success=$true; Message="Security updates installed successfully"}
        }
        else {
            Write-Host "[*] No security updates available" -ForegroundColor Yellow
            return @{Success=$true; Message="No security updates available"}
        }
    }
    catch {
        Write-Host "[!] Error installing updates: $_" -ForegroundColor Red
        return @{Success=$false; Message="Error installing updates: $_"}
    }
    finally {
        # Restore original execution policy
        Set-ExecutionPolicy -ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
        Write-Host "[*] Restored original execution policy" -ForegroundColor Cyan
    }
}

function Start-SmartDefrag {
    try {
        # Check for administrator privileges
        if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            return @{
                Success = $false
                Message = "This script requires administrator privileges."
            }
        }

        # Get all physical disks
        $disks = Get-PhysicalDisk

        foreach ($disk in $disks) {
            # Get drive letter for this physical disk
            $partition = Get-Disk | 
                Where-Object { $_.Number -eq $disk.DeviceId } | 
                Get-Partition | 
                Where-Object { $_.Type -eq "Basic" } |
                Get-Volume |
                Where-Object { $null -ne $_.DriveLetter }

            if ($null -eq $partition) {
                continue
            }

            # Check if it's an SSD or NVMe
            $isSSD = $disk.MediaType -eq "SSD" -or $disk.BusType -eq "NVMe"

            foreach ($drive in $partition) {
                $driveLetter = $drive.DriveLetter
                if ($isSSD) {
                    Write-Host "Drive ${driveLetter}: is an SSD/NVMe - Defragmentation not recommended"
                    continue
                }

                Write-Host "Analyzing fragmentation on drive ${driveLetter}:..."
                $defragAnalysis = Get-Volume -DriveLetter $driveLetter | Get-FileSystemFragmentation

                if ($defragAnalysis.FragmentationPercentage -gt 10) {
                    Write-Host "Defragmenting drive ${driveLetter}: (Fragmentation: $($defragAnalysis.FragmentationPercentage)%)"
                    Optimize-Volume -DriveLetter $driveLetter -Defrag -Verbose
                }
                else {
                    Write-Host "Drive ${driveLetter}: doesn't need defragmentation (Fragmentation: $($defragAnalysis.FragmentationPercentage)%)"
                }
            }
        }

        return @{
            Success = $true
            Message = "Defragmentation process completed."
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Error during defragmentation: $_"
        }
    }
}

function Scan-Disk {
    try {
        Start-Process "sfc" -ArgumentList "/scannow" -NoNewWindow -Wait  # Run system file checker
        return @{Success=$true; Message="Disk scan completed"}  # Return success message
    } catch {
        return @{Success=$false; Message="Error during disk scan: $_"}  # Return error message if an exception occurs
    }
}

function Quick-DefenderScan {
    try {
        $defenderPath = "${env:ProgramFiles}\Windows Defender\MpCmdRun.exe"  # Path to Windows Defender executable
        if (Test-Path $defenderPath) {
            Start-Process $defenderPath -ArgumentList "-Scan -ScanType 1" -NoNewWindow -Wait  # Run quick scan
            return @{Success=$true; Message="Windows Defender scan completed"}  # Return success message
        }
        return @{Success=$false; Message="Windows Defender not found"}  # Return error message if Defender not found
    } catch {
        return @{Success=$false; Message="Error during antivirus scan: $_"}  # Return error message if an exception occurs
    }
}

function Check-And-Start-Services {
    # Define critical services to check
    $servicesToCheck = @(
        "wuauserv",         # Windows Update
        "bits",             # Background Intelligent Transfer Service
        "WinDefend",        # Windows Defender
        "EventLog",         # Windows Event Log
        "Dhcp",             # DHCP Client
        "Dnscache",         # DNS Client
        "PlugPlay",         # Plug and Play
        "Spooler",          # Print Spooler
        "LanmanServer",     # Server
        "LanmanWorkstation",# Workstation
        "CryptSvc",         # Cryptographic Services
        "DiagTrack",        # Connected User Experiences and Telemetry
        "WerSvc",           # Windows Error Reporting Service
        "W32Time",          # Windows Time
        "IKEEXT",           # IKE and AuthIP IPsec Keying Modules
        "MSMQ",             # Message Queuing
        "Netlogon",         # Netlogon
        "TermService",      # Remote Desktop Services
        "Power",            # Power
        "SessionEnv",       # Remote Desktop Configuration
        "SENS",             # System Event Notification Service
        "SysMain",          # Superfetch
        "Themes",           # Themes
        "TrkWks",           # Distributed Link Tracking Client
        "WSearch",          # Windows Search
        "wscsvc",           # Security Center
        "Winmgmt",          # Windows Management Instrumentation
        "NlaSvc",           # Network Location Awareness
        "lmhosts",          # TCP/IP NetBIOS Helper
        "MpsSvc",           # Windows Firewall
        "Netman",           # Network Connections
        "ProfSvc",          # User Profile Service
        "Schedule",         # Task Scheduler
        "ShellHWDetection", # Shell Hardware Detection
        "wuauserv"          # Windows Update (repeated to ensure it is covered)
    )

    foreach ($serviceName in $servicesToCheck) {
        try {
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue  # Get the service
            if ($null -eq $service) {
                Write-Host "Service $serviceName not found." -ForegroundColor Red  # Service not found
                return $false
            } elseif ($service.Status -ne 'Running') {
                Write-Host "Service $serviceName is not running. Attempting to start..." -ForegroundColor Yellow  # Service not running
                Start-Service -Name $serviceName  # Attempt to start the service
                Write-Host "Service $serviceName started successfully." -ForegroundColor Green  # Service started successfully
                return $true
            } else {
                Write-Host "Service $serviceName is already running." -ForegroundColor Green  # Service already running
                return $true
            }
        } catch {
            Write-Host "Failed to start service $serviceName. Error: $_" -ForegroundColor Red  # Error starting service
            return $false
        }
    }
}

function Get-UnusedPrograms {
    try {
        # Get all installed programs
        $installedPrograms = Get-WmiObject -Class Win32_Product | 
            Select-Object Name, Version, InstallDate, Vendor

        # Get current date
        $currentDate = Get-Date

        # Track unused programs
        $unusedPrograms = @()

        foreach ($program in $installedPrograms) {
            # Get last used date from registry
            $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths\$($program.Name).exe"
            $lastUsed = $null

            # Check program usage through different methods
            try {
                # Method 1: Check Program Files directories for last write time
                $programPath = @(
                    "${env:ProgramFiles}\$($program.Name)",
                    "${env:ProgramFiles(x86)}\$($program.Name)"
                )
                
                foreach ($path in $programPath) {
                    if (Test-Path $path) {
                        $lastWriteTime = (Get-Item $path).LastWriteTime
                        if ($null -eq $lastUsed -or $lastWriteTime -gt $lastUsed) {
                            $lastUsed = $lastWriteTime
                        }
                    }
                }

                # Method 2: Check registry last used time
                if (Test-Path $regPath) {
                    $regLastUsed = (Get-Item $regPath).LastWriteTime
                    if ($null -eq $lastUsed -or $regLastUsed -gt $lastUsed) {
                        $lastUsed = $regLastUsed
                    }
                }

                # If no last used date found, use install date
                if ($null -eq $lastUsed -and $program.InstallDate) {
                    $lastUsed = [DateTime]::ParseExact($program.InstallDate, "yyyyMMdd", $null)
                }

                # If program hasn't been used in 1 month, add to list
                if ($null -ne $lastUsed) {
                    $monthsSinceUsed = (($currentDate - $lastUsed).Days / 30)
                    if ($monthsSinceUsed -gt 1) {
                        $unusedPrograms += [PSCustomObject]@{
                            Name = $program.Name
                            Vendor = $program.Vendor
                            Version = $program.Version
                            LastUsed = $lastUsed
                            MonthsUnused = [math]::Round($monthsSinceUsed, 1)
                        }
                    }
                }
            }
            catch {
                Write-Host "Error processing $($program.Name): $_" -ForegroundColor Yellow
                continue
            }
        }

        # Display unused programs and ask for action
        if ($unusedPrograms.Count -eq 0) {
            return @{
                Success = $true
                Message = "No unused programs found"
            }
        }

        # Create and show form for program selection
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Unused Programs"  # Set the form title
        $form.Size = New-Object System.Drawing.Size(800, 600)  # Set the form size
        $form.StartPosition = "CenterScreen"  # Center the form on the screen

        $listView = New-Object System.Windows.Forms.ListView
        $listView.Size = New-Object System.Drawing.Size(760, 480)  # Set the list view size
        $listView.Location = New-Object System.Drawing.Point(20, 20)  # Set the list view position
        $listView.View = [System.Windows.Forms.View]::Details  # Set the list view to display details
        $listView.CheckBoxes = $true  # Enable checkboxes
        $listView.FullRowSelect = $true  # Enable full row selection

        # Add columns to the list view
        $listView.Columns.Add("Program", 250)
        $listView.Columns.Add("Vendor", 200)
        $listView.Columns.Add("Version", 100)
        $listView.Columns.Add("Months Unused", 100)

        foreach ($program in $unusedPrograms) {
            $item = New-Object System.Windows.Forms.ListViewItem($program.Name)
            $item.SubItems.Add($program.Vendor)
            $item.SubItems.Add($program.Version)
            $item.SubItems.Add($program.MonthsUnused.ToString())
            $listView.Items.Add($item)
        }

        $btnUninstall = New-Object System.Windows.Forms.Button
        $btnUninstall.Location = New-Object System.Drawing.Point(20, 520)  # Set the button position
        $btnUninstall.Size = New-Object System.Drawing.Size(150, 30)  # Set the button size
        $btnUninstall.Text = "Uninstall Selected"  # Set the button text
        $btnUninstall.Add_Click({
            $selectedPrograms = $listView.CheckedItems
            foreach ($item in $selectedPrograms) {
                $programName = $item.Text
                try {
                    $program = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $programName }
                    if ($program) {
                        $program.Uninstall()  # Uninstall the selected program
                        Write-Host "Uninstalled: $programName" -ForegroundColor Green
                    }
                }
                catch {
                    Write-Host "Error uninstalling $programName : $_" -ForegroundColor Red
                }
            }
            $form.Close()  # Close the form
        })

        $form.Controls.Add($listView)
        $form.Controls.Add($btnUninstall)
        $form.ShowDialog()  # Show the form

        return @{
            Success = $true
            Message = "Program cleanup completed"
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Error during program cleanup: $_"
        }
    }
}

function Analyze-ResourceUsage {
    $duration = 600  # Duration to monitor in seconds (10 minutes)
    $interval = 5    # Interval between samples in seconds

    # Create a dictionary to store resource usage data
    $resourceData = @{}

    Write-Host "Analyzing resource usage for $duration seconds. Please wait..."

    for ($i = 0; $i -lt ($duration / $interval); $i++) {
        # Get the list of processes and their resource usage
        $processes = Get-Process | Select-Object Id, ProcessName, CPU, WorkingSet, @{Name='IO'; Expression = { $_.IOReadOperations + $_.IOWriteOperations }}

        foreach ($process in $processes) {
            if (-not $resourceData.ContainsKey($process.Id)) {
                $resourceData[$process.Id] = @{
                    ProcessName = $process.ProcessName
                    CPU = 0
                    Memory = 0
                    IO = 0
                    Samples = 0
                }
            }

            $resourceData[$process.Id].CPU += $process.CPU
            $resourceData[$process.Id].Memory += $process.WorkingSet
            $resourceData[$process.Id].IO += $process.IO
            $resourceData[$process.Id].Samples++
        }

        Start-Sleep -Seconds $interval  # Wait for the specified interval before taking the next sample
    }

    # Calculate average resource usage
    $averageData = $resourceData.Values | ForEach-Object {
        [PSCustomObject]@{
            ProcessName = $_.ProcessName
            AverageCPU = $_.CPU / $_.Samples
            AverageMemory = $_.Memory / $_.Samples
            AverageIO = $_.IO / $_.Samples
        }
    }

    # Determine the threshold for high resource usage
    $highCpuThreshold = 15  # CPU usage in percentage
    $highMemoryThreshold = 200MB  # Memory usage in megabytes
    $highIoThreshold = 1000  # IO operations

    # Filter processes with high resource usage
    $highUsageProcesses = $averageData | Where-Object {
        $_.AverageCPU -gt $highCpuThreshold -or
        $_.AverageMemory -gt ($highMemoryThreshold * 1MB) -or
        $_.AverageIO -gt $highIoThreshold
    }

    # Save results to a file
    $filePath = ".\performance-results.txt"
    $output = @()

    $output += "Average resource usage of all processes:"
    $output += $averageData | Format-Table -Property ProcessName, AverageCPU, AverageMemory, AverageIO -AutoSize | Out-String

    if ($highUsageProcesses.Count -eq 0) {
        $output += "No processes with high resource usage detected."
    } else {
        $output += "Processes with high resource usage:"
        $output += $highUsageProcesses | Format-Table -Property ProcessName, AverageCPU, AverageMemory, AverageIO -AutoSize | Out-String
    }

    $output | Out-File -FilePath $filePath  # Save the output to a file

    Write-Host "Results saved to $filePath"  # Notify the user that the results have been saved
}

# Create progress panel with improved visuals
$progressPanel = New-Object System.Windows.Forms.Panel
$progressPanel.Size = New-Object System.Drawing.Size(820, 80)  # Set the size of the progress panel
$progressPanel.Location = New-Object System.Drawing.Point(20, 670)  # Set the position of the progress panel
$progressPanel.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)  # Set the background color of the progress panel

# Progress label with better styling
$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Text = "Ready to start maintenance"  # Set the initial text of the progress label
$progressLabel.Font = $controlFont  # Set the font of the progress label
$progressLabel.ForeColor = [System.Drawing.Color]::White  # Set the font color of the progress label
$progressLabel.AutoSize = $true  # Allow the label to resize automatically
$progressLabel.Location = New-Object System.Drawing.Point(20, 6)  # Set the position of the progress label
$progressPanel.Controls.Add($progressLabel)  # Add the progress label to the progress panel

# Create Run button with modern styling
$buttonRun = New-Object System.Windows.Forms.Button
$buttonRun.Text = "Start Maintenance"  # Set the text of the button
$buttonRun.Size = New-Object System.Drawing.Size(200, 40)  # Set the size of the button
$buttonRun.Location = New-Object System.Drawing.Point(310, 20)  # Set the position of the button
$buttonRun.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)  # Set the font of the button
$buttonRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)  # Set the background color of the button
$buttonRun.ForeColor = [System.Drawing.Color]::White  # Set the font color of the button
$buttonRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat  # Set the button style to flat
$buttonRun.Cursor = [System.Windows.Forms.Cursors]::Hand  # Set the cursor to hand on hover

# Add hover effects
$buttonRun.Add_MouseEnter({
    $this.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)  # Change background color on hover
})
$buttonRun.Add_MouseLeave({
    $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)  # Revert background color when not hovering
})

$progressPanel.Controls.Add($buttonRun)  # Add the Run button to the progress panel

# Add panels to form
$mainPanel.Controls.Add($checkboxesPanel)  # Add the checkboxes panel to the main panel
$mainPanel.Controls.Add($progressPanel)  # Add the progress panel to the main panel
$form.Controls.Add($mainPanel)  # Add the main panel to the form

# Enhanced task execution with progress tracking
$buttonRun.Add_Click({
    $buttonRun.Enabled = $false  # Disable the button to prevent multiple clicks
    $selectedTasks = $checkboxControls.Keys | Where-Object { $checkboxControls[$_].Checked }  # Get selected tasks
    
    # Reorganize to ensure RestorePoint is first
    $orderedTasks = @()
    if ($selectedTasks -contains "RestorePoint") {
        $orderedTasks += "RestorePoint"
        $selectedTasks = $selectedTasks | Where-Object { $_ -ne "RestorePoint" }
    }
    $orderedTasks += $selectedTasks
    
    $totalTasks = $orderedTasks.Count  # Get the total number of tasks
    $currentTask = 0  # Initialize current task counter
    
    $logPath = Join-Path $env:USERPROFILE "WinCare_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"  # Define log file path
    "WinCare Maintenance Log - $(Get-Date)" | Out-File $logPath  # Write log header
    
    foreach ($task in $orderedTasks) {
        # Loop through each task and execute
        $currentTask++
        $progress = [math]::Round(($currentTask / $totalTasks) * 100)  # Calculate progress percentage
        $progressLabel.Text = "Running: $task ($progress%)"  # Update progress label
        $progressLabel.Refresh()  # Refresh label to show updated text
        
        try {
            $result = switch($task) {
                "UpdateWindows" { Update-WindowsSecurity }
                "CloseBrowsers" { Close-And-Clear-Browsers }
                "AnalyseDiskUsage" { Show-DiskUsage }
                "ScanDisk" { Scan-Disk }
                "AntivirusScan" { Quick-DefenderScan }
                "CloudProtection" { Enable-CloudProtection }
                "CheckServices" { Check-And-Start-Services }
                "DefragDisk" { Start-SmartDefrag }
                "UnusedPrograms" { Get-UnusedPrograms }
                "AnalyzeResourceUsage" { Analyze-ResourceUsage }
                "RestorePoint" { Manage-RestorePoints }
                "CleanupTemp" { Remove-TempFiles }
            }
            "[$task] - $(Get-Date)" | Out-File $logPath -Append  # Log task start time
            "Result: $($result.Message)" | Out-File $logPath -Append  # Log task result message
            "------------------" | Out-File $logPath -Append  # Log separator
        } catch {
            "[$task] - ERROR: $_" | Out-File $logPath -Append  # Log error message if exception occurs
            "------------------" | Out-File $logPath -Append  # Log separator
        }
    }
    
    $progressLabel.Text = "Maintenance completed. Log saved to: $logPath"  # Update progress label on completion
    $buttonRun.Enabled = $true  # Re-enable the button
    $form.Refresh()  # Refresh the form to update UI
})

# Show form
[Windows.Forms.Application]::Run($form)  # Run the form application
