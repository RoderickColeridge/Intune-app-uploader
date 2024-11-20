######################################### 
# Intune App Uploader
# Created by: Roderick Coleridge
# Version: 1.0.0
# Date: 26-10-2024
#########################################
# Version 1.0.1
# Added group assignment to the app upload  
# Date: 01-11-2024
#########################################
# Version 1.0.2
# Fixed an issue with detection script 
# Date: 04-11-2024
#########################################
# Version 1.0.3
# Added auto update function
# Date: 05-11-2024
#########################################
# Version 1.0.4
# Added Fixed search function
# Date: 06-11-2024
#########################################
# Version 1.0.5
# Added auto admin consent
# Date: 06-11-2024
#########################################
# Version 1.0.6
# Added Altered Log Directory
# Date: 14-11-2024
#########################################
# Version 1.0.7
# Added proactive remediation script per app to check for updates on a daily base
# Date: 19-11-2024
#########################################
# Version 1.0.8
# Updated Del Credentials button to remove all credentials and azure ad app
# Date: 20-11-2024
#########################################


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Auto-update script

# Define the URL of your script repository
$repoUrl = "https://wingetscript.blob.core.windows.net/mainscript/Winget_Apps.ps1"
$versionFileUrl = "https://wingetscript.blob.core.windows.net/mainscript/version.txt"

# Current version of the script
$currentVersion = "1.0.8"

# Get the directory of the current script
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$localScriptPath = Join-Path -Path $scriptRoot -ChildPath "Winget_Apps.ps1"

# Function to check for updates
function Check-ForUpdates {
    # Get the remote version
    $remoteVersion = (Invoke-WebRequest -Uri $versionFileUrl).Content.Trim()

    if ($currentVersion -ne $remoteVersion) {
        Write-Output "New version available. Updating script..."
        # Download the new script
        Invoke-WebRequest -Uri $repoUrl -OutFile $localScriptPath
        # Update the current version variable
        $script:currentVersion = $remoteVersion
        Write-Output "Script updated to version $remoteVersion"
    } else {
        Write-Output "Script is up to date."
    }
}

# Call the update function
Check-ForUpdates

# Define general variables
# Get script directory
$scriptDirectory = if ($PSScriptRoot) { 
    $PSScriptRoot 
} elseif ($psISE) { 
    Split-Path -Parent $psISE.CurrentFile.FullPath 
} else { 
    Split-Path -Parent $MyInvocation.MyCommand.Path 
}

# Define log directory and files
$logDirectory = Join-Path $scriptDirectory "Logs"
$LogFilePath = Join-Path $logDirectory "IntuneUploadLog.txt"

# Create log directory if it doesn't exist
if (-not (Test-Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

$IntuneWinAppUtilUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/refs/heads/master/IntuneWinAppUtil.exe"
$IntuneWinAppUtilPath = "C:\IntunePackages\IntuneWinAppUtil.exe"
$TempDir = "C:\IntunePackages"

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Intune App Uploader"
$form.Size = New-Object System.Drawing.Size(800,600)
$form.StartPosition = "CenterScreen"

# Create ListView for apps
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10,10)
$listView.Size = New-Object System.Drawing.Size(760,200)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.Columns.Add("Display Name", 200)
$listView.Columns.Add("Winget ID", 200)
$listView.Columns.Add("Publisher", 200)
$form.Controls.Add($listView)

# Create buttons
$appRegButton = New-Object System.Windows.Forms.Button
$appRegButton.Location = New-Object System.Drawing.Point(10,220)
$appRegButton.Size = New-Object System.Drawing.Size(75,23)
$appRegButton.Text = "App Reg"
$form.Controls.Add($appRegButton)

$removeCredButton = New-Object System.Windows.Forms.Button
$removeCredButton.Location = New-Object System.Drawing.Point(90,220)
$removeCredButton.Size = New-Object System.Drawing.Size(100,23)
$removeCredButton.Text = "Del Credentials"
$form.Controls.Add($removeCredButton)

$addButton = New-Object System.Windows.Forms.Button
$addButton.Location = New-Object System.Drawing.Point(195,220)
$addButton.Size = New-Object System.Drawing.Size(75,23)
$addButton.Text = "Add"
$form.Controls.Add($addButton)

$editButton = New-Object System.Windows.Forms.Button
$editButton.Location = New-Object System.Drawing.Point(275,220)
$editButton.Size = New-Object System.Drawing.Size(75,23)
$editButton.Text = "Edit"
$form.Controls.Add($editButton)

$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Location = New-Object System.Drawing.Point(355,220)
$removeButton.Size = New-Object System.Drawing.Size(75,23)
$removeButton.Text = "Remove"
$form.Controls.Add($removeButton)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Location = New-Object System.Drawing.Point(695,220)
$runButton.Size = New-Object System.Drawing.Size(75,23)
$runButton.Text = "Run"
$form.Controls.Add($runButton)

$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Location = New-Object System.Drawing.Point(435,220)
$searchButton.Size = New-Object System.Drawing.Size(100,23)
$searchButton.Text = "Search Winget"
$form.Controls.Add($searchButton)

# Create log text area
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10,250)
$logTextBox.Size = New-Object System.Drawing.Size(760,300)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = "Vertical"
$form.Controls.Add($logTextBox)

# Initialize config structure
$script:config = @{
    Apps = @()
    Credentials = @{
        TenantID = ""
        ClientID = ""
        ClientSecret = ""
        AppID = ""
    }
}

# Load existing config if it exists
$configPath = Join-Path -Path $scriptDirectory -ChildPath "config.json"
if (Test-Path $configPath) {
    $loadedConfig = Get-Content -Path $configPath | ConvertFrom-Json
    if ($loadedConfig.Apps) {
        $script:config.Apps = $loadedConfig.Apps
    }
    if ($loadedConfig.Credentials) {
        $script:config.Credentials = $loadedConfig.Credentials
    }
}

# Function to load apps from config
function Load-Apps {
    $listView.Items.Clear()
    
    # Try to get the script path
    if ($PSScriptRoot) {
        $scriptRoot = $PSScriptRoot
    } elseif ($psISE) {
        $scriptRoot = Split-Path -Parent $psISE.CurrentFile.FullPath
    } else {
        $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    
    # If we still don't have a valid path, use the current directory
    if (-not $scriptRoot) {
        $scriptRoot = Get-Location
    }
    
    $configPath = Join-Path -Path $scriptRoot -ChildPath "config.json"
    
    if (Test-Path $configPath) {
        try {
            $loadedConfig = Get-Content -Path $configPath | ConvertFrom-Json
            
            # Initialize the config structure if loading from an older version
            if (-not $script:config) {
                $script:config = @{
                    Apps = @()
                    Credentials = @{
                        TenantID = ""
                        ClientID = ""
                        ClientSecret = ""
                        AppId = ""
                    }
                }
            }
            
            # Copy apps from loaded config
            if ($loadedConfig.Apps) {
                $script:config.Apps = $loadedConfig.Apps
            }
            
            # Copy credentials if they exist
            if ($loadedConfig.Credentials) {
                $script:config.Credentials = $loadedConfig.Credentials
            }
            
            foreach ($app in $script:config.Apps) {
                $item = New-Object System.Windows.Forms.ListViewItem($app.DisplayName)
                $item.SubItems.Add($app.WingetId)
                $item.SubItems.Add($app.Publisher)
                $listView.Items.Add($item)
            }
            Log-Message "Loaded apps from config file successfully."
        } catch {
            Log-Message "Error loading config file: $_" "ERROR"
        }
    } else {
        Log-Message "Config file not found at $configPath" "ERROR"
    }
}

# Function to add a new app
function Add-App {
    $addForm = New-Object System.Windows.Forms.Form
    $addForm.Text = "Add New App"
    $addForm.Size = New-Object System.Drawing.Size(300,250)
    $addForm.StartPosition = "CenterScreen"

    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Location = New-Object System.Drawing.Point(10,20)
    $nameLabel.Size = New-Object System.Drawing.Size(100,20)
    $nameLabel.Text = "Display Name:"
    $addForm.Controls.Add($nameLabel)

    $nameTextBox = New-Object System.Windows.Forms.TextBox
    $nameTextBox.Location = New-Object System.Drawing.Point(120,20)
    $nameTextBox.Size = New-Object System.Drawing.Size(150,20)
    $addForm.Controls.Add($nameTextBox)

    $wingetLabel = New-Object System.Windows.Forms.Label
    $wingetLabel.Location = New-Object System.Drawing.Point(10,50)
    $wingetLabel.Size = New-Object System.Drawing.Size(100,20)
    $wingetLabel.Text = "Winget ID:"
    $addForm.Controls.Add($wingetLabel)

    $wingetTextBox = New-Object System.Windows.Forms.TextBox
    $wingetTextBox.Location = New-Object System.Drawing.Point(120,50)
    $wingetTextBox.Size = New-Object System.Drawing.Size(150,20)
    $addForm.Controls.Add($wingetTextBox)

    $publisherLabel = New-Object System.Windows.Forms.Label
    $publisherLabel.Location = New-Object System.Drawing.Point(10,80)
    $publisherLabel.Size = New-Object System.Drawing.Size(100,20)
    $publisherLabel.Text = "Publisher:"
    $addForm.Controls.Add($publisherLabel)

    $publisherTextBox = New-Object System.Windows.Forms.TextBox
    $publisherTextBox.Location = New-Object System.Drawing.Point(120,80)
    $publisherTextBox.Size = New-Object System.Drawing.Size(150,20)
    $addForm.Controls.Add($publisherTextBox)

    $descriptionLabel = New-Object System.Windows.Forms.Label
    $descriptionLabel.Location = New-Object System.Drawing.Point(10,110)
    $descriptionLabel.Size = New-Object System.Drawing.Size(100,20)
    $descriptionLabel.Text = "Description:"
    $addForm.Controls.Add($descriptionLabel)

    $descriptionTextBox = New-Object System.Windows.Forms.TextBox
    $descriptionTextBox.Location = New-Object System.Drawing.Point(120,110)
    $descriptionTextBox.Size = New-Object System.Drawing.Size(150,60)
    $descriptionTextBox.Multiline = $true
    $addForm.Controls.Add($descriptionTextBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(120,180)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $addForm.Controls.Add($okButton)

    $addForm.AcceptButton = $okButton

    $result = $addForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Try to get the script path
        if ($PSScriptRoot) {
            $scriptRoot = $PSScriptRoot
        } elseif ($psISE) {
            $scriptRoot = Split-Path -Parent $psISE.CurrentFile.FullPath
        } else {
            $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
        }
        
        # If we still don't have a valid path, use the current directory
        if (-not $scriptRoot) {
            $scriptRoot = Get-Location
        }
        
        $configPath = Join-Path -Path $scriptRoot -ChildPath "config.json"
        
        if (Test-Path $configPath) {
            $config = Get-Content -Path $configPath | ConvertFrom-Json
        } else {
            $config = @{
                Apps = @()
            }
        }

        # Create a safe file name by replacing spaces with underscores
        $safeFileName = $nameTextBox.Text -replace '\s', '_'

        # Generate install and uninstall commands using the safe file name
        $installCommand = "powershell.exe -ExecutionPolicy Bypass -File .\$safeFileName.ps1"
        $uninstallCommand = "powershell.exe -ExecutionPolicy Bypass -File .\${safeFileName}_uninstall.ps1"

        $newApp = @{
            DisplayName = $nameTextBox.Text
            WingetId = $wingetTextBox.Text
            Publisher = $publisherTextBox.Text
            Description = $descriptionTextBox.Text
            InstallCommand = $installCommand
            UninstallCommand = $uninstallCommand
        }

        if (-not $config.Apps) {
            $config | Add-Member -NotePropertyName Apps -NotePropertyValue @()
        }

        $config.Apps += $newApp
        $config | ConvertTo-Json | Set-Content -Path $configPath
        Load-Apps
        Log-Message "Added new app: $($nameTextBox.Text)"
    }
}

# Function to remove a selected app
function Remove-App {
    $selectedItems = $listView.SelectedItems
    if ($selectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select an app to remove.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $selectedApp = $selectedItems[0]
    $appName = $selectedApp.Text

    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove $appName?", "Confirm Removal", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Try to get the script path
        if ($PSScriptRoot) {
            $scriptRoot = $PSScriptRoot
        } elseif ($psISE) {
            $scriptRoot = Split-Path -Parent $psISE.CurrentFile.FullPath
        } else {
            $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
        }
        
        # If we still don't have a valid path, use the current directory
        if (-not $scriptRoot) {
            $scriptRoot = Get-Location
        }
        
        $configPath = Join-Path -Path $scriptRoot -ChildPath "config.json"
        
        if (Test-Path $configPath) {
            $config = Get-Content -Path $configPath | ConvertFrom-Json
            $config.Apps = $config.Apps | Where-Object { $_.DisplayName -ne $appName }
            $config | ConvertTo-Json | Set-Content -Path $configPath
            Load-Apps
            Log-Message "Removed app: $appName"
        } else {
            Log-Message "Config file not found. Unable to remove app." "ERROR"
        }
    }
}

function Remove-StoredCredentials {
    try {
        Log-Message "Starting credential removal process..."
        
        # Log current state (without showing sensitive data)
        $hasCredentials = -not [string]::IsNullOrEmpty($script:config.Credentials.TenantID)
        Log-Message "Current credentials exist: $hasCredentials"
        
        # Clear credentials from memory
        Log-Message "Clearing credentials from memory..."
        $script:config.Credentials.TenantID = ""
        $script:config.Credentials.ClientID = ""
        $script:config.Credentials.ClientSecret = ""
        $script:config.Credentials.AppId = ""
        Log-Message "Credentials cleared from memory"

        # Save the updated config to file
        $configPath = Join-Path -Path $scriptDirectory -ChildPath "config.json"
        Log-Message "Saving updated config to: $configPath"
        
        try {
            $script:config | ConvertTo-Json | Set-Content -Path $configPath
            Log-Message "Config file updated successfully"
        }
        catch {
            Log-Message "Error saving config file: $($_.Exception.Message)" "ERROR"
            throw
        }

        Log-Message "Credential removal completed successfully"
    }
    catch {
        Log-Message "Error removing credentials: $($_.Exception.Message)" "ERROR"
    }
}

function Remove-GraphCredentials {
    try {
        Log-Message "Starting Graph credential removal process..."
        
        try {
            # Disconnect from Graph if connected
            if (Get-MgContext) {
                Log-Message "Disconnecting from Microsoft Graph..."
                Disconnect-MgGraph
                Log-Message "Successfully disconnected from Microsoft Graph"
            }
            else {
                Log-Message "No active Microsoft Graph connection found"
            }

            # Clear token cache
            $tokenCachePath = "$env:LOCALAPPDATA\Microsoft\TokenCache"
            if (Test-Path $tokenCachePath) {
                Log-Message "Clearing token cache at: $tokenCachePath"
                Remove-Item -Path "$tokenCachePath\*" -Force -ErrorAction SilentlyContinue
                Log-Message "Token cache cleared successfully"
            }
            else {
                Log-Message "No token cache directory found at: $tokenCachePath"
            }

            # Remove stored Graph credentials from Windows Credential Manager
            Log-Message "Removing Graph credentials from Windows Credential Manager..."
            $credentialList = cmdkey /list | Select-String "Microsoft Graph PowerShell"
            foreach ($cred in $credentialList) {
                $targetName = ($cred -split "Target:\s+")[-1]
                cmdkey /delete:$targetName
                Log-Message "Removed credential: $targetName"
            }

            # Clear module cache
            Log-Message "Removing Graph PowerShell modules from session..."
            Get-Module Microsoft.Graph* | Remove-Module -Force
            Log-Message "Graph PowerShell modules removed from session"

            Log-Message "Graph credential removal completed successfully"
        }
        catch {
            Log-Message "Error during Graph credential removal: $($_.Exception.Message)" "ERROR"
            throw
        }
    }
    catch {
        Log-Message "Critical error in Graph credential removal: $($_.Exception.Message)" "ERROR"
    }
}

# Function to edit a selected app
function Edit-App {
    $selectedItems = $listView.SelectedItems
    if ($selectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select an app to edit.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $selectedApp = $selectedItems[0]
    $appName = $selectedApp.Text

    $editForm = New-Object System.Windows.Forms.Form
    $editForm.Text = "Edit App"
    $editForm.Size = New-Object System.Drawing.Size(300,250)
    $editForm.StartPosition = "CenterScreen"

    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Location = New-Object System.Drawing.Point(10,20)
    $nameLabel.Size = New-Object System.Drawing.Size(100,20)
    $nameLabel.Text = "Display Name:"
    $editForm.Controls.Add($nameLabel)

    $nameTextBox = New-Object System.Windows.Forms.TextBox
    $nameTextBox.Location = New-Object System.Drawing.Point(120,20)
    $nameTextBox.Size = New-Object System.Drawing.Size(150,20)
    $nameTextBox.Text = $selectedApp.SubItems[0].Text
    $editForm.Controls.Add($nameTextBox)

    $wingetLabel = New-Object System.Windows.Forms.Label
    $wingetLabel.Location = New-Object System.Drawing.Point(10,50)
    $wingetLabel.Size = New-Object System.Drawing.Size(100,20)
    $wingetLabel.Text = "Winget ID:"
    $editForm.Controls.Add($wingetLabel)

    $wingetTextBox = New-Object System.Windows.Forms.TextBox
    $wingetTextBox.Location = New-Object System.Drawing.Point(120,50)
    $wingetTextBox.Size = New-Object System.Drawing.Size(150,20)
    $wingetTextBox.Text = $selectedApp.SubItems[1].Text
    $editForm.Controls.Add($wingetTextBox)

    $publisherLabel = New-Object System.Windows.Forms.Label
    $publisherLabel.Location = New-Object System.Drawing.Point(10,80)
    $publisherLabel.Size = New-Object System.Drawing.Size(100,20)
    $publisherLabel.Text = "Publisher:"
    $editForm.Controls.Add($publisherLabel)

    $publisherTextBox = New-Object System.Windows.Forms.TextBox
    $publisherTextBox.Location = New-Object System.Drawing.Point(120,80)
    $publisherTextBox.Size = New-Object System.Drawing.Size(150,20)
    $publisherTextBox.Text = $selectedApp.SubItems[2].Text
    $editForm.Controls.Add($publisherTextBox)

    $descriptionLabel = New-Object System.Windows.Forms.Label
    $descriptionLabel.Location = New-Object System.Drawing.Point(10,110)
    $descriptionLabel.Size = New-Object System.Drawing.Size(100,20)
    $descriptionLabel.Text = "Description:"
    $editForm.Controls.Add($descriptionLabel)

    $descriptionTextBox = New-Object System.Windows.Forms.TextBox
    $descriptionTextBox.Location = New-Object System.Drawing.Point(120,110)
    $descriptionTextBox.Size = New-Object System.Drawing.Size(150,60)
    $descriptionTextBox.Multiline = $true
    $descriptionTextBox.Text = $script:config.Apps | Where-Object { $_.DisplayName -eq $appName } | Select-Object -ExpandProperty Description
    $editForm.Controls.Add($descriptionTextBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(120,180)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $editForm.Controls.Add($okButton)

    $editForm.AcceptButton = $okButton

    $result = $editForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Try to get the script path
        if ($PSScriptRoot) {
            $scriptRoot = $PSScriptRoot
        } elseif ($psISE) {
            $scriptRoot = Split-Path -Parent $psISE.CurrentFile.FullPath
        } else {
            $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
        }
        
        # If we still don't have a valid path, use the current directory
        if (-not $scriptRoot) {
            $scriptRoot = Get-Location
        }
        
        $configPath = Join-Path -Path $scriptRoot -ChildPath "config.json"
        
        if (Test-Path $configPath) {
            $config = Get-Content -Path $configPath | ConvertFrom-Json
            $appToEdit = $config.Apps | Where-Object { $_.DisplayName -eq $appName }
            
            if ($appToEdit) {
                $appToEdit.DisplayName = $nameTextBox.Text
                $appToEdit.WingetId = $wingetTextBox.Text
                $appToEdit.Publisher = $publisherTextBox.Text
                $appToEdit.Description = $descriptionTextBox.Text
                
                # Update install and uninstall commands if the name has changed
                if ($appToEdit.DisplayName -ne $appName) {
                    $safeFileName = $nameTextBox.Text -replace '\s', '_'
                    $appToEdit.InstallCommand = "powershell.exe -ExecutionPolicy Bypass -File .\$safeFileName.ps1"
                    $appToEdit.UninstallCommand = "powershell.exe -ExecutionPolicy Bypass -File .\${safeFileName}_uninstall.ps1"
                }
                
                $config | ConvertTo-Json | Set-Content -Path $configPath
                Load-Apps
                Log-Message "Updated app: $($nameTextBox.Text)"
            } else {
                Log-Message "App not found in config. Unable to edit." "ERROR"
            }
        } else {
            Log-Message "Config file not found. Unable to edit app." "ERROR"
        }
    }
}

# Define the Log-Message function
function Log-Message {
    param (
        [string]$Message,
        [string]$LogType = "INFO",
        [System.Management.Automation.ErrorRecord]$ErrorRecord = $null
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$LogType] - $Message"
    if ($ErrorRecord) {
        $logEntry += "`nError Details: $($ErrorRecord.Exception.Message)`nStack Trace: $($ErrorRecord.ScriptStackTrace)"
    }
    $logTextBox.AppendText("$logEntry`r`n")
    Add-Content -Path $LogFilePath -Value $logEntry
}

# Define cleanup function
function Cleanup-TempFiles {
    param (
        [string]$TempDir
    )
    try {
        Get-ChildItem -Path $TempDir -File | ForEach-Object {
            Remove-Item -Path $_.FullName -Force -ErrorAction Stop
            Log-Message "Removed temporary file: $($_.FullName)"
        }
        Remove-Item -Path $TempDir -Force -Recurse -ErrorAction Stop
        Log-Message "Removed temporary directory: $TempDir"
    } catch {
        Log-Message "Error during cleanup" "ERROR" $_
    }
}

function Register-IntuneApp {
    try {
        Log-Message "Starting app registration process..."

        # Install required modules if not present
        $requiredModules = @(
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Applications',
            'Microsoft.Graph.Identity.DirectoryManagement'
        )

        foreach ($module in $requiredModules) {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                Log-Message "Installing $module module..."
                Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
            }
            Import-Module -Name $module -Force
            Log-Message "Loaded $module module"
        }

        # Connect to Microsoft Graph with required permissions
        Log-Message "Connecting to Microsoft Graph..."
        $graphScopes = @(
            'Application.ReadWrite.All',
            'Directory.ReadWrite.All',
            'AppRoleAssignment.ReadWrite.All',
            'RoleAssignmentSchedule.ReadWrite.Directory',
            'Domain.Read.All',
            'Domain.ReadWrite.All',
            'Directory.Read.All',
            'Policy.ReadWrite.ConditionalAccess',
            'DeviceManagementApps.ReadWrite.All',
            'DeviceManagementConfiguration.ReadWrite.All',
            'DeviceManagementManagedDevices.ReadWrite.All'
            
        )
        Connect-MgGraph -Scopes $graphScopes | Out-Null

        # Create Azure AD Application
        $appName = "Winget-Intune-App-Deployment"
        Log-Message "Creating Azure AD Application: $appName"

        # Define Microsoft Graph permissions
        $msGraphSpn = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
        
        $requiredResourceAccess = @(
            @{
                ResourceAppId = $msGraphSpn.AppId
                ResourceAccess = @(
                    @{
                        Id = ($msGraphSpn.AppRoles | Where-Object { $_.Value -eq "DeviceManagementApps.ReadWrite.All" }).Id
                        Type = "Role"
                    }
                    @{
                        Id = ($msGraphSpn.AppRoles | Where-Object { $_.Value -eq "DeviceManagementConfiguration.ReadWrite.All" }).Id
                        Type = "Role"
                    }
                    @{
                        Id = ($msGraphSpn.AppRoles | Where-Object { $_.Value -eq "DeviceManagementManagedDevices.ReadWrite.All" }).Id
                        Type = "Role"
                    }
                    @{
                        Id = ($msGraphSpn.AppRoles | Where-Object { $_.Value -eq "Directory.Read.All" }).Id
                        Type = "Role"
                    }
                    @{
                        Id = ($msGraphSpn.AppRoles | Where-Object { $_.Value -eq "Group.ReadWrite.All" }).Id
                        Type = "Role"
                    }
                    @{
                        Id = ($msGraphSpn.AppRoles | Where-Object { $_.Value -eq "DeviceManagementRBAC.Read.All" }).Id
                        Type = "Role"
                    }
                    @{
                        Id = ($msGraphSpn.AppRoles | Where-Object { $_.Value -eq "DeviceManagementRBAC.ReadWrite.All" }).Id
                        Type = "Role"
                    }
                    @{
                        Id = ($msGraphSpn.AppRoles | Where-Object { $_.Value -eq "Application.ReadWrite.All" }).Id
                        Type = "Role"
                    }
                )
            }
        )

        # Create the application
        $params = @{
            DisplayName = $appName
            SignInAudience = "AzureADMyOrg"
            RequiredResourceAccess = $requiredResourceAccess
            Web = @{
                RedirectUris = @("https://login.microsoftonline.com/common/oauth2/nativeclient")
            }
        }

        $app = New-MgApplication @params
        Log-Message "Created Azure AD Application"

        # Create service principal
        Log-Message "Creating service principal..."
        $sp = New-MgServicePrincipal -AppId $app.AppId
        Log-Message "Created service principal"

        # Create client secret
        Log-Message "Creating client secret..."
        $secretEndDate = (Get-Date).AddYears(2)
        $passwordCred = @{
            displayName = "WinGet Deployment Secret"
            endDateTime = $secretEndDate
        }
        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential $passwordCred

        # Get tenant details
        $org = Get-MgOrganization
        $tenantId = $org.Id

        # Automatically grant admin consent
        Log-Message "Granting admin consent for application permissions..."
        foreach ($resource in $requiredResourceAccess) {
            foreach ($permission in $resource.ResourceAccess) {
                if ($permission.Type -eq "Role") {
                    try {
                        $appRole = $msGraphSpn.AppRoles | Where-Object { $_.Id -eq $permission.Id }
                        
                        $appRoleAssignment = @{
                            PrincipalId = $sp.Id
                            ResourceId = $msGraphSpn.Id
                            AppRoleId = $permission.Id
                        }

                        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -BodyParameter $appRoleAssignment
                        Log-Message "Granted consent for $($appRole.Value)"
                    }
                    catch {
                        Log-Message "Error granting consent for role: $($appRole.Value)" "ERROR" $_
                    }
                }
            }
        }

        # Store the values in config
        if (-not $script:config) {
            $script:config = @{
                Apps = @()
                Credentials = @{
                    TenantID = ""
                    ClientID = ""
                    ClientSecret = ""
                    AppId = ""
                }
            }
        }

        $script:config.Credentials.TenantID = $tenantId
        $script:config.Credentials.ClientID = $app.AppId
        $script:config.Credentials.ClientSecret = $secret.SecretText
        $script:config.Credentials.AppId = $app.Id

        # Save to config file
        $configPath = Join-Path -Path $scriptDirectory -ChildPath "config.json"
        $script:config | ConvertTo-Json | Set-Content -Path $configPath
        Log-Message "Saved credentials to config file"

        # Show success message
        [System.Windows.Forms.MessageBox]::Show(
            "Application registered and configured successfully!`n`n" +
            "Tenant ID: $tenantId`n" +
            "Client ID: $($app.AppId)`n" +
            "Client Secret: $($secret.SecretText)`n`n" +
            "Admin consent has been automatically granted.",
            "Registration Success",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)

    }
    catch {
        Log-Message "Error during application registration: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Error during application registration: $($_.Exception.Message)",
            "Registration Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Create-UpdateRemediationScript {
    param (
        [string]$AppName,
        [string]$WingetId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$TenantId,
        [bool]$AssignToAllDevices
    )

    try {
        Log-Message "Creating update remediation script for $AppName..."

        # Authentication
        $body = @{
            grant_type    = "client_credentials"
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = "https://graph.microsoft.com/.default"
        }
     
        $response = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body
        $accessToken = $response.access_token

        # Connect to Graph with the access token
        $version = (Get-Module Microsoft.Graph.Authentication | Select-Object -ExpandProperty Version).Major
        if ($version -eq 2) {
            $accessTokenFinal = ConvertTo-SecureString -String $accessToken -AsPlainText -Force
        } else {
            Select-MgProfile -Name Beta
            $accessTokenFinal = $accessToken
        }
        Connect-MgGraph -AccessToken $accessTokenFinal | Out-Null

        # Create detection script content
        $detect = @"
`$PackageName = "$WingetId"

`$ResolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
if (`$ResolveWingetPath) {
    `$WingetPath = `$ResolveWingetPath[-1].Path
    `$wingetExe = Join-Path -Path `$WingetPath -ChildPath "winget.exe"
    
    if (Test-Path `$wingetExe) {
        Set-Location -Path `$WingetPath
        `$upgradeResult = & `$wingetExe upgrade --id `$PackageName --accept-source-agreements --accept-package-agreements
        
        if (`$upgradeResult -match "No applicable update found") {
            Write-Output "No updates available for `$PackageName"
            Exit 0
        } else {
            Write-Output "Update available for `$PackageName"
            Exit 1
        }
    } else {
        Write-Warning "Winget executable not found"
        Exit 1
    }
} else {
    Write-Warning "Could not find Winget installation path"
    Exit 1
}
"@

        # Create remediation script content
        $remediate = @"
`$PackageName = "$WingetId"

`$ResolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
if (`$ResolveWingetPath) {
    `$WingetPath = `$ResolveWingetPath[-1].Path
    `$wingetExe = Join-Path -Path `$WingetPath -ChildPath "winget.exe"
    
    if (Test-Path `$wingetExe) {
        Set-Location -Path `$WingetPath
        & `$wingetExe upgrade --id `$PackageName --silent --accept-source-agreements --accept-package-agreements
        Exit `$LASTEXITCODE
    } else {
        Write-Warning "Winget executable not found"
        Exit 1
    }
} else {
    Write-Warning "Could not find Winget installation path"
    Exit 1
}
"@

        # Parameters for the proactive remediation
        $params = @{
            "@odata.type" = "#microsoft.graph.deviceHealthScript"
            displayName = $AppName
            description = "Checks and installs updates for $AppName using Winget"
            publisher = "Winget Auto-Update"
            runAs32Bit = $false
            runAsAccount = "system"
            enforceSignatureCheck = $false
            detectionScriptContent = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($detect))
            remediationScriptContent = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($remediate))
            roleScopeTagIds = @("0")
        }

        # Create the proactive remediation using Graph API
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/deviceHealthScripts"
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"

        # Create the remediation script
        $proactive = Invoke-MGGraphRequest -Uri $uri -Method Post -Body ($params | ConvertTo-Json -Depth 10) -ContentType "application/json"
        Log-Message "Created proactive remediation script for $AppName"

        if ($AssignToAllDevices) {
            $scheduleParams = @{
                deviceHealthScriptAssignments = @(
                    @{
                        target = @{
                            "@odata.type" = "#microsoft.graph.allDevicesAssignmentTarget"
                        }
                        runRemediationScript = $true
                        runSchedule = @{
                            "@odata.type" = "#microsoft.graph.deviceHealthScriptDailySchedule"
                            interval = 1
                            time = "13:00"
                            useUtc = $false
                        }
                    }
                )
            }

            # Assign the script
            $remediationId = $proactive.id
            $assignUri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$remediationId/assign"
            
            $assignment = Invoke-MGGraphRequest -Uri $assignUri -Method Post -Body ($scheduleParams | ConvertTo-Json -Depth 10) -ContentType "application/json"
            Log-Message "Assigned update remediation script for $AppName to all devices"
        }

        return $proactive.id
    }
    catch {
        Log-Message "Failed to create update remediation script: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Remove-AzureADApp {
    try {
        Log-Message "Starting Azure AD app removal process..."
        
        # Check if we have the app ID in the config
        if (-not $script:config.Credentials.AppId) {
            Log-Message "No App ID found in config, skipping app removal"
            return
        }

        Log-Message "Found App ID: $($script:config.Credentials.AppId)"

        # Install and import required module if not present
        if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Applications')) {
            Log-Message "Installing Microsoft.Graph.Applications module..."
            Install-Module -Name 'Microsoft.Graph.Applications' -Force -AllowClobber -Scope CurrentUser
            Log-Message "Microsoft.Graph.Applications module installed successfully"
        }
        Import-Module -Name 'Microsoft.Graph.Applications' -Force
        Log-Message "Microsoft.Graph.Applications module imported"
        

        # Connect to Microsoft Graph with required scope
        Log-Message "Connecting to Microsoft Graph..."
        Connect-MgGraph -Scopes "Application.ReadWrite.All" | Out-Null
        Log-Message "Successfully connected to Microsoft Graph"

        # Remove the application
        try {
            Log-Message "Attempting to remove Azure AD application..."
            Remove-MgApplication -ApplicationId $script:config.Credentials.AppId -ErrorAction Stop
            Log-Message "Successfully removed Azure AD application"
        } catch {
            if ($_.Exception.Message -match 'Request_ResourceNotFound') {
                Log-Message "Application not found, may have been already removed"
            } else {
                Log-Message "Error removing Azure AD application: $($_.Exception.Message)" "ERROR"
                throw
            }
        }
    } catch {
        Log-Message "Critical error removing Azure AD application: $($_.Exception.Message)" "ERROR"
    }
}

# Main script function
function Run-MainScript {
    try {
        # Start transcript logging
        $transcriptFileName = "MainScript_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $transcriptPath = Join-Path -Path $LogDirectory -ChildPath $transcriptFileName
        Start-Transcript -Path $transcriptPath
        Log-Message "Started transcript logging to: $transcriptPath"

        # Load config if it exists
        $configPath = Join-Path -Path $scriptDirectory -ChildPath "config.json"
        if (Test-Path $configPath) {
            Log-Message "Loading existing configuration from: $configPath"
            $loadedConfig = Get-Content -Path $configPath | ConvertFrom-Json
            
            # Convert PSCustomObject to hashtable for credentials
            $script:config.Credentials = @{
                TenantID = $loadedConfig.Credentials.TenantID
                ClientID = $loadedConfig.Credentials.ClientID
                ClientSecret = $loadedConfig.Credentials.ClientSecret
                AppId = $loadedConfig.Credentials.AppId
            }
            Log-Message "Configuration loaded successfully"
        } else {
            Log-Message "No existing configuration found at: $configPath"
        }

        # Get assignment preference
        $assignmentChoice = [System.Windows.Forms.MessageBox]::Show(
            "Would you like to assign this app to all devices?`n`nYes = Assign to all devices`nNo = Do not assign to any group",
            "Assignment Selection",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)

        Log-Message "User selected assignment choice: $assignmentChoice"

        # Validate credentials
        if ([string]::IsNullOrWhiteSpace($script:config.Credentials.TenantID) -or 
            [string]::IsNullOrWhiteSpace($script:config.Credentials.ClientID) -or 
            [string]::IsNullOrWhiteSpace($script:config.Credentials.ClientSecret)) {
            
            [System.Windows.Forms.MessageBox]::Show(
                "Please register the application first using the 'App Reg' button.",
                "Missing Credentials",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Process apps
        $selectedItems = $listView.SelectedItems
        if ($selectedItems.Count -eq 0) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "No apps selected. Do you want to process all apps?",
                "Confirm Action",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            
            if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                Log-Message "Operation cancelled. No apps processed."
                return
            }
            $appsToProcess = $script:config.Apps
        } else {
            $appsToProcess = $selectedItems | ForEach-Object {
                $displayName = $_.Text
                $script:config.Apps | Where-Object { $_.DisplayName -eq $displayName }
            }
        }

        $totalApps = $appsToProcess.Count
        if ($totalApps -eq 0) {
            Log-Message "No apps to process. Please add apps before running."
            return
        }

        try {
            # Install the IntuneWin32App module
            Log-Message "Installing IntuneWin32App module..."
            Install-Module -Name IntuneWin32App -Force -ErrorAction Stop
            Log-Message "IntuneWin32App module installed successfully."

            # Ensure C:\IntunePackages exists
            if (-not (Test-Path -Path $TempDir)) {
                New-Item -Path $TempDir -ItemType Directory -ErrorAction Stop | Out-Null
                Log-Message "Created C:\IntunePackages directory."
            } else {
                Log-Message "C:\IntunePackages directory already exists."
            }

            # Download IntuneWinAppUtil.exe if it doesn't exist
            if (-not (Test-Path $IntuneWinAppUtilPath)) {
                Invoke-WebRequest -Uri $IntuneWinAppUtilUrl -OutFile $IntuneWinAppUtilPath -ErrorAction Stop
                Log-Message "Downloaded IntuneWinAppUtil.exe successfully."
            } else {
                Log-Message "IntuneWinAppUtil.exe already exists."
            }

            # Connect to Intune Graph API
            Connect-MSIntuneGraph -TenantID $script:config.Credentials.TenantID -ClientID $script:config.Credentials.ClientID -ClientSecret $script:config.Credentials.ClientSecret -ErrorAction Stop
            Log-Message "Connected to Intune Graph API successfully."

            # Process each application
            foreach ($app in $appsToProcess) {
                try {
                    Log-Message "Starting process for $($app.DisplayName)"
                    
                    # Create a safe file name by replacing spaces with underscores
                    $safeFileName = $app.DisplayName -replace '\s', '_'

                    # Create Winget install script
                    $wingetInstallScriptPath = Join-Path -Path $TempDir -ChildPath "$safeFileName.ps1"
                    $wingetInstallScriptContent = @"
# Start Transcript
`$logDirectory = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs"
if (-not (Test-Path `$logDirectory)) {
    New-Item -ItemType Directory -Path `$logDirectory -Force | Out-Null
}
`$logPath = Join-Path -Path `$logDirectory -ChildPath "${safeFileName}_Install_`$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path `$logPath -Append

# Define the Winget Package Name
`$PackageName = '$($app.WingetId)'

`$AppInstaller = Get-AppxProvisionedPackage -Online | Where-Object DisplayName -eq Microsoft.DesktopAppInstaller

If(`$AppInstaller.Version -lt "2022.506.16.0") {
    Write-Host "Winget is not installed, trying to install latest version from Github" -ForegroundColor Yellow

    Try {
        Write-Host "Creating Winget Packages Folder" -ForegroundColor Yellow

        if (!(Test-Path -Path C:\ProgramData\WinGetPackages)) {
            New-Item -Path C:\ProgramData\WinGetPackages -Force -ItemType Directory
        }

        Set-Location C:\ProgramData\WinGetPackages

        #Downloading Packagefiles
        Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.7.0" -OutFile "C:\ProgramData\WinGetPackages\microsoft.ui.xaml.2.7.0.zip"
        Expand-Archive C:\ProgramData\WinGetPackages\microsoft.ui.xaml.2.7.0.zip -Force
        Invoke-WebRequest -Uri "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx" -OutFile "C:\ProgramData\WinGetPackages\Microsoft.VCLibs.x64.14.00.Desktop.appx"
        Invoke-WebRequest -Uri "https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" -OutFile "C:\ProgramData\WinGetPackages\Winget.msixbundle"

        Add-ProvisionedAppxPackage -online -PackagePath:.\Winget.msixbundle -DependencyPackagePath .\Microsoft.VCLibs.x64.14.00.Desktop.appx,.\microsoft.ui.xaml.2.7.0\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.7.Appx -SkipLicense

        Write-Host "Starting sleep for Winget to initiate" -Foregroundcolor Yellow
        Start-Sleep 2
    }
    Catch {
        Throw "Failed to install Winget"
        Break
    }
}
Else {
    Write-Host "Winget already installed, moving on" -ForegroundColor Green
}

#Trying to install Package with Winget
IF (`$PackageName){
    try {
        Write-Host "Installing `$PackageName via Winget" -ForegroundColor Green

        `$ResolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
        if (`$ResolveWingetPath){
               `$WingetPath = `$ResolveWingetPath[-1].Path
        }

        cd `$wingetpath

        .\winget.exe install `$PackageName --silent --accept-source-agreements --accept-package-agreements
    }
    Catch {
        Throw "Failed to install package $($_)"
    }
}
Else {
    Write-Host "Package `$PackageName not available" -ForegroundColor Yellow
}
# Stop Transcript
Stop-Transcript    
"@
                    $wingetInstallScriptContent | Out-File -FilePath $wingetInstallScriptPath -Encoding UTF8 -ErrorAction Stop
                    Log-Message "Created Winget install script for $($app.DisplayName)."

                    # Create Winget uninstall script
                    $wingetUninstallScriptPath = Join-Path -Path $TempDir -ChildPath "${safeFileName}_uninstall.ps1"
                    $wingetUninstallScriptContent = @"
# Start Transcript
`$logDirectory = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs"
if (-not (Test-Path `$logDirectory)) {
    New-Item -ItemType Directory -Path `$logDirectory -Force | Out-Null
}
`$logPath = Join-Path -Path `$logDirectory -ChildPath "${safeFileName}_Uninstall_`$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path `$logPath -Append

# Define the Winget Package Name
`$PackageName = '$($app.WingetId)'

# Find Winget executable path
`$ResolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
if (`$ResolveWingetPath) {
    `$WingetPath = `$ResolveWingetPath[-1].Path
    `$wingetExe = Join-Path -Path `$WingetPath -ChildPath 'winget.exe'
    
    if (Test-Path `$wingetExe) {
        Write-Host "Found Winget at: `$wingetExe"
        
        # Check if app is installed
        `$InstalledApps = & `$wingetExe list --id `$PackageName

        if (`$InstalledApps -match `$PackageName) {
            Write-Host "Trying to uninstall `$PackageName"
            try {
                # Change to Winget directory and run uninstall
                Set-Location -Path `$WingetPath
                & `$wingetExe uninstall --id `$PackageName --silent
                
                if (`$LASTEXITCODE -eq 0) {
                    Write-Host "Successfully uninstalled `$PackageName"
                    Exit 0
                } else {
                    Write-Host "Failed to uninstall `$PackageName. Exit code: `$LASTEXITCODE"
                    Exit 1
                }
            }
            catch {
                Write-Host "Error during uninstall: `$(`$_.Exception.Message)"
                Exit 1
            }
        }
        else {
            Write-Host "`$PackageName is not installed or detected"
            Exit 0
        }
    } else {
        Write-Host "Winget executable not found at expected location"
        Exit 1
    }
} else {
    Write-Host "Could not find Winget installation path"
    Exit 1
}

# Stop Transcript
Stop-Transcript
"@
                    $wingetUninstallScriptContent | Out-File -FilePath $wingetUninstallScriptPath -Encoding UTF8 -ErrorAction Stop
                    Log-Message "Created Winget uninstall script for $($app.DisplayName)."

                    # Package the script using IntuneWinAppUtil.exe
                    $outputFolder = "C:\IntunePackages"
                    $intuneWinFile = Join-Path -Path $outputFolder -ChildPath "$safeFileName.intunewin"
                    $arguments = @("-c", $TempDir, "-s", "$safeFileName.ps1", "-o", $outputFolder)
                    Start-Process -FilePath $IntuneWinAppUtilPath -ArgumentList $arguments -Wait -NoNewWindow -ErrorAction Stop
                    Log-Message "Packaged $($app.DisplayName) into $intuneWinFile successfully."

                    # Create detection script
                    $detectionScriptPath = Join-Path -Path $outputFolder -ChildPath "${safeFileName}_DetectionScript.ps1"
                    $detectionScriptContent = @"
# Detection Script for $($app.DisplayName)

`$PackageName = '$($app.WingetId)'

`$ResolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
    if (`$ResolveWingetPath){
           `$WingetPath = `$ResolveWingetPath[-1].Path
    }

`$config
cd `$wingetpath

`$InstalledApps = .\winget.exe list --id `$PackageName

if (`$InstalledApps -match `$PackageName) {
    Write-Host `$PackageName is installed
    Exit 0
}
else {
    Write-Host "`$PackageName is not installed"
    Exit 1
}
"@
                    $detectionScriptContent | Out-File -FilePath $detectionScriptPath -Encoding UTF8 -ErrorAction Stop
                    Log-Message "Created detection script for $($app.DisplayName)."

                    # Get metadata and create rules
                    $intuneWinMetaData = Get-IntuneWin32AppMetaData -FilePath $intuneWinFile
                    $requirementRule = New-IntuneWin32AppRequirementRule -Architecture "All" -MinimumSupportedWindowsRelease "W10_1607"
                    $detectionRule = New-IntuneWin32AppDetectionRuleScript -ScriptFile $detectionScriptPath -EnforceSignatureCheck $false -RunAs32Bit $false

                    # Single upload block based on assignment choice
                    Log-Message "Uploading $($app.DisplayName) to Intune..."
                    $intuneApp = Add-IntuneWin32App -FilePath $intuneWinFile `
                        -DisplayName $app.DisplayName `
                        -Description $app.Description `
                        -Publisher $app.Publisher `
                        -InstallExperience "system" `
                        -RestartBehavior "suppress" `
                        -DetectionRule $detectionRule `
                        -RequirementRule $requirementRule `
                        -InstallCommandLine $app.InstallCommand `
                        -UninstallCommandLine $app.UninstallCommand `
                        -Verbose -ErrorAction Stop

                    Log-Message "Successfully uploaded $($app.DisplayName) to Intune"
                    Log-Message "App ID: $($intuneApp.id)"
                    Log-Message "Display Name: $($intuneApp.displayName)"
                    Log-Message "Description: $($intuneApp.description)"
                    Log-Message "Publisher: $($intuneApp.publisher)"

                    # Add assignment only if user selected Yes
                    if ($assignmentChoice -eq [System.Windows.Forms.DialogResult]::Yes) {
                        Log-Message "Adding 'All Devices' assignment to $($app.DisplayName)"
                        Add-IntuneWin32AppAssignmentAllDevices -ID $intuneApp.id -Intent "required" -Notification "showAll" -Verbose
                        Log-Message "Successfully added 'All Devices' assignment to $($app.DisplayName)"
                    } else {
                        Log-Message "Skipping group assignment for $($app.DisplayName) as per user choice"
                    }

                    # Create update remediation script
                    Log-Message "Creating update remediation script..."
                    Create-UpdateRemediationScript `
                        -AppName $intuneApp.displayName `
                        -WingetId $app.WingetId `
                        -ClientId $script:config.Credentials.ClientID `
                        -ClientSecret $script:config.Credentials.ClientSecret `
                        -TenantId $script:config.Credentials.TenantID `
                        -AssignToAllDevices ($assignmentChoice -eq [System.Windows.Forms.DialogResult]::Yes)
                    Log-Message "Update remediation script created"

                } catch {
                    Log-Message "Error processing $($app.DisplayName)" "ERROR" $_
                }
            }

        } catch {
            Log-Message "Critical error in main script execution" "ERROR" $_
        } finally {
            # Perform cleanup operations
            Log-Message "Starting cleanup operations..."
            Cleanup-TempFiles -TempDir $TempDir
            Log-Message "Script execution completed. Intune Graph session will close automatically."
        }

        # Show completion message
        [System.Windows.Forms.MessageBox]::Show(
            "Process completed. Check the logs for details.",
            "Process Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        Log-Message "Critical error in main script execution" "ERROR" $_
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred. Please check the error log for details.`n`nError: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        # Stop transcript before cleanup
        try {
            Stop-Transcript
            Log-Message "Stopped transcript logging"
        }
        catch {
            Log-Message "Error stopping transcript" "ERROR" $_
        }
    }
}

# Define the Log-Message function
function Log-Message {
    param (
        [string]$Message,
        [string]$LogType = "INFO",
        [System.Management.Automation.ErrorRecord]$ErrorRecord = $null
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$LogType] - $Message"
    if ($ErrorRecord) {
        $logEntry += "`nError Details: $($ErrorRecord.Exception.Message)`nStack Trace: $($ErrorRecord.ScriptStackTrace)"
    }
    $logTextBox.AppendText("$logEntry`r`n")
    Add-Content -Path $LogFilePath -Value $logEntry
}

# Define cleanup function
function Cleanup-TempFiles {
    param (
        [string]$TempDir
    )
    try {
        Get-ChildItem -Path $TempDir -File | ForEach-Object {
            Remove-Item -Path $_.FullName -Force -ErrorAction Stop
            Log-Message "Removed temporary file: $($_.FullName)"
        }
        Remove-Item -Path $TempDir -Force -Recurse -ErrorAction Stop
        Log-Message "Removed temporary directory: $TempDir"
    } catch {
        Log-Message "Error during cleanup" "ERROR" $_
    }
}

# Define the Search-WingetApps function
function Search-WingetApps {
    $searchForm = New-Object System.Windows.Forms.Form
    $searchForm.Text = "Search Winget Apps"
    $searchForm.Size = New-Object System.Drawing.Size(600,400)
    $searchForm.StartPosition = "CenterScreen"

    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Location = New-Object System.Drawing.Point(10,20)
    $searchLabel.Size = New-Object System.Drawing.Size(100,20)
    $searchLabel.Text = "Search Term:"
    $searchForm.Controls.Add($searchLabel)

    $searchTextBox = New-Object System.Windows.Forms.TextBox
    $searchTextBox.Location = New-Object System.Drawing.Point(120,20)
    $searchTextBox.Size = New-Object System.Drawing.Size(350,20)
    $searchForm.Controls.Add($searchTextBox)

    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Location = New-Object System.Drawing.Point(480,20)
    $searchButton.Size = New-Object System.Drawing.Size(75,23)
    $searchButton.Text = "Search"
    $searchForm.Controls.Add($searchButton)

    $resultListView = New-Object System.Windows.Forms.ListView
    $resultListView.Location = New-Object System.Drawing.Point(10,50)
    $resultListView.Size = New-Object System.Drawing.Size(560,250)
    $resultListView.View = [System.Windows.Forms.View]::Details
    $resultListView.FullRowSelect = $true
    $resultListView.Columns.Add("Name", 200)
    $resultListView.Columns.Add("ID", 150)
    $resultListView.Columns.Add("Version", 100)
    $resultListView.Columns.Add("Source", 100)
    $searchForm.Controls.Add($resultListView)

    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Location = New-Object System.Drawing.Point(250,310)
    $addButton.Size = New-Object System.Drawing.Size(75,23)
    $addButton.Text = "Add"
    $addButton.Enabled = $false
    $searchForm.Controls.Add($addButton)

    # Create the search function
    $performSearch = {
        $resultListView.Items.Clear()
        $searchTerm = $searchTextBox.Text
        
        try {
            # Get winget search results
            $wingetOutput = winget search $searchTerm
            
            # Find the header line that contains "Name", "Id", "Version", and "Source"
            $headerLine = $wingetOutput | Where-Object { $_ -match "Name\s+Id\s+Version\s+Source" }
            if (-not $headerLine) {
                Log-Message "No results found or unexpected winget output format" "WARNING"
                return
            }
            
            # Get the positions of each column
            $namePos = $headerLine.IndexOf("Name")
            $idPos = $headerLine.IndexOf("Id")
            $versionPos = $headerLine.IndexOf("Version")
            $sourcePos = $headerLine.IndexOf("Source")
            
            # Process each line after the separator line (usually dashes)
            $startProcessing = $false
            foreach ($line in $wingetOutput) {
                # Skip until we find the separator line
                if ($line -match "^-+") {
                    $startProcessing = $true
                    continue
                }
                
                if ($startProcessing -and $line.Trim()) {
                    try {
                        # Extract each field based on position
                        $name = $line.Substring($namePos, $idPos - $namePos).Trim()
                        $id = $line.Substring($idPos, $versionPos - $idPos).Trim()
                        $version = $line.Substring($versionPos, $sourcePos - $versionPos).Trim()
                        $source = $line.Substring($sourcePos).Trim()
                        
                        # Create and add the item to the list view
                        $item = New-Object System.Windows.Forms.ListViewItem($name)
                        $item.SubItems.Add($id)
                        $item.SubItems.Add($version)
                        $item.SubItems.Add($source)
                        $resultListView.Items.Add($item)
                    }
                    catch {
                        # Skip malformed lines
                        Log-Message "Skipped malformed line: $line" "WARNING"
                        continue
                    }
                }
            }
        }
        catch {
            Log-Message "Error during winget search: $($_.Exception.Message)" "ERROR"
        }
    }

    # Add Enter key handler for the search textbox
    $searchTextBox.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true  # Prevents the "ding" sound
            & $performSearch
        }
    })

    # Add click handler for the search button
    $searchButton.Add_Click($performSearch)

    $resultListView.Add_SelectedIndexChanged({
        $addButton.Enabled = $resultListView.SelectedItems.Count -gt 0
    })

    $addButton.Add_Click({
        if ($resultListView.SelectedItems.Count -gt 0) {
            $selectedApp = $resultListView.SelectedItems[0]
            $displayName = $selectedApp.Text
            $wingetId = $selectedApp.SubItems[1].Text
            $version = $selectedApp.SubItems[2].Text

            $addForm = New-Object System.Windows.Forms.Form
            $addForm.Text = "Add App Details"
            $addForm.Size = New-Object System.Drawing.Size(300,250)
            $addForm.StartPosition = "CenterScreen"

            $publisherLabel = New-Object System.Windows.Forms.Label
            $publisherLabel.Location = New-Object System.Drawing.Point(10,20)
            $publisherLabel.Size = New-Object System.Drawing.Size(100,20)
            $publisherLabel.Text = "Publisher:"
            $addForm.Controls.Add($publisherLabel)

            $publisherTextBox = New-Object System.Windows.Forms.TextBox
            $publisherTextBox.Location = New-Object System.Drawing.Point(120,20)
            $publisherTextBox.Size = New-Object System.Drawing.Size(150,20)
            $addForm.Controls.Add($publisherTextBox)

            $descriptionLabel = New-Object System.Windows.Forms.Label
            $descriptionLabel.Location = New-Object System.Drawing.Point(10,50)
            $descriptionLabel.Size = New-Object System.Drawing.Size(100,20)
            $descriptionLabel.Text = "Description:"
            $addForm.Controls.Add($descriptionLabel)

            $descriptionTextBox = New-Object System.Windows.Forms.TextBox
            $descriptionTextBox.Location = New-Object System.Drawing.Point(120,50)
            $descriptionTextBox.Size = New-Object System.Drawing.Size(150,60)
            $descriptionTextBox.Multiline = $true
            $addForm.Controls.Add($descriptionTextBox)

            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Location = New-Object System.Drawing.Point(120,120)
            $okButton.Size = New-Object System.Drawing.Size(75,23)
            $okButton.Text = "OK"
            $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $addForm.Controls.Add($okButton)

            $addForm.AcceptButton = $okButton

            $result = $addForm.ShowDialog()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $safeFileName = $displayName -replace '\s', '_'
                $installCommand = "powershell.exe -ExecutionPolicy Bypass -File .\$safeFileName.ps1"
                $uninstallCommand = "powershell.exe -ExecutionPolicy Bypass -File .\${safeFileName}_uninstall.ps1"

                $newApp = @{
                    DisplayName = $displayName
                    WingetId = $wingetId
                    Publisher = $publisherTextBox.Text
                    Description = $descriptionTextBox.Text
                    InstallCommand = $installCommand
                    UninstallCommand = $uninstallCommand
                }

                $script:config.Apps += $newApp
                Save-Config
                Load-Apps
                Log-Message "Added new app from Winget: $displayName"
                $searchForm.Close()
            }
        }
    })

    $searchForm.ShowDialog()
}

# Define the Save-Config function
function Save-Config {
    if ($PSScriptRoot) {
        $scriptRoot = $PSScriptRoot
    } elseif ($psISE) {
        $scriptRoot = Split-Path -Parent $psISE.CurrentFile.FullPath
    } else {
        $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    
    if (-not $scriptRoot) {
        $scriptRoot = Get-Location
    }
    
    $configPath = Join-Path -Path $scriptRoot -ChildPath "config.json"
    $script:config | ConvertTo-Json | Set-Content -Path $configPath
}
# Add event handlers
$appRegButton.Add_Click({ Register-IntuneApp })
$removeCredButton.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to remove all stored credentials?",
        "Confirm Removal",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Remove-AzureADApp
            Remove-StoredCredentials
            Remove-GraphCredentials

            $script:config = $null

            [System.Windows.Forms.MessageBox]::Show(
                "All credentials have been removed successfully.",
                "Cleanup Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            Log-Message "Error during credential removal: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Error removing credentials: $($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})
$addButton.Add_Click({ Add-App })
$editButton.Add_Click({ Edit-App })
$removeButton.Add_Click({ Remove-App })
$runButton.Add_Click({ Run-MainScript })
$searchButton.Add_Click({ Search-WingetApps })

# Load apps when form is shown
$form.Add_Shown({Load-Apps})

# Add form closing event handler (add this before $form.ShowDialog())
$form.Add_FormClosing({
    param($sender, $e)
    
    [System.Windows.Forms.MessageBox]::Show(
        "All credentials will be removed.",
        "Closing Application",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)
    
    # Execute cleanup in specified order
    Remove-AzureADApp
    Remove-StoredCredentials
    Remove-GraphCredentials
})

# Show the form
$form.ShowDialog()
