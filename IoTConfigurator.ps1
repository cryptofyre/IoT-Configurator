using namespace System.Windows.Forms
using namespace System.Drawing

# Script Parameters
param(
    [switch]$DryRun = $false
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Check if running as administrator and self-elevate if necessary
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    try {
        $argList = @("-NoProfile", "-File", $MyInvocation.MyCommand.Path)
        Start-Process powershell.exe -ArgumentList $argList -Verb RunAs -Wait
        exit
    }
    catch {
        Write-Error "Failed to elevate permissions: $_"
        exit 1
    }
}

# Color scheme definition
$script:Colors = @{
    Dark = @{
        Background = [System.Drawing.Color]::FromArgb(32, 32, 32)
        ForegroundDim = [System.Drawing.Color]::FromArgb(153, 153, 153)
        Foreground = [System.Drawing.Color]::FromArgb(240, 240, 240)
        AccentPrimary = [System.Drawing.Color]::FromArgb(86, 156, 214)
        AccentSecondary = [System.Drawing.Color]::FromArgb(78, 201, 176)
        Border = [System.Drawing.Color]::FromArgb(61, 61, 61)
        ConsoleBackground = [System.Drawing.Color]::FromArgb(12, 12, 12)
        ConsoleForeground = [System.Drawing.Color]::FromArgb(204, 204, 204)
        TabBackground = [System.Drawing.Color]::FromArgb(45, 45, 45)
        ControlBackground = [System.Drawing.Color]::FromArgb(51, 51, 51)
    }
    Light = @{
        Background = [System.Drawing.Color]::FromArgb(249, 249, 249)
        ForegroundDim = [System.Drawing.Color]::FromArgb(109, 109, 109)
        Foreground = [System.Drawing.Color]::FromArgb(23, 23, 23)
        AccentPrimary = [System.Drawing.Color]::FromArgb(0, 120, 212)
        AccentSecondary = [System.Drawing.Color]::FromArgb(0, 153, 123)
        Border = [System.Drawing.Color]::FromArgb(225, 225, 225)
        ConsoleBackground = [System.Drawing.Color]::FromArgb(255, 255, 255)
        ConsoleForeground = [System.Drawing.Color]::FromArgb(0, 0, 0)
        TabBackground = [System.Drawing.Color]::FromArgb(240, 240, 240)
        ControlBackground = [System.Drawing.Color]::FromArgb(255, 255, 255)
    }
}

# Initialize progress tracking
$script:totalSteps = 0
$script:currentStep = 0
$script:tempDir = Join-Path $env:TEMP "Win11IoTSetup"

# Get the current script path
$scriptPath = $MyInvocation.MyCommand.Path

# Self-elevation function
function Start-ElevatedSession {
    # Check if running as administrator
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        try {
            # Get current script parameters
            $scriptArgs = $MyInvocation.BoundParameters.GetEnumerator() | ForEach-Object {
                "-$($_.Key)", "$($_.Value)"
            }
            
            # Build the argument list
            $argList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$scriptPath`"")
            if ($scriptArgs) {
                $argList += $scriptArgs
            }

            # Start new elevated process
            $process = Start-Process -FilePath "pwsh.exe" -ArgumentList $argList -Verb RunAs -PassThru -Wait
            
            # Check if elevation was successful
            if ($process.ExitCode -ne 0) {
                throw "Failed to run with elevated privileges. Exit code: $($process.ExitCode)"
            }

            # Exit the current non-elevated session
            exit
        }
        catch {
            Write-Error "Failed to elevate permissions: $_"
            exit 1
        }
    }
}

function Update-Progress {
    param(
        [string]$Status,
        [int]$Step
    )

    # Calculate percentage based on total steps
    $percentage = [math]::Min(100, [math]::Round(($Step / $script:totalSteps) * 100))
    
    # Update UI elements if they exist
    if ($UI -and $UI.SetProgress) {
        & $UI.SetProgress $percentage $Status
    }
    
    # Write to console if available
    if ($UI -and $UI.WriteToConsole) {
        & $UI.WriteToConsole $Status
    } else {
        Write-Host $Status
    }
}

# Check and set execution policy
function Set-RequiredExecutionPolicy {
    try {
        $currentPolicy = Get-ExecutionPolicy
        if ($currentPolicy -ne "Bypass" -and $currentPolicy -ne "Unrestricted") {
            Set-ExecutionPolicy Bypass -Scope Process -Force
        }
    }
    catch {
        Write-Error "Failed to set execution policy: $_"
        exit 1
    }
}

# Verify Windows version and edition
function Test-WindowsCompatibility {
    $os = Get-WmiObject -Class Win32_OperatingSystem
    $windowsVersion = [System.Environment]::OSVersion.Version
    
    if ($windowsVersion.Major -lt 10 -or ($windowsVersion.Major -eq 10 -and $windowsVersion.Build -lt 22000)) {
        Write-Error "This script requires Windows 11 or later."
        exit 1
    }
}

# Initialize environment
function Initialize-Environment {
    # Enable TLS 1.2 for downloads
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Create temp directory for downloads if it doesn't exist
    if (-not (Test-Path $script:tempDir)) {
        New-Item -ItemType Directory -Path $script:tempDir -Force | Out-Null
    }

    # Clean up any leftover files from previous runs
    Get-ChildItem -Path $script:tempDir -File | Remove-Item -Force
}

# Function to load required assemblies
function Initialize-Requirements {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}

# Function to get system theme
function Get-SystemTheme {
    try {
        $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $lightTheme = Get-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -ErrorAction Stop
        return $lightTheme.AppsUseLightTheme -eq 0
    }
    catch {
        return $false  # Default to light theme if registry key not found
    }
}

# Color scheme definition
$script:Colors = @{
    Dark = @{
        Background = [System.Drawing.Color]::FromArgb(32, 32, 32)
        ForegroundDim = [System.Drawing.Color]::FromArgb(153, 153, 153)
        Foreground = [System.Drawing.Color]::FromArgb(240, 240, 240)
        AccentPrimary = [System.Drawing.Color]::FromArgb(86, 156, 214)
        AccentSecondary = [System.Drawing.Color]::FromArgb(78, 201, 176)
        Border = [System.Drawing.Color]::FromArgb(61, 61, 61)
        ConsoleBackground = [System.Drawing.Color]::FromArgb(12, 12, 12)
        ConsoleForeground = [System.Drawing.Color]::FromArgb(204, 204, 204)
        TabBackground = [System.Drawing.Color]::FromArgb(45, 45, 45)
        ControlBackground = [System.Drawing.Color]::FromArgb(51, 51, 51)
    }
    Light = @{
        Background = [System.Drawing.Color]::FromArgb(249, 249, 249)
        ForegroundDim = [System.Drawing.Color]::FromArgb(109, 109, 109)
        Foreground = [System.Drawing.Color]::FromArgb(23, 23, 23)
        AccentPrimary = [System.Drawing.Color]::FromArgb(0, 120, 212)
        AccentSecondary = [System.Drawing.Color]::FromArgb(0, 153, 123)
        Border = [System.Drawing.Color]::FromArgb(225, 225, 225)
        ConsoleBackground = [System.Drawing.Color]::FromArgb(255, 255, 255)
        ConsoleForeground = [System.Drawing.Color]::FromArgb(0, 0, 0)
        TabBackground = [System.Drawing.Color]::FromArgb(240, 240, 240)
        ControlBackground = [System.Drawing.Color]::FromArgb(255, 255, 255)
    }
}

function New-SetupUI {
    # Enable DPI awareness
    Add-Type -TypeDefinition @"
        using System.Runtime.InteropServices;
        public class DpiAwareness {
            [DllImport("user32.dll")]
            public static extern bool SetProcessDPIAware();
        }
"@
    [DpiAwareness]::SetProcessDPIAware()

    # Determine theme
    $isDarkMode = Get-SystemTheme
    # Set $theme as a script-level variable
    $script:theme = if ($isDarkMode) { $script:Colors.Dark } else { $script:Colors.Light }

    # Create main form
    $form = New-Object Form
    $form.Text = "IoT Configurator (Preview) by cryptofyre"
    $form.Size = New-Object Drawing.Size(1200, 900)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object Drawing.Font("Segoe UI", 10)
    $form.BackColor = $script:theme.Background
    $form.ForeColor = $script:theme.Foreground
    $form.MinimumSize = New-Object Drawing.Size(1000, 800)

    # Create main TableLayoutPanel
    $mainLayout = New-Object TableLayoutPanel
    $mainLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainLayout.RowCount = 2
    $mainLayout.ColumnCount = 2
    $mainLayout.ColumnStyles.Clear()
    $mainLayout.ColumnStyles.Add((New-Object ColumnStyle([System.Windows.Forms.SizeType]::Percent, 70)))
    $mainLayout.ColumnStyles.Add((New-Object ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
    $mainLayout.RowStyles.Clear()
    $mainLayout.RowStyles.Add((New-Object RowStyle([System.Windows.Forms.SizeType]::Percent, 80)))  # Increased main content height
    $mainLayout.RowStyles.Add((New-Object RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))  # Increased progress section
    $mainLayout.Padding = New-Object System.Windows.Forms.Padding(10)
    $mainLayout.BackColor = $script:theme.Background

    # Apply border styling to main layout
    $mainLayout.BackColor = $script:theme.Border
    $mainLayout.Padding = New-Object System.Windows.Forms.Padding(1)
    $mainLayout.Margin = New-Object System.Windows.Forms.Padding(5)

    # Create TabControl
    $tabControl = New-Object TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tabControl.Font = New-Object Drawing.Font("Segoe UI", 10)
    $tabControl.BackColor = $script:theme.TabBackground
    $tabControl.ItemSize = New-Object Drawing.Size(0, 30)

    # Basic Setup Tab
    $basicTab = New-Object TabPage
    $basicTab.Text = "Basic Setup"
    $basicTab.BackColor = $script:theme.TabBackground
    $basicTab.ForeColor = $script:theme.Foreground
    $basicTab.Padding = New-Object System.Windows.Forms.Padding(15)

    # Basic Setup Layout
    $basicLayout = New-Object TableLayoutPanel
    $basicLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $basicLayout.AutoSize = $false          # Disable AutoSize to prevent clipping
    $basicLayout.AutoScroll = $true         # Enable scrolling if content overflows
    $basicLayout.ColumnCount = 1
    $basicLayout.RowCount = 3  # Increased row count to accommodate Activation group
    $basicLayout.ColumnStyles.Clear()
    $basicLayout.ColumnStyles.Add((New-Object ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $basicLayout.RowStyles.Clear()
    $basicLayout.RowStyles.Add((New-Object RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # For app group
    $basicLayout.RowStyles.Add((New-Object RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # For browser group
    $basicLayout.RowStyles.Add((New-Object RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # For activation group

    # Store Section
    $appGroup = New-Object GroupBox
    $appGroup.Text = "Install Apps"
    $appGroup.AutoSize = $true  # Enable AutoSize
    $appGroup.Dock = [System.Windows.Forms.DockStyle]::Top    # Ensure proper docking
    $appGroup.ForeColor = $script:theme.Foreground

    # Update the modern apps array to include Microsoft Store
    $modernApps = @(
        @{
            Name = "StoreCheckbox"
            Text = "Microsoft Store"
            Tooltip = "Installs the Microsoft Store and enables app installations"
        },
        @{
            Name = "TerminalCheckbox"
            Text = "Windows Terminal"
            Tooltip = "Modern terminal emulator for Windows"
        },
        @{
            Name = "NotepadCheckbox"
            Text = "Modern Notepad"
            Tooltip = "Updated version of Windows Notepad"
        },
        @{
            Name = "PaintCheckbox"
            Text = "Modern Paint"
            Tooltip = "Updated version of Microsoft Paint"
        }
    )

    # Update app group layout to include all modern apps
    $appGroupLayout = New-Object TableLayoutPanel
    $appGroupLayout.AutoSize = $true  # Enable AutoSize
    $appGroupLayout.Dock = [System.Windows.Forms.DockStyle]::Top
    $appGroupLayout.ColumnCount = 1
    $appGroupLayout.RowCount = $modernApps.Count + 1
    $appGroupLayout.Padding = New-Object System.Windows.Forms.Padding(5)

    $appGroupLabel = New-Object Label
    $appGroupLabel.Text = "Select which apps this script will install:"
    $appGroupLabel.AutoSize = $true
    $appGroupLabel.ForeColor = $script:theme.Foreground

    $appGroupLayout.Controls.Add($appGroupLabel, 0, 0)

    # Initialize a hashtable to store checkboxes
    $checkboxes = @{}

    $row = 1
    foreach ($app in $modernApps) {
        $checkbox = New-Object CheckBox
        $checkbox.Name = $app.Name
        $checkbox.Text = $app.Text
        $checkbox.AutoSize = $true
        $checkbox.ForeColor = $script:theme.Foreground
        $checkbox.Margin = New-Object System.Windows.Forms.Padding(5)
        
        $tooltip = New-Object ToolTip
        $tooltip.SetToolTip($checkbox, $app.Tooltip)
        
        $appGroupLayout.Controls.Add($checkbox, 0, $row)
        # Store the checkbox in a hashtable for later access
        $checkboxes[$app.Name] = $checkbox
        $row++
    }

    $appGroup.Controls.Add($appGroupLayout)
    $basicLayout.Controls.Add($appGroup, 0, 0)

    # Browser Section
    $browserGroup = New-Object GroupBox
    $browserGroup.Text = "Web Browser"
    $browserGroup.AutoSize = $true
    $browserGroup.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $browserGroup.Dock = [System.Windows.Forms.DockStyle]::Top
    $browserGroup.ForeColor = $script:theme.Foreground

    $browserLayout = New-Object TableLayoutPanel
    $browserLayout.AutoSize = $true
    $browserLayout.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $browserLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $browserLayout.ColumnCount = 1
    $browserLayout.RowCount = 2
    $browserLayout.Padding = New-Object System.Windows.Forms.Padding(5)

    $browserLabel = New-Object Label
    $browserLabel.Text = "Select your preferred web browser:"
    $browserLabel.AutoSize = $true
    $browserLabel.ForeColor = $script:theme.Foreground

    $browserComboBox = New-Object ComboBox
    $browserComboBox.Dock = [System.Windows.Forms.DockStyle]::Top
    $browserComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $browserComboBox.Items.AddRange(@("Zen Browser", "Firefox", "Chrome", "Edge"))
    $browserComboBox.BackColor = $script:theme.ControlBackground
    $browserComboBox.ForeColor = $script:theme.Foreground
    $browserComboBox.Height = 30
    
    $browserToolTip = New-Object ToolTip
    $browserToolTip.SetToolTip($browserComboBox, "Select your preferred web browser to install")

    # Add controls to browser layout
    $browserLayout.Controls.Add($browserLabel, 0, 0)
    $browserLayout.Controls.Add($browserComboBox, 0, 1)

    $browserGroup.Controls.Add($browserLayout)
    $basicLayout.Controls.Add($browserGroup, 0, 1)

    # Activation Section
    $activationGroup = New-Object GroupBox
    $activationGroup.Text = "Activation"
    $activationGroup.AutoSize = $true
    $activationGroup.Dock = [System.Windows.Forms.DockStyle]::Top
    $activationGroup.ForeColor = $script:theme.Foreground

    $activationLayout = New-Object TableLayoutPanel
    $activationLayout.AutoSize = $true
    $activationLayout.Dock = [System.Windows.Forms.DockStyle]::Top
    $activationLayout.ColumnCount = 1
    $activationLayout.RowCount = 1
    $activationLayout.Padding = New-Object System.Windows.Forms.Padding(5)

    $activateWindowsCheckbox = New-Object CheckBox
    $activateWindowsCheckbox.Name = "ActivateWindowsCheckbox"
    $activateWindowsCheckbox.Text = "Activate Windows"
    $activateWindowsCheckbox.AutoSize = $true
    $activateWindowsCheckbox.ForeColor = $script:theme.Foreground
    $activateWindowsCheckbox.Margin = New-Object System.Windows.Forms.Padding(5)

    $activationLayout.Controls.Add($activateWindowsCheckbox, 0, 0)
    $activationGroup.Controls.Add($activationLayout)
    $basicLayout.Controls.Add($activationGroup, 0, 2)

    $basicTab.Controls.Add($basicLayout)

    # Development Tools Tab
    $devTab = New-Object TabPage
    $devTab.Text = "Development Tools"
    $devTab.BackColor = $script:theme.TabBackground
    $devTab.ForeColor = $script:theme.Foreground
    $devTab.Padding = New-Object System.Windows.Forms.Padding(15)

    $devLayout = New-Object TableLayoutPanel
    $devLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $devLayout.ColumnCount = 2
    $devLayout.RowCount = 5
    $devLayout.ColumnStyles.Add((New-Object ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $devLayout.ColumnStyles.Add((New-Object ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

    $devTools = @{
        "Visual Studio" = @{
            Name = "VSCheckbox"
            Tooltip = "Full-featured IDE for Windows development"
        }
        "VS Code" = @{
            Name = "VSCodeCheckbox"
            Tooltip = "Lightweight, extensible code editor"
        }
        ".NET SDK" = @{
            Name = "DotNetCheckbox"
            Tooltip = "Development kit for .NET applications"
        }
        "Node.js LTS" = @{
            Name = "NodeCheckbox"
            Tooltip = "JavaScript runtime environment"
        }
        "Go" = @{
            Name = "GoCheckbox"
            Tooltip = "Go programming language toolchain"
        }
        "Rust" = @{
            Name = "RustCheckbox"
            Tooltip = "Rust programming language and cargo"
        }
        "LLVM" = @{
            Name = "LLVMCheckbox"
            Tooltip = "LLVM compiler infrastructure"
        }
        "Scoop" = @{
            Name = "ScoopCheckbox"
            Tooltip = "Command-line installer for Windows"
        }
        "Chocolatey" = @{
            Name = "ChocoCheckbox"
            Tooltip = "Package manager for Windows"
        }
    }

    $row = 0
    $col = 0
    foreach ($tool in $devTools.GetEnumerator()) {
        $checkbox = New-Object CheckBox
        $checkbox.Name = $tool.Value.Name
        $checkbox.Text = $tool.Key
        $checkbox.AutoSize = $true
        $checkbox.ForeColor = $script:theme.Foreground
        $checkbox.Margin = New-Object System.Windows.Forms.Padding(5)
        
        $tooltip = New-Object ToolTip
        $tooltip.SetToolTip($checkbox, $tool.Value.Tooltip)
        
        $devLayout.Controls.Add($checkbox, $col, $row)
        
        if ($col -eq 1) {
            $col = 0
            $row++
        } else {
            $col++
        }
    }

    $devTab.Controls.Add($devLayout)

    # Windows Features Tab
    $featuresTab = New-Object TabPage
    $featuresTab.Text = "Windows Features"
    $featuresTab.BackColor = $script:theme.TabBackground
    $featuresTab.ForeColor = $script:theme.Foreground
    $featuresTab.Padding = New-Object System.Windows.Forms.Padding(15)

    $featuresLayout = New-Object TableLayoutPanel
    $featuresLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $featuresLayout.ColumnCount = 1
    $featuresLayout.RowCount = 2

    $features = @(
        @{
            Name = "SandboxCheckbox"
            Text = "Windows Sandbox"
            Tooltip = "Lightweight desktop environment for safely running applications"
        },
        @{
            Name = "WSLCheckbox"
            Text = "Windows Subsystem for Linux"
            Tooltip = "Run Linux distributions natively on Windows"
        }
    )

    $row = 0
    foreach ($feature in $features) {
        $checkbox = New-Object CheckBox
        $checkbox.Name = $feature.Name
        $checkbox.Text = $feature.Text
        $checkbox.AutoSize = $true
        $checkbox.ForeColor = $script:theme.Foreground
        $checkbox.Margin = New-Object System.Windows.Forms.Padding(5)
        
        $tooltip = New-Object ToolTip
        $tooltip.SetToolTip($checkbox, $feature.Tooltip)
        
        $featuresLayout.Controls.Add($checkbox, 0, $row)
        $row++
    }

    $featuresTab.Controls.Add($featuresLayout)

    # Add all tabs to TabControl
    $tabControl.Controls.AddRange(@($basicTab, $devTab, $featuresTab))
    $mainLayout.Controls.Add($tabControl, 0, 0)

    # Apply theme colors to TabControl and TabPages
    $tabControl.BackColor = $script:theme.Background
    $tabControl.ForeColor = $script:theme.Foreground
    foreach ($tabPage in $tabControl.TabPages) {
        $tabPage.BackColor = $script:theme.ControlBackground
        $tabPage.ForeColor = $script:theme.Foreground
    }

    # Apply theme colors and custom drawing to TabControl
    $tabControl.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed

    # Add Paint event handler to set background behind tabs
    $tabControl.Add_Paint({
        param($sender, $e)
        # Fill the area behind the tabs with ConsoleBackground color
        $tabRect = $sender.GetTabRect(0)
        $tabRect.Height += 10  # Adjust height if necessary
        $e.Graphics.FillRectangle((New-Object System.Drawing.SolidBrush($script:theme.ConsoleBackground)), 0, 0, $sender.Width, $tabRect.Height)
    })

    # Adjust the event handler to use ConsoleHeader color styling
    $tabControl.Add_DrawItem({
        param($sender, $e)
        $g = $e.Graphics
        $tabPage = $sender.TabPages.Item($e.Index)
        $tabBounds = $sender.GetTabRect($e.Index)

        # Convert $tabBounds to RectangleF
        $tabBoundsF = [System.Drawing.RectangleF]$tabBounds

        # Use $script:theme to access the ConsoleHeader colors
        if ($sender.SelectedIndex -eq $e.Index) {
            $backColor = $script:theme.ConsoleBackground
            $foreColor = $script:theme.AccentPrimary
        } else {
            $backColor = $script:theme.ConsoleBackground
            $foreColor = $script:theme.ForegroundDim
        }

        # Fill the background of the tab header
        $g.FillRectangle((New-Object System.Drawing.SolidBrush($backColor)), $tabBounds)

        # Draw the tab text using $tabBoundsF
        $g.DrawString($tabPage.Text, $sender.Font, (New-Object System.Drawing.SolidBrush($foreColor)), $tabBoundsF, $stringFlags)
    })

    # Console Panel
    $consolePanel = New-Object TableLayoutPanel
    $consolePanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $consolePanel.RowCount = 2
    $consolePanel.ColumnCount = 1
    $consolePanel.RowStyles.Add((New-Object RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
    $consolePanel.RowStyles.Add((New-Object RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $consolePanel.BackColor = $script:theme.Border
    $consolePanel.Padding = New-Object System.Windows.Forms.Padding(1)
    $consolePanel.Margin = New-Object System.Windows.Forms.Padding(5)

    # Console Header
    $consoleHeader = New-Object Label
    $consoleHeader.Text = "Installation Progress"
    $consoleHeader.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $consoleHeader.Dock = [System.Windows.Forms.DockStyle]::Fill
    $consoleHeader.BackColor = $script:theme.ConsoleBackground
    $consoleHeader.ForeColor = $script:theme.AccentPrimary
    $consoleHeader.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)

    # Console Output
    $consoleOutput = New-Object RichTextBox
    $consoleOutput.Multiline = $true
    $consoleOutput.ReadOnly = $true
    $consoleOutput.BackColor = $script:theme.ConsoleBackground
    $consoleOutput.ForeColor = $script:theme.ConsoleForeground
    $consoleOutput.Font = New-Object Drawing.Font("Cascadia Code", 9)
    $consoleOutput.Dock = [System.Windows.Forms.DockStyle]::Fill
    $consoleOutput.ScrollBars = "Vertical"
    $consoleOutput.WordWrap = $true

    $consolePanel.Controls.Add($consoleHeader, 0, 0)
    $consolePanel.Controls.Add($consoleOutput, 0, 1)

    # Add console panel to main layout
    $mainLayout.Controls.Add($consolePanel, 1, 0)

    # Progress Section
    $progressPanel = New-Object TableLayoutPanel
    $progressPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $progressPanel.ColumnCount = 1
    $progressPanel.RowCount = 3
    $progressPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    $progressPanel.BackColor = $script:theme.Background
    # $progressPanel.AutoScroll = $true  # Remove auto-scrolling

    # Adjust RowStyles to allow Install button to be visible
    $progressPanel.RowStyles.Clear()
    $progressPanel.RowStyles.Add((New-Object RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # Status Label
    $progressPanel.RowStyles.Add((New-Object RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # Progress Bar
    $progressPanel.RowStyles.Add((New-Object RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # Install Button

    # Status Label
    $statusLabel = New-Object Label
    $statusLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $statusLabel.Text = "Ready to begin setup..."
    $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $statusLabel.ForeColor = $script:theme.Foreground
    $statusLabel.Font = New-Object Drawing.Font("Segoe UI", 10)

    # Progress Bar
    $progressBar = New-Object ProgressBar
    $progressBar.Dock = [System.Windows.Forms.DockStyle]::Top
    $progressBar.Style = "Continuous"
    $progressBar.Height = 25
    if ($isDarkMode) {
        $progressBar.ForeColor = $script:theme.AccentPrimary
    }

    # Install Button Container
    $buttonContainer = New-Object TableLayoutPanel
    $buttonContainer.Dock = [System.Windows.Forms.DockStyle]::Top  # Change from Fill to Top
    $buttonContainer.BackColor = $script:theme.Background
    $buttonContainer.ColumnCount = 1
    $buttonContainer.RowCount = 1
    $buttonContainer.ColumnStyles.Add((New-Object ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $buttonContainer.RowStyles.Add((New-Object RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    # Install Button
    $script:installButton = New-Object Button
    $script:installButton.Text = "Start Installation"
    $script:installButton.Size = New-Object Drawing.Size(150, 40)
    $script:installButton.BackColor = $script:theme.AccentPrimary
    $script:installButton.ForeColor = $script:theme.Background
    $script:installButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $script:installButton.Font = New-Object Drawing.Font("Segoe UI Semibold", 10)
    $script:installButton.Anchor = [System.Windows.Forms.AnchorStyles]::None

    # Add Install Button to TableLayoutPanel
    $buttonContainer.Controls.Add($script:installButton, 0, 0)

    # Add elements to progress panel
    $progressPanel.Controls.Add($statusLabel, 0, 0)
    $progressPanel.Controls.Add($progressBar, 0, 1)
    $progressPanel.Controls.Add($buttonContainer, 0, 2)

    # Add progress panel to main layout
    $mainLayout.Controls.Add($progressPanel, 0, 1)
    $mainLayout.SetColumnSpan($progressPanel, 2)

    # Add main layout to form
    $form.Controls.Add($mainLayout)

    # Console write function
    function Write-ToConsole {
        param(
            [string]$Message,
            [string]$Type = "Info"  # Info, Success, Warning, Error
        )
        
        $timestamp = Get-Date -Format "HH:mm:ss"
        $color = switch ($Type) {
            "Success" { $script:theme.AccentSecondary }
            "Warning" { [System.Drawing.Color]::Orange }
            "Error" { [System.Drawing.Color]::Red }
            default { $script:theme.ConsoleForeground }
        }

        $consoleOutput.SelectionStart = $consoleOutput.TextLength
        $consoleOutput.SelectionLength = 0
        $consoleOutput.SelectionColor = $script:theme.ForegroundDim
        $consoleOutput.AppendText("[$timestamp] ")
        $consoleOutput.SelectionColor = $color
        $consoleOutput.AppendText("$Message`r`n")
        $consoleOutput.ScrollToCaret()
    }

    # Create initial console message
    Write-ToConsole "Setup initialized. Waiting for user input..." "Info"

    # Return all UI elements and functions
    return @{
        # Main UI Elements
        Form = $form
        ProgressBar = $progressBar
        StatusLabel = $statusLabel
        ConsoleOutput = $consoleOutput
        WriteToConsole = ${function:Write-ToConsole}
        
        # Basic Setup Tab Elements
        StoreCheckbox = $checkboxes["StoreCheckbox"]
        BrowserComboBox = $browserComboBox
        ActivateWindowsCheckbox = $activateWindowsCheckbox
        
        # Modern Apps Tab Elements
        TerminalCheckbox = $checkboxes["TerminalCheckbox"]
        NotepadCheckbox = $checkboxes["NotepadCheckbox"]
        PaintCheckbox = $checkboxes["PaintCheckbox"]
        
        # Windows Features Tab Elements
        SandboxCheckbox = $featuresLayout.Controls["SandboxCheckbox"]
        WSLCheckbox = $featuresLayout.Controls["WSLCheckbox"]
        
        # Development Tools Tab Elements
        VSCheckbox = $devLayout.Controls["VSCheckbox"]
        VSCodeCheckbox = $devLayout.Controls["VSCodeCheckbox"]
        DotNetCheckbox = $devLayout.Controls["DotNetCheckbox"]
        NodeCheckbox = $devLayout.Controls["NodeCheckbox"]
        GoCheckbox = $devLayout.Controls["GoCheckbox"]
        RustCheckbox = $devLayout.Controls["RustCheckbox"]
        LLVMCheckbox = $devLayout.Controls["LLVMCheckbox"]
        ScoopCheckbox = $devLayout.Controls["ScoopCheckbox"]
        ChocoCheckbox = $devLayout.Controls["ChocoCheckbox"]
        
        # Control Elements
        InstallButton = $script:installButton
        MainLayout = $mainLayout
        TabControl = $tabControl
        ConsolePanel = $consolePanel
        ProgressPanel = $progressPanel
        
        # Theme Information
        Theme = $script:theme
        IsDarkMode = $isDarkMode
        
        # Layout Panels
        BasicLayout = $basicLayout
        AppsLayout = $appGroupLayout
        DevLayout = $devLayout
        FeaturesLayout = $featuresLayout
        
        # Helper Functions
        RefreshUI = {
            param($sender, $e)
            $form.Refresh()
        }
        
        ClearConsole = {
            $consoleOutput.Clear()
            Write-ToConsole "Console cleared." "Info"
        }
        
        SetProgress = {
            param(
                [int]$Value,
                [string]$Status
            )
            $progressBar.Value = $Value
            $statusLabel.Text = $Status
            $form.Refresh()
        }
    }
}

# Function to install basic requirements
function Install-BasicRequirements {
    Update-Progress "Installing PowerShell 7..." 1
    if (-not $DryRun) {
        winget install --id Microsoft.Powershell --source winget
    }

    Update-Progress "Installing Uniget..." 2
    if (-not $DryRun) {
        # Download and install UniGet from the new URL
        Invoke-WebRequest -Uri "https://github.com/marticliment/UniGetUI/releases/latest/download/WingetUI.Installer.exe" -OutFile "$script:tempDir\WingetUI.Installer.exe"
        Start-Process -FilePath "$script:tempDir\WingetUI.Installer.exe" -Wait
    }
}

# Function to install media components
function Install-MediaComponents {
    Update-Progress "Installing HEVC Extensions..." 4
    if (-not $DryRun) {
        $hevcUrl = "http://tlu.dl.delivery.mp.microsoft.com/filestreamingservice/files/94091873-fe5d-49a1-89c2-5f8194d3715f?P1=1730712157&P2=404&P3=2&P4=K6x5FwuILwHyd69Uj6G29jkBppk323rCWF2NOTnaVKfanrNQfyE6fc4gO5fp02uW7Pf9FSwAskO8%2fBLN21SWXg%3d%3d"
        $hevcPath = Join-Path $script:tempDir "HEVCVideoExtension.appxbundle"
        Invoke-WebRequest -Uri $hevcUrl -OutFile $hevcPath
        Add-AppxPackage -Path $hevcPath
    }

    Update-Progress "Installing Web Media Extensions..." 5
    if (-not $DryRun) {
        $webMediaUrl = "http://tlu.dl.delivery.mp.microsoft.com/filestreamingservice/files/a91cad94-456f-4333-8bd4-e9403e83fcd0?P1=1730712178&P2=404&P3=2&P4=BLAA8MwwJFmh9w3sFiX%2bGHIeeEPLtFd6bz6KIh9FTaeN6Vwru%2fA1FqheB%2bU3VBXBgRtAfC%2bTSmEXiKqS8D%2b%2blQ%3d%3d"
        $webMediaPath = Join-Path $script:tempDir "WebMediaExtensions.appxbundle"
        Invoke-WebRequest -Uri $webMediaUrl -OutFile $webMediaPath
        Add-AppxPackage -Path $webMediaPath
    }

    Update-Progress "Installing Media Feature Pack..." 6
    if (-not $DryRun) {
        Get-WindowsCapability -online | Where-Object -Property name -like "*media*" | Add-WindowsCapability -Online
    }

    # Install HEIF Image Extensions
    Update-Progress "Installing HEIF Image Extensions..." 9
    if (-not $DryRun) {
        # Download HEIF Image Extensions installer
        $heifUrl = "http://tlu.dl.delivery.mp.microsoft.com/filestreamingservice/files/4ffb7a09-8860-48d1-9d8d-5cb7655af9c5?P1=1730712725&P2=404&P3=2&P4=WLlbuxdJQtRKBMss45KxO7Uupl3DPRGPYMqKzXdbPnkDtkAVq%2fsLDAi8ECoXYLKNpaoKX5RZ3rlfMSQw1WiTDA%3d%3d"
        $heifPath = Join-Path $script:tempDir "HEIFImageExtensions.appxbundle"
        Invoke-WebRequest -Uri $heifUrl -OutFile $heifPath
        Add-AppxPackage -Path $heifPath
    }

    # Install VP9 Video Extensions
    Update-Progress "Installing VP9 Video Extensions..." 10
    if (-not $DryRun) {
        # Download VP9 Video Extensions installer
        $vp9Url = "http://tlu.dl.delivery.mp.microsoft.com/filestreamingservice/files/fbcd5fa5-922e-49ee-8a11-39eb78f45764?P1=1730712770&P2=404&P3=2&P4=dms%2fMtk1IrzcH1j220oQ7alpHLKwPhS%2bqbYKP3ZPyb8KV6n1cUWqLX%2fJPPqyq2p%2fDOCriTh2yAc4HRZTDDi2CA%3d%3d"
        $vp9Path = Join-Path $script:tempDir "VP9VideoExtensions.appxbundle"
        Invoke-WebRequest -Uri $vp9Url -OutFile $vp9Path
        Add-AppxPackage -Path $vp9Path
    }

    Update-Progress "Installing .NET Runtime..." 7
    if (-not $DryRun) {
        winget install --id Microsoft.DotNet.Runtime.6 --source winget
    }

    Update-Progress "Installing VC++ Redistributable..." 8
    if (-not $DryRun) {
        winget install --id Microsoft.VCRedist.2015+.x64 --source winget
    }
}

# Function to install selected browser
function Install-SelectedBrowser {
    param(
        [string]$BrowserChoice
    )
    
    Update-Progress "Installing $BrowserChoice..." 7
    if (-not $DryRun) {
        switch ($BrowserChoice) {
            "Zen Browser" { winget install --id Zen-Team.Zen-Browser }
            "Firefox" { winget install --id Mozilla.Firefox }
            "Chrome" { winget install --id Google.Chrome }
            "Edge" { winget install --id Microsoft.Edge }
        }
    }
}

# Function to install modern Windows apps
function Install-ModernApps {
    if ($UI.StoreCheckbox.Checked) {
        Update-Progress "Installing Microsoft Store..." 3
        if (-not $DryRun) {
            Add-AppxPackage -RegisterByFamilyName -MainPackage "Microsoft.WindowsStore_8wekyb3d8bbwe"
        }
    }

    if ($UI.TerminalCheckbox.Checked) {
        Update-Progress "Installing Windows Terminal..." 4
        if (-not $DryRun) {
            winget install --id Microsoft.WindowsTerminal --source winget
        }
    }

    if ($UI.NotepadCheckbox.Checked) {
        Update-Progress "Installing Modern Notepad..." 5
        if (-not $DryRun) {
            winget install --id Microsoft.Notepad --source winget
        }
    }

    if ($UI.PaintCheckbox.Checked) {
        Update-Progress "Installing Modern Paint..." 6
        if (-not $DryRun) {
            winget install --id Microsoft.Paint --source winget
        }
    }
}

# Function to enable Windows features
function Enable-WindowsFeatures {
    # Enable Hyper-V and Virtual Machine Platform if either Sandbox or WSL is selected
    if ($UI.SandboxCheckbox.Checked -or $UI.WSLCheckbox.Checked) {
        Update-Progress "Enabling Hyper-V..." 10
        if (-not $DryRun) {
            Enable-WindowsOptionalFeature -Online -FeatureName "VirtualMachinePlatform" -All -NoRestart  # Enable Virtual Machine Platform
            Enable-WindowsOptionalFeature -Online -FeatureName "Microsoft-Hyper-V-All" -All -NoRestart
        }
    }

    if ($UI.SandboxCheckbox.Checked) {
        Update-Progress "Enabling Windows Sandbox..." 11
        if (-not $DryRun) {
            Enable-WindowsOptionalFeature -Online -FeatureName "Containers-DisposableClientVM" -All -NoRestart
        }
    }

    if ($UI.WSLCheckbox.Checked) {
        Update-Progress "Enabling WSL..." 12
        if (-not $DryRun) {
            Enable-WindowsOptionalFeature -Online -FeatureName "Microsoft-Windows-Subsystem-Linux" -All -NoRestart  # Enable WSL
        }
    }
}

# Function to install development tools
function Install-DevelopmentTools {
    if ($UI.VSCheckbox.Checked) {
        Update-Progress "Installing Visual Studio..." 13
        if (-not $DryRun) {
            winget install --id Microsoft.VisualStudio.2022.Community --source winget
        }
    }

    if ($UI.VSCodeCheckbox.Checked) {
        Update-Progress "Installing VS Code..." 14
        if (-not $DryRun) {
            winget install --id Microsoft.VSCode --source winget
        }
    }

    if ($UI.DotNetCheckbox.Checked) {
        $dotnetVersions = @("8.0", "7.0", "6.0")
        foreach ($version in $dotnetVersions) {
            Update-Progress "Installing .NET SDK $version..." 15
            if (-not $DryRun) {
                winget install --id Microsoft.DotNet.SDK.$version --source winget
            }
        }
    }

    if ($UI.NodeCheckbox.Checked) {
        Update-Progress "Installing Node.js LTS..." 16
        if (-not $DryRun) {
            winget install --id OpenJS.NodeJS.LTS --source winget
        }
    }

    if ($UI.GoCheckbox.Checked) {
        Update-Progress "Installing Go..." 17
        if (-not $DryRun) {
            winget install --id GoLang.Go --source winget
        }
    }

    if ($UI.RustCheckbox.Checked) {
        Update-Progress "Installing Rust..." 18
        if (-not $DryRun) {
            $rustupPath = Join-Path $script:tempDir "rustup-init.exe"
            Invoke-WebRequest -Uri https://win.rustup.rs -OutFile $rustupPath
            Start-Process -FilePath $rustupPath -ArgumentList "-y" -Wait
        }
    }

    if ($UI.LLVMCheckbox.Checked) {
        Update-Progress "Installing LLVM..." 19
        if (-not $DryRun) {
            winget install --id LLVM.LLVM --source winget
        }
    }

    if ($UI.ScoopCheckbox.Checked) {
        Update-Progress "Installing Scoop..." 20
        if (-not $DryRun) {
            Invoke-Expression "& {$(Invoke-RestMethod get.scoop.sh)}"
        }
    }

    if ($UI.ChocoCheckbox.Checked) {
        Update-Progress "Installing Chocolatey..." 21
        if (-not $DryRun) {
            Set-ExecutionPolicy Bypass -Scope Process -Force
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
            Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
        }
    }
}

# Function to check for winget installation and install if missing
function Ensure-WingetInstalled {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Update-Progress "Installing Winget..." 0
        if (-not $DryRun) {
            $wingetUrl = "https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller.msixbundle"
            $wingetPath = Join-Path $script:tempDir "Microsoft.DesktopAppInstaller.msixbundle"
            Invoke-WebRequest -Uri $wingetUrl -OutFile $wingetPath
            Add-AppxPackage -Path $wingetPath
        }
    }
}

function Install-WindowsActivation {
    Update-Progress "Activating Windows..." 22
    if (-not $DryRun) {
        & ([ScriptBlock]::Create((irm https://get.activated.win))) /HWID
    }
}

function Start-InstallationProcess {
    # Calculate total steps based on selected options
    $script:currentStep = 0
    $script:totalSteps = 24  # Increased total steps by 1
    
    try {
        # Ensure winget is installed
        Ensure-WingetInstalled
        
        # Core installations
        Install-BasicRequirements
        Install-MediaComponents
        
        # Browser installation
        if ($UI.BrowserComboBox.SelectedItem) {
            Install-SelectedBrowser -BrowserChoice $UI.BrowserComboBox.SelectedItem
        }
        
        # Optional installations
        Install-ModernApps
        Enable-WindowsFeatures
        Install-DevelopmentTools

        # Windows Activation
        if ($UI.ActivateWindowsCheckbox.Checked) {
            Install-WindowsActivation
        }

        Update-Progress "Setup completed successfully!" $script:totalSteps
        
        # Show completion message
        if (-not $DryRun) {
            [System.Windows.Forms.MessageBox]::Show(
                "Installation completed successfully! Some changes may require a restart.",
                "Setup Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
    catch {
        Write-Error "An error occurred during installation: $_"
        if (-not $DryRun) {
            [System.Windows.Forms.MessageBox]::Show(
                "An error occurred during installation: $_`nCheck the log for details.",
                "Installation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Main execution flow
try {
    # Start elevation if needed
    Start-ElevatedSession

    # Set required execution policy
    Set-RequiredExecutionPolicy

    # Check Windows compatibility
    Test-WindowsCompatibility

    # Initialize environment
    Initialize-Environment

    # Initialize UI requirements
    Initialize-Requirements

    # Create and show UI
    $UI = New-SetupUI
    
    # Add click handler for install button
    $UI.InstallButton.Add_Click({
        $UI.InstallButton.Enabled = $false
        Start-InstallationProcess
        $UI.InstallButton.Enabled = $true
    })

    # Show the form
    [System.Windows.Forms.Application]::Run($UI.Form)
}
catch {
    Write-Error "A critical error occurred during setup initialization: $_"
    exit 1
}
finally {
    # Cleanup
    if (Test-Path $script:tempDir) {
        Remove-Item -Path $script:tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}