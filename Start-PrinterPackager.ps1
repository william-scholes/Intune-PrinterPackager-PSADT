<#
.SYNOPSIS
    Intune Printer Packager for PSADT v4 (WPF GUI)
.DESCRIPTION
    WPF GUI application that discovers locally installed printers, extracts driver files,
    captures local printer preferences via PrintBrm (full -NOBIN backup -> unpack -> filter XML manifests -> prune unused files -> repack),
    packages direct IP printers for deployment via Intune or Standalone Local Install using PSADT v4, and optionally compiles .intunewin and .zip packages.
.NOTES
    Pure PowerShell - avoids cscript, VBScript, and batch scripts.
#>

# Helper: BOM-less UTF-8 File Writer
function Write-BOMFreeUtf8File {
    param([string]$Path, [string]$Content)
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

# Resolve Shell Path Robustly
$powershell = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
if (-not $powershell -or -not (Test-Path $powershell)) {
    $powershell = (Get-Command powershell, pwsh -ErrorAction SilentlyContinue)[0].Path
}

# Check Administrator Elevation
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [System.Security.Principal.WindowsPrincipal]$identity
$isAdmin = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Warning "Administrator privileges required for PrintBrm and DriverStore extraction. Requesting elevation..."
    Start-Process -FilePath $powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$($MyInvocation.MyCommand.Path)`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

# Ensure STA mode for WPF if running directly
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Warning "Script must run in Single Threaded Apartment (STA) mode. Re-launching PowerShell in STA mode..."
    & $powershell -NoProfile -ExecutionPolicy Bypass -STA -File "$MyInvocation.MyCommand.Path"
    return
}

# Dynamic Base Paths & Auto-Discovery
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir) { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$AppsDir = Split-Path -Parent $ScriptDir

# Auto-discover PSADT v4 Template Directory
$candidateTemplatePaths = @(
    (Join-Path -Path $AppsDir -ChildPath "PSAppDeployToolkit_Template_v4"),
    (Join-Path -Path $ScriptDir -ChildPath "PSAppDeployToolkit_Template_v4"),
    (Join-Path -Path $ScriptDir -ChildPath "template"),
    (Join-Path -Path $AppsDir -ChildPath "template")
)
$PSADTTemplatePath = ($candidateTemplatePaths | Where-Object { Test-Path $_ } | Select-Object -First 1)
if (-not $PSADTTemplatePath) { $PSADTTemplatePath = Join-Path -Path $AppsDir -ChildPath "PSAppDeployToolkit_Template_v4" }

# Auto-discover IntuneWinAppUtil.exe
$candidateUtilPaths = @(
    (Join-Path -Path $AppsDir -ChildPath "IntuneWinAppUtil.exe"),
    (Join-Path -Path $ScriptDir -ChildPath "IntuneWinAppUtil.exe")
)
$cmdUtil = Get-Command "IntuneWinAppUtil.exe" -ErrorAction SilentlyContinue
if ($cmdUtil) { $candidateUtilPaths += $cmdUtil.Source }
$IntuneWinUtilPath = ($candidateUtilPaths | Where-Object { Test-Path $_ } | Select-Object -First 1)
if (-not $IntuneWinUtilPath) { $IntuneWinUtilPath = Join-Path -Path $AppsDir -ChildPath "IntuneWinAppUtil.exe" }

# --- WPF XAML Layout Definition ---
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Intune Printer Packager (PSADT v4)" Height="800" Width="1080"
        WindowStartupLocation="CenterScreen" Background="#1E1E1E" Foreground="#FFFFFF"
        FontFamily="Segoe UI" FontSize="13">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#CCCCCC"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#106EBE"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#3F3F46"/>
                    <Setter Property="Foreground" Value="#888888"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#252526"/>
            <Setter Property="RowBackground" Value="#252526"/>
            <Setter Property="AlternatingRowBackground" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#3F3F46"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#333337"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#094771"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="2*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="1*"/>
        </Grid.RowDefinitions>

        <!-- Header Panel -->
        <Border Grid.Row="0" Background="#2D2D30" CornerRadius="6" Padding="14" Margin="0,0,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="Intune Printer Packager for PSADT v4" FontSize="20" FontWeight="Bold" Foreground="#FFFFFF"/>
                    <TextBlock Text="Package direct IP printer queues, drivers, ports, and DevMode preferences for Intune deployment." Margin="0,4,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button Name="BtnHelp" Content="Help / How-To Guide" Background="#0078D4" Margin="0,0,8,0"/>
                    <Button Name="BtnRefreshPrinters" Content="Refresh Printers" Background="#3A3D41"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Printer List / Selection DataGrid -->
        <Grid Grid.Row="1" Margin="0,0,0,12">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            
            <DockPanel Grid.Row="0" Margin="0,0,0,8">
                <TextBlock Text="Locally Installed Printers" FontSize="14" FontWeight="SemiBold" Foreground="#FFFFFF" VerticalAlignment="Center"/>
                <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                    <TextBlock Text="Search: " VerticalAlignment="Center" Margin="0,0,6,0"/>
                    <TextBox Name="TxtFilter" Width="200"/>
                </StackPanel>
            </DockPanel>

            <DataGrid Name="GridPrinters" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" SelectionMode="Single">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Printer Name" Binding="{Binding Name}" Width="1.5*"/>
                    <DataGridTextColumn Header="Driver Name" Binding="{Binding DriverName}" Width="2*"/>
                    <DataGridTextColumn Header="Port Name" Binding="{Binding PortName}" Width="1.2*"/>
                    <DataGridTextColumn Header="IP / Host" Binding="{Binding PrinterHostAddress}" Width="1.2*"/>
                    <DataGridTextColumn Header="Protocol" Binding="{Binding ProtocolStr}" Width="0.8*"/>
                    <DataGridTextColumn Header="Port #" Binding="{Binding PortNumber}" Width="0.6*"/>
                    <DataGridTextColumn Header="Location" Binding="{Binding Location}" Width="1*"/>
                    <DataGridTextColumn Header="Comment" Binding="{Binding Comment}" Width="1.5*"/>
                </DataGrid.Columns>
            </DataGrid>
        </Grid>

        <!-- Configuration Settings Panel -->
        <Border Grid.Row="2" Background="#2D2D30" CornerRadius="6" Padding="14" Margin="0,0,0,12">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="1*"/>
                    <ColumnDefinition Width="1.1*"/>
                    <ColumnDefinition Width="1.6*"/>
                </Grid.ColumnDefinitions>

                <!-- Column 1: App Metadata -->
                <StackPanel Grid.Row="0" Grid.Column="0" Margin="0,0,12,8">
                    <TextBlock Text="App Vendor:" Margin="0,0,0,4"/>
                    <TextBox Name="TxtAppVendor" Text="Company"/>
                </StackPanel>
                <StackPanel Grid.Row="1" Grid.Column="0" Margin="0,0,12,8">
                    <TextBlock Text="App Name:" Margin="0,0,0,4"/>
                    <TextBox Name="TxtAppName" Text="Printer Package"/>
                </StackPanel>
                <StackPanel Grid.Row="2" Grid.Column="0" Margin="0,0,12,0">
                    <TextBlock Text="App Version:" Margin="0,0,0,4"/>
                    <TextBox Name="TxtAppVersion" Text="1.0.0"/>
                </StackPanel>

                <!-- Column 2: Package Options -->
                <StackPanel Grid.Row="0" Grid.RowSpan="3" Grid.Column="1" Margin="6,0,12,0">
                    <TextBlock Text="Packaging Options:" FontWeight="SemiBold" Foreground="#FFFFFF" Margin="0,0,0,8"/>
                    <CheckBox Name="ChkCapturePreferences" Content="Capture DevMode Preferences (PrintBrm)" IsChecked="True" Margin="0,0,0,8"/>
                    <CheckBox Name="ChkExtractDriver" Content="Extract Driver Store Files (FileRepository)" IsChecked="True" Margin="0,0,0,8"/>
                    <CheckBox Name="ChkCreateIntuneWin" Content="Compile .intunewin Package (IntuneWinAppUtil)" IsChecked="True" Margin="0,0,0,8"/>
                    <CheckBox Name="ChkCreateZip" Content="Create Standalone Zip Archive (.zip)" IsChecked="True" Margin="0,0,0,8"/>
                    <CheckBox Name="ChkSetDefault" Content="Set as Default Printer on Target Device" IsChecked="False" Margin="0,0,0,8"/>
                </StackPanel>

                <!-- Column 3: Configurable Paths & Browse Buttons -->
                <StackPanel Grid.Row="0" Grid.RowSpan="3" Grid.Column="2" Margin="6,0,0,0">
                    <!-- Output Directory -->
                    <TextBlock Text="Output Directory:" Margin="0,0,0,3"/>
                    <Grid Margin="0,0,0,6">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Name="TxtOutputDir" Grid.Column="0"/>
                        <Button Name="BtnBrowseOutput" Content="Browse..." Grid.Column="1" Margin="6,0,0,0" Background="#3A3D41"/>
                    </Grid>

                    <!-- PSADT Template Path -->
                    <TextBlock Text="PSADT v4 Template Path:" Margin="0,0,0,3"/>
                    <Grid Margin="0,0,0,6">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Name="TxtTemplatePath" Grid.Column="0"/>
                        <Button Name="BtnBrowseTemplate" Content="Browse..." Grid.Column="1" Margin="6,0,0,0" Background="#3A3D41"/>
                    </Grid>

                    <!-- IntuneWinAppUtil Executable Path -->
                    <TextBlock Text="IntuneWinAppUtil.exe Path:" Margin="0,0,0,3"/>
                    <Grid Margin="0,0,0,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Name="TxtIntuneWinUtilPath" Grid.Column="0"/>
                        <Button Name="BtnBrowseIntuneWinUtil" Content="Browse..." Grid.Column="1" Margin="6,0,0,0" Background="#3A3D41"/>
                    </Grid>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Action Buttons Bar -->
        <DockPanel Grid.Row="3" Margin="0,0,0,12">
            <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
                <Button Name="BtnGenerateDetection" Content="Generate Detection Script Only" Background="#3A3D41" Margin="0,0,8,0"/>
                <Button Name="BtnPackageLocal" Content="Create Local Standalone Package" Background="#0078D4" Margin="0,0,8,0" Padding="14,8"/>
                <Button Name="BtnPackagePrinter" Content="Package Printer for Intune (PSADT v4)" Background="#107C41" Padding="16,8"/>
            </StackPanel>
            <TextBlock Name="TxtStatus" Text="Ready." VerticalAlignment="Center" FontWeight="SemiBold" Foreground="#0078D4"/>
        </DockPanel>

        <!-- Output Log Window -->
        <Border Grid.Row="4" Background="#1E1E1E" BorderBrush="#3F3F46" BorderThickness="1" CornerRadius="4">
            <DockPanel>
                <Border DockPanel.Dock="Top" Background="#2D2D30" Padding="8,4">
                    <TextBlock Text="Activity Console &amp; Build Logs" FontWeight="SemiBold" Foreground="#CCCCCC"/>
                </Border>
                <TextBox Name="TxtLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" 
                         Background="#1E1E1E" Foreground="#00FF66" FontFamily="Consolas" FontSize="12" BorderThickness="0" Padding="8"/>
            </DockPanel>
        </Border>
    </Grid>
</Window>
"@

# Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Element Controls Mapping
$GridPrinters           = $window.FindName('GridPrinters')
$BtnHelp                = $window.FindName('BtnHelp')
$BtnRefreshPrinters     = $window.FindName('BtnRefreshPrinters')
$TxtFilter              = $window.FindName('TxtFilter')
$TxtAppVendor           = $window.FindName('TxtAppVendor')
$TxtAppName              = $window.FindName('TxtAppName')
$TxtAppVersion          = $window.FindName('TxtAppVersion')
$ChkCapturePreferences  = $window.FindName('ChkCapturePreferences')
$ChkExtractDriver       = $window.FindName('ChkExtractDriver')
$ChkCreateIntuneWin     = $window.FindName('ChkCreateIntuneWin')
$ChkCreateZip           = $window.FindName('ChkCreateZip')
$ChkSetDefault          = $window.FindName('ChkSetDefault')
$TxtOutputDir           = $window.FindName('TxtOutputDir')
$BtnBrowseOutput        = $window.FindName('BtnBrowseOutput')
$TxtTemplatePath        = $window.FindName('TxtTemplatePath')
$BtnBrowseTemplate      = $window.FindName('BtnBrowseTemplate')
$TxtIntuneWinUtilPath   = $window.FindName('TxtIntuneWinUtilPath')
$BtnBrowseIntuneWinUtil = $window.FindName('BtnBrowseIntuneWinUtil')
$BtnGenerateDetection   = $window.FindName('BtnGenerateDetection')
$BtnPackageLocal        = $window.FindName('BtnPackageLocal')
$BtnPackagePrinter      = $window.FindName('BtnPackagePrinter')
$TxtStatus              = $window.FindName('TxtStatus')
$TxtLog                 = $window.FindName('TxtLog')

# Set default values
$TxtTemplatePath.Text = $PSADTTemplatePath
$TxtIntuneWinUtilPath.Text = $IntuneWinUtilPath
$TxtOutputDir.Text = Join-Path -Path $ScriptDir -ChildPath "Output"

# Global Storage for Printer List
$Script:MasterPrinterList = [System.Collections.Generic.List[PSObject]]::new()

# Helper: Logging Function
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] [$Level] $Message"
    
    $TxtLog.Dispatcher.Invoke([Action]{
        $TxtLog.AppendText("$logLine`r`n")
        $TxtLog.ScrollToEnd()
    })
}

# Function: Query Installed Printers
function Load-Printers {
    Write-Log "Scanning installed printers on local system..." -Level INFO
    $TxtStatus.Text = "Scanning printers..."
    
    $Script:MasterPrinterList.Clear()
    
    try {
        $printers = Get-Printer -ErrorAction Stop
        $ports = Get-PrinterPort -ErrorAction SilentlyContinue
        $drivers = Get-PrinterDriver -ErrorAction SilentlyContinue
        
        foreach ($p in $printers) {
            # Skip Microsoft software virtual printers unless requested
            if ($p.Name -match "Microsoft XPS|Microsoft Print to PDF|Fax|OneNote") {
                continue
            }
            
            $portObj = $ports | Where-Object { $_.Name -eq $p.PortName } | Select-Object -First 1
            $driverObj = $drivers | Where-Object { $_.Name -eq $p.DriverName } | Select-Object -First 1
            
            $hostAddr = if ($portObj) { $portObj.PrinterHostAddress } else { "" }
            $portNum = if ($portObj -and $portObj.PortNumber) { $portObj.PortNumber } else { 9100 }
            
            $protocolStr = "RAW"
            if ($portObj) {
                if ($portObj.Protocol -eq 2 -or $portObj.Protocol -eq "LPR") {
                    $protocolStr = "LPR"
                } elseif ($portObj.Protocol -eq 1 -or $portObj.Protocol -eq "RAW") {
                    $protocolStr = "RAW"
                } else {
                    $protocolStr = "$($portObj.Protocol)"
                }
            }
            
            # Retrieve basic print configuration
            $printConfig = Get-PrintConfiguration -PrinterName $p.Name -ErrorAction SilentlyContinue
            $duplexing = if ($printConfig) { "$($printConfig.DuplexingMode)" } else { "" }
            $color = if ($printConfig -and $printConfig.Color -ne $null) { [bool]$printConfig.Color } else { $true }
            $paperSize = if ($printConfig -and $printConfig.PaperSize) { "$($printConfig.PaperSize)" } else { "" }

            $item = [PSCustomObject]@{
                Name               = $p.Name
                DriverName         = $p.DriverName
                PortName           = $p.PortName
                PrinterHostAddress = $hostAddr
                ProtocolStr        = $protocolStr
                Protocol           = if ($portObj) { $portObj.Protocol } else { 1 }
                PortNumber         = $portNum
                Location           = $p.Location
                Comment            = $p.Comment
                Shared             = $p.Shared
                InfPath            = if ($driverObj) { $driverObj.InfPath } else { "" }
                DuplexingMode      = $duplexing
                Color              = $color
                PaperSize          = $paperSize
            }
            
            $Script:MasterPrinterList.Add($item)
        }
        
        $GridPrinters.ItemsSource = $Script:MasterPrinterList
        Write-Log "Discovered $($Script:MasterPrinterList.Count) local hardware/network printers." -Level SUCCESS
        $TxtStatus.Text = "Ready. Discovered $($Script:MasterPrinterList.Count) printers."
    }
    catch {
        Write-Log "Failed to query printers: $_" -Level ERROR
        $TxtStatus.Text = "Error scanning printers."
    }
}

# Filter Printer List on Search Input (Null-Safe Property Access)
$TxtFilter.Add_TextChanged({
    $filter = $TxtFilter.Text.Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($filter)) {
        $GridPrinters.ItemsSource = $Script:MasterPrinterList
    } else {
        $filtered = $Script:MasterPrinterList | Where-Object {
            ($_.Name -and $_.Name.ToLower().Contains($filter)) -or
            ($_.DriverName -and $_.DriverName.ToLower().Contains($filter)) -or
            ($_.PortName -and $_.PortName.ToLower().Contains($filter)) -or
            ($_.PrinterHostAddress -and $_.PrinterHostAddress.ToLower().Contains($filter))
        }
        $GridPrinters.ItemsSource = $filtered
    }
})

# Grid Selection Handler: Auto-populate App Name
$GridPrinters.Add_SelectionChanged({
    $selected = $GridPrinters.SelectedItem
    if ($selected) {
        $safeName = $selected.Name -replace '[^\w\-\.]', '_'
        $TxtAppName.Text = "Printer - $safeName"
    }
})

# Refresh Printers Button Event
$BtnRefreshPrinters.Add_Click({
    Load-Printers
})

# Browse Output Directory Event
$BtnBrowseOutput.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select Output Folder for PSADT Package"
    if (Test-Path $TxtOutputDir.Text) { $dialog.SelectedPath = $TxtOutputDir.Text }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TxtOutputDir.Text = $dialog.SelectedPath
    }
})

# Browse PSADT Template Directory Event
$BtnBrowseTemplate.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select PSADT v4 Template Directory"
    if (Test-Path $TxtTemplatePath.Text) { $dialog.SelectedPath = $TxtTemplatePath.Text }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TxtTemplatePath.Text = $dialog.SelectedPath
    }
})

# Browse IntuneWinAppUtil Executable Event
$BtnBrowseIntuneWinUtil.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select IntuneWinAppUtil.exe Executable"
    $dialog.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
    if (Test-Path $TxtIntuneWinUtilPath.Text) {
        $dialog.InitialDirectory = Split-Path -Parent $TxtIntuneWinUtilPath.Text
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TxtIntuneWinUtilPath.Text = $dialog.FileName
    }
})

# Function: Export Driver Files from Driver Store / System
function Export-PrinterDriverFiles {
    param(
        [string]$DriverName,
        [string]$InfPath,
        [string]$DestinationDir
    )
    
    Write-Log "Attempting driver file extraction for '$DriverName'..." -Level INFO
    New-Item -Path $DestinationDir -ItemType Directory -Force | Out-Null
    
    $extracted = $false
    
    # 1. Check if InfPath points to DriverStore FileRepository (Fixed -like pattern)
    if ($InfPath -and (Test-Path $InfPath)) {
        $parentFolder = Split-Path -Parent $InfPath
        if ($parentFolder -like "*\DriverStore\FileRepository*") {
            Write-Log "Found DriverStore repository folder: $parentFolder" -Level INFO
            Copy-Item -Path "$parentFolder\*" -Destination $DestinationDir -Recurse -Force -ErrorAction SilentlyContinue
            $extracted = $true
            Write-Log "Successfully copied driver package from DriverStore." -Level SUCCESS
        }
    }
    
    # 2. Fallback: Use PnPUtil specific export if InfPath is an OEM INF
    if (-not $extracted) {
        Write-Log "Using PnPUtil driver export fallback..." -Level INFO
        $tempExport = Join-Path -Path $env:TEMP -ChildPath "PnPUtil_Export_$([Guid]::NewGuid().Guid)"
        New-Item -Path $tempExport -ItemType Directory -Force | Out-Null
        
        try {
            $targetInf = "*"
            if ($InfPath -and (Test-Path $InfPath) -and ($InfPath -like "*.inf")) {
                $targetInf = Split-Path -Leaf $InfPath
            }
            
            & "pnputil.exe" /export-driver $targetInf "$tempExport"
            $matchedInf = Get-ChildItem -Path $tempExport -Filter "*.inf" -Recurse | Where-Object {
                Select-String -Path $_.FullName -Pattern [regex]::Escape($DriverName) -Quiet
            } | Select-Object -First 1
            
            if ($matchedInf) {
                $matchFolder = Split-Path -Parent $matchedInf.FullName
                Copy-Item -Path "$matchFolder\*" -Destination $DestinationDir -Recurse -Force
                $extracted = $true
                Write-Log "PnPUtil matched and exported driver package for $DriverName." -Level SUCCESS
            }
        }
        catch {
            Write-Log "PnPUtil export warning: $_" -Level WARNING
        }
        finally {
            Remove-Item -Path $tempExport -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    
    return $extracted
}

# Function: Export full printer preferences (vendor-specific DevMode) using PrintUI.dll /Ss
function Export-PrinterSettingsDat {
    param(
        [string]$PrinterName,
        [string]$DestinationFile
    )
    
    Write-Log "Exporting full vendor-specific printer preferences for '$PrinterName' using PrintUI.dll /Ss..." -Level INFO
    
    try {
        # Use rundll32 printui.dll,PrintUIEntry /Ss to export global settings (g) and driver-specific data (d)
        $printUiArgs = @(
            "printui.dll,PrintUIEntry",
            "/Ss",
            "/n", "`"$PrinterName`"",
            "/a", "`"$DestinationFile`"",
            "g", "d"
        )
        $proc = Start-Process -FilePath "rundll32.exe" -ArgumentList $printUiArgs -WindowStyle Hidden -Wait -PassThru
        
        if ((Test-Path $DestinationFile) -and (Get-Item $DestinationFile).Length -gt 0) {
            $sizeKb = [math]::Round(((Get-Item $DestinationFile).Length / 1KB), 1)
            Write-Log "PrintUI.dll /Ss export successful for '$PrinterName' (${sizeKb} KB)." -Level SUCCESS
            return $true
        } else {
            Write-Log "PrintUI.dll /Ss export completed but output file is missing or empty." -Level WARNING
            return $false
        }
    }
    catch {
        Write-Log "PrintUI.dll /Ss export error: $_" -Level WARNING
        return $false
    }
}

# Function: Export Printer DevMode Preferences using PrintBrm (Full Backup -> Unpack -> Filter XMLs & Prune Files -> Repack)
function Export-PrinterPreferences {
    param(
        [string]$PrinterName,
        [string]$PortName,
        [string]$DriverName,
        [string]$DestinationFile
    )
    
    Write-Log "Exporting printer queue & DevMode preferences for '$PrinterName' using PrintBrm..." -Level INFO
    
    $printBrm = "$env:SystemRoot\System32\spool\tools\PrintBrm.exe"
    if (-not (Test-Path $printBrm)) {
        Write-Log "PrintBrm.exe not found on system at $printBrm." -Level WARNING
        return $false
    }
    
    # Space-Safe Temp Directory for PrintBrm (bypasses space pathing issues)
    $guidStr = [Guid]::NewGuid().Guid.Substring(0,8)
    $tempWorkDir = Join-Path -Path $env:TEMP -ChildPath "PrinterPackager_Temp_$guidStr"
    if (Test-Path $tempWorkDir) { Remove-Item -Path $tempWorkDir -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -Path $tempWorkDir -ItemType Directory -Force | Out-Null
    
    $tempFullCab = Join-Path -Path $tempWorkDir -ChildPath "full_backup.printerExport"
    $tempUnpackDir = Join-Path -Path $tempWorkDir -ChildPath "unpacked"
    
    try {
        # Step 1: Export full machine print backup using -NOBIN
        Write-Log "Step 1/5: Performing initial -NOBIN print server backup..." -Level INFO
        & $printBrm -B -NOBIN -F "$tempFullCab"
        if (-not (Test-Path $tempFullCab)) {
            Write-Log "PrintBrm initial backup failed." -Level WARNING
            return $false
        }
        
        # Step 2: Unpack CAB to working directory
        Write-Log "Step 2/5: Unpacking backup payload to temporary workspace..." -Level INFO
        & $printBrm -R -F "$tempFullCab" -D "$tempUnpackDir"
        if (-not (Test-Path $tempUnpackDir)) {
            Write-Log "PrintBrm unpack failed." -Level WARNING
            return $false
        }
        
        # Step 3: Modify XML manifests to keep ONLY the selected printer queue, port, and driver
        Write-Log "Step 3/5: Filtering XML manifests for Printer '$PrinterName', Port '$PortName', Driver '$DriverName'..." -Level INFO
        
        # 3a. Filter BrmPrinters.xml & track surviving printer XML file
        $survivingXmlFiles = [System.Collections.Generic.List[string]]::new()
        $brmPrintersXmlPath = Join-Path -Path $tempUnpackDir -ChildPath "BrmPrinters.xml"
        if (Test-Path $brmPrintersXmlPath) {
            [xml]$xmlP = Get-Content -Path $brmPrintersXmlPath
            if ($xmlP.PRINTERS -and $xmlP.PRINTERS.PRINTQUEUE) {
                $queues = @($xmlP.PRINTERS.PRINTQUEUE)
                foreach ($q in $queues) {
                    if ($q.PrinterName -eq $PrinterName) {
                        if ($q.FileName) { $survivingXmlFiles.Add($q.FileName) }
                    } else {
                        [void]$xmlP.PRINTERS.RemoveChild($q)
                    }
                }
                $xmlP.Save($brmPrintersXmlPath)
                Write-Log "BrmPrinters.xml filtered successfully." -Level SUCCESS
            }
        }
        
        # 3b. Filter BrmPorts.xml
        $brmPortsXmlPath = Join-Path -Path $tempUnpackDir -ChildPath "BrmPorts.xml"
        if ($PortName -and (Test-Path $brmPortsXmlPath)) {
            [xml]$xmlPt = Get-Content -Path $brmPortsXmlPath
            if ($xmlPt.PRINTERPORTS -and $xmlPt.PRINTERPORTS.SPM) {
                $ports = @($xmlPt.PRINTERPORTS.SPM)
                foreach ($pt in $ports) {
                    if ($pt.PortName -ne $PortName) {
                        [void]$xmlPt.PRINTERPORTS.RemoveChild($pt)
                    }
                }
                $xmlPt.Save($brmPortsXmlPath)
                Write-Log "BrmPorts.xml filtered successfully." -Level SUCCESS
            }
        }
        
        # 3c. Filter BrmDrivers.xml
        $brmDriversXmlPath = Join-Path -Path $tempUnpackDir -ChildPath "BrmDrivers.xml"
        if ($DriverName -and (Test-Path $brmDriversXmlPath)) {
            [xml]$xmlDr = Get-Content -Path $brmDriversXmlPath
            if ($xmlDr.PRINTERDRIVERS -and $xmlDr.PRINTERDRIVERS.DRIVER) {
                $drivers = @($xmlDr.PRINTERDRIVERS.DRIVER)
                foreach ($dr in $drivers) {
                    if ($dr.DriverName -ne $DriverName) {
                        [void]$xmlDr.PRINTERDRIVERS.RemoveChild($dr)
                    }
                }
                $xmlDr.Save($brmDriversXmlPath)
                Write-Log "BrmDrivers.xml filtered successfully." -Level SUCCESS
            }
        }
        
        # Step 4: Prune unused queue XML files from Printers/ directory
        Write-Log "Step 4/5: Pruning unused queue definition files from Printers/ folder..." -Level INFO
        $printersDir = Join-Path -Path $tempUnpackDir -ChildPath "Printers"
        if (Test-Path $printersDir) {
            $printerFiles = Get-ChildItem -Path $printersDir -Filter "*.xml" -ErrorAction SilentlyContinue
            foreach ($pf in $printerFiles) {
                if ($survivingXmlFiles.Count -gt 0 -and -not ($survivingXmlFiles.Contains($pf.Name))) {
                    Remove-Item -Path $pf.FullName -Force -ErrorAction SilentlyContinue
                }
            }
            Write-Log "Printers/ folder pruned successfully." -Level SUCCESS
        }
        
        # Step 5: Repack filtered & pruned directory to destination .printerExport file
        Write-Log "Step 5/5: Repacking single-printer export package via -NOBIN..." -Level INFO
        & $printBrm -B -NOBIN -F "$DestinationFile" -D "$tempUnpackDir"
        if (Test-Path $DestinationFile) {
            $sizeKb = [math]::Round(((Get-Item $DestinationFile).Length / 1KB), 1)
            Write-Log "PrintBrm single-printer export successful for '$PrinterName' (Package Size: ${sizeKb} KB)." -Level SUCCESS
            return $true
        } else {
            Write-Log "PrintBrm repack failed." -Level WARNING
            return $false
        }
    }
    catch {
        Write-Log "PrintBrm export error: $_" -Level WARNING
        return $false
    }
    finally {
        Remove-Item -Path $tempWorkDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Function: Show In-App Help / User Guide Dialog
function Show-UserGuide {
    [xml]$helpXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Printer Packager - User Guide and Recommended Workflow" Height="740" Width="860"
        WindowStartupLocation="CenterOwner" Background="#1E1E1E" Foreground="#FFFFFF"
        FontFamily="Segoe UI" FontSize="13">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,15">
            <TextBlock Text="Intune &amp; Standalone Printer Packager Guide" FontSize="20" FontWeight="Bold" Foreground="#0078D4"/>
            <TextBlock Text="Step-by-step instructions, option details, external tool download links, and deployment modes." Foreground="#AAAAAA" Margin="0,4,0,0"/>
        </StackPanel>

        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
            <StackPanel Margin="0,0,10,0">

                <Border Background="#2D2D30" CornerRadius="6" Padding="14" Margin="0,0,0,12">
                    <StackPanel>
                        <TextBlock Text="Prerequisites &amp; Toolkit Download Links" FontSize="15" FontWeight="Bold" Foreground="#00FF66" Margin="0,0,0,8"/>
                        <TextBlock Text="Before packaging, ensure you have downloaded the required toolkits:" Foreground="#CCCCCC" Margin="0,0,0,6" TextWrapping="Wrap"/>
                        
                        <TextBlock Margin="0,2,0,6">
                            <Run Text="1. PSAppDeployToolkit (Tested with v4.1.8): " FontWeight="SemiBold" Foreground="#FFFFFF"/>
                            <Hyperlink NavigateUri="https://github.com/PSAppDeployToolkit/PSAppDeployToolkit/releases/tag/4.1.8" Foreground="#569CD6">
                                https://github.com/PSAppDeployToolkit/PSAppDeployToolkit (v4.1.8)
                            </Hyperlink>
                        </TextBlock>
                        
                        <TextBlock Margin="0,0,0,4">
                            <Run Text="2. Microsoft Win32 Content Prep Tool (Tested with v1.8.7): " FontWeight="SemiBold" Foreground="#FFFFFF"/>
                            <Hyperlink NavigateUri="https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/releases/tag/v1.8.7" Foreground="#569CD6">
                                https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool (v1.8.7)
                            </Hyperlink>
                        </TextBlock>
                    </StackPanel>
                </Border>

                <Border Background="#2D2D30" CornerRadius="6" Padding="14" Margin="0,0,0,12">
                    <StackPanel>
                        <TextBlock Text="Recommended Workflow (Step-by-Step)" FontSize="15" FontWeight="Bold" Foreground="#00FF66" Margin="0,0,0,8"/>
                        <TextBlock Text="Step 1: Select Printer - Click 'Refresh Printers' and select a locally installed printer queue from the grid." Margin="0,0,0,6" TextWrapping="Wrap"/>
                        <TextBlock Text="Step 2: Set App Metadata - Enter App Vendor (e.g. Company), App Name, and Version (e.g. 1.0.0)." Margin="0,0,0,6" TextWrapping="Wrap"/>
                        <TextBlock Text="Step 3: Select Packaging Options - Choose preference capture, driver extraction, .intunewin, or standalone .zip." Margin="0,0,0,6" TextWrapping="Wrap"/>
                        <TextBlock Text="Step 4: Execute Packaging - Click 'Package Printer for Intune' or 'Create Local Standalone Package'." Margin="0,0,0,4" TextWrapping="Wrap"/>
                    </StackPanel>
                </Border>

                <Border Background="#2D2D30" CornerRadius="6" Padding="14" Margin="0,0,0,12">
                    <StackPanel>
                        <TextBlock Text="Packaging Options Breakdown" FontSize="15" FontWeight="Bold" Foreground="#00FF66" Margin="0,0,0,8"/>
                        
                        <TextBlock Text="- Capture DevMode Preferences (PrintBrm &amp; PrintUI)" FontWeight="SemiBold" Foreground="#FFFFFF" Margin="0,4,0,2"/>
                        <TextBlock Text="Exports binary printer queue configurations, duplexing defaults, paper sizes, color settings, and tray selections so target computers inherit exact local preferences." Foreground="#CCCCCC" TextWrapping="Wrap" Margin="0,0,0,8"/>

                        <TextBlock Text="- Extract Driver Store Files (FileRepository)" FontWeight="SemiBold" Foreground="#FFFFFF" Margin="0,4,0,2"/>
                        <TextBlock Text="Extracts required driver INFs and dependency binaries directly from System32\DriverStore so target clients can install the driver without internet or print server access." Foreground="#CCCCCC" TextWrapping="Wrap" Margin="0,0,0,8"/>

                        <TextBlock Text="- Compile .intunewin Package (IntuneWinAppUtil)" FontWeight="SemiBold" Foreground="#FFFFFF" Margin="0,4,0,2"/>
                        <TextBlock Text="Runs IntuneWinAppUtil.exe to generate a ready-to-upload .intunewin package for Intune Win32 App deployment." Foreground="#CCCCCC" TextWrapping="Wrap" Margin="0,0,0,8"/>

                        <TextBlock Text="- Create Standalone Zip Archive (.zip)" FontWeight="SemiBold" Foreground="#FFFFFF" Margin="0,4,0,2"/>
                        <TextBlock Text="Compresses the complete PSADT folder and local helper scripts into a single .zip archive for easy distribution to IT staff or non-Intune endpoints." Foreground="#CCCCCC" TextWrapping="Wrap" Margin="0,0,0,8"/>

                        <TextBlock Text="- Set as Default Printer on Target Device" FontWeight="SemiBold" Foreground="#FFFFFF" Margin="0,4,0,2"/>
                        <TextBlock Text="Configures the target workstation to automatically set this printer as the default system printer upon deployment." Foreground="#CCCCCC" TextWrapping="Wrap"/>
                    </StackPanel>
                </Border>

                <Border Background="#2D2D30" CornerRadius="6" Padding="14" Margin="0,0,0,12">
                    <StackPanel>
                        <TextBlock Text="Deployment Use Cases" FontSize="15" FontWeight="Bold" Foreground="#00FF66" Margin="0,0,0,8"/>

                        <TextBlock Text="Option A: Deployment via Microsoft Intune" FontWeight="SemiBold" Foreground="#FFFFFF" Margin="0,4,0,2"/>
                        <TextBlock Text="1. Upload the generated .intunewin file to Intune (Apps -> Windows app (Win32))." Foreground="#CCCCCC" TextWrapping="Wrap"/>
                        <TextBlock Text="2. Install Command: Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Silent" Foreground="#CCCCCC" TextWrapping="Wrap"/>
                        <TextBlock Text="3. Uninstall Command: Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent" Foreground="#CCCCCC" TextWrapping="Wrap"/>
                        <TextBlock Text="4. Detection Rule: Upload Detection.ps1 from the generated package folder." Foreground="#CCCCCC" TextWrapping="Wrap" Margin="0,0,0,8"/>

                        <TextBlock Text="Option B: Local Standalone Deployment (No Intune)" FontWeight="SemiBold" Foreground="#FFFFFF" Margin="0,4,0,2"/>
                        <TextBlock Text="1. Extract the generated package folder or .zip file on the target computer." Foreground="#CCCCCC" TextWrapping="Wrap"/>
                        <TextBlock Text="2. To Install: Right-click 'Install-PrinterLocal.cmd' and select 'Run as Administrator' (or double-click)." Foreground="#CCCCCC" TextWrapping="Wrap"/>
                        <TextBlock Text="3. To Uninstall: Right-click 'Uninstall-PrinterLocal.cmd' and select 'Run as Administrator' (or double-click)." Foreground="#CCCCCC" TextWrapping="Wrap"/>
                        <TextBlock Text="4. Progress and logs will display interactively on screen during deployment." Foreground="#CCCCCC" TextWrapping="Wrap"/>
                    </StackPanel>
                </Border>

            </StackPanel>
        </ScrollViewer>

        <Button Name="BtnCloseHelp" Grid.Row="2" Content="Close Guide" Background="#0078D4" Width="120" HorizontalAlignment="Right" Margin="0,10,0,0"/>
    </Grid>
</Window>
"@
    $helpReader = (New-Object System.Xml.XmlNodeReader $helpXaml)
    $helpWindow = [System.Windows.Markup.XamlReader]::Load($helpReader)
    $helpWindow.Owner = $window

    $helpWindow.AddHandler([System.Windows.Documents.Hyperlink]::RequestNavigateEvent, [System.Windows.Navigation.RequestNavigateEventHandler]{
        param($sender, $e)
        try {
            [System.Diagnostics.Process]::Start((New-Object System.Diagnostics.ProcessStartInfo($e.Uri.AbsoluteUri)))
        } catch {
            [System.Diagnostics.Process]::Start("explorer.exe", "`"$($e.Uri.AbsoluteUri)`"")
        }
        $e.Handled = $true
    })

    $btnCloseHelp = $helpWindow.FindName('BtnCloseHelp')
    $btnCloseHelp.Add_Click({ $helpWindow.Close() })
    $helpWindow.ShowDialog() | Out-Null
}

# Function: Generate Standalone Local Install/Uninstall Helper Scripts
function Generate-LocalInstallScripts {
    param(
        [string]$PackageDir,
        [string]$PrinterName
    )
    
    Write-Log "Generating local standalone execution scripts..." -Level INFO
    
    # 1. Install-PrinterLocal.ps1
    $installPs1Path = Join-Path -Path $PackageDir -ChildPath "Install-PrinterLocal.ps1"
    $installPs1Content = @"
<#
.SYNOPSIS
    Local Interactive Printer Installer for '$PrinterName'
.DESCRIPTION
    Launches PSADT v4 deployment for printer installation on the local device.
#>
`$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
`$principal = [System.Security.Principal.WindowsPrincipal]`$identity
if (-not `$principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Administrator privileges required. Requesting elevation..."
    `$powershell = (Get-Command powershell, pwsh -ErrorAction SilentlyContinue)[0].Path
    Start-Process -FilePath `$powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"`$PSScriptRoot\Install-PrinterLocal.ps1`"" -Verb RunAs
    exit
}

`$scriptDir = `$PSScriptRoot
`$exePath = Join-Path -Path `$scriptDir -ChildPath "Invoke-AppDeployToolkit.exe"
`$ps1Path = Join-Path -Path `$scriptDir -ChildPath "Invoke-AppDeployToolkit.ps1"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Installing Printer: $PrinterName" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

if (Test-Path `$exePath) {
    Write-Host "Launching PSADT executable runner..." -ForegroundColor Green
    Start-Process -FilePath `$exePath -ArgumentList "-DeploymentType Install -DeployMode Interactive" -Wait
} elseif (Test-Path `$ps1Path) {
    Write-Host "Launching PSADT script directly..." -ForegroundColor Green
    `$powershell = (Get-Command powershell, pwsh -ErrorAction SilentlyContinue)[0].Path
    & `$powershell -NoProfile -ExecutionPolicy Bypass -File `$ps1Path -DeploymentType Install -DeployMode Interactive
} else {
    Write-Error "Could not find Invoke-AppDeployToolkit.exe or Invoke-AppDeployToolkit.ps1 in `$scriptDir"
}
"@
    Write-BOMFreeUtf8File -Path $installPs1Path -Content $installPs1Content

    # 2. Uninstall-PrinterLocal.ps1
    $uninstallPs1Path = Join-Path -Path $PackageDir -ChildPath "Uninstall-PrinterLocal.ps1"
    $uninstallPs1Content = @"
<#
.SYNOPSIS
    Local Interactive Printer Uninstaller for '$PrinterName'
.DESCRIPTION
    Launches PSADT v4 deployment for printer uninstallation on the local device.
#>
`$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
`$principal = [System.Security.Principal.WindowsPrincipal]`$identity
if (-not `$principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Administrator privileges required. Requesting elevation..."
    `$powershell = (Get-Command powershell, pwsh -ErrorAction SilentlyContinue)[0].Path
    Start-Process -FilePath `$powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"`$PSScriptRoot\Uninstall-PrinterLocal.ps1`"" -Verb RunAs
    exit
}

`$scriptDir = `$PSScriptRoot
`$exePath = Join-Path -Path `$scriptDir -ChildPath "Invoke-AppDeployToolkit.exe"
`$ps1Path = Join-Path -Path `$scriptDir -ChildPath "Invoke-AppDeployToolkit.ps1"

Write-Host "==========================================================" -ForegroundColor Yellow
Write-Host " Uninstalling Printer: $PrinterName" -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Yellow

if (Test-Path `$exePath) {
    Write-Host "Launching PSADT executable runner..." -ForegroundColor Yellow
    Start-Process -FilePath `$exePath -ArgumentList "-DeploymentType Uninstall -DeployMode Interactive" -Wait
} elseif (Test-Path `$ps1Path) {
    Write-Host "Launching PSADT script directly..." -ForegroundColor Yellow
    `$powershell = (Get-Command powershell, pwsh -ErrorAction SilentlyContinue)[0].Path
    & `$powershell -NoProfile -ExecutionPolicy Bypass -File `$ps1Path -DeploymentType Uninstall -DeployMode Interactive
} else {
    Write-Error "Could not find Invoke-AppDeployToolkit.exe or Invoke-AppDeployToolkit.ps1 in `$scriptDir"
}
"@
    Write-BOMFreeUtf8File -Path $uninstallPs1Path -Content $uninstallPs1Content

    # 3. Install-PrinterLocal.cmd
    $installCmdPath = Join-Path -Path $PackageDir -ChildPath "Install-PrinterLocal.cmd"
    $installCmdContent = @"
@echo off
:: Self-elevating launcher script for Install-PrinterLocal.ps1 ($PrinterName)
net session >nul 2>&1
if %errorlevel% == 0 (
    goto :admin
) else (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process cmd -ArgumentList '/c `"%~dp0Install-PrinterLocal.cmd`"' -Verb RunAs"
    exit /b
)
:admin
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-PrinterLocal.ps1"
pause
"@
    Write-BOMFreeUtf8File -Path $installCmdPath -Content $installCmdContent

    # 4. Uninstall-PrinterLocal.cmd
    $uninstallCmdPath = Join-Path -Path $PackageDir -ChildPath "Uninstall-PrinterLocal.cmd"
    $uninstallCmdContent = @"
@echo off
:: Self-elevating launcher script for Uninstall-PrinterLocal.ps1 ($PrinterName)
net session >nul 2>&1
if %errorlevel% == 0 (
    goto :admin
) else (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process cmd -ArgumentList '/c `"%~dp0Uninstall-PrinterLocal.cmd`"' -Verb RunAs"
    exit /b
)
:admin
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall-PrinterLocal.ps1"
pause
"@
    Write-BOMFreeUtf8File -Path $uninstallCmdPath -Content $uninstallCmdContent

    Write-Log "Generated standalone local scripts: Install-PrinterLocal.cmd/.ps1 and Uninstall-PrinterLocal.cmd/.ps1" -Level SUCCESS
}

# Help / How-To Button Event
if ($BtnHelp) {
    $BtnHelp.Add_Click({
        Show-UserGuide
    })
}

# Generate Detection Script Button Click
$BtnGenerateDetection.Add_Click({
    $selected = $GridPrinters.SelectedItem
    if (-not $selected) {
        [System.Windows.Forms.MessageBox]::Show("Please select a printer from the list first.", "No Printer Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $outputFolder = $TxtOutputDir.Text.Trim()
    if (-not (Test-Path $outputFolder)) {
        New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null
    }
    
    $detectionPath = Join-Path -Path $outputFolder -ChildPath "Detection_$($selected.Name -replace '[^\w\-\.]', '_').ps1"
    
    $detectionScript = @"
# Intune Win32 App Detection Script for Printer: $($selected.Name)
`$PrinterName = "$($selected.Name)"
`$PortName = "$($selected.PortName)"

`$printer = Get-Printer -Name `$PrinterName -ErrorAction SilentlyContinue
`$port = Get-PrinterPort -Name `$PortName -ErrorAction SilentlyContinue

if (`$printer -and `$port) {
    Write-Output "Installed: Printer '$PrinterName' and Port '$PortName' exist."
    Exit 0
} else {
    Exit 1
}
"@
    
    Write-BOMFreeUtf8File -Path $detectionPath -Content $detectionScript
    Write-Log "Generated Detection Script at $detectionPath" -Level SUCCESS
    [System.Windows.Forms.MessageBox]::Show("Detection script generated successfully at:`n$detectionPath", "Detection Script Created", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Core Package Builder Function
function Start-PrinterPackageBuild {
    param(
        [bool]$IsStandaloneMode = $false
    )
    
    $selected = $GridPrinters.SelectedItem
    if (-not $selected) {
        [System.Windows.Forms.MessageBox]::Show("Please select a printer from the list to package.", "No Printer Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Check IP address on selected printer
    if ([string]::IsNullOrWhiteSpace($selected.PrinterHostAddress)) {
        $warnResult = [System.Windows.Forms.MessageBox]::Show(
            "The selected printer '$($selected.Name)' does not have an IP address/HostAddress specified.`n`nDo you still want to proceed with packaging?",
            "Missing IP Address Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($warnResult -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }
    
    $templatePath = $TxtTemplatePath.Text.Trim()
    if (-not (Test-Path $templatePath)) {
        [System.Windows.Forms.MessageBox]::Show("PSADT v4 Template directory not found at:`n$templatePath`n`nPlease specify a valid PSADT template path.", "Template Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $intuneWinUtilPath = $TxtIntuneWinUtilPath.Text.Trim()
    if ($ChkCreateIntuneWin.IsChecked -and -not (Test-Path $intuneWinUtilPath)) {
        [System.Windows.Forms.MessageBox]::Show("IntuneWinAppUtil.exe not found at:`n$intuneWinUtilPath`n`nPlease specify a valid path to IntuneWinAppUtil.exe or uncheck the IntuneWin compilation option.", "IntuneWinAppUtil Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $appVendor = $TxtAppVendor.Text.Trim()
    $appName = $TxtAppName.Text.Trim()
    $appVersion = $TxtAppVersion.Text.Trim()
    $outDir = $TxtOutputDir.Text.Trim()
    
    if (-not $appName) { $appName = "Printer - $($selected.Name)" }
    if (-not $appVendor) { $appVendor = "Company" }
    if (-not $appVersion) { $appVersion = "1.0.0" }
    
    $safePackageFolder = "$appVendor - $appName - $appVersion" -replace '[^\w\-\.\s]', '_'
    $packageDir = Join-Path -Path $outDir -ChildPath $safePackageFolder
    
    Write-Log "Starting PSADT v4 packaging for printer '$($selected.Name)' (Standalone Mode: $IsStandaloneMode)..." -Level INFO
    Write-Log "Package Output Directory: $packageDir" -Level INFO
    $TxtStatus.Text = "Building PSADT package..."
    
    try {
        # 1. Clone PSADT Template
        if (Test-Path $packageDir) {
            Write-Log "Cleaning existing output folder: $packageDir" -Level WARNING
            Remove-Item -Path $packageDir -Recurse -Force
        }
        
        Write-Log "Cloning PSADT v4 template files from $templatePath..." -Level INFO
        try {
            Copy-Item -Path $templatePath -Destination $packageDir -Recurse -Force -ErrorAction Stop
        } catch {
            throw "Failed to copy PSADT template files: $_"
        }
        
        $filesDir = Join-Path -Path $packageDir -ChildPath "Files"
        $driversSubDir = Join-Path -Path $filesDir -ChildPath "Drivers"
        New-Item -Path $driversSubDir -ItemType Directory -Force | Out-Null
        
        # 2. Extract Driver Files if checked
        $driverFolderName = ""
        if ($ChkExtractDriver.IsChecked) {
            $safeDriverName = $selected.DriverName -replace '[^\w\-\.]', '_'
            $targetDriverDir = Join-Path -Path $driversSubDir -ChildPath $safeDriverName
            $exported = Export-PrinterDriverFiles -DriverName $selected.DriverName -InfPath $selected.InfPath -DestinationDir $targetDriverDir
            if ($exported) {
                $driverFolderName = "Drivers\$safeDriverName"
                $drvSizeMb = [math]::Round(((Get-ChildItem -Path $targetDriverDir -Recurse | Measure-Object -Sum Length).Sum / 1MB), 2)
                Write-Log "Extracted Driver Package Size: $drvSizeMb MB" -Level INFO
            }
        }
        
        # 3a. Export Printer Preferences (PrintBrm) if checked
        $usePrintBrm = $false
        if ($ChkCapturePreferences.IsChecked) {
            $exportFilePath = Join-Path -Path $filesDir -ChildPath "printer.printerExport"
            $usePrintBrm = Export-PrinterPreferences -PrinterName $selected.Name -PortName $selected.PortName -DriverName $selected.DriverName -DestinationFile $exportFilePath
        }
        
        # 3b. Export full vendor-specific printer preferences (PrintUI.dll /Ss) if checked
        $useSettingsDat = $false
        if ($ChkCapturePreferences.IsChecked) {
            $settingsDatPath = Join-Path -Path $filesDir -ChildPath "printer_settings.dat"
            $useSettingsDat = Export-PrinterSettingsDat -PrinterName $selected.Name -DestinationFile $settingsDatPath
        }
        
        # 4. Create Metadata Config JSON
        $configObject = [PSCustomObject]@{
            AppName            = $appName
            AppVendor          = $appVendor
            AppVersion         = $appVersion
            PrinterName        = $selected.Name
            DriverName         = $selected.DriverName
            DriverFolder       = $driverFolderName
            PortName           = $selected.PortName
            PrinterHostAddress = $selected.PrinterHostAddress
            PortNumber         = $selected.PortNumber
            Protocol           = $selected.Protocol
            Comment            = $selected.Comment
            Location           = $selected.Location
            IsDefault          = [bool]$ChkSetDefault.IsChecked
            UsePrintBrm        = $usePrintBrm
            UseSettingsDat     = $useSettingsDat
            DuplexingMode      = $selected.DuplexingMode
            Color              = $selected.Color
            PaperSize          = $selected.PaperSize
        }
        
        $configJsonPath = Join-Path -Path $filesDir -ChildPath "printer_config.json"
        $configJsonStr = $configObject | ConvertTo-Json -Depth 5
        Write-BOMFreeUtf8File -Path $configJsonPath -Content $configJsonStr
        Write-Log "Created payload configuration: printer_config.json" -Level SUCCESS
        
        # 5. Inject Logic into Invoke-AppDeployToolkit.ps1
        $psadtScriptPath = Join-Path -Path $packageDir -ChildPath "Invoke-AppDeployToolkit.ps1"
        if (Test-Path $psadtScriptPath) {
            $psadtCode = Get-Content -Path $psadtScriptPath -Raw
            
            # Literal String Replacements for Metadata
            $psadtCode = $psadtCode.Replace("AppVendor = ''", "AppVendor = '$appVendor'")
            $psadtCode = $psadtCode.Replace("AppName = ''", "AppName = '$appName'")
            $psadtCode = $psadtCode.Replace("AppVersion = ''", "AppVersion = '$appVersion'")
            
            # Install Phase Insertion (Using StringArray @(...) for Start-ADTProcess -ArgumentList)
            $installCode = @"
    ## ================================================
    ## MARK: Install
    ## ================================================
    `$adtSession.InstallPhase = `$adtSession.DeploymentType

    Write-ADTLogEntry -Message "Starting printer deployment via PSADT..."

    # Resolve Files Directory robustly for PSADT v4
    `$dirFiles = if (`$adtSession.DirFiles) { `$adtSession.DirFiles } else { Join-Path -Path `$adtSession.ScriptDirectory -ChildPath "Files" }

    # Load Config JSON
    `$configFile = Join-Path -Path `$dirFiles -ChildPath "printer_config.json"
    if (Test-Path -Path `$configFile) {
        `$config = Get-Content -Path `$configFile -Raw | ConvertFrom-Json
    } else {
        throw "printer_config.json not found in `$dirFiles"
    }

    # Space-Safe Staging Directory for PrintBRM & Pnputil
    `$stagingDir = Join-Path -Path `$env:TEMP -ChildPath "PSADT_Printer_Deploy"
    if (Test-Path -Path `$stagingDir) { Remove-Item -Path `$stagingDir -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -Path `$stagingDir -ItemType Directory -Force | Out-Null

    # Resolve native 64-bit System32 directory to bypass WOW64 redirection on 32-bit (x86) processes
    `$system32Dir = if (Test-Path -Path "`$env:SystemRoot\Sysnative") { "`$env:SystemRoot\Sysnative" } else { Join-Path -Path `$env:SystemRoot -ChildPath "System32" }
    `$pnputilExe = Join-Path -Path `$system32Dir -ChildPath "pnputil.exe"
    `$printBrmExe = Join-Path -Path `$system32Dir -ChildPath "spool\tools\PrintBrm.exe"
    `$rundll32Exe = Join-Path -Path `$system32Dir -ChildPath "rundll32.exe"

    # 1. Install Driver into Driver Store & Spooler (Copies FULL driver folder to staging)
    if (`$config.DriverFolder) {
        `$driverPath = Join-Path -Path `$dirFiles -ChildPath `$config.DriverFolder
        if (Test-Path -Path `$driverPath) {
            Write-ADTLogEntry -Message "Staging driver package from `$driverPath..."
            `$stagedDriverDir = Join-Path -Path `$stagingDir -ChildPath "DriverPackage"
            Copy-Item -Path `$driverPath -Destination `$stagedDriverDir -Recurse -Force
            `$infFiles = Get-ChildItem -Path `$stagedDriverDir -Filter "*.inf" -Recurse
            foreach (`$inf in `$infFiles) {
                Start-ADTProcess -FilePath `$pnputilExe -ArgumentList @("/add-driver", "`$(`$inf.FullName)", "/install") -WindowStyle Hidden
            }
        }
    }

    if (-not (Get-PrinterDriver -Name `$config.DriverName -ErrorAction SilentlyContinue)) {
        Write-ADTLogEntry -Message "Adding Printer Driver to Spooler: `$(`$config.DriverName)"
        Add-PrinterDriver -Name `$config.DriverName -ErrorAction SilentlyContinue
    }

    # 2. Create Printer Port
    if (`$config.PortName -and -not (Get-PrinterPort -Name `$config.PortName -ErrorAction SilentlyContinue)) {
        Write-ADTLogEntry -Message "Creating TCP/IP Printer Port: `$(`$config.PortName) (`$(`$config.PrinterHostAddress))"
        if (`$config.Protocol -eq 2) {
            Add-PrinterPort -Name `$config.PortName -LprHostAddress `$config.PrinterHostAddress -LprQueueName "lp" -ErrorAction SilentlyContinue
        } else {
            `$portNum = if (`$config.PortNumber) { `$config.PortNumber } else { 9100 }
            Add-PrinterPort -Name `$config.PortName -PrinterHostAddress `$config.PrinterHostAddress -PortNumber `$portNum -ErrorAction SilentlyContinue
        }
    }

    # 3. Restore Printer Configuration & DevMode Preferences via PrintBRM
    `$exportFile = Join-Path -Path `$dirFiles -ChildPath "printer.printerExport"
    if (`$config.UsePrintBrm -and (Test-Path -Path `$exportFile)) {
        Write-ADTLogEntry -Message "Restoring printer settings & DevMode via PrintBrm..."
        `$stagedExport = Join-Path -Path `$stagingDir -ChildPath "printer.printerExport"
        Copy-Item -Path `$exportFile -Destination `$stagedExport -Force
        
        if (Test-Path -Path `$printBrmExe) {
            Start-ADTProcess -FilePath `$printBrmExe -ArgumentList @("-R", "-F", "`$stagedExport", "-O", "FORCE") -WindowStyle Hidden
        }
    }

    # 4. Ensure Printer Queue exists and parameters match
    if (-not (Get-Printer -Name `$config.PrinterName -ErrorAction SilentlyContinue)) {
        Write-ADTLogEntry -Message "Creating Printer Queue: `$(`$config.PrinterName)"
        `$printerParams = @{
            Name = `$config.PrinterName
            DriverName = `$config.DriverName
            PortName = `$config.PortName
        }
        if (`$config.Comment) { `$printerParams.Comment = `$config.Comment }
        if (`$config.Location) { `$printerParams.Location = `$config.Location }
        Add-Printer @printerParams -ErrorAction SilentlyContinue
    } else {
        Set-Printer -Name `$config.PrinterName -PortName `$config.PortName -ErrorAction SilentlyContinue
    }

    # 5. Restore full vendor-specific printer preferences via PrintUI.dll /Sr
    `$settingsDat = Join-Path -Path `$dirFiles -ChildPath "printer_settings.dat"
    if (`$config.UseSettingsDat -and (Test-Path -Path `$settingsDat)) {
        Write-ADTLogEntry -Message "Restoring full vendor-specific preferences via PrintUI.dll /Sr..."
        `$stagedDat = Join-Path -Path `$stagingDir -ChildPath "printer_settings.dat"
        Copy-Item -Path `$settingsDat -Destination `$stagedDat -Force
        try {
            Start-ADTProcess -FilePath `$rundll32Exe -ArgumentList @("printui.dll,PrintUIEntry", "/Sr", "/n", "`$(`$config.PrinterName)", "/a", "`$stagedDat", "g", "d", "r") -WindowStyle Hidden
        } catch {
            Write-ADTLogEntry -Message "PrintUI.dll /Sr preferences restore notice: `$_"
        }
    }

    # 6. Enforce High-Level Print Configuration (Duplex, Color, Paper Size)
    if (`$config.DuplexingMode -or `$config.PaperSize) {
        Write-ADTLogEntry -Message "Applying Print Configuration settings..."
        `$setConfigParams = @{ PrinterName = `$config.PrinterName }
        if (`$config.DuplexingMode) { `$setConfigParams.Add('DuplexingMode', `$config.DuplexingMode) }
        if (`$config.Color -ne `$null) { `$setConfigParams.Add('Color', [bool]`$config.Color) }
        if (`$config.PaperSize) { `$setConfigParams.Add('PaperSize', `$config.PaperSize) }
        Set-PrintConfiguration @setConfigParams -ErrorAction SilentlyContinue
    }

    # 7. Set Default Printer Option (CIM approach)
    if (`$config.IsDefault) {
        Write-ADTLogEntry -Message "Setting `$(`$config.PrinterName) as default printer..."
        `$safePName = `$config.PrinterName -replace "'", "''"
        `$cimPrinter = Get-CimInstance -ClassName Win32_Printer -Filter "Name='`$safePName'" -ErrorAction SilentlyContinue
        if (`$cimPrinter) {
            Invoke-CimMethod -InputObject `$cimPrinter -MethodName SetDefaultPrinter -ErrorAction SilentlyContinue | Out-Null
        }
    }

    # Cleanup Staging
    Remove-Item -Path `$stagingDir -Recurse -Force -ErrorAction SilentlyContinue
"@
            
            # Uninstall Phase Insertion (Includes driver removal)
            $uninstallCode = @"
    ## ================================================
    ## MARK: Uninstall
    ## ================================================
    `$adtSession.InstallPhase = `$adtSession.DeploymentType

    `$dirFiles = if (`$adtSession.DirFiles) { `$adtSession.DirFiles } else { Join-Path -Path `$adtSession.ScriptDirectory -ChildPath "Files" }
    `$configFile = Join-Path -Path `$dirFiles -ChildPath "printer_config.json"
    if (Test-Path -Path `$configFile) {
        `$config = Get-Content -Path `$configFile -Raw | ConvertFrom-Json
        
        Write-ADTLogEntry -Message "Removing Printer Queue: `$(`$config.PrinterName)"
        Remove-Printer -Name `$config.PrinterName -ErrorAction SilentlyContinue
        
        Write-ADTLogEntry -Message "Removing Printer Port: `$(`$config.PortName)"
        Remove-PrinterPort -Name `$config.PortName -ErrorAction SilentlyContinue

        Write-ADTLogEntry -Message "Removing Printer Driver: `$(`$config.DriverName)"
        Remove-PrinterDriver -Name `$config.DriverName -ErrorAction SilentlyContinue
    }
"@
            
            # Robust Marker Replacement Strategy
            $psadtCode = $psadtCode -replace '## MARK: Install[\s\S]*?## MARK: Post-Install', "## MARK: Install`r`n    ##================================================`r`n$installCode`r`n`r`n    ## MARK: Post-Install"
            $psadtCode = $psadtCode -replace '## MARK: Uninstall[\s\S]*?## MARK: Post-Uninstallation', "## MARK: Uninstall`r`n    ##================================================`r`n$uninstallCode`r`n`r`n    ## MARK: Post-Uninstallation"
            
            # Post-Injection Validation
            if (-not ($psadtCode.Contains("Starting printer deployment via PSADT..."))) {
                throw "PSADT v4 code injection validation failed. Could not locate '## MARK: Install' block in Invoke-AppDeployToolkit.ps1."
            }
            
            Write-BOMFreeUtf8File -Path $psadtScriptPath -Content $psadtCode
            Write-Log "Updated Invoke-AppDeployToolkit.ps1 with printer deployment logic." -Level SUCCESS
        }
        
        # 6a. Generate Detection Script inside Package
        $detectionScriptPath = Join-Path -Path $packageDir -ChildPath "Detection.ps1"
        $detCode = @"
# Intune Detection Script for Printer: $($selected.Name)
`$PrinterName = "$($selected.Name)"
`$PortName = "$($selected.PortName)"

`$printer = Get-Printer -Name `$PrinterName -ErrorAction SilentlyContinue
`$port = Get-PrinterPort -Name `$PortName -ErrorAction SilentlyContinue

if (`$printer -and `$port) {
    Write-Output "Installed: Printer '$PrinterName' and Port '$PortName' exist."
    Exit 0
} else {
    Exit 1
}
"@
        Write-BOMFreeUtf8File -Path $detectionScriptPath -Content $detCode
        Write-Log "Generated Detection.ps1 in package folder." -Level SUCCESS
        
        # 6b. Generate Local Standalone Install/Uninstall Helpers (.ps1 & .cmd)
        Generate-LocalInstallScripts -PackageDir $packageDir -PrinterName $selected.Name
        
        # Measure total uncompressed folder size before archiving
        $uncompressedMb = [math]::Round(((Get-ChildItem -Path $packageDir -Recurse | Measure-Object -Sum Length).Sum / 1MB), 2)
        Write-Log "Total Uncompressed Payload Folder Size: $uncompressedMb MB" -Level INFO

        # 7. Compile .intunewin package directly inside packageDir, then move to outDir with dynamic name
        $createdIntuneWinPath = ""
        if ($ChkCreateIntuneWin.IsChecked -and -not $IsStandaloneMode) {
            if (Test-Path $intuneWinUtilPath) {
                Write-Log "Compiling .intunewin package using IntuneWinAppUtil..." -Level INFO
                
                & "$intuneWinUtilPath" -c "$packageDir" -s "Invoke-AppDeployToolkit.exe" -o "$packageDir" -q
                
                $generatedFile = Join-Path -Path $packageDir -ChildPath "Invoke-AppDeployToolkit.intunewin"
                
                if (Test-Path $generatedFile) {
                    $safePrinterName = $selected.Name -replace '[^\w\-\.]', '_'
                    $dynamicIntuneWinName = "$appVendor - Printer - $safePrinterName - $appVersion.intunewin" -replace '[^\w\-\.\s]', '_'
                    $targetIntuneWinPath = Join-Path -Path $outDir -ChildPath $dynamicIntuneWinName
                    
                    if (Test-Path $targetIntuneWinPath) {
                        Remove-Item -Path $targetIntuneWinPath -Force -ErrorAction SilentlyContinue
                    }
                    
                    Move-Item -Path $generatedFile -Destination $targetIntuneWinPath -Force
                    
                    $sizeMb = [math]::Round(((Get-Item $targetIntuneWinPath).Length / 1MB), 2)
                    $createdIntuneWinPath = $targetIntuneWinPath
                    Write-Log "Compiled & Moved .intunewin package to: $dynamicIntuneWinName ($sizeMb MB)" -Level SUCCESS
                } else {
                    Write-Log "IntuneWinAppUtil completed, but Invoke-AppDeployToolkit.intunewin was not found in $packageDir." -Level WARNING
                }
            } else {
                Write-Log "IntuneWinAppUtil.exe not found at $intuneWinUtilPath" -Level WARNING
            }
        }
        
        # 8. Create Standalone ZIP Archive if checked or Standalone Mode
        $createdZipPath = ""
        if ($ChkCreateZip.IsChecked -or $IsStandaloneMode) {
            $safePrinterName = $selected.Name -replace '[^\w\-\.]', '_'
            $dynamicZipName = "$appVendor - Printer - $safePrinterName - $appVersion.zip" -replace '[^\w\-\.\s]', '_'
            $targetZipPath = Join-Path -Path $outDir -ChildPath $dynamicZipName
            
            if (Test-Path $targetZipPath) {
                Remove-Item -Path $targetZipPath -Force -ErrorAction SilentlyContinue
            }
            
            Write-Log "Compressing package folder into standalone ZIP archive..." -Level INFO
            Compress-Archive -Path "$packageDir\*" -DestinationPath $targetZipPath -Force
            
            if (Test-Path $targetZipPath) {
                $zipMb = [math]::Round(((Get-Item $targetZipPath).Length / 1MB), 2)
                $createdZipPath = $targetZipPath
                Write-Log "Created Standalone ZIP Archive: $dynamicZipName ($zipMb MB)" -Level SUCCESS
            } else {
                Write-Log "Failed to create Standalone ZIP Archive at $targetZipPath." -Level WARNING
            }
        }
        
        Write-Log "Package processing completed successfully!" -Level SUCCESS
        $TxtStatus.Text = "Package created at $packageDir"
        
        $msg = "Printer Package created successfully!`n`nPackage Directory:`n$packageDir`n`nUncompressed Payload Size: $uncompressedMb MB`n`nLocal Helper Launchers:`n- $packageDir\Install-PrinterLocal.cmd`n- $packageDir\Uninstall-PrinterLocal.cmd"
        if ($createdIntuneWinPath) {
            $msg += "`n`nCompiled IntuneWin Package:`n$createdIntuneWinPath"
        }
        if ($createdZipPath) {
            $msg += "`n`nStandalone ZIP Archive:`n$createdZipPath"
        }
        
        [System.Windows.Forms.MessageBox]::Show($msg, "Packaging Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        Write-Log "Packaging failed with error: $_" -Level ERROR
        $TxtStatus.Text = "Packaging failed."
        [System.Windows.Forms.MessageBox]::Show("Failed to create package:`n$_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Package Printer (PSADT v4 for Intune) Button Click
$BtnPackagePrinter.Add_Click({
    Start-PrinterPackageBuild -IsStandaloneMode $false
})

# Create Local Standalone Package Button Click
if ($BtnPackageLocal) {
    $BtnPackageLocal.Add_Click({
        Start-PrinterPackageBuild -IsStandaloneMode $true
    })
}

# Window Loaded Event: Load Printers automatically
$window.Add_Loaded({
    Load-Printers
})

# Show WPF Window
$window.ShowDialog() | Out-Null
