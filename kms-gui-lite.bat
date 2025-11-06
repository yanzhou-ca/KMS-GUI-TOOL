@echo off
setlocal
set "SCRIPT_PATH=%~f0"
powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "$content = Get-Content -LiteralPath '%SCRIPT_PATH%' -Encoding UTF8; $start = 0; for ($i = 0; $i -lt $content.Count; $i++) { if ($content[$i] -eq '###PS###') { $start = $i + 1; break } }; if ($start -ge $content.Count) { throw 'PowerShell marker not found!' }; $script = ($content | Select-Object -Skip $start) -join [Environment]::NewLine; if ([string]::IsNullOrWhiteSpace($script)) { throw 'No PowerShell script found after marker!' }; Invoke-Expression $script"
exit /b
###PS###
<#
.SYNOPSIS
    A professional PowerShell GUI for activating Windows and Office products via a Key Management Service (KMS).

.DESCRIPTION
    This tool provides a polished, responsive, and robust graphical interface for IT professionals to perform KMS activation.
    The UI is fully asynchronous, preventing any freezing or "Not Responding" states during operation. Products are neatly
    categorized into 'Windows' and 'Office' tabs for easy navigation, and the layout is fully resizable, including the
    activation log panel.

    The script features advanced feedback mechanisms, including a progress bar and status updates, along with intelligent
    error handling for common activation issues. It is built with best practices in mind, including resource management to
    prevent memory leaks and optimized code for performance and stability.

    Administrative privileges are required to execute the necessary 'slmgr.vbs' commands.

.AUTHOR
    Y.Z
.VERSION
    1.0
#>

#Requires -RunAsAdministrator

#region 1. Initialization
function Main {
    <#
    .SYNOPSIS
        The main entry point for the script.
    .DESCRIPTION
        This function serves as the orchestrator for the entire application. It loads the required WPF assemblies,
        builds the GUI by calling New-KmsGui, caches the product data for performance, initializes the user interface
        with product information, registers all necessary event handlers for interactivity, and finally displays the
        main window to the user.
    #>
    
    # Load required .NET assemblies for WPF.
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

    # Create the GUI from XAML and get references to all interactive controls.
    $script:gui = New-KmsGui
    if (-not $script:gui) {
        Write-Error "Failed to create the GUI. Exiting."
        return
    }

    # Cache the product data in a script-level variable. This is a performance optimization to ensure
    # the GVLK data is read only once at startup, rather than multiple times during UI initialization.
    $script:productData = Get-GvlkData

    # Populate the Windows and Office TreeView controls with the cached product data.
    Initialize-ProductTreeViews

    # Connect UI events (button clicks, text changes) to their corresponding PowerShell functions.
    Register-GuiEventHandlers
    
    # Perform an initial check to set the enabled/disabled state of the 'Activate' button.
    Update-ActivateButtonState

    # Display the main window and wait for it to be closed.
    $script:gui.Window.ShowDialog() | Out-Null
}
#endregion Initialization

#region 2. UI Definition
function New-KmsGui {
    <#
    .SYNOPSIS
        Builds the complete WPF Graphical User Interface from an embedded XAML string.
    .DESCRIPTION
        This function defines the entire visual structure, layout, and styling of the application. The layout is built
        using a Grid with proportional sizing (*) to ensure the UI is fully resizable. It includes a GridSplitter, allowing
        the user to dynamically resize the log panel. Advanced styling is applied to all controls for a modern, dark-theme
        aesthetic. The function returns a PSCustomObject containing named references to all key UI elements that the script
        needs to interact with, such as text boxes, buttons, and the status bar.
    #>
    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="KMS Activation Tool - Lite" Height="650" Width="600"
        MinWidth="550" MinHeight="500"
        WindowStartupLocation="CenterScreen" WindowStyle="SingleBorderWindow"
        Background="#2D2D30" ResizeMode="CanResizeWithGrip"
        UseLayoutRounding="True">
    <Window.Resources>
        <!-- Style for section header TextBlocks -->
        <Style x:Key="HeaderStyle" TargetType="{x:Type TextBlock}">
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>

        <!-- Style for the main TextBox input -->
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="Background" Value="#3F3F46"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="CaretBrush" Value="#F1F1F1"/>
        </Style>
        
        <!-- Style for the main Button -->
        <Style TargetType="{x:Type Button}">
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#009CFF"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#555555"/>
                    <Setter Property="Foreground" Value="#999999"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Style for the TreeView and its items -->
        <Style TargetType="{x:Type TreeView}">
            <Setter Property="Background" Value="#3F3F46"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="5"/>
        </Style>
        <Style TargetType="{x:Type TreeViewItem}">
            <Setter Property="Padding" Value="3"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="Background" Value="Transparent"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#555555"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <!-- Style for TabControl and TabItems -->
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Background" Value="#3F3F46"/>
            <Setter Property="Foreground" Value="#DCDCDC"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border Name="Border" Background="{TemplateBinding Background}" BorderBrush="Transparent" BorderThickness="1,1,1,0" CornerRadius="4,4,0,0" Margin="0,0,2,0">
                            <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center" HorizontalAlignment="Center" ContentSource="Header" Margin="10,2"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#007ACC" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsSelected" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#3F3F46" />
                                <Setter Property="Foreground" Value="#DCDCDC" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Style for the log TextBox -->
        <Style x:Key="LogStyle" TargetType="{x:Type TextBox}">
            <Setter Property="Background" Value="#1E1E1E"/>
            <Setter Property="Foreground" Value="#DCDCDC"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="TextWrapping" Value="Wrap"/>
        </Style>

        <!-- Style for the StatusBar -->
        <Style TargetType="{x:Type StatusBar}">
            <Setter Property="Background" Value="#007ACC"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>      <!-- 0: KMS Header -->
            <RowDefinition Height="Auto"/>      <!-- 1: KMS TextBox -->
            <RowDefinition Height="Auto"/>      <!-- 2: Product Header -->
            <RowDefinition Height="3*"/>        <!-- 3: Product TabControl (Resizable Top Panel, gets 3 shares of space) -->
            <RowDefinition Height="Auto"/>      <!-- 4: Activate Button -->
            <RowDefinition Height="5"/>         <!-- 5: GridSplitter -->
            <RowDefinition Height="Auto"/>      <!-- 6: Log Header -->
            <RowDefinition Height="*"/>         <!-- 7: Log TextBox (Resizable Bottom Panel, gets 1 share of space) -->
            <RowDefinition Height="Auto"/>      <!-- 8: Status Bar -->
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="1. Enter KMS Server Address" Style="{StaticResource HeaderStyle}"/>
        <TextBox x:Name="kmsServerTextBox" Grid.Row="1" Margin="0,5,0,15" ToolTip="Enter the hostname or IP address of your KMS server (e.g., kms.contoso.com:1688)"/>

        <TextBlock Grid.Row="2" Text="2. Select Product to Activate" Style="{StaticResource HeaderStyle}" Margin="0,0,0,5"/>
        <TabControl x:Name="productTabControl" Grid.Row="3" Margin="0,5,0,0">
            <TabItem Header="Windows">
                <TreeView x:Name="windowsProductTreeView" Margin="5"/>
            </TabItem>
            <TabItem Header="Office">
                <TreeView x:Name="officeProductTreeView" Margin="5"/>
            </TabItem>
        </TabControl>
        
        <Button x:Name="activateButton" Grid.Row="4" Content="ACTIVATE" Margin="0,15,0,0" Height="40"/>
        
        <GridSplitter Grid.Row="5" Height="5" HorizontalAlignment="Stretch" Background="#555555" ResizeDirection="Rows" Margin="0,10,0,10"/>
        
        <TextBlock Grid.Row="6" Text="Activation Log" Style="{StaticResource HeaderStyle}" Margin="0,5,0,5"/>
        <TextBox x:Name="logTextBox" Grid.Row="7" Margin="0,0,0,0" Style="{StaticResource LogStyle}"/>

        <StatusBar Grid.Row="8" Margin="0,10,0,-10">
            <StatusBarItem>
                <ProgressBar x:Name="statusProgressBar" Width="150" Height="15" IsIndeterminate="False" Visibility="Collapsed" Margin="0,0,10,0"/>
            </StatusBarItem>
            <StatusBarItem>
                <TextBlock x:Name="statusTextBlock" Foreground="White">Ready</TextBlock>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
'@
    
    try {
        # Read XAML
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)

        # Get references to named controls
        $controls = @{
            Window                 = $window
            KmsServerTextBox       = $window.FindName("kmsServerTextBox")
            ProductTabControl      = $window.FindName("productTabControl")
            WindowsProductTreeView = $window.FindName("windowsProductTreeView")
            OfficeProductTreeView  = $window.FindName("officeProductTreeView")
            ActivateButton         = $window.FindName("activateButton")
            LogTextBox             = $window.FindName("logTextBox")
            StatusTextBlock        = $window.FindName("statusTextBlock")
            StatusProgressBar      = $window.FindName("statusProgressBar")
        }
        
        return [PSCustomObject]$controls
    }
    catch {
        Write-Error "Error parsing XAML: $($_.Exception.Message)"
        return $null
    }
}

function Initialize-ProductTreeViews {
    Initialize-WindowsTreeView
    Initialize-OfficeTreeView
}

function Initialize-WindowsTreeView {
    $windowsData = $script:productData | Where-Object { $_.Category -like '*Windows*' }
    $groupedData = $windowsData | Group-Object -Property Category

    $script:gui.WindowsProductTreeView.Items.Clear()

    foreach ($group in $groupedData) {
        $categoryItem = New-Object System.Windows.Controls.TreeViewItem
        $categoryItem.Header = $group.Name
        $categoryItem.IsExpanded = $false
        $categoryItem.FontWeight = 'Bold'
        
        foreach ($product in $group.Group) {
            $productItem = New-Object System.Windows.Controls.TreeViewItem
            $productItem.Header = $product.DisplayName
            $productItem.Tag = $product
            $categoryItem.Items.Add($productItem) | Out-Null
        }
        $script:gui.WindowsProductTreeView.Items.Add($categoryItem) | Out-Null
    }
}

function Initialize-OfficeTreeView {
    $officeData = $script:productData | Where-Object { $_.Category -like '*Office*' }
    $groupedData = $officeData | Group-Object -Property Category

    $script:gui.OfficeProductTreeView.Items.Clear()

    foreach ($group in $groupedData) {
        $categoryItem = New-Object System.Windows.Controls.TreeViewItem
        $categoryItem.Header = $group.Name
        $categoryItem.IsExpanded = $false
        $categoryItem.FontWeight = 'Bold'
        
        foreach ($product in $group.Group) {
            $productItem = New-Object System.Windows.Controls.TreeViewItem
            $productItem.Header = $product.DisplayName
            $productItem.Tag = $product
            $categoryItem.Items.Add($productItem) | Out-Null
        }
        $script:gui.OfficeProductTreeView.Items.Add($categoryItem) | Out-Null
    }
}
#endregion UI Definition

#region 3. Event Handlers
function Register-GuiEventHandlers {
    $script:gui.ActivateButton.add_Click({ Invoke-KmsActivationWorkflow })
    $script:gui.KmsServerTextBox.add_TextChanged({ Update-ActivateButtonState })
    $script:gui.WindowsProductTreeView.add_SelectedItemChanged({ Update-ActivateButtonState })
    $script:gui.OfficeProductTreeView.add_SelectedItemChanged({ Update-ActivateButtonState })
    $script:gui.ProductTabControl.add_SelectionChanged({ Update-ActivateButtonState })
}

function Invoke-KmsActivationWorkflow {
    $kmsServer = $script:gui.KmsServerTextBox.Text.Trim()
    $selectedItem = $null

    if ($script:gui.ProductTabControl.SelectedIndex -eq 0) {
        $selectedItem = $script:gui.WindowsProductTreeView.SelectedItem
    }
    else {
        $selectedItem = $script:gui.OfficeProductTreeView.SelectedItem
    }

    if (-not $selectedItem -or -not $selectedItem.Tag) {
        Add-LogEntry -Message "Please select a specific product version to activate." -IsError
        return
    }

    if ([string]::IsNullOrWhiteSpace($kmsServer)) {
        Add-LogEntry -Message "KMS server address cannot be empty." -IsError
        return
    }
    if ($kmsServer -match '[^a-zA-Z0-9.\-:\[\]]') {
        Add-LogEntry -Message "Invalid characters in KMS server address. Only letters, digits, '.', '-', ':', and '[]' are allowed." -IsError
        return
    }

    $product = $selectedItem.Tag

    $script:gui.ActivateButton.IsEnabled = $false
    $script:gui.KmsServerTextBox.IsEnabled = $false
    $script:gui.ProductTabControl.IsEnabled = $false
    $script:gui.Window.Cursor = [System.Windows.Input.Cursors]::Wait
    $script:gui.StatusProgressBar.Visibility = "Visible"
    $script:gui.StatusProgressBar.IsIndeterminate = $true

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("product", $product)
    $runspace.SessionStateProxy.SetVariable("kmsServer", $kmsServer)
    $runspace.SessionStateProxy.SetVariable("gui", $script:gui)

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    
    $scriptBlock = {
        function Add-LogEntry {
            param([string]$Message, [switch]$IsStatus, [switch]$IsError)
            try {
                $gui.LogTextBox.Dispatcher.Invoke([action]{
                    $gui.LogTextBox.AppendText("[$([datetime]::now.ToString('HH:mm:ss'))] $Message`r`n")
                    $gui.LogTextBox.ScrollToEnd()
                    $text = $gui.LogTextBox.Text
                    $lines = $text.Split(@("`r`n"), [System.StringSplitOptions]::None)
                    if ($lines.Length -gt 500) {
                        $keep = @()
                        for ($i = $lines.Length - 300; $i -lt $lines.Length; $i++) {
                            $keep += $lines[$i]
                        }
                        $gui.LogTextBox.Text = ($keep -join "`r`n")
                        $gui.LogTextBox.ScrollToEnd()
                    }
                    if ($IsStatus) { $gui.StatusTextBlock.Text = $Message }
                    if ($IsError) { 
                        $gui.StatusTextBlock.Text = $Message
                        try { $gui.StatusTextBlock.Background = [System.Windows.Media.Brushes]::Red } catch { }
                    }
                }, "Normal")
            }
            catch { }
        }

        function Invoke-ActivationCommand {
            param(
                [string]$ProductType,
                [string[]]$Arguments
            )
            try {
                if ([System.Environment]::Is64BitOperatingSystem -and -not [System.Environment]::Is64BitProcess) {
                    $systemPath = [System.IO.Path]::Combine($env:SystemRoot, "Sysnative")
                } else {
                    $systemPath = [System.IO.Path]::Combine($env:SystemRoot, "System32")
                }

                if ($ProductType -eq 'Windows') {
                    $exePath = [System.IO.Path]::Combine($systemPath, "cscript.exe")
                    $scriptPath = [System.IO.Path]::Combine($systemPath, "slmgr.vbs")
                }
                else {
                    $officePaths = @(
                        "${env:ProgramFiles}\Microsoft Office\Office16",
                        "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
                        "${env:ProgramFiles}\Microsoft Office\Office15",
                        "${env:ProgramFiles(x86)}\Microsoft Office\Office15"
                    )
                    $scriptPath = $null
                    foreach ($path in $officePaths) {
                        $candidate = [System.IO.Path]::Combine($path, "ospp.vbs")
                        if (Test-Path $candidate) {
                            $scriptPath = $candidate
                            break
                        }
                    }
                    if (-not $scriptPath) {
                        throw "Office installation not found. Is Office installed?"
                    }
                    $exePath = [System.IO.Path]::Combine($systemPath, "cscript.exe")
                }

                # Build arguments string
                $argString = ""
                if ($ProductType -eq 'Windows') {
                    $argString = "//nologo `"$scriptPath`" $($Arguments -join ' ')"
                }
                else {
                    $argsList = @()
                    for ($i = 0; $i -lt $Arguments.Count; $i++) {
                        if ($Arguments[$i] -eq '/sethst') {
                            if ($i + 1 -lt $Arguments.Count) {
                                $argsList += "/sethst:`"$($Arguments[$i+1])`""
                                $i++
                            }
                        }
                        elseif ($Arguments[$i] -eq '/inpkey') {
                            if ($i + 1 -lt $Arguments.Count) {
                                $argsList += "/inpkey:`"$($Arguments[$i+1])`""
                                $i++
                            }
                        }
                        else {
                            $argsList += $Arguments[$i]
                        }
                    }
                    $argString = "//nologo `"$scriptPath`" $($argsList -join ' ')"
                }

                $pstart = New-Object System.Diagnostics.ProcessStartInfo
                $pstart.FileName = $exePath
                $pstart.Arguments = $argString
                $pstart.UseShellExecute = $false
                $pstart.RedirectStandardOutput = $true
                $pstart.RedirectStandardError = $true
                $pstart.WindowStyle = "Hidden"
                $pstart.CreateNoWindow = $true
                $pstart.StandardOutputEncoding = [System.Text.Encoding]::UTF8
                $pstart.StandardErrorEncoding = [System.Text.Encoding]::UTF8

                $process = [System.Diagnostics.Process]::Start($pstart)
                $stdout = $process.StandardOutput.ReadToEnd()
                $stderr = $process.StandardError.ReadToEnd()
                $process.WaitForExit()

                # --- CLEAN FILTERED OUTPUT ---
                $lines = ($stdout + "`n" + $stderr).Trim() -split "`r?`n"
                $cleanLines = @()
                foreach ($line in $lines) {
                    $line = $line.Trim()
                    # Mask full Windows product key
                    $line = $line -replace '\b([A-Z0-9]{5}-){4}[A-Z0-9]{5}\b', '*****-*****-*****-*****-*****'

                    # Skip visual noise
                    if ($line -eq "" -or
                        $line -like "*---*" -or
                        $line -like "---------------------------------------*" -or
                        $line -like "Successfully applied setting." -or
                        $line -like "Installed product key*" -or
                        $line -like "<Product key installation successful>" -or
                        $line -like "<Product activation successful>" -or
                        $line -like "Installed product key detected*" -or
                        $line -match 'LICENSE (NAME|DESCRIPTION):' -or
                        $line -match 'Last 5 characters of installed product key:'
                    ) {
                        if ($line -match 'LICENSE NAME:\s*(.+)') {
                            $cleanLines += "Activating: $($matches[1].Trim())"
                        }
                        elseif ($line -like "*<Product key installation successful>*" -or
                                $line -like "*Installed product key*successfully*") {
                            $cleanLines += "Product key installed successfully."
                        }
                        elseif ($line -like "*<Product activation successful>*") {
                            $cleanLines += "Product activated successfully."
                        }
                        elseif ($line -match 'Last 5 characters of installed product key:\s*(.+)') {
                            $cleanLines += "Key ending in: $($matches[1].Trim())"
                        }
                    }
                    else {
                        if ($line -ne "") {
                            $cleanLines += $line
                        }
                    }
                }

                $cleanOutput = ($cleanLines -join "`r`n").Trim()

                return New-Object PSObject -Property @{
                    Success = ($process.ExitCode -eq 0)
                    Output = if ($cleanOutput -eq "") { "Command completed successfully." } else { $cleanOutput }
                }
            }
            catch {
                return New-Object PSObject -Property @{
                    Success = $false
                    Output = "Exception: $($_.Exception.Message)"
                }
            }
        }

        try {
            Add-LogEntry -Message "----- Activation process started -----" -IsStatus

            $isOffice = $product.Category -like '*Office*'
            $productType = if ($isOffice) { 'Office' } else { 'Windows' }

            Add-LogEntry -Message "Setting KMS server to '$kmsServer'..." -IsStatus
            if ($isOffice) {
                $result = Invoke-ActivationCommand -ProductType 'Office' -Arguments @('/sethst', $kmsServer)
                Add-LogEntry -Message $result.Output
            } else {
                $result = Invoke-ActivationCommand -ProductType 'Windows' -Arguments @('/skms', $kmsServer)
                if (-not $result.Success) { throw "Failed to set KMS server." }
                Add-LogEntry -Message "KMS server configured successfully." -IsStatus
            }
            if (-not $result.Success) { throw "Failed to set KMS server." }

            $last5 = $product.Gvlk.Substring($product.Gvlk.Length - 5)
            Add-LogEntry -Message "Installing key for '$($product.DisplayName)' (ending in: $last5)..." -IsStatus
            if ($isOffice) {
                $result = Invoke-ActivationCommand -ProductType 'Office' -Arguments @('/inpkey', $product.Gvlk)
                Add-LogEntry -Message $result.Output
            } else {
                $result = Invoke-ActivationCommand -ProductType 'Windows' -Arguments @('/ipk', $product.Gvlk)
                if (-not $result.Success) { throw "Failed to install product key." }
                Add-LogEntry -Message "Product key installed successfully." -IsStatus
            }
            if (-not $result.Success) { throw "Failed to install product key." }

            Add-LogEntry -Message "Attempting activation..." -IsStatus
            if ($isOffice) {
                $result = Invoke-ActivationCommand -ProductType 'Office' -Arguments @('/act')
                Add-LogEntry -Message $result.Output
            } else {
                $result = Invoke-ActivationCommand -ProductType 'Windows' -Arguments @('/ato')
                if (-not $result.Success) { throw "Activation failed." }
                Add-LogEntry -Message "Windows $($product.DisplayName) activated successfully." -IsStatus
            }
            if (-not $result.Success) {
                $msg = "Activation failed."
                if ($result.Output -match '0xC004F074') {
                    $msg += " The KMS host could not be contacted. Verify server address and network connectivity."
                }
                elseif ($result.Output -match '0xC004F038') {
                    $msg += " The product could not be activated. Ensure the KMS server has enough clients (minimum 5 for Windows, 25 for Office)."
                }
                throw $msg
            }

            Add-LogEntry -Message "----- PRODUCT ACTIVATED SUCCESSFULLY! -----" -IsStatus
        }
        catch {
            Add-LogEntry -Message "ERROR: $($_.Exception.Message)" -IsError
        }
        finally {
            try {
                $gui.Window.Dispatcher.Invoke([action]{
                    $gui.KmsServerTextBox.IsEnabled = $true
                    $gui.ProductTabControl.IsEnabled = $true
                    $gui.Window.Cursor = [System.Windows.Input.Cursors]::Arrow
                    $gui.StatusProgressBar.Visibility = "Collapsed"
                    $gui.StatusProgressBar.IsIndeterminate = $false
                    
                    $isServerValid = -not [string]::IsNullOrWhiteSpace($gui.KmsServerTextBox.Text)
                    $selectedItem = $null
                    if ($gui.ProductTabControl.SelectedIndex -eq 0) {
                        $selectedItem = $gui.WindowsProductTreeView.SelectedItem
                    } else {
                        $selectedItem = $gui.OfficeProductTreeView.SelectedItem
                    }
                    $isProductValid = ($null -ne $selectedItem) -and ($null -ne $selectedItem.Tag)
                    $gui.ActivateButton.IsEnabled = ($isServerValid -and $isProductValid)
                }, "Normal")
            }
            catch { }
        }
    }

    [void]$ps.AddScript($scriptBlock)
    [void]$ps.BeginInvoke()
}

function Update-ActivateButtonState {
    $isServerValid = -not [string]::IsNullOrWhiteSpace($script:gui.KmsServerTextBox.Text)
    $selectedItem = $null
    if ($script:gui.ProductTabControl.SelectedIndex -eq 0) {
        $selectedItem = $script:gui.WindowsProductTreeView.SelectedItem
    }
    else {
        $selectedItem = $script:gui.OfficeProductTreeView.SelectedItem
    }
    $isProductValid = ($null -ne $selectedItem) -and ($null -ne $selectedItem.Tag)
    $script:gui.ActivateButton.IsEnabled = ($isServerValid -and $isProductValid)
}

function Add-LogEntry {
    param([string]$Message, [switch]$IsStatus, [switch]$IsError)
    $script:gui.LogTextBox.Dispatcher.Invoke({
        $script:gui.LogTextBox.AppendText("[$([datetime]::now.ToString('HH:mm:ss'))] $Message`r`n")
        $script:gui.LogTextBox.ScrollToEnd()
        if ($script:gui.LogTextBox.LineCount -gt 500) {
            $lines = $script:gui.LogTextBox.Text.Split([string[]]@("`r`n"), "None") | Select-Object -Last 300
            $script:gui.LogTextBox.Text = $lines -join "`r`n"
        }
        if ($IsStatus.IsPresent) {
            $script:gui.StatusTextBlock.Text = $Message
            $script:gui.StatusTextBlock.Background = [System.Windows.Media.Brushes]::Transparent
        }
        if ($IsError.IsPresent) {
            $script:gui.StatusTextBlock.Text = $Message
            $script:gui.StatusTextBlock.Background = [System.Windows.Media.Brushes]::Red
        }
    })
}
#endregion Event Handlers

#region 5. Data Storage
function Get-GvlkData {
    return @(
        # Office 2016
        [PSCustomObject]@{ Category = 'Office 2016'; DisplayName = 'Professional Plus'; Gvlk = 'XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99' }
        [PSCustomObject]@{ Category = 'Office 2016'; DisplayName = 'Standard'; Gvlk = 'JNRGM-WHDWX-FJJG3-K47QV-DRTFM' }
        [PSCustomObject]@{ Category = 'Office 2016'; DisplayName = 'Project Professional'; Gvlk = 'YG9NW-3K39V-2T3HJ-93F3Q-G83KT' }
        [PSCustomObject]@{ Category = 'Office 2016'; DisplayName = 'Project Standard'; Gvlk = 'GNFHQ-F6YQM-KQDGJ-327XX-KQBVC' }
        [PSCustomObject]@{ Category = 'Office 2016'; DisplayName = 'Visio Professional'; Gvlk = 'PD3PC-RHNGV-FXJ29-8JK7D-RJRJK' }
        [PSCustomObject]@{ Category = 'Office 2016'; DisplayName = 'Visio Standard'; Gvlk = '7WHWN-4T7MP-G96JF-G33KR-W8GF4' }
        
        # Office 2019
        [PSCustomObject]@{ Category = 'Office 2019'; DisplayName = 'Professional Plus'; Gvlk = 'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP' }
        [PSCustomObject]@{ Category = 'Office 2019'; DisplayName = 'Standard'; Gvlk = '6NWWJ-YQWMR-K87HH-KY3K2-GY2CH' }
        [PSCustomObject]@{ Category = 'Office 2019'; DisplayName = 'Project Professional'; Gvlk = 'B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B' }
        [PSCustomObject]@{ Category = 'Office 2019'; DisplayName = 'Project Standard'; Gvlk = 'C4F7P-NCP8C-6CQPT-MQHV9-JXD2M' }
        [PSCustomObject]@{ Category = 'Office 2019'; DisplayName = 'Visio Professional'; Gvlk = '9BGNQ-K37YR-RQHF2-38RQ3-7VCBB' }
        [PSCustomObject]@{ Category = 'Office 2019'; DisplayName = 'Visio Standard'; Gvlk = '7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2' }

        # Office LTSC 2021
        [PSCustomObject]@{ Category = 'Office LTSC 2021'; DisplayName = 'Professional Plus'; Gvlk = 'FXYTK-NJJ8C-GB6DW-3YQT3-6T3PY' }
        [PSCustomObject]@{ Category = 'Office LTSC 2021'; DisplayName = 'Standard'; Gvlk = 'KDX7X-BNVR8-TXXGX-4Q7Y8-TRV2Y' }
        [PSCustomObject]@{ Category = 'Office LTSC 2021'; DisplayName = 'Project Professional'; Gvlk = 'FTNWT-C6WBT-8HMGF-K9PRX-QV9H8' }
        [PSCustomObject]@{ Category = 'Office LTSC 2021'; DisplayName = 'Project Standard'; Gvlk = 'J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T' }
        [PSCustomObject]@{ Category = 'Office LTSC 2021'; DisplayName = 'Visio LTSC Professional'; Gvlk = 'KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4' }
        [PSCustomObject]@{ Category = 'Office LTSC 2021'; DisplayName = 'Visio LTSC Standard'; Gvlk = 'MJVNY-BYWPY-CWV6J-2RKRT-4M8QG' }

        # Office LTSC 2024
        [PSCustomObject]@{ Category = 'Office LTSC 2024'; DisplayName = 'Professional Plus'; Gvlk = 'XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB' }

        # Windows Client
        [PSCustomObject]@{ Category = 'Windows 11 / 10'; DisplayName = 'Pro'; Gvlk = 'W269N-WFGWX-YVC9B-4J6C9-T83GX' }
        [PSCustomObject]@{ Category = 'Windows 11 / 10'; DisplayName = 'Pro N'; Gvlk = 'MH37W-N47XK-V7XM9-C7227-GCQG9' }
        [PSCustomObject]@{ Category = 'Windows 11 / 10'; DisplayName = 'Pro for Workstations'; Gvlk = 'NRG8B-VKK3Q-CX5G3-CTQCR-XXP6H' }
        [PSCustomObject]@{ Category = 'Windows 11 / 10'; DisplayName = 'Enterprise'; Gvlk = 'NPPR9-FWDCX-D2C8J-H872K-2YT43' }
        [PSCustomObject]@{ Category = 'Windows 11 / 10'; DisplayName = 'Enterprise G'; Gvlk = 'YYVX9-NTFWV-6MDM3-9PT4T-4M68B' }
        [PSCustomObject]@{ Category = 'Windows 11 / 10'; DisplayName = 'Enterprise LTSC'; Gvlk = 'M7XTQ-FN8P6-TTKYV-9D4CC-J462D' }
        [PSCustomObject]@{ Category = 'Windows 11 / 10'; DisplayName = 'Education'; Gvlk = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2' }
        

        # Windows Server (LTSC versions)        
        [PSCustomObject]@{ Category = 'Windows Server 2016'; DisplayName = 'Standard'; Gvlk = 'WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY' }
        [PSCustomObject]@{ Category = 'Windows Server 2016'; DisplayName = 'Datacenter'; Gvlk = 'CB7KF-BWN84-R7R2Y-793K2-8XDDG' }
        [PSCustomObject]@{ Category = 'Windows Server 2016'; DisplayName = 'Essentials'; Gvlk = 'JCKRF-N37P4-C2D82-9YXRT-4M63B' }

        [PSCustomObject]@{ Category = 'Windows Server 2022'; DisplayName = 'Standard'; Gvlk = 'VDYBN-27WPP-V4HQT-9VMD4-VMK7H' }
        [PSCustomObject]@{ Category = 'Windows Server 2022'; DisplayName = 'Datacenter'; Gvlk = 'WX4NM-KYWYW-QJJR4-XV3QB-6VM33' }
        [PSCustomObject]@{ Category = 'Windows Server 2022'; DisplayName = 'Datacenter: Azure Edition'; Gvlk = 'NTBV8-9K7Q8-V27C6-M2BTV-KHMXV' }
        
        [PSCustomObject]@{ Category = 'Windows Server 2019'; DisplayName = 'Standard'; Gvlk = 'N69G4-B89J2-4G8F4-WWYCC-J464C' }
        [PSCustomObject]@{ Category = 'Windows Server 2019'; DisplayName = 'Datacenter'; Gvlk = 'WMDGN-G9PQG-XVVXX-R3X43-63DFG' }
        [PSCustomObject]@{ Category = 'Windows Server 2019'; DisplayName = 'Essentials'; Gvlk = 'WVDHN-86M7X-466P6-VHXV7-YY726' }

        [PSCustomObject]@{ Category = 'Windows Server 2025'; DisplayName = 'Standard'; Gvlk = 'TVRH6-WHNXV-R9WG3-9XRFY-MY832' }
        [PSCustomObject]@{ Category = 'Windows Server 2025'; DisplayName = 'Datacenter'; Gvlk = 'D764K-2NDRG-47T6Q-P8T8W-YP6DF' }
        [PSCustomObject]@{ Category = 'Windows Server 2025'; DisplayName = 'Azure Edition'; Gvlk = 'XGN3F-F394H-FD2MY-PP6FD-8MCRC' }
    )
}
#endregion Data Storage

# --- Script Entry Point ---
Main