#requires -Version 5.1
# PowerApps Troubleshooter GUI Tool
# Select Technology Branded - Purple (#3C1053) and Orange (#E87722)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Web

# Install/Import modules upfront (not in background)
if (!(Get-Module -ListAvailable -Name Microsoft.Xrm.Tooling.CrmConnector.PowerShell)) {
    Write-Host "Installing Microsoft.Xrm.Tooling.CrmConnector.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.Xrm.Tooling.CrmConnector.PowerShell -Scope CurrentUser -Force -AllowClobber
}

if (!(Get-Module -ListAvailable -Name Microsoft.Xrm.Data.PowerShell)) {
    Write-Host "Installing Microsoft.Xrm.Data.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.Xrm.Data.PowerShell -Scope CurrentUser -Force -AllowClobber
}

Import-Module Microsoft.Xrm.Tooling.CrmConnector.PowerShell -ErrorAction SilentlyContinue
Import-Module Microsoft.Xrm.Data.PowerShell -ErrorAction SilentlyContinue

# XAML Definition for the GUI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PowerTool - Select Technology"
        Height="650" Width="950"
        MinHeight="500" MinWidth="700"
        WindowStartupLocation="CenterScreen"
        WindowState="Maximized"
        Background="#F5F5F5">
    <Window.Resources>
        <!-- Select Technology Colors -->
        <SolidColorBrush x:Key="PrimaryPurple" Color="#3C1053"/>
        <SolidColorBrush x:Key="SecondaryOrange" Color="#E87722"/>
        <SolidColorBrush x:Key="LightPurple" Color="#5C2D73"/>
        <SolidColorBrush x:Key="DarkOrange" Color="#C86518"/>
        <SolidColorBrush x:Key="BackgroundLight" Color="#F5F5F5"/>
        <SolidColorBrush x:Key="TextLight" Color="#FFFFFF"/>

        <!-- Button Style -->
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource SecondaryOrange}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{StaticResource DarkOrange}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#CCCCCC"/>
                                <Setter Property="Foreground" Value="#888888"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SecondaryButton" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource PrimaryPurple}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{StaticResource LightPurple}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#CCCCCC"/>
                                <Setter Property="Foreground" Value="#888888"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- TextBox Style -->
        <Style x:Key="InputTextBox" TargetType="TextBox">
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
            <Style.Triggers>
                <Trigger Property="IsFocused" Value="True">
                    <Setter Property="BorderBrush" Value="{StaticResource SecondaryOrange}"/>
                    <Setter Property="BorderThickness" Value="2"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- CheckBox Style -->
        <Style x:Key="ScanCheckBox" TargetType="CheckBox">
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Margin" Value="0,2"/>
            <Setter Property="Foreground" Value="#333333"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="{StaticResource PrimaryPurple}" Padding="15,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="PowerTool - Select Technology"
                               Foreground="White"
                               FontSize="18"
                               FontWeight="Bold"/>
                    <TextBlock Text="Discover flows, plugins, and processes in your Dataverse environment"
                               Foreground="#CCBBDD"
                               FontSize="10"
                               Margin="0,3,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                    <TextBlock Text="SELECT" Foreground="{StaticResource SecondaryOrange}" FontSize="14" FontWeight="Bold" HorizontalAlignment="Right"/>
                    <TextBlock Text="TECHNOLOGY" Foreground="White" FontSize="10" FontWeight="SemiBold" HorizontalAlignment="Right"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Configuration Panel -->
        <Border Grid.Row="1" Background="White" Padding="12,10" BorderBrush="#DDDDDD" BorderThickness="0,0,0,1">
            <ScrollViewer HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Disabled">
                <Grid MinWidth="650">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*" MinWidth="280"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <!-- Left: Connection Settings -->
                    <StackPanel Grid.Column="0" Margin="0,0,15,0">
                        <TextBlock Text="Connection Settings" FontWeight="SemiBold" FontSize="12" Foreground="{StaticResource PrimaryPurple}" Margin="0,0,0,6"/>

                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Environment URL:" FontSize="10" Margin="0,0,0,2"/>
                            <TextBox x:Name="txtEnvironmentUrl" Grid.Row="1" Grid.Column="0"
                                     Style="{StaticResource InputTextBox}"
                                     Text=""
                                     Margin="0,0,8,6"/>

                            <TextBlock Grid.Row="0" Grid.Column="1" Text="Search Keywords (comma-separated):" FontSize="10" Margin="0,0,0,2"/>
                            <TextBox x:Name="txtKeywords" Grid.Row="1" Grid.Column="1"
                                     Style="{StaticResource InputTextBox}"
                                     Text="project,approval,approve,email,notification"
                                     Margin="0,0,0,6"/>

                            <TextBlock Grid.Row="2" Grid.Column="0" Text="Breakage Date:" FontSize="10" Margin="0,0,0,2"/>
                            <DatePicker x:Name="dpBreakageDate" Grid.Row="3" Grid.Column="0"
                                        Padding="6" FontSize="11" Margin="0,0,8,0"
                                        SelectedDate="{x:Static sys:DateTime.Today}"
                                        xmlns:sys="clr-namespace:System;assembly=mscorlib"/>
                        </Grid>
                    </StackPanel>

                    <!-- Middle: Scan Options (2 columns) -->
                    <StackPanel Grid.Column="1" Margin="0,0,15,0">
                        <TextBlock Text="Scan Options" FontWeight="SemiBold" FontSize="12" Foreground="{StaticResource PrimaryPurple}" Margin="0,0,0,6"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0" Margin="0,0,10,0">
                                <CheckBox x:Name="chkCloudFlows" Content="Cloud Flows" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkClassicWorkflows" Content="Classic Workflows" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkPlugins" Content="Plugins" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkWebResources" Content="Web Resources" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1">
                                <CheckBox x:Name="chkCanvasApps" Content="Canvas Apps" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkModelDrivenApps" Content="Model-Driven Apps" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkCustomActions" Content="Custom Actions" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkServiceEndpoints" Content="Service Endpoints" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkSLAs" Content="SLAs" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkBPFs" Content="Business Process Flows" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkTableSchemas" Content="Table Schemas" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkConnections" Content="Connection References" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkEnvVariables" Content="Environment Variables" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                            </StackPanel>
                        </Grid>
                    </StackPanel>

                    <!-- Right: Action Buttons -->
                    <StackPanel Grid.Column="2" VerticalAlignment="Center">
                        <Button x:Name="btnConnect" Content="Connect" Style="{StaticResource SecondaryButton}" Margin="0,0,0,8" Width="120"/>
                        <Button x:Name="btnScan" Content="Start Scan" Style="{StaticResource PrimaryButton}" IsEnabled="False" Margin="0,0,0,8" Width="120"/>
                        <Button x:Name="btnFlowHistory" Content="Get Flow History" Style="{StaticResource SecondaryButton}" IsEnabled="False" Margin="0,0,0,8" Width="120" FontSize="10"/>
                        <Button x:Name="btnAISettings" Content="AI Settings" Style="{StaticResource SecondaryButton}" Margin="0,0,0,4" Width="120" FontSize="10"/>
                        <Button x:Name="btnManageCredentials" Content="Manage Credentials" Style="{StaticResource SecondaryButton}" Margin="0,0,0,4" Width="120" FontSize="10"/>
                        <Button x:Name="btnSignOut" Content="Sign Out" Style="{StaticResource SecondaryButton}" Width="120" FontSize="10"/>
                    </StackPanel>
                </Grid>
            </ScrollViewer>
        </Border>

        <!-- Results Area -->
        <Grid Grid.Row="2" Margin="12">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <!-- Results Header -->
            <Grid Grid.Row="0" Margin="0,0,0,8">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="Discovery Results" FontWeight="SemiBold" FontSize="13" Foreground="{StaticResource PrimaryPurple}"/>
                    <TextBlock x:Name="txtResultCount" Text=" (0 items)" FontSize="11" Foreground="#666666" VerticalAlignment="Bottom" Margin="5,0,0,1"/>
                </StackPanel>

                <Button x:Name="btnAISummary" Grid.Column="1" Content="AI Summary" Style="{StaticResource SecondaryButton}" IsEnabled="False" Margin="0,0,8,0"/>
                <Button x:Name="btnExportCsv" Grid.Column="2" Content="Export CSV" Style="{StaticResource SecondaryButton}" IsEnabled="False" Margin="0,0,8,0"/>
                <Button x:Name="btnSaveSnapshot" Grid.Column="3" Content="Save Snapshot" Style="{StaticResource SecondaryButton}" IsEnabled="False" Margin="0,0,8,0"/>
                <Button x:Name="btnCompareSnapshots" Grid.Column="4" Content="Compare Snapshots" Style="{StaticResource SecondaryButton}" IsEnabled="False" Margin="0,0,8,0"/>
                <Button x:Name="btnViewDependencies" Grid.Column="5" Content="View Dependencies" Style="{StaticResource SecondaryButton}" IsEnabled="False" Margin="0,0,8,0"/>
                <Button x:Name="btnClearResults" Grid.Column="6" Content="Clear" Style="{StaticResource SecondaryButton}" IsEnabled="False"/>
            </Grid>

            <!-- Results DataGrid with horizontal scroll -->
            <DataGrid x:Name="dgResults" Grid.Row="1"
                      AutoGenerateColumns="False"
                      IsReadOnly="True"
                      GridLinesVisibility="Horizontal"
                      HeadersVisibility="Column"
                      BorderBrush="#DDDDDD"
                      BorderThickness="1"
                      RowHeight="28"
                      FontSize="11"
                      AlternatingRowBackground="#F9F9F9"
                      SelectionMode="Single"
                      SelectionUnit="FullRow"
                      HorizontalScrollBarVisibility="Auto"
                      VerticalScrollBarVisibility="Auto">
                <DataGrid.ColumnHeaderStyle>
                    <Style TargetType="DataGridColumnHeader">
                        <Setter Property="Background" Value="{StaticResource PrimaryPurple}"/>
                        <Setter Property="Foreground" Value="White"/>
                        <Setter Property="FontWeight" Value="SemiBold"/>
                        <Setter Property="FontSize" Value="11"/>
                        <Setter Property="Padding" Value="8,6"/>
                        <Setter Property="BorderBrush" Value="#5C2D73"/>
                        <Setter Property="BorderThickness" Value="0,0,1,0"/>
                    </Style>
                </DataGrid.ColumnHeaderStyle>
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Score" Binding="{Binding Score}" Width="50">
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="FontWeight" Value="Bold"/>
                                <Setter Property="Foreground" Value="{StaticResource SecondaryOrange}"/>
                                <Setter Property="VerticalAlignment" Value="Center"/>
                                <Setter Property="Padding" Value="6,0"/>
                            </Style>
                        </DataGridTextColumn.ElementStyle>
                    </DataGridTextColumn>
                    <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="120">
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="VerticalAlignment" Value="Center"/>
                                <Setter Property="Padding" Value="6,0"/>
                            </Style>
                        </DataGridTextColumn.ElementStyle>
                    </DataGridTextColumn>
                    <DataGridTextColumn Header="Solution" Binding="{Binding Solution}" Width="140">
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="VerticalAlignment" Value="Center"/>
                                <Setter Property="Padding" Value="6,0"/>
                                <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
                                <Setter Property="ToolTip" Value="{Binding Solution}"/>
                            </Style>
                        </DataGridTextColumn.ElementStyle>
                    </DataGridTextColumn>
                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="180">
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="VerticalAlignment" Value="Center"/>
                                <Setter Property="Padding" Value="6,0"/>
                                <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
                                <Setter Property="ToolTip" Value="{Binding Name}"/>
                            </Style>
                        </DataGridTextColumn.ElementStyle>
                    </DataGridTextColumn>
                    <DataGridTemplateColumn Header="Link" Width="45">
                        <DataGridTemplateColumn.CellTemplate>
                            <DataTemplate>
                                <Button Content="Copy" Tag="{Binding Url}"
                                        Padding="4,2" FontSize="9" Cursor="Hand"
                                        Background="#E87722" Foreground="White" BorderThickness="0">
                                    <Button.Style>
                                        <Style TargetType="Button">
                                            <Style.Triggers>
                                                <DataTrigger Binding="{Binding Url}" Value="">
                                                    <Setter Property="Visibility" Value="Collapsed"/>
                                                </DataTrigger>
                                                <DataTrigger Binding="{Binding Url}" Value="{x:Null}">
                                                    <Setter Property="Visibility" Value="Collapsed"/>
                                                </DataTrigger>
                                            </Style.Triggers>
                                        </Style>
                                    </Button.Style>
                                </Button>
                            </DataTemplate>
                        </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTextColumn Header="State" Binding="{Binding StateIndicator}" Width="60">
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="VerticalAlignment" Value="Center"/>
                                <Setter Property="Padding" Value="6,0"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                                <Setter Property="HorizontalAlignment" Value="Center"/>
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding StateIndicator}" Value="[ON]">
                                        <Setter Property="Foreground" Value="Green"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding StateIndicator}" Value="[OFF]">
                                        <Setter Property="Foreground" Value="Red"/>
                                    </DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </DataGridTextColumn.ElementStyle>
                    </DataGridTextColumn>
                    <DataGridTextColumn Header="Actions" Binding="{Binding Actions}" Width="140">
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="VerticalAlignment" Value="Center"/>
                                <Setter Property="Padding" Value="6,0"/>
                                <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
                                <Setter Property="ToolTip" Value="{Binding Actions}"/>
                                <Setter Property="FontSize" Value="10"/>
                            </Style>
                        </DataGridTextColumn.ElementStyle>
                    </DataGridTextColumn>
                    <DataGridTextColumn Header="Errors (30d)" Binding="{Binding ErrorCount}" Width="85">
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="VerticalAlignment" Value="Center"/>
                                <Setter Property="Padding" Value="6,0"/>
                                <Setter Property="HorizontalAlignment" Value="Right"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding HasErrors}" Value="True">
                                        <Setter Property="Foreground" Value="Red"/>
                                    </DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </DataGridTextColumn.ElementStyle>
                    </DataGridTextColumn>
                    <DataGridTextColumn Header="Modified" Binding="{Binding ModifiedOn}" Width="110">
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="VerticalAlignment" Value="Center"/>
                                <Setter Property="Padding" Value="6,0"/>
                            </Style>
                        </DataGridTextColumn.ElementStyle>
                    </DataGridTextColumn>
                    <DataGridTextColumn Header="Match Reasons" Binding="{Binding Reasons}" Width="*" MinWidth="250">
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="VerticalAlignment" Value="Center"/>
                                <Setter Property="Padding" Value="6,0"/>
                                <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
                                <Setter Property="ToolTip" Value="{Binding Reasons}"/>
                            </Style>
                        </DataGridTextColumn.ElementStyle>
                    </DataGridTextColumn>
                </DataGrid.Columns>
            </DataGrid>
        </Grid>

        <!-- Status Bar -->
        <Border Grid.Row="3" Background="{StaticResource PrimaryPurple}" Padding="10,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <Ellipse x:Name="statusIndicator" Width="8" Height="8" Fill="#FF6666" Margin="0,0,6,0"/>
                    <TextBlock x:Name="txtConnectionStatus" Text="Not Connected" Foreground="White" FontSize="10"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" Margin="15,0,0,0">
                    <Ellipse x:Name="flowApiStatusIndicator" Width="8" Height="8" Fill="#FF6666" Margin="0,0,6,0"/>
                    <TextBlock x:Name="txtFlowApiStatus" Text="Flow API: Not authenticated" Foreground="White" FontSize="10"/>
                </StackPanel>

                <TextBlock x:Name="txtStatus" Grid.Column="2" Text="Ready" Foreground="#CCBBDD" FontSize="10" Margin="15,0" TextAlignment="Center"/>

                <ProgressBar x:Name="progressBar" Grid.Column="3" Width="150" Height="5" IsIndeterminate="False" Value="0"
                             Background="#5C2D73" Foreground="{StaticResource SecondaryOrange}"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Create the window
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$script:txtEnvironmentUrl = $window.FindName("txtEnvironmentUrl")
$txtKeywords = $window.FindName("txtKeywords")
$dpBreakageDate = $window.FindName("dpBreakageDate")
$btnConnect = $window.FindName("btnConnect")
$btnScan = $window.FindName("btnScan")
$btnFlowHistory = $window.FindName("btnFlowHistory")
$btnAISettings = $window.FindName("btnAISettings")
$btnManageCredentials = $window.FindName("btnManageCredentials")
$btnSignOut = $window.FindName("btnSignOut")
$btnAISummary = $window.FindName("btnAISummary")
$btnExportCsv = $window.FindName("btnExportCsv")
$btnSaveSnapshot = $window.FindName("btnSaveSnapshot")
$btnCompareSnapshots = $window.FindName("btnCompareSnapshots")
$btnViewDependencies = $window.FindName("btnViewDependencies")
$btnClearResults = $window.FindName("btnClearResults")
$dgResults = $window.FindName("dgResults")
$txtResultCount = $window.FindName("txtResultCount")
$txtStatus = $window.FindName("txtStatus")
$txtConnectionStatus = $window.FindName("txtConnectionStatus")
$statusIndicator = $window.FindName("statusIndicator")
$script:txtFlowApiStatus = $window.FindName("txtFlowApiStatus")
$script:flowApiStatusIndicator = $window.FindName("flowApiStatusIndicator")
$progressBar = $window.FindName("progressBar")

# Checkboxes
$chkCloudFlows = $window.FindName("chkCloudFlows")
$chkClassicWorkflows = $window.FindName("chkClassicWorkflows")
$chkPlugins = $window.FindName("chkPlugins")
$chkWebResources = $window.FindName("chkWebResources")
$chkCanvasApps = $window.FindName("chkCanvasApps")
$chkModelDrivenApps = $window.FindName("chkModelDrivenApps")
$chkCustomActions = $window.FindName("chkCustomActions")
$chkServiceEndpoints = $window.FindName("chkServiceEndpoints")
$chkSLAs = $window.FindName("chkSLAs")
$chkBPFs = $window.FindName("chkBPFs")
$chkTableSchemas = $window.FindName("chkTableSchemas")
$chkConnections = $window.FindName("chkConnections")
$chkEnvVariables = $window.FindName("chkEnvVariables")

# Global variables - use synchronized hashtable for cross-thread access
$script:SyncHash = [hashtable]::Synchronized(@{
    Connection = $null
    Results = @()
    EnvironmentId = $null
    FlowApiToken = $null
    FlowApiTokenExpiry = $null
})

# Function to show error dialog with copy-to-clipboard option
function Show-ErrorWithCopyOption {
    param(
        [string]$Title = "Error",
        [string]$Message,
        [string]$Details = ""
    )

    $fullError = if ($Details) { "$Message`n`nDetails:`n$Details" } else { $Message }

    $errorXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Error" Height="350" Width="550"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="X" FontSize="24" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,10,0" Foreground="#C00000"/>
            <TextBlock Text="An error occurred" FontSize="16" FontWeight="SemiBold" VerticalAlignment="Center" Foreground="#C00000"/>
        </StackPanel>
        
        <TextBox Grid.Row="1" x:Name="txtErrorDetails" 
                 IsReadOnly="True" 
                 TextWrapping="Wrap" 
                 VerticalScrollBarVisibility="Auto"
                 Background="#FFF0F0"
                 BorderBrush="#C00000"
                 Padding="10"
                 FontFamily="Consolas"
                 FontSize="11"/>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="btnCopyError" Content="Copy to Clipboard" 
                    Padding="15,8" Margin="0,0,10,0"
                    Background="#3C1053" Foreground="White" FontWeight="SemiBold"
                    Cursor="Hand"/>
            <Button x:Name="btnCloseError" Content="Close" 
                    Padding="15,8"
                    Background="#E87722" Foreground="White" FontWeight="SemiBold"
                    Cursor="Hand"/>
        </StackPanel>
    </Grid>
</Window>
"@

    try {
        $errorReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($errorXaml))
        $errorWindow = [System.Windows.Markup.XamlReader]::Load($errorReader)

        $btnCopyError = $errorWindow.FindName("btnCopyError")
        $btnCloseError = $errorWindow.FindName("btnCloseError")
        $txtErrorDetails = $errorWindow.FindName("txtErrorDetails")

        # Set window title and error text
        $errorWindow.Title = $Title
        $txtErrorDetails.Text = $fullError

        # Store error text in tag for access in click handler
        $btnCopyError.Tag = $fullError

        $btnCopyError.Add_Click({
            param($sender, $e)
            $errorText = $sender.Tag
            [System.Windows.Clipboard]::SetText($errorText)
            $sender.Content = "Copied!"
            $sender.Background = [System.Windows.Media.Brushes]::Green
        })

        $btnCloseError.Add_Click({
            $errorWindow.Close()
        })

        [void]$errorWindow.ShowDialog()
    } catch {
        # Fallback to standard message box if custom dialog fails
        [System.Windows.MessageBox]::Show($fullError, $Title, "OK", "Error")
    }
}

# Function to update Flow API token status in UI
function Update-FlowApiTokenStatus {
    if ($script:SyncHash.FlowApiToken -and $script:SyncHash.FlowApiTokenExpiry) {
        $expiryTime = $script:SyncHash.FlowApiTokenExpiry
        $timeRemaining = $expiryTime - (Get-Date)

        if ($timeRemaining.TotalMinutes -gt 10) {
            $script:flowApiStatusIndicator.Fill = [System.Windows.Media.Brushes]::LimeGreen
            $script:txtFlowApiStatus.Text = "Flow API: Valid (expires $($expiryTime.ToString('HH:mm')))"
        } elseif ($timeRemaining.TotalMinutes -gt 0) {
            $script:flowApiStatusIndicator.Fill = [System.Windows.Media.Brushes]::Yellow
            $script:txtFlowApiStatus.Text = "Flow API: Expiring soon ($([int]$timeRemaining.TotalMinutes)m)"
        } else {
            $script:flowApiStatusIndicator.Fill = [System.Windows.Media.Brushes]::Red
            $script:txtFlowApiStatus.Text = "Flow API: Expired"
        }
    } else {
        $script:flowApiStatusIndicator.Fill = [System.Windows.Media.Brushes]::Red
        $script:txtFlowApiStatus.Text = "Flow API: Not authenticated"
    }
}

# Function to get list of saved client configurations
function Get-SavedClientConfigs {
    $configDir = "$env:USERPROFILE\.powersearch\clients"
    if (-not (Test-Path $configDir)) {
        return ,@()
    }

    $configs = [System.Collections.ArrayList]::new()
    Get-ChildItem -Path $configDir -Filter "*.json" | ForEach-Object {
        try {
            $configFile = Get-Content $_.FullName -Raw | ConvertFrom-Json
            $secureConfig = ConvertTo-SecureString $configFile.EncryptedConfig
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureConfig)
            $configJson = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

            $config = $configJson | ConvertFrom-Json
            $configObj = [PSCustomObject]@{
                FileName = $_.Name
                ClientName = $config.ClientName
                TenantId = $config.TenantId
                ClientId = $config.ClientId
                SecretExpiry = $config.SecretExpiry
            }
            [void]$configs.Add($configObj)
        } catch {
            Write-Host "Could not load config $($_.Name): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # Return as array to ensure it's IEnumerable even with single item
    return ,$configs.ToArray()
}

# Function to load a specific client config by filename
function Load-ClientConfig {
    param([string]$FileName)

    $configPath = "$env:USERPROFILE\.powersearch\clients\$FileName"

    if (Test-Path $configPath) {
        try {
            $configFile = Get-Content $configPath -Raw | ConvertFrom-Json
            $secureConfig = ConvertTo-SecureString $configFile.EncryptedConfig
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureConfig)
            $configJson = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

            $config = $configJson | ConvertFrom-Json
            Write-Host "Loaded client config '$($config.ClientName)' (expires: $($config.SecretExpiry))" -ForegroundColor Green
            return $config
        } catch {
            Write-Host "Could not load client config: $($_.Exception.Message)" -ForegroundColor Yellow
            return $null
        }
    }
    return $null
}

# Function to save client configuration
function Save-ClientConfig {
    param(
        [string]$ClientName,
        [string]$EnvironmentUrl,
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$SecretExpiry
    )

    $configDir = "$env:USERPROFILE\.powersearch\clients"
    if (-not (Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }

    # Create a safe filename from client name
    $safeFileName = $ClientName -replace '[^\w\s-]', '' -replace '\s+', '_'
    $configPath = "$configDir\$safeFileName.json"

    $configData = @{
        ClientName = $ClientName
        EnvironmentUrl = $EnvironmentUrl
        TenantId = $TenantId
        ClientId = $ClientId
        ClientSecret = $ClientSecret
        SecretExpiry = $SecretExpiry
        CreatedDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }

    # Encrypt and save
    $configJson = $configData | ConvertTo-Json
    $secureConfig = ConvertTo-SecureString $configJson -AsPlainText -Force
    $encryptedConfig = ConvertFrom-SecureString $secureConfig

    $configFile = @{
        EncryptedConfig = $encryptedConfig
    } | ConvertTo-Json

    $configFile | Out-File $configPath -Encoding UTF8
    Write-Host "Client config saved to: $configPath" -ForegroundColor Green

    return "$safeFileName.json"
}

# Function to load app registration config (backward compatibility with old single config)
function Load-AppConfig {
    $configPath = "$env:USERPROFILE\.powersearch\app_config.json"

    if (Test-Path $configPath) {
        try {
            $configFile = Get-Content $configPath -Raw | ConvertFrom-Json
            $secureConfig = ConvertTo-SecureString $configFile.EncryptedConfig
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureConfig)
            $configJson = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

            $config = $configJson | ConvertFrom-Json
            Write-Host "Loaded app registration config (expires: $($config.SecretExpiry))" -ForegroundColor Green
            return $config
        } catch {
            Write-Host "Could not load app config: $($_.Exception.Message)" -ForegroundColor Yellow
            return $null
        }
    }
    return $null
}

# Function to get token using client credentials (app registration)
function Get-TokenWithClientCredentials {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    try {
        $authority = "https://login.microsoftonline.com/$TenantId"
        $tokenUrl = "$authority/oauth2/v2.0/token"

        $body = @{
            client_id = $ClientId
            client_secret = $ClientSecret
            scope = "https://service.flow.microsoft.com/.default"
            grant_type = "client_credentials"
        }

        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"

        # Token expires in seconds, convert to DateTime
        $expiryTime = (Get-Date).AddSeconds($response.expires_in)
        $script:SyncHash.FlowApiTokenExpiry = $expiryTime

        Write-Host "Successfully acquired token using app registration (expires: $($expiryTime.ToString('HH:mm')))" -ForegroundColor Green
        return $response.access_token
    } catch {
        $errorDetails = $_.Exception.Message

        # Try to get more detailed error from response
        if ($_.Exception.Response) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $responseBody = $reader.ReadToEnd()
                $reader.Close()
                $errorObj = $responseBody | ConvertFrom-Json
                if ($errorObj.error_description) {
                    $errorDetails = $errorObj.error_description
                }
            } catch {
                # Ignore if can't parse error details
            }
        }

        Write-Host "Failed to get token with client credentials: $errorDetails" -ForegroundColor Red
        Write-Host "Common causes:" -ForegroundColor Yellow
        Write-Host "  - Admin consent not granted (check Azure Portal > App registrations > API permissions)" -ForegroundColor Gray
        Write-Host "  - Client secret is incorrect or expired" -ForegroundColor Gray
        Write-Host "  - Tenant ID or Client ID is incorrect" -ForegroundColor Gray
        return $null
    }
}

# Function to save token to disk (encrypted to current user)
function Save-FlowApiToken {
    param([string]$Token)

    $configDir = "$env:USERPROFILE\.powersearch"
    if (-not (Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }

    # Use Windows DPAPI to encrypt token (only current user can decrypt)
    $secureToken = ConvertTo-SecureString $Token -AsPlainText -Force
    $encryptedToken = ConvertFrom-SecureString $secureToken

    $tokenData = @{
        Token = $encryptedToken
        SavedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }

    $tokenData | ConvertTo-Json | Out-File "$configDir\flowapi_token.json" -Encoding UTF8
}

# Function to load token from disk
function Load-FlowApiToken {
    $tokenPath = "$env:USERPROFILE\.powersearch\flowapi_token.json"

    if (Test-Path $tokenPath) {
        try {
            $tokenData = Get-Content $tokenPath -Raw | ConvertFrom-Json

            # Check if token is older than 50 minutes (tokens expire after 1 hour, so refresh before that)
            $savedDate = [DateTime]::Parse($tokenData.SavedDate)
            $tokenAge = (Get-Date) - $savedDate
            $expiryTime = $savedDate.AddMinutes(60)

            if ($tokenAge.TotalMinutes -gt 50) {
                Write-Host "Cached token is $([int]$tokenAge.TotalMinutes) minutes old and likely expired" -ForegroundColor Yellow
                $script:SyncHash.FlowApiTokenExpiry = $null
                return $null
            }

            $secureToken = ConvertTo-SecureString $tokenData.Token
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureToken)
            $token = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

            # Store expiry time
            $script:SyncHash.FlowApiTokenExpiry = $expiryTime

            Write-Host "Loaded cached Flow API token from $($tokenData.SavedDate) (age: $([int]$tokenAge.TotalMinutes) minutes, expires: $($expiryTime.ToString('HH:mm')))" -ForegroundColor Green
            return $token
        } catch {
            Write-Host "Could not load cached token: $($_.Exception.Message)" -ForegroundColor Yellow
            $script:SyncHash.FlowApiTokenExpiry = $null
            return $null
        }
    }
    $script:SyncHash.FlowApiTokenExpiry = $null
    return $null
}

# Function to get Power Automate API token using device code flow
# Uses Microsoft's built-in Azure PowerShell client ID - no app registration needed!
function Get-PowerAutomateApiToken {
    param(
        [string]$TenantId = "common"
    )
    
    try {
        Write-Host "Starting Power Automate API authentication..." -ForegroundColor Cyan

        # Use Microsoft Power Apps first-party client ID
        # This supports device code flow for Power Platform APIs
        $ClientId = "a8f7a65c-f5ba-4859-b2d6-df772c264e9d"
        $Resource = "https://service.flow.microsoft.com/"

        # Use v1 OAuth endpoint with resource parameter
        $authority = "https://login.microsoftonline.com/$TenantId"

        Write-Host "Requesting device code..." -ForegroundColor Cyan
        Write-Host "Client ID: $ClientId" -ForegroundColor Gray
        Write-Host "Tenant ID: $TenantId" -ForegroundColor Gray
        Write-Host "Resource: $Resource" -ForegroundColor Gray

        # Start device code flow using v1 endpoint
        $deviceCodeResponse = Invoke-RestMethod -Uri "$authority/oauth2/devicecode" -Method Post -Body @{
            client_id = $ClientId
            resource = $Resource
        } -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop

        Write-Host "Device code received: $($deviceCodeResponse.user_code)" -ForegroundColor Green

        # Copy code to clipboard FIRST
        [System.Windows.Clipboard]::SetText($deviceCodeResponse.user_code)

        # Show device code message to user
        $message = @"
To authenticate with Power Automate API:

Device Code: $($deviceCodeResponse.user_code)
(Code copied to clipboard!)

1. Your browser will open automatically
2. Paste the code (Ctrl+V) or enter it manually
3. Sign in with your Microsoft account

Click OK and then complete the sign-in in your browser...
"@

        [System.Windows.MessageBox]::Show($message, "Power Automate Authentication", "OK", "Information")

        # Open browser after showing message
        $verificationUrl = if ($deviceCodeResponse.verification_uri) { $deviceCodeResponse.verification_uri } else { $deviceCodeResponse.verification_url }
        if ($verificationUrl) {
            Start-Process $verificationUrl
        } else {
            # Default Microsoft device login URL
            Start-Process "https://microsoft.com/devicelogin"
        }

        # Poll for token
        Write-Host "Waiting for user to complete authentication..." -ForegroundColor Cyan
        $tokenResponse = $null
        $interval = if ($deviceCodeResponse.interval) { $deviceCodeResponse.interval } else { 5 }
        $expiresIn = if ($deviceCodeResponse.expires_in) { $deviceCodeResponse.expires_in } else { 900 }
        $startTime = Get-Date
        $pollCount = 0

        while (((Get-Date) - $startTime).TotalSeconds -lt $expiresIn) {
            Start-Sleep -Seconds $interval
            $pollCount++
            Write-Host "Polling attempt $pollCount..." -ForegroundColor Gray

            try {
                # Use v1 token endpoint (matches the devicecode endpoint)
                # v1 uses grant_type=device_code (not the URN format used in v2)
                $tokenResponse = Invoke-RestMethod -Uri "$authority/oauth2/token" -Method Post -Body @{
                    grant_type = "device_code"
                    client_id = $ClientId
                    resource = $Resource
                    code = $deviceCodeResponse.device_code
                } -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop

                # Success!
                Write-Host "Authentication successful!" -ForegroundColor Green
                $token = $tokenResponse.access_token
                
                # Debug: show token info
                if ($tokenResponse.expires_in) {
                    Write-Host "Token expires in: $($tokenResponse.expires_in) seconds" -ForegroundColor Cyan
                }
                Write-Host "Token length: $($token.Length) characters" -ForegroundColor Cyan
                
                # Decode JWT to see audience and other claims
                try {
                    $tokenParts = $token.Split('.')
                    if ($tokenParts.Count -ge 2) {
                        $payload = $tokenParts[1]
                        # Add padding if needed
                        $padding = 4 - ($payload.Length % 4)
                        if ($padding -lt 4) { $payload += ('=' * $padding) }
                        $decodedPayload = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload))
                        $claims = $decodedPayload | ConvertFrom-Json
                        Write-Host "Token audience (aud): $($claims.aud)" -ForegroundColor Cyan
                        Write-Host "Token issuer (iss): $($claims.iss)" -ForegroundColor Cyan
                        Write-Host "Token app ID (appid): $($claims.appid)" -ForegroundColor Cyan
                        if ($claims.scp) { Write-Host "Token scopes (scp): $($claims.scp)" -ForegroundColor Cyan }
                    }
                } catch {
                    Write-Host "Could not decode token: $($_.Exception.Message)" -ForegroundColor Yellow
                }

                # Save token to disk for future use
                Save-FlowApiToken -Token $token

                return $token
            } catch {
                $errorMsg = $_.Exception.Message
                Write-Host "Poll response: $errorMsg" -ForegroundColor Gray

                # Check if it's "authorization_pending" which is expected
                if ($errorMsg -match "authorization_pending") {
                    # Still waiting for user, continue polling
                    continue
                } elseif ($errorMsg -match "authorization_declined") {
                    Write-Host "User declined authentication" -ForegroundColor Red
                    return $null
                } elseif ($errorMsg -match "expired_token") {
                    Write-Host "Device code expired" -ForegroundColor Red
                    return $null
                } else {
                    # Other error, but continue polling
                    continue
                }
            }
        }

        Write-Host "Authentication timed out" -ForegroundColor Red
        return $null

    } catch {
        Write-Host "Error getting Power Automate API token: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        return $null
    }
}

# Simplified function to authenticate with Power Automate API
# No app registration needed - uses Microsoft's built-in client ID
function Show-PowerAutomateApiSetup {
    try {
        $message = @"
To view Cloud Flow run history, you need to authenticate with Power Automate API.

This uses Microsoft's built-in authentication - no app registration required!

Note: You will only see run history for flows you have permission to access.
Flows owned by others may show "No access" if you don't have sharing permissions.

Click OK to start authentication.
"@

        $result = [System.Windows.MessageBox]::Show($message, "Power Automate Authentication", "OKCancel", "Information")

        if ($result -eq "OK") {
            # Get tenant ID from connection if available
            $tenantId = "common"
            if ($script:SyncHash.Connection -and $script:SyncHash.Connection.TenantId) {
                $tenantId = $script:SyncHash.Connection.TenantId.ToString()
            }
            
            # Get token using built-in Microsoft client ID
            $token = Get-PowerAutomateApiToken -TenantId $tenantId

            if ($token) {
                $script:SyncHash.FlowApiToken = $token
                # Set expiry time to 60 minutes from now
                $script:SyncHash.FlowApiTokenExpiry = (Get-Date).AddMinutes(60)

                # Update UI status
                Update-FlowApiTokenStatus

                [System.Windows.MessageBox]::Show("Authentication successful! Flow run history will now be available.", "Success", "OK", "Information")
                return $true
            } else {
                [System.Windows.MessageBox]::Show("Authentication failed. Please try again later.", "Error", "OK", "Error")
                return $false
            }
        }

        return $false
    } catch {
        [System.Windows.MessageBox]::Show("Error during authentication: $($_.Exception.Message)", "Error", "OK", "Error")
        return $false
    }
}

# Function to show credentials management dialog
function Show-ManageCredentialsDialog {
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Manage Client Credentials" Height="750" Width="700"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
        MinHeight="600" MinWidth="650"
        Background="#1E1E1E">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="3"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#1C97EA"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style TargetType="PasswordBox">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#CCCCCC"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Padding" Value="0,0,0,4"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Saved Clients Section -->
        <StackPanel Grid.Row="0" Margin="0,0,0,15">
            <Label Content="Saved Client Configurations" FontSize="14" FontWeight="Bold" Foreground="White"/>
            <ListBox x:Name="lstClients" Height="120" Background="#2D2D30"
                     Foreground="White" BorderBrush="#3F3F46" FontSize="11">
                <ListBox.ItemTemplate>
                    <DataTemplate>
                        <StackPanel Orientation="Horizontal" Margin="0,2">
                            <TextBlock Text="{Binding ClientName}" FontWeight="Bold" Width="180" Foreground="White"/>
                            <TextBlock Text="Tenant: " Foreground="#888888"/>
                            <TextBlock Text="{Binding TenantId}" Foreground="#888888" Width="230"/>
                            <TextBlock Text=" | Expires: " Foreground="#888888"/>
                            <TextBlock Text="{Binding SecretExpiry}" Foreground="#888888"/>
                        </StackPanel>
                    </DataTemplate>
                </ListBox.ItemTemplate>
            </ListBox>
            <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                <Button x:Name="btnLoadClient" Content="Load Selected" Width="110" Margin="0,0,8,0"/>
                <Button x:Name="btnDeleteClient" Content="Delete Selected" Width="110" Background="#C42B1C"/>
            </StackPanel>
        </StackPanel>

        <!-- Add/Edit Client Section with ScrollViewer -->
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
            <StackPanel>
                <Label Content="Add New Client Configuration" FontSize="14" FontWeight="Bold" Foreground="White" Margin="0,0,0,10"/>

                <Label Content="Client Name (e.g., Contoso Inc):"/>
                <TextBox x:Name="txtClientName" Margin="0,0,0,10"/>

                <Label Content="Environment URL (e.g., https://yourorg.crm.dynamics.com):"/>
                <TextBox x:Name="txtEnvironmentUrl" Margin="0,0,0,10"/>

                <Label Content="Tenant ID:"/>
                <TextBox x:Name="txtTenantId" Margin="0,0,0,10"/>

                <Label Content="Application (Client) ID:"/>
                <TextBox x:Name="txtClientId" Margin="0,0,0,10"/>

                <Label Content="Client Secret:"/>
                <PasswordBox x:Name="txtClientSecret" Margin="0,0,0,10"/>

                <Label Content="Secret Expiry Date (yyyy-MM-dd):"/>
                <TextBox x:Name="txtSecretExpiry" Margin="0,0,0,10"/>

                <TextBlock x:Name="txtStatus" Text="" Margin="0,5,0,10" FontSize="11" TextWrapping="Wrap"/>
            </StackPanel>
        </ScrollViewer>

        <!-- Action Buttons -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="btnTestConnection" Content="Test Connection" Width="120" Margin="0,0,8,0"/>
            <Button x:Name="btnSaveClient" Content="Save Client" Width="120" Margin="0,0,8,0"/>
            <Button x:Name="btnClose" Content="Close" Width="120" Background="#3F3F46"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $credDialog = [System.Windows.Markup.XamlReader]::Load($reader)

    # Get controls
    $lstClients = $credDialog.FindName("lstClients")
    $btnLoadClient = $credDialog.FindName("btnLoadClient")
    $btnDeleteClient = $credDialog.FindName("btnDeleteClient")
    $txtClientName = $credDialog.FindName("txtClientName")
    $txtEnvironmentUrl = $credDialog.FindName("txtEnvironmentUrl")
    $txtTenantId = $credDialog.FindName("txtTenantId")
    $txtClientId = $credDialog.FindName("txtClientId")
    $txtClientSecret = $credDialog.FindName("txtClientSecret")
    $txtSecretExpiry = $credDialog.FindName("txtSecretExpiry")
    $txtStatus = $credDialog.FindName("txtStatus")
    $btnTestConnection = $credDialog.FindName("btnTestConnection")
    $btnSaveClient = $credDialog.FindName("btnSaveClient")
    $btnClose = $credDialog.FindName("btnClose")

    # Load saved clients
    $savedClients = Get-SavedClientConfigs
    $lstClients.ItemsSource = $savedClients

    # Load selected client
    $btnLoadClient.Add_Click({
        $selected = $lstClients.SelectedItem
        if ($selected) {
            $config = Load-ClientConfig -FileName $selected.FileName
            if ($config) {
                # Authenticate with this client
                $token = Get-TokenWithClientCredentials -TenantId $config.TenantId -ClientId $config.ClientId -ClientSecret $config.ClientSecret
                if ($token) {
                    $script:SyncHash.FlowApiToken = $token
                    Update-FlowApiTokenStatus

                    # If the config has an Environment URL, populate it in the main window
                    if ($config.EnvironmentUrl -and $script:txtEnvironmentUrl) {
                        try {
                            $script:txtEnvironmentUrl.Text = $config.EnvironmentUrl
                            Write-Host "Environment URL set to: $($config.EnvironmentUrl)" -ForegroundColor Cyan
                        } catch {
                            Write-Host "Could not set Environment URL in main window: $($_.Exception.Message)" -ForegroundColor Yellow
                        }
                    }

                    $txtStatus.Text = "[SUCCESS] Authenticated as '$($config.ClientName)'"
                    $txtStatus.Foreground = "LimeGreen"

                    # Close the credentials dialog
                    $credDialog.Close()

                    # Show success message with next steps
                    $message = "Successfully authenticated as '$($config.ClientName)'!`n`n"
                    if ($config.EnvironmentUrl) {
                        $message += "Environment URL has been populated.`n"
                    } else {
                        $message += "Please enter your Environment URL.`n"
                    }
                    $message += "`nClick 'Connect' to establish connection to Dataverse."
                    [System.Windows.MessageBox]::Show($message, "Authentication Successful", "OK", "Information")
                } else {
                    $txtStatus.Text = "[ERROR] Failed to authenticate with '$($config.ClientName)'"
                    $txtStatus.Foreground = "Red"
                }
            }
        } else {
            [System.Windows.MessageBox]::Show("Please select a client from the list.", "No Selection", "OK", "Warning")
        }
    })

    # Delete selected client
    $btnDeleteClient.Add_Click({
        $selected = $lstClients.SelectedItem
        if ($selected) {
            $result = [System.Windows.MessageBox]::Show(
                "Are you sure you want to delete the configuration for '$($selected.ClientName)'?",
                "Confirm Delete",
                "YesNo",
                "Warning"
            )
            if ($result -eq "Yes") {
                $configPath = "$env:USERPROFILE\.powersearch\clients\$($selected.FileName)"
                if (Test-Path $configPath) {
                    Remove-Item $configPath -Force
                    $txtStatus.Text = "[OK] Deleted '$($selected.ClientName)'"
                    $txtStatus.Foreground = "Yellow"
                    # Refresh list
                    $lstClients.ItemsSource = $null
                    $lstClients.ItemsSource = Get-SavedClientConfigs
                }
            }
        } else {
            [System.Windows.MessageBox]::Show("Please select a client from the list.", "No Selection", "OK", "Warning")
        }
    })

    # Test connection
    $btnTestConnection.Add_Click({
        $tenantId = $txtTenantId.Text.Trim()
        $clientId = $txtClientId.Text.Trim()
        $clientSecret = $txtClientSecret.Password

        if (-not $tenantId -or -not $clientId -or -not $clientSecret) {
            $txtStatus.Text = "[WARNING] Please fill in Tenant ID, Client ID, and Client Secret"
            $txtStatus.Foreground = "Yellow"
            return
        }

        $txtStatus.Text = "Testing connection..."
        $txtStatus.Foreground = "Cyan"

        $token = Get-TokenWithClientCredentials -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret
        if ($token) {
            $txtStatus.Text = "[SUCCESS] Connection test successful! Credentials are valid."
            $txtStatus.Foreground = "LimeGreen"
        } else {
            $txtStatus.Text = "[ERROR] Connection test failed. Please verify your credentials."
            $txtStatus.Foreground = "Red"
        }
    })

    # Save client
    $btnSaveClient.Add_Click({
        $clientName = $txtClientName.Text.Trim()
        $environmentUrl = $txtEnvironmentUrl.Text.Trim()
        $tenantId = $txtTenantId.Text.Trim()
        $clientId = $txtClientId.Text.Trim()
        $clientSecret = $txtClientSecret.Password
        $secretExpiry = $txtSecretExpiry.Text.Trim()

        if (-not $clientName -or -not $tenantId -or -not $clientId -or -not $clientSecret) {
            $txtStatus.Text = "[WARNING] Please fill in Client Name, Tenant ID, Client ID, and Client Secret (required fields)"
            $txtStatus.Foreground = "Yellow"
            return
        }

        # Default expiry if not provided
        if (-not $secretExpiry) {
            $secretExpiry = (Get-Date).AddYears(2).ToString('yyyy-MM-dd')
        }

        $fileName = Save-ClientConfig -ClientName $clientName -EnvironmentUrl $environmentUrl -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret -SecretExpiry $secretExpiry

        $txtStatus.Text = "[SUCCESS] Client '$clientName' saved successfully!"
        $txtStatus.Foreground = "LimeGreen"

        # Clear input fields
        $txtClientName.Text = ""
        $txtEnvironmentUrl.Text = ""
        $txtTenantId.Text = ""
        $txtClientId.Text = ""
        $txtClientSecret.Password = ""
        $txtSecretExpiry.Text = ""

        # Refresh list
        $lstClients.ItemsSource = $null
        $lstClients.ItemsSource = Get-SavedClientConfigs
    })

    # Close button
    $btnClose.Add_Click({
        $credDialog.Close()
    })

    $credDialog.ShowDialog() | Out-Null
}

# Function to show AI summary in a window
function Show-AISummaryWindow {
    param(
        [string]$Title,
        [string]$Content
    )

    # Escape the title for XML
    $escapedTitle = [System.Security.SecurityElement]::Escape($Title)

    [xml]$summaryXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="$escapedTitle"
        Height="500" Width="700"
        WindowStartupLocation="CenterOwner"
        Background="#F5F5F5">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" BorderBrush="#3C1053" BorderThickness="2" Background="White">
            <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="10">
                <TextBox x:Name="txtSummary"
                         TextWrapping="Wrap"
                         IsReadOnly="True"
                         BorderThickness="0"
                         Background="Transparent"
                         FontSize="12"
                         FontFamily="Segoe UI"/>
            </ScrollViewer>
        </Border>

        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnCopy" Content="Copy to Clipboard" Width="130" Padding="8,4" Margin="0,0,8,0"
                    Background="#E87722" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand"/>
            <Button x:Name="btnExport" Content="Export to File" Width="110" Padding="8,4" Margin="0,0,8,0"
                    Background="#3C1053" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand"/>
            <Button x:Name="btnClose" Content="Close" Width="70" Padding="8,4"
                    Background="#666666" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $summaryReader = (New-Object System.Xml.XmlNodeReader $summaryXaml)
    $summaryWindow = [Windows.Markup.XamlReader]::Load($summaryReader)
    $summaryWindow.Owner = $window

    $txtSummary = $summaryWindow.FindName("txtSummary")
    $btnCopy = $summaryWindow.FindName("btnCopy")
    $btnExport = $summaryWindow.FindName("btnExport")
    $btnClose = $summaryWindow.FindName("btnClose")

    # Set the content after the window is created to avoid XML escaping issues
    $txtSummary.Text = $Content

    # Copy to clipboard
    $btnCopy.Add_Click({
        [System.Windows.Clipboard]::SetText($txtSummary.Text)
        $btnCopy.Content = "Copied!"
        Start-Sleep -Milliseconds 1000
        $btnCopy.Content = "Copy to Clipboard"
    })

    # Export to file
    $btnExport.Add_Click({
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "Text files (*.txt)|*.txt|Markdown files (*.md)|*.md|All files (*.*)|*.*"
        $saveDialog.FileName = "AI_Summary_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

        if ($saveDialog.ShowDialog() -eq $true) {
            $txtSummary.Text | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
            [System.Windows.MessageBox]::Show("Summary exported to:`n$($saveDialog.FileName)", "Export Complete", "OK", "Information")
        }
    })

    # Close button
    $btnClose.Add_Click({
        $summaryWindow.Close()
    })

    $summaryWindow.ShowDialog() | Out-Null
}

# Function to get AI summary of an item
function Start-AISummary {
    param($Item)

    if (-not $Item) {
        [System.Windows.MessageBox]::Show("Please select an item from the results grid first.", "No Selection", "OK", "Warning")
        return
    }

    $txtStatus.Text = "Fetching details for AI analysis..."
    $progressBar.IsIndeterminate = $true
    [System.Windows.Forms.Application]::DoEvents()

    try {
        # Fetch full details from Dataverse based on item type
        $fullDetails = $null
        $conn = $script:SyncHash.Connection

        switch -Wildcard ($Item.Type) {
            "Cloud Flow" {
                $fetch = @"
<fetch>
  <entity name='workflow'>
    <attribute name='name' />
    <attribute name='description' />
    <attribute name='clientdata' />
    <attribute name='statecode' />
    <attribute name='modifiedon' />
    <attribute name='createdon' />
    <attribute name='primaryentity' />
    <filter type='and'>
      <condition attribute='workflowid' operator='eq' value='$($Item.Id)' />
    </filter>
  </entity>
</fetch>
"@
                $result = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch
                if ($result.CrmRecords.Count -gt 0) {
                    $flow = $result.CrmRecords[0]
                    $fullDetails = @{
                        Name = $flow.name
                        Description = $flow.description
                        Definition = $flow.clientdata
                        State = if ($flow.statecode -eq 1) { "Active" } else { "Inactive" }
                        PrimaryEntity = $flow.primaryentity
                        Created = $flow.createdon
                        Modified = $flow.modifiedon
                    }
                }
            }
            "Classic Workflow" {
                $fetch = @"
<fetch>
  <entity name='workflow'>
    <attribute name='name' />
    <attribute name='description' />
    <attribute name='xaml' />
    <attribute name='statecode' />
    <attribute name='modifiedon' />
    <attribute name='primaryentity' />
    <filter type='and'>
      <condition attribute='workflowid' operator='eq' value='$($Item.Id)' />
    </filter>
  </entity>
</fetch>
"@
                $result = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch
                if ($result.CrmRecords.Count -gt 0) {
                    $wf = $result.CrmRecords[0]
                    $fullDetails = @{
                        Name = $wf.name
                        Description = $wf.description
                        XAML = $wf.xaml
                        State = if ($wf.statecode -eq 1) { "Active" } else { "Inactive" }
                        PrimaryEntity = $wf.primaryentity
                        Modified = $wf.modifiedon
                    }
                }
            }
            "Custom Action" {
                $fetch = @"
<fetch>
  <entity name='workflow'>
    <attribute name='name' />
    <attribute name='description' />
    <attribute name='xaml' />
    <attribute name='statecode' />
    <attribute name='modifiedon' />
    <filter type='and'>
      <condition attribute='workflowid' operator='eq' value='$($Item.Id)' />
    </filter>
  </entity>
</fetch>
"@
                $result = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch
                if ($result.CrmRecords.Count -gt 0) {
                    $action = $result.CrmRecords[0]
                    $fullDetails = @{
                        Name = $action.name
                        Description = $action.description
                        XAML = $action.xaml
                        State = if ($action.statecode -eq 1) { "Active" } else { "Inactive" }
                        Modified = $action.modifiedon
                    }
                }
            }
            "Web Resource*" {
                # Note: Not fetching content as it can be very large
                $fullDetails = @{
                    Name = $Item.Name
                    Type = $Item.Type
                    Note = "Web resource content not fetched due to size. Summary based on name and metadata only."
                }
            }
            default {
                $fullDetails = @{
                    Name = $Item.Name
                    Type = $Item.Type
                    Score = $Item.Score
                    Reasons = $Item.Reasons
                }
            }
        }

        if ($fullDetails) {
            $txtStatus.Text = "Requesting AI analysis..."
            [System.Windows.Forms.Application]::DoEvents()

            # Call Claude API for analysis
            $summary = Get-ClaudeAnalysis -ItemDetails $fullDetails -ItemType $Item.Type -MatchReasons $Item.Reasons

            # Display summary in a scrollable window
            Show-AISummaryWindow -Title "AI Analysis: $($Item.Name)" -Content $summary
            $txtStatus.Text = "AI analysis complete"
        } else {
            [System.Windows.MessageBox]::Show("Could not fetch details for this item.", "Error", "OK", "Error")
            $txtStatus.Text = "Failed to fetch item details"
        }

    } catch {
        $errorMsg = $_.Exception.Message
        [System.Windows.MessageBox]::Show("Error getting AI summary: $errorMsg", "Error", "OK", "Error")
        $txtStatus.Text = "Error: $errorMsg"
    } finally {
        $progressBar.IsIndeterminate = $false
    }
}

# Function to get saved API key
function Get-SavedApiKey {
    $configPath = "$env:USERPROFILE\.powersearch\config.json"
    if (Test-Path $configPath) {
        try {
            $config = Get-Content $configPath -Raw | ConvertFrom-Json
            if ($config.ApiKey) {
                # Decrypt the API key
                $secureString = ConvertTo-SecureString $config.ApiKey
                $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
                return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            }
        } catch {
            return $null
        }
    }
    return $null
}

# Function to save API key
function Save-ApiKey {
    param([string]$ApiKey)

    $configDir = "$env:USERPROFILE\.powersearch"
    if (-not (Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }

    # Encrypt the API key
    $secureString = ConvertTo-SecureString $ApiKey -AsPlainText -Force
    $encryptedKey = ConvertFrom-SecureString $secureString

    $config = @{
        ApiKey = $encryptedKey
        SavedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }

    $config | ConvertTo-Json | Out-File "$configDir\config.json" -Encoding UTF8
}

# Function to show API key settings window
function Show-ApiKeySettings {
    [CmdletBinding()]
    param()

    $settingsXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Claude API Settings"
        Height="320" Width="550"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5"
        ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Claude API Configuration"
                   FontSize="16" FontWeight="Bold"
                   Foreground="#3C1053" Margin="0,0,0,15"/>

        <TextBlock Grid.Row="1" TextWrapping="Wrap" Margin="0,0,0,15" FontSize="11">
            <Run Text="To use AI summaries, you need a Claude API key from Anthropic."/>
            <LineBreak/>
            <LineBreak/>
            <Run Text="Get your API key at: "/>
            <Hyperlink x:Name="linkConsole" NavigateUri="https://console.anthropic.com/">
                <Run Text="https://console.anthropic.com/"/>
            </Hyperlink>
        </TextBlock>

        <StackPanel Grid.Row="2" Margin="0,0,0,15">
            <TextBlock Text="API Key:" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,5"/>
            <PasswordBox x:Name="txtApiKey"
                         Padding="8,6"
                         FontSize="11"
                         FontFamily="Consolas"/>
            <TextBlock x:Name="txtSavedStatus"
                       Text=""
                       FontSize="10"
                       Foreground="Green"
                       Margin="0,5,0,0"/>
        </StackPanel>

        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnSaveKey" Content="Save API Key" Width="100" Padding="8,4" Margin="0,0,8,0"
                    Background="#E87722" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand"/>
            <Button x:Name="btnTestKey" Content="Test Connection" Width="110" Padding="8,4" Margin="0,0,8,0"
                    Background="#3C1053" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand"/>
            <Button x:Name="btnCloseSettings" Content="Close" Width="70" Padding="8,4"
                    Background="#666666" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $settingsReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($settingsXaml))
    $settingsWindow = [Windows.Markup.XamlReader]::Load($settingsReader)
    $settingsReader.Close()
    try {
        $settingsWindow.Owner = $window
    } catch {
        # Ignore if parent window not available
    }

    $txtApiKey = $settingsWindow.FindName("txtApiKey")
    $txtSavedStatus = $settingsWindow.FindName("txtSavedStatus")
    $btnSaveKey = $settingsWindow.FindName("btnSaveKey")
    $btnTestKey = $settingsWindow.FindName("btnTestKey")
    $btnCloseSettings = $settingsWindow.FindName("btnCloseSettings")
    $linkConsole = $settingsWindow.FindName("linkConsole")

    # Load existing key if available
    $existingKey = Get-SavedApiKey
    if ($existingKey) {
        $txtApiKey.Password = $existingKey
        $txtSavedStatus.Text = "[OK] Saved API key loaded"
    }

    # Handle hyperlink click
    if ($linkConsole) {
        $linkConsole.Add_RequestNavigate({
            param($sender, $e)
            Start-Process $e.Uri.AbsoluteUri
            $e.Handled = $true
        })
    }

    # Save button
    $btnSaveKey.Add_Click({
        $apiKey = $txtApiKey.Password
        if ($apiKey -and $apiKey.Length -gt 0) {
            Save-ApiKey -ApiKey $apiKey
            $txtSavedStatus.Text = "[OK] API key saved successfully!"
            $txtSavedStatus.Foreground = "Green"
        } else {
            $txtSavedStatus.Text = "[WARNING] Please enter an API key"
            $txtSavedStatus.Foreground = "Red"
        }
    }) | Out-Null

    # Test button
    $btnTestKey.Add_Click({
        $apiKey = $txtApiKey.Password
        if ($apiKey -and $apiKey.Length -gt 0) {
            $txtSavedStatus.Text = "Testing connection..."
            $txtSavedStatus.Foreground = "Blue"

            try {
                $headers = @{
                    "x-api-key" = $apiKey
                    "anthropic-version" = "2023-06-01"
                    "content-type" = "application/json"
                }

                $body = @{
                    model = "claude-3-5-sonnet-20241022"
                    max_tokens = 10
                    messages = @(
                        @{
                            role = "user"
                            content = "Hi"
                        }
                    )
                } | ConvertTo-Json -Depth 10

                $response = Invoke-RestMethod -Uri "https://api.anthropic.com/v1/messages" `
                    -Method Post `
                    -Headers $headers `
                    -Body $body `
                    -TimeoutSec 10

                $txtSavedStatus.Text = "[OK] Connection successful!"
                $txtSavedStatus.Foreground = "Green"
            } catch {
                $txtSavedStatus.Text = "[ERROR] Connection failed: $($_.Exception.Message)"
                $txtSavedStatus.Foreground = "Red"
            }
        } else {
            $txtSavedStatus.Text = "[WARNING] Please enter an API key"
            $txtSavedStatus.Foreground = "Red"
        }
    }) | Out-Null

    # Close button
    $btnCloseSettings.Add_Click({
        $settingsWindow.Close()
    }) | Out-Null

    $settingsWindow.ShowDialog() | Out-Null
}

# Function to get Claude session key from Claude Desktop
function Get-ClaudeSessionKey {
    # Try to find Claude Desktop's session file
    $sessionPaths = @(
        "$env:APPDATA\Claude\sessions\active_session.json",
        "$env:LOCALAPPDATA\Claude\sessions\active_session.json",
        "$env:USERPROFILE\.claude\session.json"
    )

    foreach ($path in $sessionPaths) {
        if (Test-Path $path) {
            try {
                $sessionData = Get-Content $path -Raw | ConvertFrom-Json
                if ($sessionData.sessionKey) {
                    return $sessionData.sessionKey
                }
            } catch {
                # Continue to next path
            }
        }
    }

    return $null
}

# Function to call Claude API
function Get-ClaudeAnalysis {
    param(
        [hashtable]$ItemDetails,
        [string]$ItemType,
        [string]$MatchReasons
    )

    # Build the prompt
    $detailsJson = $ItemDetails | ConvertTo-Json -Depth 10
    $prompt = @"
You are analyzing a PowerApps/Dataverse component that was flagged during troubleshooting.

Item Type: $ItemType
Why it was flagged: $MatchReasons

Component Details:
$detailsJson

Please provide a concise analysis covering:
1. What this component does (2-3 sentences)
2. Why it might be relevant to the search criteria
3. Any potential issues or concerns you notice
4. Recommendations for troubleshooting

Keep your response under 300 words and focus on actionable insights.
"@

    # Try saved API key
    $apiKey = Get-SavedApiKey

    # Fall back to environment variable
    if (-not $apiKey) {
        $apiKey = $env:ANTHROPIC_API_KEY
    }

    if (-not $apiKey) {
        # Show settings window to get API key
        Show-ApiKeySettings | Out-Null

        # Try again after settings
        $apiKey = Get-SavedApiKey
        if (-not $apiKey) {
            return "No API key configured. Please configure your API key in the settings window and try again."
        }
    }

    $headers = @{
        "x-api-key" = $apiKey
        "anthropic-version" = "2023-06-01"
        "content-type" = "application/json"
    }

    $body = @{
        model = "claude-3-5-sonnet-20241022"
        max_tokens = 2048
        messages = @(
            @{
                role = "user"
                content = $prompt
            }
        )
    } | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri "https://api.anthropic.com/v1/messages" `
            -Method Post `
            -Headers $headers `
            -Body $body `
            -TimeoutSec 30

        return $response.content[0].text
    } catch {
        $errorDetail = $_.Exception.Message
        $statusCode = ""
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
        }

        return "Error calling Claude API: $errorDetail`n`nStatus Code: $statusCode`n`nPlease check:`n1. Your API key is valid`n2. You have API credits available`n3. Your network connection is working`n`nYou can get an API key from: https://console.anthropic.com/"
    }
}

# Add click handler for Copy buttons in the DataGrid
$dgResults.AddHandler(
    [System.Windows.Controls.Button]::ClickEvent,
    [System.Windows.RoutedEventHandler]{
        param($clickSender, $e)
        if ($e.OriginalSource -is [System.Windows.Controls.Button]) {
            $button = $e.OriginalSource

            # Check if it's a Copy link button
            if ($button.Content -eq "Copy") {
                $url = $button.Tag
                if ($url -and $url -ne "") {
                    [System.Windows.Clipboard]::SetText($url)
                }
            }
        }
    }
)

# Add double-click handler for DataGrid rows to show item details
$dgResults.Add_MouseDoubleClick({
    param($senderObj, $e)
    $selectedItem = $dgResults.SelectedItem
    if ($selectedItem) {
        # Escape special characters for XAML
        $escapedName = [System.Security.SecurityElement]::Escape($selectedItem.Name)
        
        # Create a details window with tabs
        $detailsXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Item Details"
        Width="900" Height="750"
        WindowStartupLocation="CenterOwner"
        Background="White">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Item Details" FontSize="20" FontWeight="Bold" Margin="0,0,0,15"/>

        <TabControl Grid.Row="1">
            <TabItem Header="General">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="10">
                    <StackPanel>
                        <TextBlock Text="Name:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="txtName" IsReadOnly="True" TextWrapping="Wrap" Margin="0,0,0,15" Padding="5"/>

                        <TextBlock Text="Type:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="txtType" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>

                        <TextBlock Text="Solution:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="txtSolution" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>

                        <TextBlock Text="ID:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="txtId" IsReadOnly="True" FontFamily="Consolas" Margin="0,0,0,15" Padding="5"/>

                        <TextBlock Text="State:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="txtState" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>

                        <TextBlock Text="Actions:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="txtActions" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>

                        <TextBlock Text="Modified On:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="txtModifiedOn" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>

                        <TextBlock Text="Primary Entity:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="txtPrimaryEntity" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>

                        <TextBlock Text="Match Reasons:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="txtReasons" IsReadOnly="True" TextWrapping="Wrap" Height="80" Margin="0,0,0,15" Padding="5" VerticalScrollBarVisibility="Auto"/>

                        <TextBlock Text="URL:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="txtUrl" IsReadOnly="True" TextWrapping="Wrap" Margin="0,0,0,15" Padding="5"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="Errors">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="10">
                    <StackPanel>
                        <TextBlock Text="Recent Errors (Last 30 Days):" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="txtRecentErrors" IsReadOnly="True" TextWrapping="Wrap" Height="500" Margin="0,0,0,15" Padding="5" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="10" Background="#FFF5F5"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="Flow Steps">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="10">
                    <StackPanel>
                        <TextBlock Text="Flow Steps/Actions:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="txtFlowSteps" IsReadOnly="True" TextWrapping="Wrap" Height="500" Margin="0,0,0,15" Padding="5" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="10"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="DEBUG - Raw API Data" Background="#FFFFCC">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="10">
                    <StackPanel>
                        <TextBlock Text="Debug: All Properties from Dataverse API" FontWeight="Bold" Margin="0,0,0,5" Foreground="#CC6600"/>
                        <TextBlock Text="This shows ALL data returned by the Dataverse API for this item." TextWrapping="Wrap" Margin="0,0,0,10" Foreground="#666666" FontStyle="Italic"/>
                        
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <Button x:Name="btnRefreshDebug" Content="Refresh Debug Data" Width="150" Height="30" Margin="0,0,10,0" Background="#CC6600" Foreground="White"/>
                            <Button x:Name="btnCopyDebug" Content="Copy to Clipboard" Width="140" Height="30" Margin="0,0,10,0" Background="#3C1053" Foreground="White"/>
                        </StackPanel>
                        
                        <TextBlock Text="Load Large Fields (can be slow):" FontWeight="Bold" Margin="0,10,0,5"/>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <Button x:Name="btnLoadClientData" Content="Load Flow Definition (JSON)" Width="180" Height="28" Margin="0,0,10,0" Background="#0078D4" Foreground="White"/>
                            <Button x:Name="btnLoadRunHistory" Content="Load Run History" Width="140" Height="28" Margin="0,0,10,0" Background="#107C10" Foreground="White"/>
                            <Button x:Name="btnFlowReAuth" Content="🔑 Re-authenticate" Width="130" Height="28" Margin="0,0,10,0" Background="#C50F1F" Foreground="White" ToolTip="Sign in with a different account for Power Automate API"/>
                        </StackPanel>
                        
                        <TextBox x:Name="txtDebugInfo" IsReadOnly="True" TextWrapping="NoWrap" Height="400" Margin="0,0,0,15" Padding="5" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="11" Background="#FFFFF0"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>

        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button Content="Copy URL" Width="100" Height="30" Margin="0,0,10,0" Background="#E87722" Foreground="White" BorderThickness="0" x:Name="btnCopyUrl"/>
            <Button Content="Close" Width="100" Height="30" Background="#666666" Foreground="White" BorderThickness="0" x:Name="btnClose"/>
        </StackPanel>
    </Grid>
</Window>
"@

        try {
            $detailsReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($detailsXaml))
            $detailsWindow = [Windows.Markup.XamlReader]::Load($detailsReader)
            $detailsReader.Close()

            # Get buttons and text boxes
            $btnCopyUrl = $detailsWindow.FindName("btnCopyUrl")
            $btnClose = $detailsWindow.FindName("btnClose")
            $txtFlowSteps = $detailsWindow.FindName("txtFlowSteps")
            $txtRecentErrors = $detailsWindow.FindName("txtRecentErrors")
            
            # Get debug tab controls
            $txtDebugInfo = $detailsWindow.FindName("txtDebugInfo")
            $btnRefreshDebug = $detailsWindow.FindName("btnRefreshDebug")
            $btnCopyDebug = $detailsWindow.FindName("btnCopyDebug")
            $btnLoadClientData = $detailsWindow.FindName("btnLoadClientData")
            $btnLoadRunHistory = $detailsWindow.FindName("btnLoadRunHistory")
            $btnFlowReAuth = $detailsWindow.FindName("btnFlowReAuth")
            
            # Get General tab text boxes
            $txtName = $detailsWindow.FindName("txtName")
            $txtType = $detailsWindow.FindName("txtType")
            $txtSolution = $detailsWindow.FindName("txtSolution")
            $txtId = $detailsWindow.FindName("txtId")
            $txtState = $detailsWindow.FindName("txtState")
            $txtActions = $detailsWindow.FindName("txtActions")
            $txtModifiedOn = $detailsWindow.FindName("txtModifiedOn")
            $txtPrimaryEntity = $detailsWindow.FindName("txtPrimaryEntity")
            $txtReasons = $detailsWindow.FindName("txtReasons")
            $txtUrl = $detailsWindow.FindName("txtUrl")
            
            # Populate General tab fields
            if ($txtName) { $txtName.Text = $selectedItem.Name }
            if ($txtType) { $txtType.Text = $selectedItem.Type }
            if ($txtSolution) { $txtSolution.Text = if ($selectedItem.Solution) { $selectedItem.Solution } else { "N/A" } }
            if ($txtId) { $txtId.Text = if ($selectedItem.Id) { $selectedItem.Id } else { "N/A" } }
            if ($txtState) { $txtState.Text = if ($selectedItem.State) { $selectedItem.State } else { "Unknown" } }
            if ($txtActions) { $txtActions.Text = if ($selectedItem.Actions) { $selectedItem.Actions } else { "N/A" } }
            if ($txtModifiedOn) { $txtModifiedOn.Text = if ($selectedItem.ModifiedOn) { $selectedItem.ModifiedOn } else { "N/A" } }
            if ($txtPrimaryEntity) { $txtPrimaryEntity.Text = if ($selectedItem.PrimaryEntity) { $selectedItem.PrimaryEntity } else { "N/A" } }
            if ($txtReasons) { $txtReasons.Text = if ($selectedItem.Reasons) { $selectedItem.Reasons } else { "N/A" } }
            if ($txtUrl) { $txtUrl.Text = if ($selectedItem.Url) { $selectedItem.Url } else { "No URL available" } }

            # Fetch and display recent errors
            try {
                $recentErrors = Get-RecentErrorDetails -ItemId $selectedItem.Id -ItemType $selectedItem.Type -Connection $script:SyncHash.Connection
                if ($recentErrors -and $recentErrors.Count -gt 0) {
                    $errorText = "Found $($recentErrors.Count) error(s):`r`n`r`n"
                    foreach ($err in $recentErrors) {
                        $errorText += "[$($err.Timestamp)]`r`n"
                        $errorText += "Step: $($err.Step)`r`n"
                        $errorText += "Message: $($err.Message)`r`n"
                        $errorText += "$('-' * 80)`r`n"
                    }
                    $txtRecentErrors.Text = $errorText
                } else {
                    $txtRecentErrors.Text = "No errors found in the last 30 days."
                }
            } catch {
                $txtRecentErrors.Text = "Error fetching error details: $($_.Exception.Message)"
            }

            # Parse and display flow steps if available
            if ($selectedItem.ClientData) {
                try {
                    $flowSteps = ""
                    $clientDataJson = $selectedItem.ClientData | ConvertFrom-Json
                    
                    # Helper function to parse actions recursively
                    function Parse-FlowAction {
                        param($action, $actionName, $indent = "  ", $connRefs = $null)
                        
                        $actionType = if ($action.type) { $action.type } else { "Unknown" }
                        $result = "$indent|- $actionName`r`n"
                        $result += "$indent|    Type: $actionType`r`n"
                        
                        # Show connection reference if available
                        if ($action.inputs -and $action.inputs.host -and $action.inputs.host.connectionName) {
                            $connRefName = $action.inputs.host.connectionName
                            $result += "$indent|    Connection Ref: $connRefName`r`n"
                            
                            # Look up the actual connector from connection references
                            if ($connRefs -and $connRefs.$connRefName) {
                                $connRef = $connRefs.$connRefName
                                if ($connRef.api -and $connRef.api.name) {
                                    $result += "$indent|    Connector: $($connRef.api.name)`r`n"
                                }
                            }
                        }
                        
                        # Show operation for API calls
                        if ($action.inputs -and $action.inputs.host -and $action.inputs.host.operationId) {
                            $result += "$indent|    Operation: $($action.inputs.host.operationId)`r`n"
                        }
                        
                        # Show method and URI for HTTP actions
                        if ($actionType -eq "Http") {
                            if ($action.inputs.method) {
                                $result += "$indent|    Method: $($action.inputs.method)`r`n"
                            }
                            if ($action.inputs.uri) {
                                $uri = $action.inputs.uri
                                if ($uri.Length -gt 80) { $uri = $uri.Substring(0, 80) + "..." }
                                $result += "$indent|    URI: $uri`r`n"
                            }
                        }
                        
                        # Show entity/table for Dataverse actions
                        if ($action.inputs.parameters) {
                            $params = $action.inputs.parameters
                            if ($params.entityName) {
                                $result += "$indent|    Table: $($params.entityName)`r`n"
                            }
                            if ($params.'item/entityName') {
                                $result += "$indent|    Table: $($params.'item/entityName')`r`n"
                            }
                        }
                        
                        # Show run after conditions
                        if ($action.runAfter -and $action.runAfter.PSObject.Properties.Count -gt 0) {
                            $runAfterList = @()
                            foreach ($dep in $action.runAfter.PSObject.Properties) {
                                $conditions = $dep.Value -join ", "
                                $runAfterList += "$($dep.Name) [$conditions]"
                            }
                            if ($runAfterList.Count -gt 0) {
                                $result += "$indent|    Runs After: $($runAfterList -join '; ')`r`n"
                            }
                        }
                        
                        $result += "$indent|`r`n"
                        
                        # Handle Condition (If/Else)
                        if ($actionType -eq "If") {
                            # Show the condition expression if available
                            if ($action.expression) {
                                $exprStr = $action.expression | ConvertTo-Json -Depth 2 -Compress
                                if ($exprStr.Length -gt 100) { $exprStr = $exprStr.Substring(0, 100) + "..." }
                                $result += "$indent|    Condition: $exprStr`r`n"
                            }
                            if ($action.actions) {
                                $result += "$indent+-- [YES BRANCH] --+`r`n"
                                foreach ($subActionName in $action.actions.PSObject.Properties.Name) {
                                    $result += Parse-FlowAction -action $action.actions.$subActionName -actionName $subActionName -indent "$indent|   " -connRefs $connRefs
                                }
                            }
                            if ($action.else -and $action.else.actions) {
                                $result += "$indent+-- [NO BRANCH] ---+`r`n"
                                foreach ($subActionName in $action.else.actions.PSObject.Properties.Name) {
                                    $result += Parse-FlowAction -action $action.else.actions.$subActionName -actionName $subActionName -indent "$indent|   " -connRefs $connRefs
                                }
                            }
                        }
                        # Handle Switch statement
                        elseif ($actionType -eq "Switch") {
                            if ($action.cases) {
                                foreach ($caseName in $action.cases.PSObject.Properties.Name) {
                                    $case = $action.cases.$caseName
                                    $result += "$indent+-- [CASE: $caseName] --+`r`n"
                                    if ($case.actions) {
                                        foreach ($subActionName in $case.actions.PSObject.Properties.Name) {
                                            $result += Parse-FlowAction -action $case.actions.$subActionName -actionName $subActionName -indent "$indent|   " -connRefs $connRefs
                                        }
                                    }
                                }
                            }
                            if ($action.default -and $action.default.actions) {
                                $result += "$indent+-- [DEFAULT] --+`r`n"
                                foreach ($subActionName in $action.default.actions.PSObject.Properties.Name) {
                                    $result += Parse-FlowAction -action $action.default.actions.$subActionName -actionName $subActionName -indent "$indent|   " -connRefs $connRefs
                                }
                            }
                        }
                        # Handle Scope
                        elseif ($actionType -eq "Scope") {
                            if ($action.actions) {
                                $result += "$indent+-- [SCOPE] --+`r`n"
                                foreach ($subActionName in $action.actions.PSObject.Properties.Name) {
                                    $result += Parse-FlowAction -action $action.actions.$subActionName -actionName $subActionName -indent "$indent|   " -connRefs $connRefs
                                }
                            }
                        }
                        # Handle Foreach (Apply to each)
                        elseif ($actionType -eq "Foreach") {
                            if ($action.foreach) {
                                $foreachExpr = $action.foreach
                                if ($foreachExpr.Length -gt 60) { $foreachExpr = $foreachExpr.Substring(0, 60) + "..." }
                                $result += "$indent|    Iterates Over: $foreachExpr`r`n"
                            }
                            if ($action.actions) {
                                $result += "$indent+-- [LOOP BODY] --+`r`n"
                                foreach ($subActionName in $action.actions.PSObject.Properties.Name) {
                                    $result += Parse-FlowAction -action $action.actions.$subActionName -actionName $subActionName -indent "$indent|   " -connRefs $connRefs
                                }
                            }
                        }
                        # Handle Do Until
                        elseif ($actionType -eq "Until") {
                            if ($action.expression) {
                                $exprStr = $action.expression | ConvertTo-Json -Depth 2 -Compress
                                if ($exprStr.Length -gt 100) { $exprStr = $exprStr.Substring(0, 100) + "..." }
                                $result += "$indent|    Until: $exprStr`r`n"
                            }
                            if ($action.limit) {
                                $result += "$indent|    Limit: Count=$($action.limit.count), Timeout=$($action.limit.timeout)`r`n"
                            }
                            if ($action.actions) {
                                $result += "$indent+-- [UNTIL BODY] --+`r`n"
                                foreach ($subActionName in $action.actions.PSObject.Properties.Name) {
                                    $result += Parse-FlowAction -action $action.actions.$subActionName -actionName $subActionName -indent "$indent|   " -connRefs $connRefs
                                }
                            }
                        }
                        
                        return $result
                    }
                    
                    # Get connection references for lookup
                    $connRefs = $null
                    if ($clientDataJson.properties -and $clientDataJson.properties.connectionReferences) {
                        $connRefs = $clientDataJson.properties.connectionReferences
                    }
                    
                    if ($clientDataJson.properties -and $clientDataJson.properties.definition) {
                        $definition = $clientDataJson.properties.definition
                        
                        # CONNECTION REFERENCES SUMMARY
                        if ($connRefs) {
                            $flowSteps += "=" * 60 + "`r`n"
                            $flowSteps += "CONNECTION REFERENCES`r`n"
                            $flowSteps += "=" * 60 + "`r`n"
                            foreach ($connRefName in $connRefs.PSObject.Properties.Name) {
                                $connRef = $connRefs.$connRefName
                                $connectorName = if ($connRef.api -and $connRef.api.name) { $connRef.api.name } else { "Unknown" }
                                $runtimeSource = if ($connRef.runtimeSource) { $connRef.runtimeSource } else { "N/A" }
                                $flowSteps += "  - $connRefName`r`n"
                                $flowSteps += "      Connector: $connectorName`r`n"
                                $flowSteps += "      Source: $runtimeSource`r`n"
                                if ($connRef.connection -and $connRef.connection.connectionReferenceLogicalName) {
                                    $flowSteps += "      Logical Name: $($connRef.connection.connectionReferenceLogicalName)`r`n"
                                }
                            }
                            $flowSteps += "`r`n"
                        }
                        
                        # TRIGGERS
                        if ($definition.triggers) {
                            $flowSteps += "=" * 60 + "`r`n"
                            $flowSteps += "TRIGGERS (What starts this flow)`r`n"
                            $flowSteps += "=" * 60 + "`r`n"
                            foreach ($triggerName in $definition.triggers.PSObject.Properties.Name) {
                                $trigger = $definition.triggers.$triggerName
                                $triggerType = if ($trigger.type) { $trigger.type } else { "Unknown" }
                                $triggerKind = if ($trigger.kind) { " ($($trigger.kind))" } else { "" }
                                
                                $flowSteps += "  TRIGGER: $triggerName`r`n"
                                $flowSteps += "      Type: $triggerType$triggerKind`r`n"
                                
                                # Show recurrence for scheduled triggers
                                if ($trigger.recurrence) {
                                    $freq = $trigger.recurrence.frequency
                                    $interval = $trigger.recurrence.interval
                                    $flowSteps += "      Schedule: Every $interval $freq`r`n"
                                    if ($trigger.recurrence.startTime) {
                                        $flowSteps += "      Start Time: $($trigger.recurrence.startTime)`r`n"
                                    }
                                    if ($trigger.recurrence.timeZone) {
                                        $flowSteps += "      Time Zone: $($trigger.recurrence.timeZone)`r`n"
                                    }
                                }
                                
                                # Show entity/table for Dataverse triggers
                                if ($trigger.inputs -and $trigger.inputs.parameters) {
                                    $params = $trigger.inputs.parameters
                                    if ($params.entityName -or $params.'subscriptionRequest/entityname') {
                                        $entityName = if ($params.entityName) { $params.entityName } else { $params.'subscriptionRequest/entityname' }
                                        $flowSteps += "      Table: $entityName`r`n"
                                    }
                                    if ($params.'subscriptionRequest/message') {
                                        $flowSteps += "      Event: $($params.'subscriptionRequest/message')`r`n"
                                    }
                                    if ($params.filterExpression -or $params.'subscriptionRequest/filterexpression') {
                                        $filter = if ($params.filterExpression) { $params.filterExpression } else { $params.'subscriptionRequest/filterexpression' }
                                        if ($filter.Length -gt 80) { $filter = $filter.Substring(0, 80) + "..." }
                                        $flowSteps += "      Filter: $filter`r`n"
                                    }
                                    if ($params.'subscriptionRequest/scope') {
                                        $flowSteps += "      Scope: $($params.'subscriptionRequest/scope')`r`n"
                                    }
                                }
                                
                                # Show connection reference
                                if ($trigger.inputs -and $trigger.inputs.host -and $trigger.inputs.host.connectionName) {
                                    $flowSteps += "      Connection Ref: $($trigger.inputs.host.connectionName)`r`n"
                                }
                                if ($trigger.inputs -and $trigger.inputs.host -and $trigger.inputs.host.operationId) {
                                    $flowSteps += "      Operation: $($trigger.inputs.host.operationId)`r`n"
                                }
                                
                                # Show conditions/split on
                                if ($trigger.splitOn) {
                                    $flowSteps += "      Split On: $($trigger.splitOn)`r`n"
                                }
                                if ($trigger.conditions -and $trigger.conditions.Count -gt 0) {
                                    $flowSteps += "      Conditions: $($trigger.conditions.Count) condition(s)`r`n"
                                }
                                
                                $flowSteps += "`r`n"
                            }
                        }
                        
                        # ACTIONS
                        if ($definition.actions) {
                            $flowSteps += "=" * 60 + "`r`n"
                            $flowSteps += "ACTIONS (Flow steps)`r`n"
                            $flowSteps += "=" * 60 + "`r`n"
                            $actionNumber = 1
                            foreach ($actionName in $definition.actions.PSObject.Properties.Name) {
                                $action = $definition.actions.$actionName
                                $flowSteps += "`r`n$actionNumber. "
                                $flowSteps += Parse-FlowAction -action $action -actionName $actionName -connRefs $connRefs
                                $actionNumber++
                            }
                        }
                    }
                    
                    if ($flowSteps -eq "") {
                        $flowSteps = "Flow definition not available or couldn't be parsed."
                    }
                    
                    $txtFlowSteps.Text = $flowSteps
                } catch {
                    $txtFlowSteps.Text = "Error parsing flow steps: $($_.Exception.Message)"
                }
            } else {
                $txtFlowSteps.Text = "Flow steps not available for this item type."
            }

            # Load debug information
            try {
                $debugInfo = Get-FlowDebugInfo -FlowId $selectedItem.Id -Connection $script:SyncHash.Connection
                $txtDebugInfo.Text = $debugInfo
            } catch {
                $txtDebugInfo.Text = "Error loading debug information: $($_.Exception.Message)"
            }

            # Refresh Debug button handler
            $btnRefreshDebug.Add_Click({
                try {
                    $txtDebugInfo.Text = "Refreshing..."
                    $debugInfo = Get-FlowDebugInfo -FlowId $selectedItem.Id -Connection $script:SyncHash.Connection
                    $txtDebugInfo.Text = $debugInfo
                } catch {
                    $txtDebugInfo.Text = "Error refreshing debug information: $($_.Exception.Message)"
                }
            })

            # Copy Debug button handler
            $btnCopyDebug.Add_Click({
                if ($txtDebugInfo.Text -and $txtDebugInfo.Text -ne "") {
                    [System.Windows.Clipboard]::SetText($txtDebugInfo.Text)
                    [System.Windows.MessageBox]::Show("Debug info copied to clipboard!", "Success", "OK", "Information")
                }
            })

            # Load Client Data (Flow Definition JSON) button handler
            $btnLoadClientData.Add_Click({
                try {
                    $txtDebugInfo.Text = "Loading flow definition JSON...`r`n`r`n"
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    $flowId = $selectedItem.Id
                    $conn = $script:SyncHash.Connection
                    
                    $fetch = @"
<fetch>
  <entity name='workflow'>
    <attribute name='clientdata' />
    <attribute name='xaml' />
    <attribute name='inputparameters' />
    <filter>
      <condition attribute='workflowid' operator='eq' value='$flowId' />
    </filter>
  </entity>
</fetch>
"@
                    $result = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch -AllRows -ErrorAction Stop
                    
                    if ($result -and $result.CrmRecords -and $result.CrmRecords.Count -gt 0) {
                        $flow = $result.CrmRecords[0]
                        $output = [System.Text.StringBuilder]::new()
                        
                        if ($flow.clientdata) {
                            [void]$output.AppendLine("=" * 80)
                            [void]$output.AppendLine("FLOW DEFINITION (clientdata) - Full JSON")
                            [void]$output.AppendLine("=" * 80)
                            [void]$output.AppendLine("")
                            
                            # Try to pretty-print the JSON
                            try {
                                $jsonObj = $flow.clientdata | ConvertFrom-Json
                                $prettyJson = $jsonObj | ConvertTo-Json -Depth 20
                                [void]$output.AppendLine($prettyJson)
                            } catch {
                                [void]$output.AppendLine($flow.clientdata)
                            }
                            [void]$output.AppendLine("")
                        } else {
                            [void]$output.AppendLine("No clientdata (flow definition) found.")
                        }
                        
                        if ($flow.xaml) {
                            [void]$output.AppendLine("")
                            [void]$output.AppendLine("=" * 80)
                            [void]$output.AppendLine("WORKFLOW XAML (Classic Workflow Definition)")
                            [void]$output.AppendLine("=" * 80)
                            [void]$output.AppendLine("")
                            [void]$output.AppendLine($flow.xaml)
                        }
                        
                        if ($flow.inputparameters) {
                            [void]$output.AppendLine("")
                            [void]$output.AppendLine("=" * 80)
                            [void]$output.AppendLine("INPUT PARAMETERS")
                            [void]$output.AppendLine("=" * 80)
                            [void]$output.AppendLine("")
                            [void]$output.AppendLine($flow.inputparameters)
                        }
                        
                        $txtDebugInfo.Text = $output.ToString()
                    } else {
                        $txtDebugInfo.Text = "Could not load flow definition data."
                    }
                } catch {
                    $txtDebugInfo.Text = "Error loading flow definition: $($_.Exception.Message)"
                }
            })

            # Re-authenticate button handler - Force re-authentication with a different account
            $btnFlowReAuth.Add_Click({
                try {
                    # Clear existing token
                    $script:SyncHash.FlowApiToken = $null
                    $script:SyncHash.FlowApiTokenExpiry = $null
                    
                    # Delete cached token file
                    $tokenPath = "$env:USERPROFILE\.powersearch\flowapi_token.json"
                    if (Test-Path $tokenPath) {
                        Remove-Item $tokenPath -Force -ErrorAction SilentlyContinue
                    }
                    
                    $txtDebugInfo.Text = "Starting authentication...`r`n`r`nA browser window will open - enter the device code shown and sign in.`r`n`r`nMake sure to sign in with an account that has access to this flow!"
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    # Get tenant ID from connection if available
                    $tenantId = "common"
                    if ($script:SyncHash.Connection -and $script:SyncHash.Connection.TenantId) {
                        $tenantId = $script:SyncHash.Connection.TenantId.ToString()
                    }
                    
                    # Run authentication
                    $token = Get-PowerAutomateApiToken -TenantId $tenantId
                    
                    if ($token) {
                        $script:SyncHash.FlowApiToken = $token
                        $script:SyncHash.FlowApiTokenExpiry = (Get-Date).AddMinutes(60)
                        Update-FlowApiTokenStatus
                        $txtDebugInfo.Text = "Authentication successful!`r`n`r`nClick 'Load Run History' to view flow runs."
                        [System.Windows.MessageBox]::Show("Authentication successful! Click 'Load Run History' to view flow runs.", "Success", "OK", "Information")
                    } else {
                        $txtDebugInfo.Text = "Authentication failed or was cancelled."
                    }
                } catch {
                    $txtDebugInfo.Text = "Error during authentication: $($_.Exception.Message)"
                }
            })

            # Load Run History button handler - Uses Power Automate API like FlowExecutionHistory tool
            # Self-contained: handles authentication if needed
            $btnLoadRunHistory.Add_Click({
                try {
                    $flowId = $selectedItem.Id
                    $flowIdUnique = $selectedItem.WorkflowIdUnique
                    $conn = $script:SyncHash.Connection
                    $output = [System.Text.StringBuilder]::new()
                    $thirtyDaysAgo = (Get-Date).AddDays(-30)
                    $thirtyDaysAgoStr = $thirtyDaysAgo.ToString("yyyy-MM-dd")
                    
                    # For Cloud Flows, check if we need to authenticate first
                    if ($selectedItem.Type -eq "Cloud Flow" -and $flowIdUnique) {
                        # Check if we have a valid token
                        $needsAuth = $false
                        $authReason = ""
                        
                        if (-not $script:SyncHash.FlowApiToken) {
                            $needsAuth = $true
                            $authReason = "No authentication token found."
                        } elseif ($script:SyncHash.FlowApiTokenExpiry -and (Get-Date) -gt $script:SyncHash.FlowApiTokenExpiry) {
                            $needsAuth = $true
                            $authReason = "Authentication token has expired."
                        }
                        
                        if ($needsAuth) {
                            # Prompt user to authenticate
                            $authMessage = @"
$authReason

To view Cloud Flow run history, you need to authenticate with Power Automate API.

This uses Microsoft's built-in authentication - no app registration required!

Would you like to authenticate now?
"@
                            $authResult = [System.Windows.MessageBox]::Show($authMessage, "Authentication Required", "YesNo", "Question")
                            
                            if ($authResult -eq "Yes") {
                                # Clear any existing cached token first
                                $script:SyncHash.FlowApiToken = $null
                                $script:SyncHash.FlowApiTokenExpiry = $null
                                
                                # Delete cached token file to force fresh login
                                $tokenPath = "$env:USERPROFILE\.powersearch\flowapi_token.json"
                                if (Test-Path $tokenPath) {
                                    Remove-Item $tokenPath -Force -ErrorAction SilentlyContinue
                                }
                                
                                $txtDebugInfo.Text = "Starting authentication...`r`nPlease complete sign-in in your browser.`r`n`r`nA browser window will open - enter the device code shown and sign in."
                                [System.Windows.Forms.Application]::DoEvents()
                                
                                # Get tenant ID from connection if available
                                $tenantId = "common"
                                if ($conn -and $conn.TenantId) {
                                    $tenantId = $conn.TenantId.ToString()
                                }
                                
                                # Run authentication
                                $token = Get-PowerAutomateApiToken -TenantId $tenantId
                                
                                if ($token) {
                                    $script:SyncHash.FlowApiToken = $token
                                    $script:SyncHash.FlowApiTokenExpiry = (Get-Date).AddMinutes(60)
                                    Update-FlowApiTokenStatus
                                    [System.Windows.MessageBox]::Show("Authentication successful!", "Success", "OK", "Information")
                                } else {
                                    $txtDebugInfo.Text = "Authentication failed or was cancelled.`r`n`r`nYou can still view Dataverse-based run history below."
                                    [System.Windows.Forms.Application]::DoEvents()
                                }
                            }
                        }
                    }
                    
                    $txtDebugInfo.Text = "Loading run history...`r`n`r`n"
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    [void]$output.AppendLine("=" * 80)
                    [void]$output.AppendLine("RUN HISTORY for: $($selectedItem.Name)")
                    [void]$output.AppendLine("Flow ID (workflowid): $flowId")
                    [void]$output.AppendLine("Flow ID Unique (workflowidunique): $flowIdUnique")
                    [void]$output.AppendLine("Looking back 30 days from: $thirtyDaysAgoStr")
                    [void]$output.AppendLine("=" * 80)
                    [void]$output.AppendLine("")
                    
                    # FIRST: Try Power Automate Management API (like FlowExecutionHistory tool)
                    if ($script:SyncHash.FlowApiToken -and $selectedItem.Type -eq "Cloud Flow" -and $flowIdUnique) {
                        [void]$output.AppendLine(">>> POWER AUTOMATE API (Cloud Flow Run History) <<<")
                        [void]$output.AppendLine("-" * 50)
                        
                        try {
                            # Get environment ID - try multiple sources
                            $environmentId = $null
                            
                            # Debug: Show connection properties
                            if ($conn) {
                                if ($conn.TenantId) {
                                    [void]$output.AppendLine("Tenant ID: $($conn.TenantId)")
                                } else {
                                    [void]$output.AppendLine("Tenant ID: Not available from connection")
                                }
                            }
                            
                            # Method 1: Try ConnectedOrgId from CRM Tooling connection
                            if ($conn -and $conn.ConnectedOrgId) {
                                $environmentId = $conn.ConnectedOrgId.ToString()
                                [void]$output.AppendLine("Environment ID (from ConnectedOrgId): $environmentId")
                            }
                            # Method 2: Try OrganizationId property
                            elseif ($conn -and $conn.OrganizationId) {
                                $environmentId = $conn.OrganizationId.ToString()
                                [void]$output.AppendLine("Environment ID (from OrganizationId): $environmentId")
                            }
                            # Method 3: Query organization entity from Dataverse
                            else {
                                try {
                                    $orgFetch = @"
<fetch>
  <entity name='organization'>
    <attribute name='organizationid' />
  </entity>
</fetch>
"@
                                    $orgResult = Get-CrmRecordsByFetch -conn $conn -Fetch $orgFetch -AllRows -ErrorAction Stop
                                    if ($orgResult -and $orgResult.CrmRecords -and $orgResult.CrmRecords.Count -gt 0) {
                                        $environmentId = $orgResult.CrmRecords[0].organizationid.ToString()
                                        [void]$output.AppendLine("Environment ID (from Dataverse query): $environmentId")
                                    }
                                } catch {
                                    [void]$output.AppendLine("Failed to query organization entity: $($_.Exception.Message)")
                                }
                            }
                            
                            [void]$output.AppendLine("")
                            
                            if ($environmentId) {
                                # Build API URL with date filter (like FlowExecutionHistory)
                                $dateFilter = "StartTime gt " + $thirtyDaysAgo.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                                $flowApiUrl = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$environmentId/flows/$flowIdUnique/runs?api-version=2016-11-01&`$top=100&`$filter=$([System.Web.HttpUtility]::UrlEncode($dateFilter))"
                                
                                [void]$output.AppendLine("API URL: $flowApiUrl")
                                [void]$output.AppendLine("")
                                
                                $result = Invoke-RestMethod -Uri $flowApiUrl -Method Get -Headers @{
                                    "Authorization" = "Bearer " + $script:SyncHash.FlowApiToken
                                    "Accept" = "application/json"
                                } -ErrorAction Stop
                                
                                if ($result.value -and $result.value.Count -gt 0) {
                                    [void]$output.AppendLine("Found $($result.value.Count) flow run(s) from Power Automate API")
                                    [void]$output.AppendLine("")
                                    
                                    # Summary statistics
                                    $successful = ($result.value | Where-Object { $_.properties.status -eq "Succeeded" }).Count
                                    $failed = ($result.value | Where-Object { $_.properties.status -eq "Failed" }).Count
                                    $cancelled = ($result.value | Where-Object { $_.properties.status -eq "Cancelled" }).Count
                                    $running = ($result.value | Where-Object { $_.properties.status -eq "Running" }).Count
                                    
                                    [void]$output.AppendLine("Summary: Succeeded=$successful, Failed=$failed, Cancelled=$cancelled, Running=$running")
                                    [void]$output.AppendLine("")
                                    [void]$output.AppendLine("-" * 50)
                                    
                                    foreach ($run in $result.value | Select-Object -First 30) {
                                        $props = $run.properties
                                        $runId = $run.name
                                        $status = $props.status
                                        $startTime = if ($props.startTime) { [DateTime]::Parse($props.startTime).ToLocalTime() } else { $null }
                                        $endTime = if ($props.endTime) { [DateTime]::Parse($props.endTime).ToLocalTime() } else { $null }
                                        
                                        # Calculate duration
                                        $duration = ""
                                        if ($startTime -and $endTime) {
                                            $durationMs = ($endTime - $startTime).TotalMilliseconds
                                            if ($durationMs -lt 1000) {
                                                $duration = "$([int]$durationMs) ms"
                                            } elseif ($durationMs -lt 60000) {
                                                $duration = "$([math]::Round($durationMs / 1000, 1)) sec"
                                            } elseif ($durationMs -lt 3600000) {
                                                $duration = "$([math]::Round($durationMs / 60000, 1)) min"
                                            } else {
                                                $duration = "$([math]::Round($durationMs / 3600000, 2)) hrs"
                                            }
                                        }
                                        
                                        # Status indicator
                                        $statusIcon = switch ($status) {
                                            "Succeeded" { "[OK]" }
                                            "Failed" { "[FAIL]" }
                                            "Cancelled" { "[CANCEL]" }
                                            "Running" { "[RUN]" }
                                            default { "[$status]" }
                                        }
                                        
                                        [void]$output.AppendLine("$statusIcon Run: $runId")
                                        [void]$output.AppendLine("  Status: $status")
                                        if ($startTime) { [void]$output.AppendLine("  Start: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))") }
                                        if ($endTime) { [void]$output.AppendLine("  End: $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))") }
                                        if ($duration) { [void]$output.AppendLine("  Duration: $duration") }
                                        
                                        # Correlation ID
                                        if ($props.correlation -and $props.correlation.clientTrackingId) {
                                            [void]$output.AppendLine("  Correlation ID: $($props.correlation.clientTrackingId)")
                                        }
                                        
                                        # Trigger info
                                        if ($props.trigger -and $props.trigger.name) {
                                            [void]$output.AppendLine("  Trigger: $($props.trigger.name)")
                                        }
                                        
                                        # Error details for failed runs
                                        if ($status -eq "Failed" -and $props.error) {
                                            [void]$output.AppendLine("  Error Code: $($props.error.code)")
                                            if ($props.error.message) {
                                                $errMsg = $props.error.message
                                                if ($errMsg.Length -gt 300) { $errMsg = $errMsg.Substring(0, 300) + "..." }
                                                [void]$output.AppendLine("  Error Message: $errMsg")
                                            }
                                        }
                                        
                                        # Link to run details
                                        $runUrl = "https://make.powerautomate.com/environments/$environmentId/flows/$flowIdUnique/runs/$runId"
                                        [void]$output.AppendLine("  URL: $runUrl")
                                        [void]$output.AppendLine("")
                                    }
                                    
                                    if ($result.value.Count -gt 30) {
                                        [void]$output.AppendLine("... and $($result.value.Count - 30) more runs")
                                    }
                                    
                                    # Check for nextLink (pagination)
                                    if ($result.nextLink) {
                                        [void]$output.AppendLine("")
                                        [void]$output.AppendLine("More runs available via pagination (nextLink present)")
                                    }
                                } else {
                                    [void]$output.AppendLine("No runs found in Power Automate API for the last 30 days.")
                                }
                            } else {
                                [void]$output.AppendLine("Could not determine environment ID.")
                                [void]$output.AppendLine("")
                                [void]$output.AppendLine("Debug - Connection properties:")
                                if ($conn) {
                                    [void]$output.AppendLine("  ConnectedOrgId: $($conn.ConnectedOrgId)")
                                    [void]$output.AppendLine("  OrganizationId: $($conn.OrganizationId)")
                                    [void]$output.AppendLine("  ConnectedOrgUniqueName: $($conn.ConnectedOrgUniqueName)")
                                    [void]$output.AppendLine("  CrmConnectOrgUriActual: $($conn.CrmConnectOrgUriActual)")
                                } else {
                                    [void]$output.AppendLine("  Connection object is null!")
                                }
                            }
                        } catch {
                            $errorMsg = $_.Exception.Message
                            $statusCode = $null
                            $responseBody = $null
                            if ($_.Exception.Response) {
                                $statusCode = [int]$_.Exception.Response.StatusCode
                                try {
                                    $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                                    $responseBody = $reader.ReadToEnd()
                                    $reader.Close()
                                } catch { }
                            }
                            
                            if ($statusCode -eq 401) {
                                [void]$output.AppendLine("Authentication failed (401 Unauthorized).")
                                [void]$output.AppendLine("")
                                [void]$output.AppendLine("Error details: $errorMsg")
                                if ($responseBody) {
                                    [void]$output.AppendLine("Response: $responseBody")
                                }
                                [void]$output.AppendLine("")
                                [void]$output.AppendLine("Click 'Re-authenticate' to sign in again.")
                                $script:SyncHash.FlowApiToken = $null
                                $script:SyncHash.FlowApiTokenExpiry = $null
                                # Clear cached token file
                                $tokenPath = "$env:USERPROFILE\.powersearch\flowapi_token.json"
                                if (Test-Path $tokenPath) {
                                    Remove-Item $tokenPath -Force -ErrorAction SilentlyContinue
                                }
                            } elseif ($statusCode -eq 403) {
                                [void]$output.AppendLine("Access denied (403). Possible causes:")
                                [void]$output.AppendLine("  - The flow is owned by another user and not shared with you")
                                [void]$output.AppendLine("  - The flow is in a different environment")
                                [void]$output.AppendLine("  - You signed in with the wrong account")
                                [void]$output.AppendLine("")
                                [void]$output.AppendLine("Try re-authenticating with a different account.")
                                
                                # Clear current token and offer to re-authenticate
                                $script:SyncHash.FlowApiToken = $null
                                $script:SyncHash.FlowApiTokenExpiry = $null
                                # Clear cached token file
                                $tokenPath = "$env:USERPROFILE\.powersearch\flowapi_token.json"
                                if (Test-Path $tokenPath) {
                                    Remove-Item $tokenPath -Force -ErrorAction SilentlyContinue
                                }
                                
                                # Update display immediately
                                $txtDebugInfo.Text = $output.ToString()
                                [System.Windows.Forms.Application]::DoEvents()
                                
                                # Offer to re-authenticate
                                $reAuthResult = [System.Windows.MessageBox]::Show(
                                    "Access was denied.`n`nThis usually means you're signed in with an account that doesn't have permission to view this flow's run history.`n`nWould you like to sign in with a different account?",
                                    "Re-authenticate?",
                                    "YesNo",
                                    "Question"
                                )
                                
                                if ($reAuthResult -eq "Yes") {
                                    $txtDebugInfo.Text = "Starting authentication...`r`nPlease complete sign-in in your browser.`r`n`r`nA browser window will open - enter the device code shown and sign in."
                                    [System.Windows.Forms.Application]::DoEvents()
                                    
                                    # Get tenant ID from connection if available
                                    $tenantId = "common"
                                    if ($conn -and $conn.TenantId) {
                                        $tenantId = $conn.TenantId.ToString()
                                    }
                                    
                                    $newToken = Get-PowerAutomateApiToken -TenantId $tenantId
                                    if ($newToken) {
                                        $script:SyncHash.FlowApiToken = $newToken
                                        $script:SyncHash.FlowApiTokenExpiry = (Get-Date).AddMinutes(60)
                                        Update-FlowApiTokenStatus
                                        [System.Windows.MessageBox]::Show("Authentication successful! Click 'Load Run History' to try again.", "Success", "OK", "Information")
                                    }
                                }
                                return  # Exit early since we've already updated the display
                            } else {
                                [void]$output.AppendLine("Error calling Power Automate API: $errorMsg")
                                [void]$output.AppendLine("Status Code: $statusCode")
                            }
                        }
                        
                        [void]$output.AppendLine("")
                        [void]$output.AppendLine("")
                    } elseif ($selectedItem.Type -eq "Cloud Flow" -and -not $script:SyncHash.FlowApiToken) {
                        [void]$output.AppendLine(">>> POWER AUTOMATE API <<<")
                        [void]$output.AppendLine("-" * 50)
                        [void]$output.AppendLine("Authentication was not completed or was cancelled.")
                        [void]$output.AppendLine("Click 'Load Run History' again to authenticate.")
                        [void]$output.AppendLine("")
                        [void]$output.AppendLine("")
                    }
                    
                    # FALLBACK: Query Dataverse tables (flowsession, workflowlog, asyncoperation)
                    [void]$output.AppendLine(">>> DATAVERSE TABLES (Fallback) <<<")
                    [void]$output.AppendLine("-" * 50)
                    [void]$output.AppendLine("Note: Cloud Flow run history is primarily stored in Power Automate,")
                    [void]$output.AppendLine("not Dataverse. These queries may return limited or no results.")
                    [void]$output.AppendLine("")
                    
                    # Try flowsession entity (Cloud Flows)
                    [void]$output.AppendLine("FLOWSESSION TABLE:")
                    try {
                        $flowSessionFetch = @"
<fetch>
  <entity name='flowsession'>
    <all-attributes />
    <order attribute='createdon' descending='true' />
    <filter>
      <condition attribute='regardingobjectid' operator='eq' value='$flowId' />
      <condition attribute='createdon' operator='on-or-after' value='$thirtyDaysAgoStr' />
    </filter>
  </entity>
</fetch>
"@
                        $sessionResult = Get-CrmRecordsByFetch -conn $conn -Fetch $flowSessionFetch -AllRows -ErrorAction Stop
                        
                        if ($sessionResult -and $sessionResult.CrmRecords -and $sessionResult.CrmRecords.Count -gt 0) {
                            $sessions = $sessionResult.CrmRecords
                            [void]$output.AppendLine("Found $($sessions.Count) flow session(s)")
                            [void]$output.AppendLine("")
                            
                            foreach ($session in $sessions | Select-Object -First 20) {
                                [void]$output.AppendLine("Session ID: $($session.flowsessionid)")
                                [void]$output.AppendLine("  Created: $($session.createdon)")
                                [void]$output.AppendLine("  Status: $($session.statuscode)")
                                if ($session.completedon) { [void]$output.AppendLine("  Completed: $($session.completedon)") }
                                if ($session.startedon) { [void]$output.AppendLine("  Started: $($session.startedon)") }
                                if ($session.errormessage) { 
                                    $errMsg = $session.errormessage
                                    if ($errMsg.Length -gt 200) { $errMsg = $errMsg.Substring(0, 200) + "..." }
                                    [void]$output.AppendLine("  Error: $errMsg") 
                                }
                                
                                # Show all properties for debugging
                                [void]$output.AppendLine("  All Properties:")
                                foreach ($prop in $session.PSObject.Properties | Where-Object { $_.Name -notlike "*_Property" -and $_.Name -ne "original" } | Sort-Object Name) {
                                    if ($prop.Value -ne $null -and $prop.Value.ToString() -ne "") {
                                        $val = $prop.Value.ToString()
                                        if ($val.Length -gt 100) { $val = $val.Substring(0, 100) + "..." }
                                        [void]$output.AppendLine("    $($prop.Name): $val")
                                    }
                                }
                                [void]$output.AppendLine("")
                            }
                            
                            if ($sessions.Count -gt 20) {
                                [void]$output.AppendLine("... and $($sessions.Count - 20) more sessions")
                            }
                        } else {
                            [void]$output.AppendLine("No flowsession records found for this flow.")
                            [void]$output.AppendLine("Note: Cloud Flow run history is primarily stored in Power Platform and")
                            [void]$output.AppendLine("may not be fully accessible via Dataverse queries.")
                        }
                    } catch {
                        [void]$output.AppendLine("Error querying flowsession: $($_.Exception.Message)")
                    }
                    
                    [void]$output.AppendLine("")
                    [void]$output.AppendLine("")
                    
                    # Try workflowlog entity (Classic Workflows)
                    [void]$output.AppendLine("WORKFLOWLOG TABLE:")
                    try {
                        $workflowLogFetch = @"
<fetch>
  <entity name='workflowlog'>
    <all-attributes />
    <order attribute='createdon' descending='true' />
    <filter>
      <condition attribute='createdon' operator='on-or-after' value='$thirtyDaysAgoStr' />
    </filter>
    <link-entity name='asyncoperation' from='asyncoperationid' to='asyncoperationid' alias='async'>
      <filter>
        <condition attribute='workflowactivationid' operator='eq' value='$flowId' />
      </filter>
    </link-entity>
  </entity>
</fetch>
"@
                        $logResult = Get-CrmRecordsByFetch -conn $conn -Fetch $workflowLogFetch -AllRows -ErrorAction Stop
                        
                        if ($logResult -and $logResult.CrmRecords -and $logResult.CrmRecords.Count -gt 0) {
                            $logs = $logResult.CrmRecords
                            [void]$output.AppendLine("Found $($logs.Count) workflow log(s)")
                            [void]$output.AppendLine("")
                            
                            foreach ($log in $logs | Select-Object -First 20) {
                                [void]$output.AppendLine("Log ID: $($log.workflowlogid)")
                                [void]$output.AppendLine("  Created: $($log.createdon)")
                                [void]$output.AppendLine("  Status: $($log.status)")
                                if ($log.stagename) { [void]$output.AppendLine("  Stage: $($log.stagename)") }
                                if ($log.message) { 
                                    $msg = $log.message
                                    if ($msg.Length -gt 200) { $msg = $msg.Substring(0, 200) + "..." }
                                    [void]$output.AppendLine("  Message: $msg") 
                                }
                                [void]$output.AppendLine("")
                            }
                        } else {
                            [void]$output.AppendLine("No workflowlog records found for this flow.")
                            [void]$output.AppendLine("This is expected for Cloud Flows (Modern Flows).")
                        }
                    } catch {
                        [void]$output.AppendLine("Error querying workflowlog: $($_.Exception.Message)")
                    }
                    
                    [void]$output.AppendLine("")
                    [void]$output.AppendLine("")
                    
                    # Try asyncoperation entity directly
                    [void]$output.AppendLine("")
                    [void]$output.AppendLine("ASYNCOPERATION TABLE:")
                    try {
                        $asyncFetch = @"
<fetch>
  <entity name='asyncoperation'>
    <attribute name='asyncoperationid' />
    <attribute name='name' />
    <attribute name='createdon' />
    <attribute name='completedon' />
    <attribute name='startedon' />
    <attribute name='statuscode' />
    <attribute name='statecode' />
    <attribute name='message' />
    <attribute name='friendlymessage' />
    <attribute name='operationtype' />
    <order attribute='createdon' descending='true' />
    <filter>
      <condition attribute='workflowactivationid' operator='eq' value='$flowId' />
      <condition attribute='createdon' operator='on-or-after' value='$thirtyDaysAgoStr' />
    </filter>
  </entity>
</fetch>
"@
                        $asyncResult = Get-CrmRecordsByFetch -conn $conn -Fetch $asyncFetch -AllRows -ErrorAction Stop
                        
                        if ($asyncResult -and $asyncResult.CrmRecords -and $asyncResult.CrmRecords.Count -gt 0) {
                            $asyncOps = $asyncResult.CrmRecords
                            [void]$output.AppendLine("Found $($asyncOps.Count) async operation(s)")
                            [void]$output.AppendLine("")
                            
                            foreach ($async in $asyncOps | Select-Object -First 20) {
                                [void]$output.AppendLine("Async Op ID: $($async.asyncoperationid)")
                                [void]$output.AppendLine("  Name: $($async.name)")
                                [void]$output.AppendLine("  Created: $($async.createdon)")
                                [void]$output.AppendLine("  Status: $($async.statuscode) (State: $($async.statecode))")
                                if ($async.startedon) { [void]$output.AppendLine("  Started: $($async.startedon)") }
                                if ($async.completedon) { [void]$output.AppendLine("  Completed: $($async.completedon)") }
                                if ($async.message) { 
                                    $msg = $async.message
                                    if ($msg.Length -gt 200) { $msg = $msg.Substring(0, 200) + "..." }
                                    [void]$output.AppendLine("  Message: $msg") 
                                }
                                [void]$output.AppendLine("")
                            }
                        } else {
                            [void]$output.AppendLine("No asyncoperation records found for this flow.")
                        }
                    } catch {
                        [void]$output.AppendLine("Error querying asyncoperation: $($_.Exception.Message)")
                    }
                    
                    $txtDebugInfo.Text = $output.ToString()
                } catch {
                    $txtDebugInfo.Text = "Error loading run history: $($_.Exception.Message)"
                }
            })

            # Copy URL button handler
            $btnCopyUrl.Add_Click({
                if ($selectedItem.Url -and $selectedItem.Url -ne "") {
                    [System.Windows.Clipboard]::SetText($selectedItem.Url)
                    [System.Windows.MessageBox]::Show("URL copied to clipboard!", "Success", "OK", "Information")
                }
            })

            # Close button handler
            $btnClose.Add_Click({
                $detailsWindow.Close()
            })

            # Set owner and show dialog
            $detailsWindow.Owner = $window
            $detailsWindow.ShowDialog() | Out-Null

        } catch {
            [System.Windows.MessageBox]::Show("Error showing details: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    }
})

# Manage Credentials button click handler
$btnManageCredentials.Add_Click({
    Show-ManageCredentialsDialog
})

# Get Flow History button click handler
$btnFlowHistory.Add_Click({
    $message = @"
To view Cloud Flow run history, you need to authenticate interactively with your user account.

This is in addition to the app registration authentication and requires:
• Interactive browser login
• MFA if enabled on your account

Note: Run history is only available via Power Automate Management API, which requires user delegation (not available with app-only authentication).

Would you like to authenticate now?
"@

    $result = [System.Windows.MessageBox]::Show($message, "Flow Run History Authentication", "YesNo", "Question")

    if ($result -eq "Yes") {
        $authResult = Show-PowerAutomateApiSetup
        if ($authResult) {
            # Successfully authenticated - now update flow run history for all cloud flows
            $txtStatus.Text = "Updating flow run history..."
            $progressBar.IsIndeterminate = $true

            try {
                $cloudFlows = $script:SyncHash.Results | Where-Object { $_.Type -eq "Cloud Flow" }
                $conn = $script:SyncHash.Connection
                $updated = 0

                foreach ($flow in $cloudFlows) {
                    # Use WorkflowIdUnique for the API call (like FlowExecutionHistory tool)
                    $flowIdForApi = if ($flow.WorkflowIdUnique) { $flow.WorkflowIdUnique } else { $flow.Id }
                    $runHistory = Get-FlowRunHistory -FlowId $flowIdForApi -Connection $conn
                    $flow.LastRun = $runHistory.LastRun
                    $flow.RunCount30d = $runHistory.RunCount
                    $updated++
                }

                # Refresh the grid
                $dgResults.ItemsSource = $null
                $dgResults.ItemsSource = $script:SyncHash.Results

                $txtStatus.Text = "Flow run history updated for $updated flows"
                [System.Windows.MessageBox]::Show("Successfully updated run history for $updated cloud flows!", "Success", "OK", "Information")
            } catch {
                $txtStatus.Text = "Error updating flow history"
                [System.Windows.MessageBox]::Show("Error updating flow run history: $($_.Exception.Message)", "Error", "OK", "Error")
            } finally {
                $progressBar.IsIndeterminate = $false
            }
        }
    }
})

# Sign Out button click handler
$btnSignOut.Add_Click({
    try {
        # Delete cached token file
        $tokenPath = "$env:USERPROFILE\.powersearch\flowapi_token.json"
        if (Test-Path $tokenPath) {
            Remove-Item $tokenPath -Force
            Write-Host "Cached Flow API token deleted" -ForegroundColor Cyan
        }

        # Clear in-memory tokens
        $script:SyncHash.FlowApiToken = $null
        $script:SyncHash.FlowApiTokenExpiry = $null

        # Update UI status
        Update-FlowApiTokenStatus

        # Open Microsoft sign-out page
        # Note: This only clears the session in the browser, not the device code flow cache
        Start-Process "https://login.microsoftonline.com/common/oauth2/v2.0/logout"

        Write-Host "Sign out complete - all local tokens cleared" -ForegroundColor Green

        $message = @"
You have been signed out successfully.

What was cleared:
• Cached Flow API token file
• In-memory authentication tokens
• Browser opened to Microsoft sign-out page

Note: To fully clear your authentication, you may need to:
1. Close all browser windows after sign-out completes
2. Or use InPrivate/Incognito mode for the next authentication

Click OK to continue.
"@

        [System.Windows.MessageBox]::Show($message, "Signed Out", "OK", "Information")
    } catch {
        Write-Host "Error during sign out: $($_.Exception.Message)" -ForegroundColor Red
        [System.Windows.MessageBox]::Show("Error during sign out: $($_.Exception.Message)", "Error", "OK", "Error")
    }
})

# AI Settings button click handler
$btnAISettings.Add_Click({
    Show-ApiKeySettings
})

# AI Summary button click handler
$btnAISummary.Add_Click({
    $selectedItem = $dgResults.SelectedItem
    if ($selectedItem) {
        Start-AISummary -Item $selectedItem
    } else {
        [System.Windows.MessageBox]::Show("Please select an item from the results grid first.", "No Selection", "OK", "Warning")
    }
})

# Connect button click handler - runs synchronously in main thread
$btnConnect.Add_Click({
    $txtStatus.Text = "Connecting to Dataverse..."
    $progressBar.IsIndeterminate = $true
    $btnConnect.IsEnabled = $false

    # Force UI update
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $envUrl = $txtEnvironmentUrl.Text

        if (-not $envUrl.StartsWith("http")) {
            $envUrl = "https://$envUrl"
        }

        $connectionString = "AuthType=OAuth;Url=$envUrl;AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;RedirectUri=http://localhost;LoginPrompt=Auto"
        $conn = Get-CrmConnection -ConnectionString $connectionString

        if ($conn.IsReady) {
            $script:SyncHash.Connection = $conn
            $orgName = $conn.ConnectedOrgFriendlyName

            # Extract environment ID from the connection URL for building links
            $orgUrl = $conn.CrmConnectOrgUriActual.Host
            $script:SyncHash.OrgUrl = $orgUrl
            # Try to get the environment ID (used for Power Automate URLs)
            try {
                $envIdMatch = $orgUrl -replace '\.crm\d*\.dynamics\.com$', ''
                $script:SyncHash.EnvironmentId = $envIdMatch
            } catch {
                $script:SyncHash.EnvironmentId = $null
            }

            $txtConnectionStatus.Text = "Connected: $orgName"
            $statusIndicator.Fill = [System.Windows.Media.Brushes]::LimeGreen
            $txtStatus.Text = "Connected successfully!"
            $btnConnect.Content = "Reconnect"
            $btnScan.IsEnabled = $true
        } else {
            throw "Connection failed - not ready"
        }
    } catch {
        $errorMsg = $_.Exception.Message
        $txtConnectionStatus.Text = "Connection Failed"
        $statusIndicator.Fill = [System.Windows.Media.Brushes]::Red
        $txtStatus.Text = "Error: $errorMsg"
        [System.Windows.MessageBox]::Show("Connection failed: $errorMsg", "Connection Error", "OK", "Error")
    } finally {
        $progressBar.IsIndeterminate = $false
        $btnConnect.IsEnabled = $true
    }
})

# Debug function to fetch ALL properties of a flow/workflow from Dataverse
function Get-FlowDebugInfo {
    param(
        [string]$FlowId,
        $Connection
    )

    try {
        # Fetch ALL attributes from the workflow entity
        $fetch = @"
<fetch>
  <entity name='workflow'>
    <all-attributes />
    <filter>
      <condition attribute='workflowid' operator='eq' value='$FlowId' />
    </filter>
  </entity>
</fetch>
"@
        $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetch -AllRows -ErrorAction Stop
        
        if ($result -and $result.CrmRecords -and $result.CrmRecords.Count -gt 0) {
            $flow = $result.CrmRecords[0]
            $debugInfo = [System.Text.StringBuilder]::new()
            
            [void]$debugInfo.AppendLine("=" * 80)
            [void]$debugInfo.AppendLine("FLOW DEBUG INFORMATION - ALL PROPERTIES FROM DATAVERSE")
            [void]$debugInfo.AppendLine("=" * 80)
            [void]$debugInfo.AppendLine("")
            
            # Key state properties first
            [void]$debugInfo.AppendLine(">>> STATE PROPERTIES (KEY FOR ON/OFF DETECTION) <<<")
            [void]$debugInfo.AppendLine("-" * 50)
            [void]$debugInfo.AppendLine("NOTE: CRM PowerShell returns formatted strings, not numeric values")
            [void]$debugInfo.AppendLine("")
            
            # statecode - determines if flow is On/Off
            $stateCodeValue = $flow.statecode
            $isActive = ($stateCodeValue -eq "Activated")
            [void]$debugInfo.AppendLine("statecode: '$stateCodeValue' [Type: $($stateCodeValue.GetType().Name)]")
            [void]$debugInfo.AppendLine("  -> Valid values: 'Draft' (Off), 'Activated' (On), 'Suspended' (Off)")
            [void]$debugInfo.AppendLine("  -> IS ACTIVE: $isActive")
            [void]$debugInfo.AppendLine("")
            
            # statuscode
            $statusCodeValue = $flow.statuscode
            [void]$debugInfo.AppendLine("statuscode: '$statusCodeValue' [Type: $($statusCodeValue.GetType().Name)]")
            [void]$debugInfo.AppendLine("  -> Valid values: 'Draft', 'Activated'")
            [void]$debugInfo.AppendLine("")
            
            # category
            $categoryValue = $flow.category
            [void]$debugInfo.AppendLine("category: '$categoryValue' [Type: $($categoryValue.GetType().Name)]")
            [void]$debugInfo.AppendLine("  -> Valid values: 'Workflow', 'Dialog', 'Business Rule', 'Action', 'Business Process Flow', 'Modern Flow'")
            [void]$debugInfo.AppendLine("")
            
            # type
            $typeValue = $flow.type
            [void]$debugInfo.AppendLine("type: '$typeValue' [Type: $($typeValue.GetType().Name)]")
            [void]$debugInfo.AppendLine("  -> Valid values: 'Definition', 'Activation', 'Template'")
            
            [void]$debugInfo.AppendLine("")
            [void]$debugInfo.AppendLine(">>> ALL OTHER PROPERTIES <<<")
            [void]$debugInfo.AppendLine("-" * 50)
            
            # Get all properties and sort them
            $properties = $flow.PSObject.Properties | Where-Object { 
                $_.Name -notin @('statecode', 'statuscode', 'category', 'type', 'clientdata', 'xaml', 'inputparameters', 'triggerdata') 
            } | Sort-Object Name
            
            foreach ($prop in $properties) {
                $propValue = $prop.Value
                $propType = if ($propValue -ne $null) { $propValue.GetType().Name } else { "null" }
                
                # Handle OptionSetValue
                if ($propValue -is [Microsoft.Xrm.Sdk.OptionSetValue]) {
                    [void]$debugInfo.AppendLine("$($prop.Name): $($propValue.Value) [OptionSetValue]")
                }
                # Handle EntityReference
                elseif ($propValue -is [Microsoft.Xrm.Sdk.EntityReference]) {
                    [void]$debugInfo.AppendLine("$($prop.Name): $($propValue.Name) (Id: $($propValue.Id)) [EntityReference]")
                }
                # Handle AliasedValue
                elseif ($propValue -is [Microsoft.Xrm.Sdk.AliasedValue]) {
                    [void]$debugInfo.AppendLine("$($prop.Name): $($propValue.Value) [AliasedValue]")
                }
                # Handle Money
                elseif ($propValue -is [Microsoft.Xrm.Sdk.Money]) {
                    [void]$debugInfo.AppendLine("$($prop.Name): $($propValue.Value) [Money]")
                }
                # Handle DateTime
                elseif ($propValue -is [DateTime]) {
                    [void]$debugInfo.AppendLine("$($prop.Name): $($propValue.ToString('yyyy-MM-dd HH:mm:ss')) [DateTime]")
                }
                # Handle Guid
                elseif ($propValue -is [Guid]) {
                    [void]$debugInfo.AppendLine("$($prop.Name): $propValue [Guid]")
                }
                # Handle Boolean
                elseif ($propValue -is [Boolean]) {
                    [void]$debugInfo.AppendLine("$($prop.Name): $propValue [Boolean]")
                }
                # Handle strings (truncate if too long)
                elseif ($propValue -is [String]) {
                    $displayValue = if ($propValue.Length -gt 200) { $propValue.Substring(0, 200) + "... [TRUNCATED]" } else { $propValue }
                    [void]$debugInfo.AppendLine("$($prop.Name): $displayValue [String]")
                }
                # Handle other types
                else {
                    $displayValue = if ($propValue -ne $null) { $propValue.ToString() } else { "(null)" }
                    if ($displayValue.Length -gt 200) { $displayValue = $displayValue.Substring(0, 200) + "... [TRUNCATED]" }
                    [void]$debugInfo.AppendLine("$($prop.Name): $displayValue [$propType]")
                }
            }
            
            [void]$debugInfo.AppendLine("")
            [void]$debugInfo.AppendLine(">>> LARGE TEXT FIELDS (TRUNCATED) <<<")
            [void]$debugInfo.AppendLine("-" * 50)
            
            # Show truncated versions of large fields
            if ($flow.clientdata) {
                $clientDataLen = $flow.clientdata.Length
                [void]$debugInfo.AppendLine("clientdata: [$clientDataLen chars] - Contains flow definition JSON")
            }
            if ($flow.xaml) {
                $xamlLen = $flow.xaml.Length
                [void]$debugInfo.AppendLine("xaml: [$xamlLen chars] - Contains workflow XAML definition")
            }
            if ($flow.inputparameters) {
                $inputLen = $flow.inputparameters.Length
                [void]$debugInfo.AppendLine("inputparameters: [$inputLen chars]")
            }
            
            [void]$debugInfo.AppendLine("")
            [void]$debugInfo.AppendLine(">>> USE BUTTONS ABOVE TO LOAD MORE DATA <<<")
            [void]$debugInfo.AppendLine("-" * 50)
            [void]$debugInfo.AppendLine("- 'Load Flow Definition' - Shows full clientdata JSON")
            [void]$debugInfo.AppendLine("- 'Load Run History' - Queries flowsession, workflowlog, and asyncoperation tables")
            
            return $debugInfo.ToString()
        } else {
            return "No flow found with ID: $FlowId"
        }
    } catch {
        return "Error fetching flow debug info: $($_.Exception.Message)`n`nStack Trace:`n$($_.ScriptStackTrace)"
    }
}

# Scan function - runs synchronously but with UI updates
function Start-DataverseScan {
    param(
        $Connection,
        [string[]]$Keywords,
        [datetime]$BreakageDate,
        [bool]$ScanCloudFlows,
        [bool]$ScanClassicWorkflows,
        [bool]$ScanPlugins,
        [bool]$ScanWebResources,
        [bool]$ScanCanvasApps,
        [bool]$ScanModelDrivenApps,
        [bool]$ScanCustomActions,
        [bool]$ScanServiceEndpoints,
        [bool]$ScanSLAs,
        [bool]$ScanBPFs,
        [bool]$ScanTableSchemas,
        [bool]$ScanConnections,
        [bool]$ScanEnvVariables,
        [string]$OrgUrl
    )

    $conn = $Connection
    $discoveredItems = [System.Collections.ArrayList]::new()
    $totalSteps = 13
    $currentStep = 0

    # Debug: Check if connection is valid
    if (-not $conn -or -not $conn.IsReady) {
        $txtStatus.Text = "Error: Connection is not ready"
        return $discoveredItems
    }

    # Get environment ID for URL generation
    $environmentId = $null
    if ($conn -and $conn.OrganizationId) {
        $environmentId = $conn.OrganizationId
    }

    # Debug: Show keywords being used
    $keywordList = $Keywords -join ", "
    $txtStatus.Text = "Scanning with keywords: $keywordList"
    [System.Windows.Forms.Application]::DoEvents()

    # Helper function to get flow run history from Power Automate Management API
    # Function to get flow run history from Dataverse (works with app registration)
    # NOTE: Cloud Flows (Modern Flows) don't store run history in Dataverse
    # This returns "N/A" since run history requires Power Automate Management API with user delegation
    function Get-FlowRunHistoryFromDataverse {
        param(
            [string]$FlowId,
            $Connection
        )

        # Cloud flows (category=5) don't store run history in Dataverse
        # Run history is only available via Power Automate Management API which requires user delegation
        # With app registration (client credentials), we can't access this data
        return @{
            LastRun = "N/A"
            RunCount = "N/A"
        }
    }

    function Get-FlowRunHistory {
        param(
            [string]$FlowId,
            $Connection
        )

        try {
            # Check if we have a Flow API token
            if (-not $script:SyncHash.FlowApiToken) {
                # No token available - return "Click to setup"
                return @{
                    LastRun = "Setup required"
                    RunCount = "Setup required"
                }
            }

            # Get environment ID from connection - try multiple properties
            $environmentId = $null
            
            # Method 1: Try ConnectedOrgId (CRM Tooling standard property)
            if ($Connection -and $Connection.ConnectedOrgId) {
                $environmentId = $Connection.ConnectedOrgId.ToString()
                Write-Host "Got environment ID from ConnectedOrgId: $environmentId" -ForegroundColor Cyan
            }
            # Method 2: Try OrganizationId property
            elseif ($Connection -and $Connection.OrganizationId) {
                $environmentId = $Connection.OrganizationId.ToString()
                Write-Host "Got environment ID from OrganizationId: $environmentId" -ForegroundColor Cyan
            }
            # Method 3: Fallback - query organization entity
            else {
                try {
                    $orgFetch = @"
<fetch>
  <entity name='organization'>
    <attribute name='organizationid' />
  </entity>
</fetch>
"@
                    $orgResult = Get-CrmRecordsByFetch -conn $Connection -Fetch $orgFetch -AllRows -ErrorAction SilentlyContinue

                    if ($orgResult -and $orgResult.CrmRecords -and $orgResult.CrmRecords.Count -gt 0) {
                        $environmentId = $orgResult.CrmRecords[0].organizationid.ToString()
                        Write-Host "Got environment ID from Dataverse query: $environmentId" -ForegroundColor Cyan
                    }
                } catch {
                    Write-Host "Failed to query organization: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            if ($environmentId) {

                # Call Power Automate Management API with proper token
                $flowApiUrl = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$environmentId/flows/$FlowId/runs?api-version=2016-11-01&`$top=50"
                # Write-Host "Calling Flow API: $flowApiUrl" -ForegroundColor Cyan

                $result = Invoke-RestMethod -Uri $flowApiUrl -Method Get -Headers @{
                    "Authorization" = "Bearer " + $script:SyncHash.FlowApiToken
                    "Accept" = "application/json"
                } -ErrorAction Stop

                # Write-Host "API call successful, got $($result.value.Count) runs" -ForegroundColor Green

                if ($result.value -and $result.value.Count -gt 0) {
                    # Count runs in last 30 days
                    $thirtyDaysAgo = (Get-Date).AddDays(-30)
                    $recentRuns = $result.value | Where-Object {
                        $_.properties.startTime -and ([DateTime]$_.properties.startTime) -gt $thirtyDaysAgo
                    }

                    $runCount = $recentRuns.Count
                    $lastRun = ""

                    if ($result.value[0].properties.startTime) {
                        $lastRunDate = [DateTime]::Parse($result.value[0].properties.startTime)
                        $lastRun = $lastRunDate.ToString("yyyy-MM-dd HH:mm")
                    }

                    return @{
                        LastRun = if ($lastRun) { $lastRun } else { "Unknown" }
                        RunCount = $runCount
                    }
                } else {
                    return @{
                        LastRun = "No runs"
                        RunCount = "0"
                    }
                }
            }

            # Couldn't get environment ID
            return @{
                LastRun = "Error"
                RunCount = "Error"
            }

        } catch {
            # API call failed - token might be expired or insufficient permissions
            $errorMsg = $_.Exception.Message
            $statusCode = $null

            # Try to extract status code from error
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            # Only log non-403 errors (403 is expected with app registration auth)
            if ($statusCode -ne 403) {
                Write-Host "Error getting flow run history for $FlowId : $errorMsg (Status: $statusCode)" -ForegroundColor Red
                Write-Host "Stack: $($_.ScriptStackTrace)" -ForegroundColor Gray
            }

            if ($statusCode -eq 401) {
                # Unauthorized - token expired
                Write-Host "Authentication issue detected - token may be expired" -ForegroundColor Yellow
                return @{
                    LastRun = "Auth expired"
                    RunCount = "Auth expired"
                }
            } elseif ($statusCode -eq 403 -or $errorMsg -match "Forbidden") {
                # Forbidden - app registration doesn't have delegated permissions for Flow API
                # This is expected when using client credentials flow (app-only auth)
                return @{
                    LastRun = "No access"
                    RunCount = "No access"
                }
            }

            return @{
                LastRun = "Error"
                RunCount = "Error"
            }
        }
    }

    # Helper function to generate URL for an item
    function Get-ItemUrl {
        param(
            [string]$Type,
            [string]$Id,
            [string]$OrgUrl,
            [string]$EnvironmentId = $null
        )

        if (-not $Id -or -not $OrgUrl) { return "" }

        $baseUrl = "https://$OrgUrl"

        switch -Wildcard ($Type) {
            "Cloud Flow" {
                # Power Automate flow URL
                if ($EnvironmentId) {
                    return "https://make.powerautomate.com/environments/$EnvironmentId/flows/$Id/details"
                } else {
                    return "https://make.powerautomate.com/environments/Default-*/flows/$Id/details"
                }
            }
            "Classic Workflow" {
                return "$baseUrl/main.aspx?appid=&pagetype=webresource&webresourceName=msdyn_/WorkflowEditor/WorkflowEditor.html&Data=%7B%22EntityId%22%3A%22$Id%22%7D"
            }
            "Custom Action" {
                return "$baseUrl/main.aspx?appid=&pagetype=webresource&webresourceName=msdyn_/WorkflowEditor/WorkflowEditor.html&Data=%7B%22EntityId%22%3A%22$Id%22%7D"
            }
            "Plugin" {
                return "$baseUrl/main.aspx?pagetype=entityrecord&etn=sdkmessageprocessingstep&id=$Id"
            }
            "Web Resource*" {
                return "$baseUrl/main.aspx?pagetype=entityrecord&etn=webresource&id=$Id"
            }
            "Service Endpoint*" {
                return "$baseUrl/main.aspx?pagetype=entityrecord&etn=sdkmessageprocessingstep&id=$Id"
            }
            "SLA" {
                return "$baseUrl/main.aspx?pagetype=entityrecord&etn=sla&id=$Id"
            }
            "Business Process Flow" {
                return "$baseUrl/main.aspx?pagetype=entityrecord&etn=workflow&id=$Id"
            }
            "Canvas App" {
                if ($EnvironmentId) {
                    return "https://make.powerapps.com/environments/$EnvironmentId/apps/$Id/details"
                } else {
                    return "https://make.powerapps.com/environments/Default-*/apps/$Id/details"
                }
            }
            "Model-Driven App" {
                return "$baseUrl/main.aspx?appid=$Id"
            }
            default {
                return ""
            }
        }
    }

    # Helper function to get solution name for an item
    function Get-ItemSolution {
        param(
            [string]$ItemId,
            [string]$ItemType,
            $Connection
        )

        # Return early if no ItemId
        if ([string]::IsNullOrEmpty($ItemId)) {
            return "Unknown"
        }

        try {
            # Map item type to component type (Dataverse component type codes)
            $componentType = switch -Wildcard ($ItemType) {
                "Table Schema" { 1 }
                "Cloud Flow" { 29 }
                "Classic Workflow" { 29 }
                "Custom Action" { 29 }
                "Business Process Flow" { 29 }
                "SLA" { 44 }
                "Web Resource*" { 61 }
                "Plugin" { 90 }
                "Service Endpoint*" { 95 }
                "Canvas App" { 300 }
                default { 
                    return "Unknown"
                }
            }

            # Clean up GUID if needed (remove braces)
            $cleanId = $ItemId -replace '[{}]', ''

            $fetch = @"
<fetch top='1'>
  <entity name='solutioncomponent'>
    <attribute name='solutionid' />
    <filter>
      <condition attribute='objectid' operator='eq' value='$cleanId' />
      <condition attribute='componenttype' operator='eq' value='$componentType' />
    </filter>
    <link-entity name='solution' from='solutionid' to='solutionid' alias='sol'>
      <attribute name='friendlyname' />
      <attribute name='ismanaged' />
      <filter>
        <condition attribute='isvisible' operator='eq' value='1' />
      </filter>
      <order attribute='ismanaged' descending='false' />
    </link-entity>
  </entity>
</fetch>
"@

            $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetch -ErrorAction Stop

            if ($result -and $result.CrmRecords -and $result.CrmRecords.Count -gt 0) {
                $solutionName = $result.CrmRecords[0].'sol.friendlyname'
                $isManaged = $result.CrmRecords[0].'sol.ismanaged'
                if ($isManaged) {
                    return "$solutionName (Managed)"
                } else {
                    return $solutionName
                }
            }

            return "Default Solution"
        } catch {
            # Silently return Unknown on error (component may not be in a solution)
            return "Unknown"
        }
    }

    # Helper function to get error count for an item
    function Get-ItemErrors {
        param(
            [string]$ItemId,
            [string]$ItemType,
            $Connection
        )

        try {
            $thirtyDaysAgo = (Get-Date).AddDays(-30).ToString("yyyy-MM-dd")
            $errorCount = 0

            if ($ItemType -like "*Flow*" -or $ItemType -like "*Workflow*" -or $ItemType -eq "Custom Action" -or $ItemType -eq "Business Process Flow") {
                # Query workflow logs for errors
                $fetch = @"
<fetch aggregate='true'>
  <entity name='workflowlog'>
    <attribute name='workflowlogid' alias='count' aggregate='count' />
    <filter>
      <condition attribute='regardingobjectid' operator='eq' value='$ItemId' />
      <condition attribute='status' operator='eq' value='4' />
      <condition attribute='createdon' operator='on-or-after' value='$thirtyDaysAgo' />
    </filter>
  </entity>
</fetch>
"@
                $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetch -ErrorAction SilentlyContinue
                if ($result -and $result.CrmRecords -and $result.CrmRecords.Count -gt 0) {
                    $errorCount = [int]$result.CrmRecords[0].count
                }
            } elseif ($ItemType -eq "Plugin") {
                # Query plugin trace logs for errors
                # Note: This is less reliable as plugin trace logs may not be enabled
                $fetch = @"
<fetch aggregate='true'>
  <entity name='plugintracelog'>
    <attribute name='plugintracelogid' alias='count' aggregate='count' />
    <filter>
      <condition attribute='pluginstepid' operator='eq' value='$ItemId' />
      <condition attribute='exceptiondetails' operator='not-null' />
      <condition attribute='createdon' operator='on-or-after' value='$thirtyDaysAgo' />
    </filter>
  </entity>
</fetch>
"@
                $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetch -ErrorAction SilentlyContinue
                if ($result -and $result.CrmRecords -and $result.CrmRecords.Count -gt 0) {
                    $errorCount = [int]$result.CrmRecords[0].count
                }
            }

            return @{
                Count = $errorCount
                HasErrors = $errorCount -gt 0
            }
        } catch {
            return @{
                Count = 0
                HasErrors = $false
            }
        }
    }

    # Helper function to get recent error details for an item
    function Get-RecentErrorDetails {
        param(
            [string]$ItemId,
            [string]$ItemType,
            $Connection,
            [int]$MaxErrors = 10
        )

        try {
            $thirtyDaysAgo = (Get-Date).AddDays(-30).ToString("yyyy-MM-dd")
            $errors = @()

            if ($ItemType -like "*Flow*" -or $ItemType -like "*Workflow*" -or $ItemType -eq "Custom Action" -or $ItemType -eq "Business Process Flow") {
                # Query workflow logs for recent errors
                $fetch = @"
<fetch top='$MaxErrors'>
  <entity name='workflowlog'>
    <attribute name='createdon' />
    <attribute name='description' />
    <attribute name='message' />
    <attribute name='activityname' />
    <order attribute='createdon' descending='true' />
    <filter>
      <condition attribute='regardingobjectid' operator='eq' value='$ItemId' />
      <condition attribute='status' operator='eq' value='4' />
      <condition attribute='createdon' operator='on-or-after' value='$thirtyDaysAgo' />
    </filter>
  </entity>
</fetch>
"@
                $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetch -ErrorAction SilentlyContinue
                if ($result -and $result.CrmRecords) {
                    foreach ($log in $result.CrmRecords) {
                        $errors += @{
                            Timestamp = if ($log.createdon) { $log.createdon } else { "Unknown" }
                            Step = if ($log.activityname) { $log.activityname } else { "N/A" }
                            Message = if ($log.message) { $log.message } else { if ($log.description) { $log.description } else { "No error message" } }
                        }
                    }
                }
            } elseif ($ItemType -eq "Plugin") {
                # Query plugin trace logs for recent errors
                $fetch = @"
<fetch top='$MaxErrors'>
  <entity name='plugintracelog'>
    <attribute name='createdon' />
    <attribute name='exceptiondetails' />
    <attribute name='messagename' />
    <order attribute='createdon' descending='true' />
    <filter>
      <condition attribute='pluginstepid' operator='eq' value='$ItemId' />
      <condition attribute='exceptiondetails' operator='not-null' />
      <condition attribute='createdon' operator='on-or-after' value='$thirtyDaysAgo' />
    </filter>
  </entity>
</fetch>
"@
                $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetch -ErrorAction SilentlyContinue
                if ($result -and $result.CrmRecords) {
                    foreach ($log in $result.CrmRecords) {
                        $errors += @{
                            Timestamp = if ($log.createdon) { $log.createdon } else { "Unknown" }
                            Step = if ($log.messagename) { $log.messagename } else { "N/A" }
                            Message = if ($log.exceptiondetails) { $log.exceptiondetails } else { "No error details" }
                        }
                    }
                }
            }

            return $errors
        } catch {
            return @()
        }
    }

    # Helper to update progress
    $updateProgress = {
        param([string]$Message, [int]$Step)
        $percent = [int](($Step / $totalSteps) * 100)
        $txtStatus.Text = $Message
        $progressBar.Value = $percent
        [System.Windows.Forms.Application]::DoEvents()
    }

    # 1. Cloud Flows
    if ($ScanCloudFlows) {
        $currentStep++
        & $updateProgress "Scanning Cloud Flows..." $currentStep

        $flowsFetch = @"
<fetch>
  <entity name='workflow'>
    <attribute name='name' />
    <attribute name='workflowid' />
    <attribute name='workflowidunique' />
    <attribute name='description' />
    <attribute name='statecode' />
    <attribute name='statuscode' />
    <attribute name='category' />
    <attribute name='primaryentity' />
    <attribute name='createdon' />
    <attribute name='modifiedon' />
    <attribute name='clientdata' />
    <link-entity name='solution' from='solutionid' to='solutionid' alias='sol' link-type='outer'>
      <attribute name='friendlyname' />
      <attribute name='ismanaged' />
    </link-entity>
    <filter type='and'>
      <condition attribute='category' operator='eq' value='5' />
      <condition attribute='type' operator='eq' value='1' />
    </filter>
    <order attribute='modifiedon' descending='true' />
  </entity>
</fetch>
"@

        try {
            $flows = Get-CrmRecordsByFetch -conn $conn -Fetch $flowsFetch
            & $updateProgress "Scanning Cloud Flows... (found $($flows.CrmRecords.Count))" $currentStep

            foreach ($flow in $flows.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                $flowName = if ($flow.name) { $flow.name.ToLower() } else { "" }
                foreach ($keyword in $Keywords) {
                    # Exact name match gets highest priority
                    if ($flowName -eq $keyword) {
                        $matchScore += 15
                        $matchReasons += "[EXACT MATCH] Name exactly matches '$keyword'"
                    } elseif ($flowName -like "*$keyword*") {
                        $matchScore += 8
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                # Check if primary entity matches keywords (high value)
                if ($flow.primaryentity) {
                    $primaryEntity = $flow.primaryentity.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($primaryEntity -eq $keyword) {
                            $matchScore += 12
                            $matchReasons += "[PRIMARY ENTITY] Works with entity '$keyword'"
                        } elseif ($primaryEntity -like "*$keyword*") {
                            $matchScore += 6
                            $matchReasons += "Primary entity contains '$keyword'"
                        }
                    }
                }

                if ($flow.description) {
                    $flowDesc = $flow.description.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($flowDesc -like "*$keyword*") {
                            $matchScore += 2
                            $matchReasons += "Description contains '$keyword'"
                        }
                    }
                }

                # Track action types for better troubleshooting (low score - informational only)
                $hasSharePoint = $false
                $hasEmail = $false
                $hasFileOps = $false
                $actionDetails = @()

                if ($flow.clientdata) {
                    $clientData = $flow.clientdata.ToLower()

                    # Detect email actions (informational, low score)
                    if ($clientData -match 'send.*email|office365outlook|smtp|mail\.send|sendemailv2|sendemailnotification') {
                        $matchScore += 1
                        $hasEmail = $true
                        $matchReasons += "[EMAIL] Contains email sending action"
                        $actionDetails += "Email"
                    }

                    # Detect SharePoint actions (informational, low score)
                    if ($clientData -match 'sharepoint|sharepointonline|sp\.') {
                        $matchScore += 1
                        $hasSharePoint = $true
                        $matchReasons += "[SHAREPOINT] Contains SharePoint action"
                        $actionDetails += "SharePoint"
                        
                        # Detect specific file operations in SharePoint
                        if ($clientData -match 'createfile|create.*file|uploadfile|copyfile|copy.*file') {
                            $hasFileOps = $true
                            $matchReasons += "[FILE CREATE] Creates/uploads files in SharePoint"
                            $actionDetails += "File Create"
                        }
                        if ($clientData -match 'updatefile|update.*file|modifyfile') {
                            $hasFileOps = $true
                            $matchReasons += "[FILE MODIFY] Modifies files in SharePoint"
                            $actionDetails += "File Modify"
                        }
                        if ($clientData -match 'deletefile|delete.*file') {
                            $hasFileOps = $true
                            $matchReasons += "[FILE DELETE] Deletes files in SharePoint"
                            $actionDetails += "File Delete"
                        }
                        if ($clientData -match 'getfilecontent|get.*file.*content|downloadfile') {
                            $matchReasons += "[FILE READ] Reads files from SharePoint"
                            $actionDetails += "File Read"
                        }
                    }

                    # Detect OneDrive file operations (informational)
                    if ($clientData -match 'onedrive') {
                        $matchScore += 1
                        $matchReasons += "[ONEDRIVE] Contains OneDrive action"
                        $actionDetails += "OneDrive"
                        
                        if ($clientData -match 'createfile|uploadfile|updatefile|deletefile') {
                            $hasFileOps = $true
                            $matchReasons += "[FILE OPS] File operations in OneDrive"
                            $actionDetails += "File Ops"
                        }
                    }

                    if ($clientData -match 'msdyn_project|project') {
                        $matchScore += 1
                        $matchReasons += "References project entity"
                    }

                    if ($clientData -match 'approval|approve|reject') {
                        $matchScore += 1
                        $matchReasons += "Contains approval logic"
                    }

                    # Check if clientdata contains keyword - boost if it looks like entity/table reference
                    foreach ($keyword in $Keywords) {
                        if ($clientData -like "*$keyword*") {
                            # Higher score if keyword appears as entity reference in triggers/actions
                            if ($clientData -match "entityname.*$keyword|entity.*:.*$keyword|tablename.*$keyword") {
                                $matchScore += 5
                                $matchReasons += "References entity '$keyword' in actions"
                            } else {
                                $matchScore += 1
                            }
                        }
                    }
                }

                if ($flow.modifiedon -and $BreakageDate) {
                    try {
                        $modDate = [DateTime]::Parse($flow.modifiedon.ToString())
                        $daysSinceBreakage = ($modDate - $BreakageDate).Days
                        if ([Math]::Abs($daysSinceBreakage) -le 30) {
                            $matchScore += 2
                            $matchReasons += "Modified around breakage date ($daysSinceBreakage days)"
                        }
                    } catch { }
                }

                if ($matchScore -gt 0) {
                    # For Cloud Flows (category 5 workflows):
                    # statecode: 0 = Draft, 1 = Activated, 2 = Suspended
                    # CRM PowerShell module returns formatted strings like "Activated" or "Draft"
                    $stateValue = $flow.statecode
                    if ($stateValue -is [Microsoft.Xrm.Sdk.OptionSetValue]) {
                        $stateValue = $stateValue.Value
                        $isOn = ($stateValue -eq 1)
                    } else {
                        # Handle string value returned by formatted view
                        $isOn = ($stateValue -eq "Activated")
                    }
                    $state = if ($isOn) { "On" } else { "Off" }
                    $itemType = "Cloud Flow"
                    $itemId = $flow.workflowid

                    # Get flow run history from Dataverse (works with app registration)
                    $runHistory = Get-FlowRunHistoryFromDataverse -FlowId $itemId -Connection $conn

                    # Get error count
                    $errorInfo = Get-ItemErrors -ItemId $itemId -ItemType $itemType -Connection $conn

                    # Get solution name directly from the flow data
                    $solutionName = if ($flow.'sol.friendlyname') {
                        $solName = $flow.'sol.friendlyname'
                        $isManaged = $flow.'sol.ismanaged'
                        if ($isManaged) { "$solName (Managed)" } else { $solName }
                    } else { "Default Solution" }

                    # Build action indicators for the name display
                    $actionIndicators = ""
                    if ($actionDetails.Count -gt 0) {
                        $actionIndicators = " [" + ($actionDetails -join ", ") + "]"
                    }

                    # State indicator for State column only
                    $stateIcon = if ($isOn) { "[ON]" } else { "[OFF]" }
                    $displayName = "$($flow.name)$actionIndicators"

                    # Prepare action list for column display
                    $actionList = if ($actionDetails.Count -gt 0) { $actionDetails -join ", " } else { "" }

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $displayName
                        Id = $itemId
                        WorkflowIdUnique = if ($flow.workflowidunique) { $flow.workflowidunique.ToString() } else { $null }
                        State = $state
                        StateIndicator = $stateIcon
                        Actions = $actionList
                        Solution = $solutionName
                        ErrorCount = if ($errorInfo.Count -gt 0) { $errorInfo.Count.ToString() } else { "0" }
                        HasErrors = $errorInfo.HasErrors
                        ModifiedOn = if ($flow.modifiedon) { $flow.modifiedon.ToString() } else { "" }
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = $flow.primaryentity
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl -EnvironmentId $environmentId
                        ClientData = if ($flow.clientdata) { $flow.clientdata } else { "" }
                        StatusCode = $flow.statuscode
                        StateCode = $stateValue
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning Cloud Flows: $($_.Exception.Message)"
        }
    }

    # 2. Classic Workflows
    if ($ScanClassicWorkflows) {
        $currentStep++
        & $updateProgress "Scanning Classic Workflows..." $currentStep

        $classicFetch = @"
<fetch>
  <entity name='workflow'>
    <attribute name='name' />
    <attribute name='workflowid' />
    <attribute name='description' />
    <attribute name='statecode' />
    <attribute name='xaml' />
    <attribute name='modifiedon' />
    <link-entity name='solution' from='solutionid' to='solutionid' alias='sol' link-type='outer'>
      <attribute name='friendlyname' />
      <attribute name='ismanaged' />
    </link-entity>
    <filter type='and'>
      <condition attribute='type' operator='eq' value='1' />
      <condition attribute='category' operator='ne' value='5' />
      <condition attribute='category' operator='ne' value='4' />
    </filter>
    <order attribute='modifiedon' descending='true' />
  </entity>
</fetch>
"@

        try {
            $workflows = Get-CrmRecordsByFetch -conn $conn -Fetch $classicFetch

            foreach ($wf in $workflows.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                $wfName = if ($wf.name) { $wf.name.ToLower() } else { "" }
                foreach ($keyword in $Keywords) {
                    # Exact name match gets highest priority
                    if ($wfName -eq $keyword) {
                        $matchScore += 15
                        $matchReasons += "[EXACT MATCH] Name exactly matches '$keyword'"
                    } elseif ($wfName -like "*$keyword*") {
                        $matchScore += 8
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                # Check primary entity (high value)
                if ($wf.primaryentity) {
                    $primaryEntity = $wf.primaryentity.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($primaryEntity -eq $keyword) {
                            $matchScore += 12
                            $matchReasons += "[PRIMARY ENTITY] Works with entity '$keyword'"
                        } elseif ($primaryEntity -like "*$keyword*") {
                            $matchScore += 6
                            $matchReasons += "Primary entity contains '$keyword'"
                        }
                    }
                }

                # Track action types for classic workflows (informational, low score)
                $hasEmail = $false
                $actionDetails = @()

                if ($wf.xaml) {
                    $xaml = $wf.xaml.ToLower()

                    if ($xaml -match 'sendemail|createemail') {
                        $matchScore += 1
                        $hasEmail = $true
                        $matchReasons += "[EMAIL] Contains email action"
                        $actionDetails += "Email"
                    }

                    if ($xaml -match 'msdyn_project') {
                        $matchScore += 1
                        $matchReasons += "Works with project entity"
                    }

                    foreach ($keyword in $Keywords) {
                        if ($xaml -like "*$keyword*") {
                            $matchScore += 1
                        }
                    }
                }

                if ($wf.modifiedon -and $BreakageDate) {
                    try {
                        $modDate = [DateTime]::Parse($wf.modifiedon.ToString())
                        $daysSinceBreakage = ($modDate - $BreakageDate).Days
                        if ([Math]::Abs($daysSinceBreakage) -le 30) {
                            $matchScore += 2
                            $matchReasons += "Modified around breakage date ($daysSinceBreakage days)"
                        }
                    } catch { }
                }

                if ($matchScore -gt 0) {
                    # For Classic Workflows, check statecode
                    # statecode 1 = Activated
                    # CRM PowerShell module returns formatted strings like "Activated" or "Draft"
                    $stateValue = $wf.statecode
                    if ($stateValue -is [Microsoft.Xrm.Sdk.OptionSetValue]) {
                        $stateValue = $stateValue.Value
                        $isOn = ($stateValue -eq 1)
                    } else {
                        # Handle string value returned by formatted view
                        $isOn = ($stateValue -eq "Activated")
                    }
                    $state = if ($isOn) { "On" } else { "Off" }
                    $itemType = "Classic Workflow"
                    $itemId = $wf.workflowid

                    # Get error count
                    $errorInfo = Get-ItemErrors -ItemId $itemId -ItemType $itemType -Connection $conn

                    # Get solution name directly from the workflow data
                    $solutionName = if ($wf.'sol.friendlyname') {
                        $solName = $wf.'sol.friendlyname'
                        $isManaged = $wf.'sol.ismanaged'
                        if ($isManaged) { "$solName (Managed)" } else { $solName }
                    } else { "Default Solution" }

                    # Build action indicators
                    $actionIndicators = ""
                    if ($actionDetails.Count -gt 0) {
                        $actionIndicators = " [" + ($actionDetails -join ", ") + "]"
                    }

                    # State indicator for State column only
                    $stateIcon = if ($isOn) { "[ON]" } else { "[OFF]" }
                    $displayName = "$($wf.name)$actionIndicators"

                    # Prepare action list for column display
                    $actionList = if ($actionDetails.Count -gt 0) { $actionDetails -join ", " } else { "" }

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $displayName
                        Id = $itemId
                        State = $state
                        StateIndicator = $stateIcon
                        Actions = $actionList
                        Solution = $solutionName
                        ErrorCount = if ($errorInfo.Count -gt 0) { $errorInfo.Count.ToString() } else { "0" }
                        HasErrors = $errorInfo.HasErrors
                        ModifiedOn = if ($wf.modifiedon) { $wf.modifiedon.ToString() } else { "" }
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = ""
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning Classic Workflows: $($_.Exception.Message)"
        }
    }

    # 3. Plugins
    if ($ScanPlugins) {
        $currentStep++
        & $updateProgress "Scanning Plugins..." $currentStep

        $pluginsFetch = @"
<fetch>
  <entity name='plugintype'>
    <attribute name='typename' />
    <attribute name='plugintypeid' />
    <attribute name='friendlyname' />
    <link-entity name='sdkmessageprocessingstep' from='plugintypeid' to='plugintypeid' alias='step'>
      <attribute name='name' />
      <attribute name='sdkmessageprocessingstepid' />
      <attribute name='statecode' />
      <attribute name='stage' />
      <attribute name='mode' />
      <link-entity name='solution' from='solutionid' to='solutionid' alias='sol' link-type='outer'>
        <attribute name='friendlyname' />
        <attribute name='ismanaged' />
      </link-entity>
      <link-entity name='sdkmessagefilter' from='sdkmessagefilterid' to='sdkmessagefilterid' alias='filter'>
        <attribute name='primaryobjecttypecode' />
      </link-entity>
      <link-entity name='sdkmessage' from='sdkmessageid' to='sdkmessageid' alias='message'>
        <attribute name='name' />
      </link-entity>
    </link-entity>
  </entity>
</fetch>
"@

        try {
            $plugins = Get-CrmRecordsByFetch -conn $conn -Fetch $pluginsFetch

            foreach ($plugin in $plugins.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                # Check if primary object type matches keyword (high priority)
                $entityCode = $plugin.'filter.primaryobjecttypecode'
                if ($entityCode -eq 10054) {
                    $matchScore += 1
                    $matchReasons += "Registered on Project entity"
                }

                if ($plugin.'step.name') {
                    $stepName = $plugin.'step.name'.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($stepName -eq $keyword) {
                            $matchScore += 15
                            $matchReasons += "[EXACT MATCH] Step name exactly matches '$keyword'"
                        } elseif ($stepName -like "*$keyword*") {
                            $matchScore += 8
                            $matchReasons += "Step name contains '$keyword'"
                        }
                    }
                }

                # Check if plugin is registered on entity matching keyword
                if ($plugin.'filter.primaryobjecttypecode_name') {
                    $entityName = $plugin.'filter.primaryobjecttypecode_name'.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($entityName -eq $keyword) {
                            $matchScore += 12
                            $matchReasons += "[PRIMARY ENTITY] Registered on entity '$keyword'"
                        } elseif ($entityName -like "*$keyword*") {
                            $matchScore += 6
                            $matchReasons += "Registered on entity containing '$keyword'"
                        }
                    }
                }

                $messageName = $plugin.'message.name'
                if ($messageName -in @('Create', 'Update', 'Delete')) {
                    if ($matchScore -gt 0) {
                        $matchReasons += "Triggers on $messageName"
                    }
                }

                if ($matchScore -gt 0) {
                    $state = if ($plugin.'step.statecode' -eq 0) { "On" } else { "Off" }
                    $stage = switch ($plugin.'step.stage') {
                        10 { "Pre-validation" }
                        20 { "Pre-operation" }
                        40 { "Post-operation" }
                        default { "Unknown" }
                    }
                    $itemType = "Plugin"
                    $itemId = $plugin.'step.sdkmessageprocessingstepid'

                    # Get error count
                    $errorInfo = Get-ItemErrors -ItemId $itemId -ItemType $itemType -Connection $conn

                    # Get solution name directly from the plugin step data
                    $solutionName = if ($plugin.'sol.friendlyname') {
                        $solName = $plugin.'sol.friendlyname'
                        $isManaged = $plugin.'sol.ismanaged'
                        if ($isManaged) { "$solName (Managed)" } else { $solName }
                    } else { "Default Solution" }

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = "$($plugin.typename) - $($plugin.'step.name')"
                        Id = $itemId
                        State = $state
                        StateIndicator = if ($plugin.'step.statecode' -eq 0) { "[ON]" } else { "[OFF]" }
                        Actions = ""
                        Solution = $solutionName
                        ErrorCount = if ($errorInfo.Count -gt 0) { $errorInfo.Count.ToString() } else { "0" }
                        HasErrors = $errorInfo.HasErrors
                        ModifiedOn = ""
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ") + " | Stage: $stage | Message: $messageName"
                        PrimaryEntity = $entityCode
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning Plugins: $($_.Exception.Message)"
        }
    }

    # 4. Web Resources (skip content to avoid timeout - just match by name)
    if ($ScanWebResources) {
        $currentStep++
        & $updateProgress "Scanning Web Resources..." $currentStep

        # Note: Not fetching 'content' attribute as it can be very large and cause timeouts
        $webResourcesFetch = @"
<fetch>
  <entity name='webresource'>
    <attribute name='name' />
    <attribute name='webresourceid' />
    <attribute name='displayname' />
    <attribute name='webresourcetype' />
    <attribute name='modifiedon' />
    <link-entity name='solution' from='solutionid' to='solutionid' alias='sol' link-type='outer'>
      <attribute name='friendlyname' />
      <attribute name='ismanaged' />
    </link-entity>
    <filter type='and'>
      <filter type='or'>
        <condition attribute='webresourcetype' operator='eq' value='3' />
        <condition attribute='webresourcetype' operator='eq' value='1' />
        <condition attribute='webresourcetype' operator='eq' value='2' />
      </filter>
    </filter>
    <order attribute='modifiedon' descending='true' />
  </entity>
</fetch>
"@

        try {
            $webResources = Get-CrmRecordsByFetch -conn $conn -Fetch $webResourcesFetch

            foreach ($wr in $webResources.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                $wrName = if ($wr.name) { $wr.name.ToLower() } else { "" }
                $wrDisplayName = if ($wr.displayname) { $wr.displayname.ToLower() } else { "" }

                foreach ($keyword in $Keywords) {
                    if ($wrName -eq $keyword -or $wrDisplayName -eq $keyword) {
                        $matchScore += 15
                        $matchReasons += "[EXACT MATCH] Name exactly matches '$keyword'"
                    } elseif ($wrName -like "*$keyword*" -or $wrDisplayName -like "*$keyword*") {
                        $matchScore += 8
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                if ($wr.modifiedon -and $BreakageDate) {
                    try {
                        $modDate = [DateTime]::Parse($wr.modifiedon.ToString())
                        $daysSinceBreakage = ($modDate - $BreakageDate).Days
                        if ([Math]::Abs($daysSinceBreakage) -le 30) {
                            $matchScore += 2
                            $matchReasons += "Modified around breakage date ($daysSinceBreakage days)"
                        }
                    } catch { }
                }

                if ($matchScore -gt 0) {
                    $wrType = switch ($wr.webresourcetype_Property.Value.Value) {
                        1 { "HTML" }
                        2 { "CSS" }
                        3 { "JavaScript" }
                        default { "Unknown" }
                    }
                    $itemType = "Web Resource ($wrType)"
                    $itemId = $wr.webresourceid

                    # Get solution name directly from the web resource data
                    $solutionName = if ($wr.'sol.friendlyname') {
                        $solName = $wr.'sol.friendlyname'
                        $isManaged = $wr.'sol.ismanaged'
                        if ($isManaged) { "$solName (Managed)" } else { $solName }
                    } else { "Default Solution" }

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $wr.displayname
                        Id = $itemId
                        State = "N/A"
                        StateIndicator = ""
                        Actions = ""
                        Solution = $solutionName
                        ErrorCount = "0"
                        HasErrors = $false
                        ModifiedOn = if ($wr.modifiedon) { $wr.modifiedon.ToString() } else { "" }
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = ""
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning Web Resources: $($_.Exception.Message)"
        }
    }

    # 5. Canvas Apps
    if ($ScanCanvasApps) {
        $currentStep++
        & $updateProgress "Scanning Canvas Apps..." $currentStep

        $canvasAppsFetch = @"
<fetch>
  <entity name='canvasapp'>
    <attribute name='name' />
    <attribute name='canvasappid' />
    <attribute name='displayname' />
    <attribute name='description' />
    <attribute name='modifiedon' />
    <attribute name='tags' />
    <attribute name='appdefinition' />
    <attribute name='appcomponents' />
    <attribute name='publishedappversion' />
    <link-entity name='solution' from='solutionid' to='solutionid' alias='sol' link-type='outer'>
      <attribute name='friendlyname' />
      <attribute name='ismanaged' />
    </link-entity>
    <order attribute='modifiedon' descending='true' />
  </entity>
</fetch>
"@

        try {
            $canvasApps = Get-CrmRecordsByFetch -conn $conn -Fetch $canvasAppsFetch

            foreach ($app in $canvasApps.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                # Check app name (highest priority)
                $appName = if ($app.displayname) { $app.displayname.ToLower() } else { "" }
                foreach ($keyword in $Keywords) {
                    if ($appName -eq $keyword) {
                        $matchScore += 15
                        $matchReasons += "[EXACT MATCH] Name exactly matches '$keyword'"
                    } elseif ($appName -like "*$keyword*") {
                        $matchScore += 8
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                # Check description
                if ($app.description) {
                    $appDesc = $app.description.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($appDesc -like "*$keyword*") {
                            $matchScore += 2
                            $matchReasons += "Description contains '$keyword'"
                        }
                    }
                }

                # Check tags
                if ($app.tags) {
                    $appTags = $app.tags.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($appTags -like "*$keyword*") {
                            $matchScore += 2
                            $matchReasons += "Tags contain '$keyword'"
                        }
                    }
                }

                # Parse app definition for connectors and flows
                if ($app.appdefinition) {
                    try {
                        $appDefJson = $app.appdefinition | ConvertFrom-Json

                        # Check for Power Automate connections (lower priority - informational)
                        if ($appDefJson.properties.connectionReferences) {
                            $connRefs = $appDefJson.properties.connectionReferences
                            $connRefNames = $connRefs.PSObject.Properties.Name

                            foreach ($connRef in $connRefNames) {
                                $connRefData = $connRefs.$connRef
                                if ($connRefData.displayName) {
                                    $connName = $connRefData.displayName.ToLower()
                                    foreach ($keyword in $Keywords) {
                                        if ($connName -like "*$keyword*") {
                                            $matchScore += 1
                                            $matchReasons += "Uses connector: $($connRefData.displayName)"
                                        }
                                    }
                                }
                            }
                        }

                        # Check for embedded flows
                        if ($appDefJson.properties.embeddedFlows) {
                            $flows = $appDefJson.properties.embeddedFlows
                            $flowNames = $flows.PSObject.Properties.Name

                            foreach ($flowName in $flowNames) {
                                $flowData = $flows.$flowName
                                if ($flowData.displayName) {
                                    $flowDisplayName = $flowData.displayName.ToLower()
                                    foreach ($keyword in $Keywords) {
                                        if ($flowDisplayName -like "*$keyword*") {
                                            $matchScore += 3
                                            $matchReasons += "Contains flow: $($flowData.displayName)"
                                        }
                                    }
                                }
                            }
                        }
                    } catch {
                        # Ignore JSON parsing errors
                    }
                }

                if ($matchScore -gt 0) {
                    # Canvas apps use publishedappversion to determine if published
                    $state = if ($app.publishedappversion) { "Published" } else { "Unpublished" }
                    $itemType = "Canvas App"
                    $itemId = $app.canvasappid
                    $displayName = if ($app.displayname) { $app.displayname } else { $app.name }

                    # Get solution name directly from the canvas app data
                    $solutionName = if ($app.'sol.friendlyname') {
                        $solName = $app.'sol.friendlyname'
                        $isManaged = $app.'sol.ismanaged'
                        if ($isManaged) { "$solName (Managed)" } else { $solName }
                    } else { "Default Solution" }

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $displayName
                        Id = $itemId
                        State = $state
                        StateIndicator = if ($state -eq "Published") { "[PUB]" } else { "[UNPUB]" }
                        Actions = ""
                        Solution = $solutionName
                        ErrorCount = "0"
                        HasErrors = $false
                        ModifiedOn = if ($app.modifiedon) { $app.modifiedon.ToString() } else { "" }
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = ""
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning Canvas Apps: $($_.Exception.Message)"
        }
    }

    # 5.5. Model-Driven Apps
    if ($ScanModelDrivenApps) {
        $currentStep++
        & $updateProgress "Scanning Model-Driven Apps..." $currentStep

        $appFetch = @"
<fetch>
  <entity name='appmodule'>
    <attribute name='appmoduleid' />
    <attribute name='name' />
    <attribute name='uniquename' />
    <attribute name='description' />
    <attribute name='publishedon' />
    <attribute name='modifiedon' />
    <link-entity name='solution' from='solutionid' to='solutionid' alias='sol' link-type='outer'>
      <attribute name='friendlyname' />
      <attribute name='ismanaged' />
    </link-entity>
  </entity>
</fetch>
"@

        try {
            $apps = Get-CrmRecordsByFetch -conn $conn -Fetch $appFetch
            & $updateProgress "Scanning Model-Driven Apps... (found $($apps.CrmRecords.Count))" $currentStep

            foreach ($app in $apps.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                $appName = if ($app.name) { $app.name.ToLower() } else { "" }
                $appDesc = if ($app.description) { $app.description.ToLower() } else { "" }

                foreach ($keyword in $Keywords) {
                    if ($appName -eq $keyword) {
                        $matchScore += 15
                        $matchReasons += "[EXACT MATCH] Name exactly matches '$keyword'"
                    } elseif ($appName -like "*$keyword*") {
                        $matchScore += 8
                        $matchReasons += "Name contains '$keyword'"
                    }
                    if ($appDesc -like "*$keyword*") {
                        $matchScore += 2
                        $matchReasons += "Description contains '$keyword'"
                    }
                }

                # Query forms associated with this app
                $formFetch = @"
<fetch>
  <entity name='systemform'>
    <attribute name='name' />
    <attribute name='formxml' />
    <link-entity name='appmoduleroles' from='objectid' to='formid' link-type='inner' alias='amr'>
      <link-entity name='appmodule' from='appmoduleid' to='appmoduleid' link-type='inner'>
        <filter>
          <condition attribute='appmoduleid' operator='eq' value='$($app.appmoduleid)' />
        </filter>
      </link-entity>
    </link-entity>
  </entity>
</fetch>
"@
                
                $forms = Get-CrmRecordsByFetch -conn $conn -Fetch $formFetch -ErrorAction SilentlyContinue -TopCount 50
                $webResources = @()
                
                if ($forms -and $forms.CrmRecords) {
                    foreach ($form in $forms.CrmRecords) {
                        if ($form.formxml) {
                            # Parse form XML to find JavaScript web resources
                            try {
                                $formXml = [xml]$form.formxml
                                $events = $formXml.SelectNodes("//event[@name='onload' or @name='onsave' or @name='onchange']//*[@libraryname]")
                                foreach ($event in $events) {
                                    $libName = $event.GetAttribute("libraryname")
                                    if ($libName -and $webResources -notcontains $libName) {
                                        $webResources += $libName
                                        
                                        # Check if web resource name matches keywords
                                        foreach ($keyword in $Keywords) {
                                            if ($libName.ToLower() -like "*$keyword*") {
                                                $matchScore += 2
                                                $matchReasons += "Form uses web resource '$libName' matching '$keyword'"
                                            }
                                        }
                                    }
                                }
                            } catch {
                                # Ignore XML parse errors
                            }
                        }
                    }
                }

                if ($matchScore -gt 0) {
                    $state = if ($app.publishedon) { "Published" } else { "Unpublished" }
                    $itemType = "Model-Driven App"
                    $itemId = $app.appmoduleid
                    $displayName = if ($app.name) { $app.name } else { $app.uniquename }

                    $solutionName = if ($app.'sol.friendlyname') {
                        $solName = $app.'sol.friendlyname'
                        $isManaged = $app.'sol.ismanaged'
                        if ($isManaged) { "$solName (Managed)" } else { $solName }
                    } else { "Default Solution" }

                    $actions = if ($webResources.Count -gt 0) {
                        "$($webResources.Count) JavaScript libraries"
                    } else {
                        "No custom JS"
                    }

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $displayName
                        Id = $itemId
                        State = $state
                        StateIndicator = if ($state -eq "Published") { "[PUB]" } else { "[UNPUB]" }
                        Actions = $actions
                        Solution = $solutionName
                        ErrorCount = "0"
                        HasErrors = $false
                        ModifiedOn = if ($app.modifiedon) { $app.modifiedon.ToString() } else { "" }
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = ""
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl
                        ClientData = ""
                        StatusCode = ""
                        StateCode = ""
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning Model-Driven Apps: $($_.Exception.Message)"
        }
    }

    # 6. Custom Actions
    if ($ScanCustomActions) {
        $currentStep++
        & $updateProgress "Scanning Custom Actions..." $currentStep

        $actionsFetch = @"
<fetch>
  <entity name='workflow'>
    <attribute name='name' />
    <attribute name='workflowid' />
    <attribute name='description' />
    <attribute name='statecode' />
    <attribute name='xaml' />
    <attribute name='modifiedon' />
    <link-entity name='solution' from='solutionid' to='solutionid' alias='sol' link-type='outer'>
      <attribute name='friendlyname' />
      <attribute name='ismanaged' />
    </link-entity>
    <filter type='and'>
      <condition attribute='type' operator='eq' value='1' />
      <condition attribute='category' operator='eq' value='3' />
    </filter>
    <order attribute='modifiedon' descending='true' />
  </entity>
</fetch>
"@

        try {
            $actions = Get-CrmRecordsByFetch -conn $conn -Fetch $actionsFetch

            foreach ($action in $actions.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                $actionName = if ($action.name) { $action.name.ToLower() } else { "" }
                foreach ($keyword in $Keywords) {
                    if ($actionName -eq $keyword) {
                        $matchScore += 15
                        $matchReasons += "[EXACT MATCH] Name exactly matches '$keyword'"
                    } elseif ($actionName -like "*$keyword*") {
                        $matchScore += 8
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                # Check primary entity (high value)
                if ($action.primaryentity) {
                    $primaryEntity = $action.primaryentity.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($primaryEntity -eq $keyword) {
                            $matchScore += 12
                            $matchReasons += "[PRIMARY ENTITY] Works with entity '$keyword'"
                        } elseif ($primaryEntity -like "*$keyword*") {
                            $matchScore += 6
                            $matchReasons += "Primary entity contains '$keyword'"
                        }
                    }
                }

                if ($action.xaml) {
                    $xaml = $action.xaml.ToLower()

                    if ($xaml -match 'sendemail|createemail') {
                        $matchScore += 1
                        $matchReasons += "Contains email action"
                    }

                    if ($xaml -match 'msdyn_project') {
                        $matchScore += 1
                        $matchReasons += "Works with project entity"
                    }

                    foreach ($keyword in $Keywords) {
                        if ($xaml -like "*$keyword*") {
                            $matchScore += 1
                        }
                    }
                }

                if ($action.modifiedon -and $BreakageDate) {
                    try {
                        $modDate = [DateTime]::Parse($action.modifiedon.ToString())
                        $daysSinceBreakage = ($modDate - $BreakageDate).Days
                        if ([Math]::Abs($daysSinceBreakage) -le 30) {
                            $matchScore += 2
                            $matchReasons += "Modified around breakage date ($daysSinceBreakage days)"
                        }
                    } catch { }
                }

                if ($matchScore -gt 0) {
                    # For Custom Actions, statecode 1 = Activated
                    # CRM PowerShell module returns formatted strings like "Activated" or "Draft"
                    $stateValue = $action.statecode
                    if ($stateValue -is [Microsoft.Xrm.Sdk.OptionSetValue]) {
                        $isOn = ($stateValue.Value -eq 1)
                    } else {
                        $isOn = ($stateValue -eq "Activated")
                    }
                    $state = if ($isOn) { "On" } else { "Off" }
                    $itemType = "Custom Action"
                    $itemId = $action.workflowid

                    # Get error count
                    $errorInfo = Get-ItemErrors -ItemId $itemId -ItemType $itemType -Connection $conn

                    # Get solution name directly from the action data
                    $solutionName = if ($action.'sol.friendlyname') {
                        $solName = $action.'sol.friendlyname'
                        $isManaged = $action.'sol.ismanaged'
                        if ($isManaged) { "$solName (Managed)" } else { $solName }
                    } else { "Default Solution" }

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $action.name
                        Id = $itemId
                        State = $state
                        StateIndicator = if ($isOn) { "[ON]" } else { "[OFF]" }
                        Actions = ""
                        Solution = $solutionName
                        ErrorCount = if ($errorInfo.Count -gt 0) { $errorInfo.Count.ToString() } else { "0" }
                        HasErrors = $errorInfo.HasErrors
                        ModifiedOn = if ($action.modifiedon) { $action.modifiedon.ToString() } else { "" }
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = ""
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl
                        ClientData = ""
                        StatusCode = if ($action.statuscode) { $action.statuscode } else { "" }
                        StateCode = if ($action.statecode) { $action.statecode } else { "" }
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning Custom Actions: $($_.Exception.Message)"
        }
    }

    # 7. Service Endpoints
    if ($ScanServiceEndpoints) {
        $currentStep++
        & $updateProgress "Scanning Service Endpoints..." $currentStep

        $endpointsFetch = @"
<fetch>
  <entity name='serviceendpoint'>
    <attribute name='name' />
    <attribute name='serviceendpointid' />
    <attribute name='description' />
    <attribute name='url' />
    <attribute name='contract' />
    <attribute name='modifiedon' />
    <link-entity name='sdkmessageprocessingstep' from='eventhandler' to='serviceendpointid' alias='step'>
      <attribute name='name' />
      <attribute name='sdkmessageprocessingstepid' />
      <attribute name='statecode' />
      <link-entity name='solution' from='solutionid' to='solutionid' alias='sol' link-type='outer'>
        <attribute name='friendlyname' />
        <attribute name='ismanaged' />
      </link-entity>
      <link-entity name='sdkmessagefilter' from='sdkmessagefilterid' to='sdkmessagefilterid' alias='filter'>
        <attribute name='primaryobjecttypecode' />
      </link-entity>
    </link-entity>
  </entity>
</fetch>
"@

        try {
            $endpoints = Get-CrmRecordsByFetch -conn $conn -Fetch $endpointsFetch

            foreach ($endpoint in $endpoints.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                $entityCode = $endpoint.'filter.primaryobjecttypecode'
                if ($entityCode -eq 10054) {
                    $matchScore += 1
                    $matchReasons += "Registered on Project entity"
                }

                $endpointName = if ($endpoint.name) { $endpoint.name.ToLower() } else { "" }
                foreach ($keyword in $Keywords) {
                    if ($endpointName -eq $keyword) {
                        $matchScore += 15
                        $matchReasons += "[EXACT MATCH] Name exactly matches '$keyword'"
                    } elseif ($endpointName -like "*$keyword*") {
                        $matchScore += 8
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                # Check if registered on entity matching keyword
                if ($endpoint.'filter.primaryobjecttypecode_name') {
                    $entityName = $endpoint.'filter.primaryobjecttypecode_name'.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($entityName -eq $keyword) {
                            $matchScore += 12
                            $matchReasons += "[PRIMARY ENTITY] Registered on entity '$keyword'"
                        } elseif ($entityName -like "*$keyword*") {
                            $matchScore += 6
                            $matchReasons += "Registered on entity containing '$keyword'"
                        }
                    }
                }

                if ($endpoint.url) {
                    $url = $endpoint.url.ToLower()
                    if ($url -match 'logic.*app|function.*app|servicebus|webhook') {
                        $matchScore += 1
                        $matchReasons += "Connects to external automation"
                    }
                }

                if ($matchScore -gt 0) {
                    $state = if ($endpoint.'step.statecode' -eq 0) { "On" } else { "Off" }
                    $contract = switch ($endpoint.contract) {
                        1 { "OneWay" }
                        2 { "Queue" }
                        3 { "Rest" }
                        4 { "TwoWay" }
                        5 { "Topic" }
                        6 { "Webhook" }
                        7 { "EventHub" }
                        8 { "EventGrid" }
                        default { "Unknown" }
                    }
                    $itemType = "Service Endpoint ($contract)"
                    $itemId = $endpoint.'step.sdkmessageprocessingstepid'

                    # Get solution name directly from the step data
                    $solutionName = if ($endpoint.'sol.friendlyname') {
                        $solName = $endpoint.'sol.friendlyname'
                        $isManaged = $endpoint.'sol.ismanaged'
                        if ($isManaged) { "$solName (Managed)" } else { $solName }
                    } else { "Default Solution" }

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = "$($endpoint.name) - $($endpoint.'step.name')"
                        Id = $itemId
                        State = $state
                        StateIndicator = if ($endpoint.'step.statecode' -eq 0) { "[ON]" } else { "[OFF]" }
                        Actions = ""
                        Solution = $solutionName
                        ErrorCount = "0"
                        HasErrors = $false
                        ModifiedOn = if ($endpoint.modifiedon) { $endpoint.modifiedon.ToString() } else { "" }
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = $entityCode
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning Service Endpoints: $($_.Exception.Message)"
        }
    }

    # 8. SLAs
    if ($ScanSLAs) {
        $currentStep++
        & $updateProgress "Scanning SLAs..." $currentStep

        $slasFetch = @"
<fetch>
  <entity name='sla'>
    <attribute name='name' />
    <attribute name='slaid' />
    <attribute name='description' />
    <attribute name='statecode' />
    <attribute name='modifiedon' />
    <attribute name='objecttypecode' />
    <link-entity name='solution' from='solutionid' to='solutionid' alias='sol' link-type='outer'>
      <attribute name='friendlyname' />
      <attribute name='ismanaged' />
    </link-entity>
  </entity>
</fetch>
"@

        try {
            $slas = Get-CrmRecordsByFetch -conn $conn -Fetch $slasFetch

            foreach ($sla in $slas.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                if ($sla.objecttypecode -eq 10054) {
                    $matchScore += 1
                    $matchReasons += "Applied to Project entity"
                }

                $slaName = if ($sla.name) { $sla.name.ToLower() } else { "" }
                foreach ($keyword in $Keywords) {
                    if ($slaName -eq $keyword) {
                        $matchScore += 15
                        $matchReasons += "[EXACT MATCH] Name exactly matches '$keyword'"
                    } elseif ($slaName -like "*$keyword*") {
                        $matchScore += 8
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                # Check if applied to entity matching keyword
                if ($sla.objecttypecode_name) {
                    $entityName = $sla.objecttypecode_name.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($entityName -eq $keyword) {
                            $matchScore += 12
                            $matchReasons += "[PRIMARY ENTITY] Applied to entity '$keyword'"
                        } elseif ($entityName -like "*$keyword*") {
                            $matchScore += 6
                            $matchReasons += "Applied to entity containing '$keyword'"
                        }
                    }
                }

                if ($matchScore -gt 0) {
                    $state = if ($sla.statecode -eq 0) { "Active" } else { "Inactive" }
                    $itemType = "SLA"
                    $itemId = $sla.slaid

                    # Get solution name directly from the SLA data
                    $solutionName = if ($sla.'sol.friendlyname') {
                        $solName = $sla.'sol.friendlyname'
                        $isManaged = $sla.'sol.ismanaged'
                        if ($isManaged) { "$solName (Managed)" } else { $solName }
                    } else { "Default Solution" }

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $sla.name
                        Id = $itemId
                        State = $state
                        StateIndicator = if ($sla.statecode -eq 0) { "[ACTIVE]" } else { "[INACTIVE]" }
                        Actions = ""
                        Solution = $solutionName
                        ErrorCount = "0"
                        HasErrors = $false
                        ModifiedOn = if ($sla.modifiedon) { $sla.modifiedon.ToString() } else { "" }
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = $sla.objecttypecode
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning SLAs: $($_.Exception.Message)"
        }
    }

    # 9. Business Process Flows
    if ($ScanBPFs) {
        $currentStep++
        & $updateProgress "Scanning Business Process Flows..." $currentStep

        $bpfsFetch = @"
<fetch>
  <entity name='workflow'>
    <attribute name='name' />
    <attribute name='workflowid' />
    <attribute name='statecode' />
    <attribute name='modifiedon' />
    <link-entity name='solution' from='solutionid' to='solutionid' alias='sol' link-type='outer'>
      <attribute name='friendlyname' />
      <attribute name='ismanaged' />
    </link-entity>
    <filter type='and'>
      <condition attribute='category' operator='eq' value='4' />
    </filter>
    <order attribute='modifiedon' descending='true' />
  </entity>
</fetch>
"@

        try {
            $bpfs = Get-CrmRecordsByFetch -conn $conn -Fetch $bpfsFetch

            foreach ($bpf in $bpfs.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                $bpfName = if ($bpf.name) { $bpf.name.ToLower() } else { "" }
                foreach ($keyword in $Keywords) {
                    if ($bpfName -eq $keyword) {
                        $matchScore += 15
                        $matchReasons += "[EXACT MATCH] Name exactly matches '$keyword'"
                    } elseif ($bpfName -like "*$keyword*") {
                        $matchScore += 8
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                # Check primary entity (high value)
                if ($bpf.primaryentity) {
                    $primaryEntity = $bpf.primaryentity.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($primaryEntity -eq $keyword) {
                            $matchScore += 12
                            $matchReasons += "[PRIMARY ENTITY] Works with entity '$keyword'"
                        } elseif ($primaryEntity -like "*$keyword*") {
                            $matchScore += 6
                            $matchReasons += "Primary entity contains '$keyword'"
                        }
                    }
                }

                if ($matchScore -gt 0) {
                    # For BPFs, statecode 1 = Activated
                    # CRM PowerShell module returns formatted strings like "Activated" or "Draft"
                    $stateValue = $bpf.statecode
                    if ($stateValue -is [Microsoft.Xrm.Sdk.OptionSetValue]) {
                        $isOn = ($stateValue.Value -eq 1)
                    } else {
                        $isOn = ($stateValue -eq "Activated")
                    }
                    $state = if ($isOn) { "On" } else { "Off" }
                    $itemType = "Business Process Flow"
                    $itemId = $bpf.workflowid

                    $errorInfo = Get-ItemErrors -ItemId $itemId -ItemType $itemType -Connection $conn

                    # Get solution name directly from the BPF data
                    $solutionName = if ($bpf.'sol.friendlyname') {
                        $solName = $bpf.'sol.friendlyname'
                        $isManaged = $bpf.'sol.ismanaged'
                        if ($isManaged) { "$solName (Managed)" } else { $solName }
                    } else { "Default Solution" }

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $bpf.name
                        Id = $itemId
                        State = $state
                        StateIndicator = if ($isOn) { "[ON]" } else { "[OFF]" }
                        Actions = ""
                        Solution = $solutionName
                        ErrorCount = if ($errorInfo.Count -gt 0) { $errorInfo.Count.ToString() } else { "0" }
                        HasErrors = $errorInfo.HasErrors
                        ModifiedOn = if ($bpf.modifiedon) { $bpf.modifiedon.ToString() } else { "" }
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = ""
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl
                        ClientData = ""
                        StatusCode = if ($bpf.statuscode) { $bpf.statuscode } else { "" }
                        StateCode = if ($bpf.statecode) { $bpf.statecode } else { "" }
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning BPFs: $($_.Exception.Message)"
        }
    }

    # 10. Table Schemas (Entities)
    if ($ScanTableSchemas) {
        $currentStep++
        & $updateProgress "Scanning Table Schemas..." $currentStep

        try {
            # Use Get-CrmEntityAllMetadata to retrieve entity metadata (not FetchXML)
            # Entity metadata is not accessible via the 'entity' table in FetchXML
            $entities = Get-CrmEntityAllMetadata -conn $conn -OnlyPublished $true -EntityFilters Entity
            & $updateProgress "Scanning Table Schemas... (found $($entities.Count))" $currentStep

            foreach ($entity in $entities) {
                $matchScore = 0
                $matchReasons = @()

                $entityLogicalName = if ($entity.LogicalName) { $entity.LogicalName.ToLower() } else { "" }
                # Safely access DisplayName - check each nested property
                $entityDisplayName = ""
                if ($entity.DisplayName) {
                    if ($entity.DisplayName.UserLocalizedLabel -and $entity.DisplayName.UserLocalizedLabel.Label) {
                        $entityDisplayName = $entity.DisplayName.UserLocalizedLabel.Label.ToLower()
                    }
                }
                $entitySchemaName = if ($entity.SchemaName) { $entity.SchemaName.ToLower() } else { "" }

                foreach ($keyword in $Keywords) {
                    # Exact logical name match gets highest priority
                    if ($entityLogicalName -and $entityLogicalName -eq $keyword) {
                        $matchScore += 20
                        $matchReasons += "[EXACT MATCH] Logical name exactly matches '$keyword'"
                    } elseif ($entityLogicalName -and $entityLogicalName -like "*$keyword*") {
                        $matchScore += 10
                        $matchReasons += "Logical name contains '$keyword'"
                    }
                    if ($entityDisplayName -and $entityDisplayName -eq $keyword) {
                        $matchScore += 15
                        $matchReasons += "[EXACT MATCH] Display name exactly matches '$keyword'"
                    } elseif ($entityDisplayName -and $entityDisplayName -like "*$keyword*") {
                        $matchScore += 5
                        $matchReasons += "Display name contains '$keyword'"
                    }
                    if ($entitySchemaName -and $entitySchemaName -like "*$keyword*") {
                        $matchScore += 2
                        $matchReasons += "Schema name contains '$keyword'"
                    }
                }

                if ($matchScore -gt 0) {
                    # Safely get display name for result
                    $displayName = $entity.LogicalName
                    if ($entity.DisplayName -and $entity.DisplayName.UserLocalizedLabel -and $entity.DisplayName.UserLocalizedLabel.Label) {
                        $displayName = $entity.DisplayName.UserLocalizedLabel.Label
                    }
                    $itemType = "Table Schema"
                    # Use MetadataId for solution lookup
                    $itemId = $entity.MetadataId

                    # Build URL to table in maker portal
                    $tableUrl = "https://make.powerapps.com/environments/Default-*/solutions/tables/$($entity.LogicalName)"

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = "$displayName ($entityLogicalName)"
                        Id = $itemId
                        State = "N/A"
                        StateIndicator = ""
                        Actions = ""
                        Solution = Get-ItemSolution -ItemId $itemId -ItemType $itemType -Connection $conn
                        ErrorCount = "0"
                        HasErrors = $false
                        ModifiedOn = ""
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = $entity.LogicalName
                        Url = $tableUrl
                    })
                }
            }
        } catch {
            $errorMsg = $_.Exception.Message
            $errorStack = $_.ScriptStackTrace
            $txtStatus.Text = "Error scanning table schemas: $errorMsg"
            Write-Host "Table Schema Error: $errorMsg" -ForegroundColor Red
            Write-Host "Stack: $errorStack" -ForegroundColor DarkRed
        }
    }

    # 11. Connection References
    if ($ScanConnections) {
        $currentStep++
        & $updateProgress "Scanning Connection References..." $currentStep

        $connectionsFetch = @"
<fetch>
  <entity name='connectionreference'>
    <attribute name='connectionreferenceid' />
    <attribute name='connectionreferencelogicalname' />
    <attribute name='connectorid' />
    <attribute name='connectionid' />
    <attribute name='statecode' />
    <attribute name='statuscode' />
    <link-entity name='solution' from='solutionid' to='solutionid' alias='sol' link-type='outer'>
      <attribute name='friendlyname' />
      <attribute name='ismanaged' />
    </link-entity>
  </entity>
</fetch>
"@

        try {
            $connections = Get-CrmRecordsByFetch -conn $conn -Fetch $connectionsFetch
            & $updateProgress "Scanning Connection References... (found $($connections.CrmRecords.Count))" $currentStep

            foreach ($connRef in $connections.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                $connName = if ($connRef.connectionreferencelogicalname) { $connRef.connectionreferencelogicalname.ToLower() } else { "" }
                
                foreach ($keyword in $Keywords) {
                    if ($connName -eq $keyword) {
                        $matchScore += 15
                        $matchReasons += "[EXACT MATCH] Name exactly matches '$keyword'"
                    } elseif ($connName -like "*$keyword*") {
                        $matchScore += 8
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                # Check if connection is broken (no connectionid) - only add if keyword matched
                if (-not $connRef.connectionid -and $matchScore -gt 0) {
                    $matchScore += 3
                    $matchReasons += "[BROKEN] No connection associated"
                }

                # Only add if keyword actually matched
                if ($matchScore -gt 0) {
                    $itemType = "Connection Reference"
                    $itemId = $connRef.connectionreferenceid
                    $connectorType = if ($connRef.connectorid) { 
                        $connRef.connectorid -replace '/providers/Microsoft.PowerApps/apis/', ''
                    } else { "Unknown" }

                    # Get solution name
                    $solutionName = if ($connRef.'sol.friendlyname') {
                        $solName = $connRef.'sol.friendlyname'
                        $isManaged = $connRef.'sol.ismanaged'
                        if ($isManaged) { "$solName (Managed)" } else { $solName }
                    } else { "Default Solution" }

                    $state = if ($connRef.connectionid) { "Connected" } else { "Not Connected" }
                    $stateIcon = if ($connRef.connectionid) { "[OK]" } else { "[BROKEN]" }

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = "$stateIcon $connName"
                        Id = $itemId
                        State = $state
                        StateIndicator = $stateIcon
                        Actions = $connectorType
                        Solution = $solutionName
                        ErrorCount = if ($connRef.connectionid) { "0" } else { "1" }
                        HasErrors = -not $connRef.connectionid
                        ModifiedOn = ""
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = ""
                        Url = ""
                        ClientData = ""
                        StatusCode = if ($connRef.statuscode) { $connRef.statuscode } else { "" }
                        StateCode = if ($connRef.statecode) { $connRef.statecode } else { "" }
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning Connection References: $($_.Exception.Message)"
        }
    }

    # 12. Environment Variables
    if ($ScanEnvVariables) {
        $currentStep++
        & $updateProgress "Scanning Environment Variables..." $currentStep

        $envVarsFetch = @"
<fetch>
  <entity name='environmentvariabledefinition'>
    <attribute name='environmentvariabledefinitionid' />
    <attribute name='schemaname' />
    <attribute name='displayname' />
    <attribute name='type' />
    <attribute name='defaultvalue' />
    <link-entity name='environmentvariablevalue' from='environmentvariabledefinitionid' to='environmentvariabledefinitionid' alias='val' link-type='outer'>
      <attribute name='value' />
    </link-entity>
    <link-entity name='solution' from='solutionid' to='solutionid' alias='sol' link-type='outer'>
      <attribute name='friendlyname' />
      <attribute name='ismanaged' />
    </link-entity>
  </entity>
</fetch>
"@

        try {
            $envVars = Get-CrmRecordsByFetch -conn $conn -Fetch $envVarsFetch
            & $updateProgress "Scanning Environment Variables... (found $($envVars.CrmRecords.Count))" $currentStep

            foreach ($envVar in $envVars.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                $varName = if ($envVar.schemaname) { $envVar.schemaname.ToLower() } else { "" }
                $displayName = if ($envVar.displayname) { $envVar.displayname } else { $envVar.schemaname }
                
                foreach ($keyword in $Keywords) {
                    if ($varName -eq $keyword) {
                        $matchScore += 15
                        $matchReasons += "[EXACT MATCH] Name exactly matches '$keyword'"
                    } elseif ($varName -like "*$keyword*") {
                        $matchScore += 8
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                # Check if variable has no value set - only flag if keyword matched
                $hasValue = $envVar.'val.value' -or $envVar.defaultvalue
                if (-not $hasValue -and $matchScore -gt 0) {
                    $matchScore += 2
                    $matchReasons += "[NO VALUE] Environment variable not configured"
                }

                # Only add if keyword actually matched
                if ($matchScore -gt 0) {
                    $itemType = "Environment Variable"
                    $itemId = $envVar.environmentvariabledefinitionid
                    
                    # Determine variable type
                    $varType = switch ($envVar.type_Property.Value.Value) {
                        100000000 { "String" }
                        100000001 { "Number" }
                        100000002 { "Boolean" }
                        100000003 { "JSON" }
                        100000004 { "Data Source" }
                        default { "Unknown" }
                    }

                    # Get solution name
                    $solutionName = if ($envVar.'sol.friendlyname') {
                        $solName = $envVar.'sol.friendlyname'
                        $isManaged = $envVar.'sol.ismanaged'
                        if ($isManaged) { "$solName (Managed)" } else { $solName }
                    } else { "Default Solution" }

                    $currentValue = if ($envVar.'val.value') { 
                        $envVar.'val.value' 
                    } elseif ($envVar.defaultvalue) { 
                        "$($envVar.defaultvalue) (default)" 
                    } else { 
                        "[NOT SET]" 
                    }

                    $state = if ($hasValue) { "Configured" } else { "Not Configured" }
                    $stateIcon = if ($hasValue) { "[SET]" } else { "[EMPTY]" }

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = "$stateIcon $displayName"
                        Id = $itemId
                        State = $state
                        StateIndicator = $stateIcon
                        Actions = "$varType : $currentValue"
                        Solution = $solutionName
                        ErrorCount = if ($hasValue) { "0" } else { "1" }
                        HasErrors = -not $hasValue
                        ModifiedOn = ""
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = ""
                        Url = ""
                        ClientData = ""
                        StatusCode = ""
                        StateCode = ""
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning Environment Variables: $($_.Exception.Message)"
        }
    }

    return $discoveredItems
}

# Scan button click handler
$btnScan.Add_Click({
    $txtStatus.Text = "Starting scan..."
    $progressBar.Value = 0
    $btnScan.IsEnabled = $false
    $dgResults.ItemsSource = $null

    # Force UI update
    [System.Windows.Forms.Application]::DoEvents()

    try {
        # Get values from UI
        $keywordsText = $txtKeywords.Text
        $breakageDate = $dpBreakageDate.SelectedDate
        $keywords = $keywordsText -split ',' | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ -ne "" }

        if (-not $breakageDate) {
            $breakageDate = Get-Date
        }

        # Run the scan (use -eq $true to handle nullable booleans from WPF)
        $results = Start-DataverseScan `
            -Connection $script:SyncHash.Connection `
            -Keywords $keywords `
            -BreakageDate $breakageDate `
            -ScanCloudFlows ($chkCloudFlows.IsChecked -eq $true) `
            -ScanClassicWorkflows ($chkClassicWorkflows.IsChecked -eq $true) `
            -ScanPlugins ($chkPlugins.IsChecked -eq $true) `
            -ScanWebResources ($chkWebResources.IsChecked -eq $true) `
            -ScanCanvasApps ($chkCanvasApps.IsChecked -eq $true) `
            -ScanModelDrivenApps ($chkModelDrivenApps.IsChecked -eq $true) `
            -ScanCustomActions ($chkCustomActions.IsChecked -eq $true) `
            -ScanServiceEndpoints ($chkServiceEndpoints.IsChecked -eq $true) `
            -ScanSLAs ($chkSLAs.IsChecked -eq $true) `
            -ScanBPFs ($chkBPFs.IsChecked -eq $true) `
            -ScanTableSchemas ($chkTableSchemas.IsChecked -eq $true) `
            -ScanConnections ($chkConnections.IsChecked -eq $true) `
            -ScanEnvVariables ($chkEnvVariables.IsChecked -eq $true) `
            -OrgUrl $script:SyncHash.OrgUrl

        # Sort and display results
        $sorted = $results | Sort-Object -Property Score -Descending
        $script:SyncHash.Results = $sorted

        $dgResults.ItemsSource = $sorted
        $txtResultCount.Text = " ($($sorted.Count) items)"
        $txtStatus.Text = "Scan complete - Found $($sorted.Count) matching items"
        $progressBar.Value = 100
        $btnAISummary.IsEnabled = $sorted.Count -gt 0
        $btnExportCsv.IsEnabled = $sorted.Count -gt 0
        $btnSaveSnapshot.IsEnabled = $sorted.Count -gt 0
        $btnViewDependencies.IsEnabled = $sorted.Count -gt 0
        $btnClearResults.IsEnabled = $sorted.Count -gt 0

        # Enable "Compare Snapshots" if we have at least 2 snapshots
        $snapshotCount = (Get-ChildItem -Path "$env:USERPROFILE\.powersearch\snapshots" -Filter "Snapshot_*.json" -ErrorAction SilentlyContinue).Count
        $btnCompareSnapshots.IsEnabled = $snapshotCount -ge 2

        # Enable "Get Flow History" button if we have cloud flows
        $hasCloudFlows = ($sorted | Where-Object { $_.Type -eq "Cloud Flow" }).Count -gt 0
        $btnFlowHistory.IsEnabled = $hasCloudFlows

        # Check if any flows show "Setup required" and prompt user
        $flowsNeedingSetup = $sorted | Where-Object { $_.Type -eq "Cloud Flow" -and $_.LastRun -eq "Setup required" }
        if ($flowsNeedingSetup -and $flowsNeedingSetup.Count -gt 0 -and -not $script:SyncHash.FlowApiToken) {
            $setupPrompt = [System.Windows.MessageBox]::Show(
                "Flow run history requires Power Automate API authentication.`n`nWould you like to set it up now to see flow run statistics?",
                "Power Automate API Setup",
                "YesNo",
                "Question"
            )

            if ($setupPrompt -eq "Yes") {
                $setupResult = Show-PowerAutomateApiSetup
                if ($setupResult) {
                    # User authenticated successfully - offer to rescan
                    $rescan = [System.Windows.MessageBox]::Show(
                        "Authentication successful! Would you like to rescan flows to load run history?",
                        "Rescan Flows",
                        "YesNo",
                        "Question"
                    )
                    if ($rescan -eq "Yes") {
                        # Trigger scan button click
                        $btnScan.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
                    }
                }
            }
        }

    } catch {
        $errorMsg = $_.Exception.Message
        $errorStack = $_.ScriptStackTrace
        $txtStatus.Text = "Scan failed: $errorMsg"
        Show-ErrorWithCopyOption -Title "Scan Error" -Message "Scan failed: $errorMsg" -Details $errorStack
    } finally {
        $btnScan.IsEnabled = $true
    }
})

# Export CSV button click handler
$btnExportCsv.Add_Click({
    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveDialog.FileName = "PowerAppsDiscovery_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

    if ($saveDialog.ShowDialog() -eq $true) {
        $script:SyncHash.Results | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
        $txtStatus.Text = "Results exported to: $($saveDialog.FileName)"
        [System.Windows.MessageBox]::Show("Results exported successfully!", "Export Complete", "OK", "Information")
    }
})

# Initialize snapshot directory
$snapshotDir = "$env:USERPROFILE\.powersearch\snapshots"
if (-not (Test-Path $snapshotDir)) {
    New-Item -ItemType Directory -Path $snapshotDir -Force | Out-Null
}

# Save Snapshot button click handler
$btnSaveSnapshot.Add_Click({
    if ($script:SyncHash.Results.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No results to save. Please run a scan first.", "No Results", "OK", "Warning")
        return
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $snapshotPath = "$snapshotDir\Snapshot_$timestamp.json"

    # Create snapshot object with metadata
    $snapshot = @{
        Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Environment = $script:txtEnvironmentUrl.Text
        Keywords = $txtKeywords.Text
        ItemCount = $script:SyncHash.Results.Count
        Items = $script:SyncHash.Results
    }

    $snapshot | ConvertTo-Json -Depth 10 | Out-File -FilePath $snapshotPath -Encoding UTF8
    $txtStatus.Text = "Snapshot saved: Snapshot_$timestamp.json"
    [System.Windows.MessageBox]::Show("Snapshot saved successfully!`n`nFile: Snapshot_$timestamp.json`nItems: $($script:SyncHash.Results.Count)", "Snapshot Saved", "OK", "Information")
})

# Compare Snapshots button click handler
$btnCompareSnapshots.Add_Click({
    # Get all snapshots
    $snapshots = Get-ChildItem -Path $snapshotDir -Filter "Snapshot_*.json" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending

    if ($snapshots.Count -lt 2) {
        [System.Windows.MessageBox]::Show("You need at least 2 snapshots to compare. Current snapshots: $($snapshots.Count)", "Insufficient Snapshots", "OK", "Warning")
        return
    }

    # Create snapshot comparison window
    $compareXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Compare Snapshots"
        Height="700" Width="1000"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5">
    <Window.Resources>
        <SolidColorBrush x:Key="PrimaryPurple" Color="#3C1053"/>
        <SolidColorBrush x:Key="SecondaryOrange" Color="#E87722"/>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <TextBlock Grid.Row="0" Text="Compare Snapshots" FontSize="18" FontWeight="Bold"
                   Foreground="{StaticResource PrimaryPurple}" Margin="0,0,0,15"/>

        <!-- Snapshot Selection -->
        <Grid Grid.Row="1" Margin="0,0,0,15">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="20"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="20"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <StackPanel Grid.Column="0">
                <TextBlock Text="Baseline Snapshot (Older):" FontWeight="SemiBold" Margin="0,0,0,5"/>
                <ComboBox x:Name="cmbBaseline" Height="30" DisplayMemberPath="DisplayName"/>
            </StackPanel>

            <StackPanel Grid.Column="2">
                <TextBlock Text="Current Snapshot (Newer):" FontWeight="SemiBold" Margin="0,0,0,5"/>
                <ComboBox x:Name="cmbCurrent" Height="30" DisplayMemberPath="DisplayName"/>
            </StackPanel>

            <Button x:Name="btnCompare" Grid.Column="4" Content="Compare"
                    Background="{StaticResource SecondaryOrange}" Foreground="White"
                    FontWeight="SemiBold" Padding="20,8" VerticalAlignment="Bottom"/>
        </Grid>

        <!-- Comparison Results -->
        <TabControl Grid.Row="2" x:Name="tabResults">
            <TabItem Header="Summary">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <TextBlock x:Name="txtSummary" Padding="10" FontFamily="Consolas" FontSize="12" TextWrapping="Wrap"/>
                </ScrollViewer>
            </TabItem>
            <TabItem Header="New Items">
                <DataGrid x:Name="dgNewItems" AutoGenerateColumns="True" IsReadOnly="True"/>
            </TabItem>
            <TabItem Header="Removed Items">
                <DataGrid x:Name="dgRemovedItems" AutoGenerateColumns="True" IsReadOnly="True"/>
            </TabItem>
            <TabItem Header="Modified Items">
                <DataGrid x:Name="dgModifiedItems" AutoGenerateColumns="True" IsReadOnly="True"/>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@

    $compareReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($compareXaml))
    $compareWindow = [Windows.Markup.XamlReader]::Load($compareReader)
    $compareReader.Close()

    $cmbBaseline = $compareWindow.FindName("cmbBaseline")
    $cmbCurrent = $compareWindow.FindName("cmbCurrent")
    $btnCompare = $compareWindow.FindName("btnCompare")
    $txtSummary = $compareWindow.FindName("txtSummary")
    $dgNewItems = $compareWindow.FindName("dgNewItems")
    $dgRemovedItems = $compareWindow.FindName("dgRemovedItems")
    $dgModifiedItems = $compareWindow.FindName("dgModifiedItems")

    # Populate combo boxes
    $snapshotList = $snapshots | ForEach-Object {
        $content = Get-Content $_.FullName -Raw | ConvertFrom-Json
        [PSCustomObject]@{
            DisplayName = "$($_.BaseName) - $($content.Timestamp) ($($content.ItemCount) items)"
            Path = $_.FullName
            Data = $content
        }
    }

    $cmbBaseline.ItemsSource = $snapshotList
    $cmbCurrent.ItemsSource = $snapshotList

    if ($snapshotList.Count -ge 2) {
        $cmbBaseline.SelectedIndex = 1  # Second oldest
        $cmbCurrent.SelectedIndex = 0   # Most recent
    }

    # Compare button click
    $btnCompare.Add_Click({
        $baseline = $cmbBaseline.SelectedItem
        $current = $cmbCurrent.SelectedItem

        if (-not $baseline -or -not $current) {
            [System.Windows.MessageBox]::Show("Please select both snapshots to compare.", "Selection Required", "OK", "Warning")
            return
        }

        if ($baseline.Path -eq $current.Path) {
            [System.Windows.MessageBox]::Show("Please select different snapshots to compare.", "Same Snapshot", "OK", "Warning")
            return
        }

        $baselineItems = $baseline.Data.Items
        $currentItems = $current.Data.Items

        # Create lookup dictionaries by ID
        $baselineDict = @{}
        foreach ($item in $baselineItems) {
            $baselineDict[$item.Id] = $item
        }

        $currentDict = @{}
        foreach ($item in $currentItems) {
            $currentDict[$item.Id] = $item
        }

        # Find new items (in current but not in baseline)
        $newItems = $currentItems | Where-Object { -not $baselineDict.ContainsKey($_.Id) }

        # Find removed items (in baseline but not in current)
        $removedItems = $baselineItems | Where-Object { -not $currentDict.ContainsKey($_.Id) }

        # Find modified items (in both but with differences)
        $modifiedItems = @()
        foreach ($item in $currentItems) {
            if ($baselineDict.ContainsKey($item.Id)) {
                $baseItem = $baselineDict[$item.Id]
                $changes = @()

                # Check for state changes
                if ($item.State -ne $baseItem.State) {
                    $changes += "State: $($baseItem.State) -> $($item.State)"
                }

                # Check for error count changes
                if ($item.ErrorCount -ne $baseItem.ErrorCount) {
                    $changes += "Errors: $($baseItem.ErrorCount) -> $($item.ErrorCount)"
                }

                # Check for solution changes
                if ($item.Solution -ne $baseItem.Solution) {
                    $changes += "Solution: $($baseItem.Solution) -> $($item.Solution)"
                }

                if ($changes.Count -gt 0) {
                    $modifiedItems += [PSCustomObject]@{
                        Type = $item.Type
                        Name = $item.Name
                        Id = $item.Id
                        Changes = ($changes -join "; ")
                    }
                }
            }
        }

        # Update UI
        $dgNewItems.ItemsSource = $newItems
        $dgRemovedItems.ItemsSource = $removedItems
        $dgModifiedItems.ItemsSource = $modifiedItems

        # Generate summary
        $summary = @"
SNAPSHOT COMPARISON SUMMARY
===========================

Baseline: $($baseline.Data.Timestamp)
Environment: $($baseline.Data.Environment)
Items: $($baselineItems.Count)

Current: $($current.Data.Timestamp)
Environment: $($current.Data.Environment)
Items: $($currentItems.Count)

CHANGES DETECTED
================
New Items: $($newItems.Count)
Removed Items: $($removedItems.Count)
Modified Items: $($modifiedItems.Count)
Unchanged Items: $($currentItems.Count - $newItems.Count - $modifiedItems.Count)

NET CHANGE: $(if ($currentItems.Count - $baselineItems.Count -gt 0) { "+$($currentItems.Count - $baselineItems.Count)" } else { "$($currentItems.Count - $baselineItems.Count)" }) items

"@

        if ($newItems.Count -gt 0) {
            $summary += "`nNEW ITEMS BY TYPE:`n"
            $newItems | Group-Object Type | ForEach-Object {
                $summary += "  $($_.Name): $($_.Count)`n"
            }
        }

        if ($removedItems.Count -gt 0) {
            $summary += "`nREMOVED ITEMS BY TYPE:`n"
            $removedItems | Group-Object Type | ForEach-Object {
                $summary += "  $($_.Name): $($_.Count)`n"
            }
        }

        if ($modifiedItems.Count -gt 0) {
            $summary += "`nMODIFIED ITEMS BY TYPE:`n"
            $modifiedItems | Group-Object Type | ForEach-Object {
                $summary += "  $($_.Name): $($_.Count)`n"
            }
        }

        $txtSummary.Text = $summary
    })

    $compareWindow.ShowDialog() | Out-Null
})

# View Dependencies button click handler
$btnViewDependencies.Add_Click({
    if ($script:SyncHash.Results.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No results to analyze. Please run a scan first.", "No Results", "OK", "Warning")
        return
    }

    # Helper function to extract dependencies from flow/workflow JSON
    function Get-ItemDependencies {
        param(
            [string]$ItemId,
            [string]$ItemType,
            $Connection,
            $AllItems
        )

        $dependencies = @()

        try {
            # Only process flows and workflows that have JSON definitions
            if ($ItemType -eq "Cloud Flow") {
                # Get flow definition from Dataverse
                $flowFetch = @"
<fetch top='1'>
  <entity name='workflow'>
    <attribute name='clientdata' />
    <filter>
      <condition attribute='workflowid' operator='eq' value='$ItemId' />
    </filter>
  </entity>
</fetch>
"@
                $flowData = Get-CrmRecordsByFetch -conn $Connection -Fetch $flowFetch -ErrorAction SilentlyContinue

                if ($flowData -and $flowData.CrmRecords -and $flowData.CrmRecords[0].clientdata) {
                    $flowJson = $flowData.CrmRecords[0].clientdata | ConvertFrom-Json -ErrorAction SilentlyContinue

                    if ($flowJson) {
                        # Extract flow steps/actions
                        if ($flowJson.properties.definition.actions) {
                            $actions = $flowJson.properties.definition.actions
                            $actionNames = $actions.PSObject.Properties.Name

                            foreach ($actionName in $actionNames) {
                                $action = $actions.$actionName
                                $actionType = if ($action.type) { $action.type } else { "Unknown" }

                                # Get a friendly description
                                $description = $actionName
                                if ($action.metadata.operationMetadataId) {
                                    $description = "$actionName"
                                } elseif ($action.inputs) {
                                    # Try to get more context from inputs
                                    if ($action.inputs.host.operationId) {
                                        $description = "$actionName ($($action.inputs.host.operationId))"
                                    }
                                }

                                $dependencies += [PSCustomObject]@{
                                    Type = "Flow Steps"
                                    TargetType = $actionType
                                    TargetName = $description
                                    TargetId = $actionName
                                }
                            }
                        } else {
                            # Debug: Add a note if no actions found
                            $dependencies += [PSCustomObject]@{
                                Type = "Debug Info"
                                TargetType = "Info"
                                TargetName = "No actions found in flow definition"
                                TargetId = ""
                            }
                        }

                        # Find trigger information
                        if ($flowJson.properties.definition.triggers) {
                            $triggers = $flowJson.properties.definition.triggers
                            $triggerNames = $triggers.PSObject.Properties.Name

                            foreach ($triggerName in $triggerNames) {
                                $trigger = $triggers.$triggerName
                                $triggerType = if ($trigger.type) { $trigger.type } else { "Unknown" }

                                $dependencies += [PSCustomObject]@{
                                    Type = "Trigger"
                                    TargetType = $triggerType
                                    TargetName = $triggerName
                                    TargetId = $triggerName
                                }
                            }
                        }

                        # Look for child flow references in the definition
                        $flowJsonString = $flowData.CrmRecords[0].clientdata

                        # Find workflow references (child flows)
                        if ($flowJsonString -match '"workflow":\s*"([^"]+)"') {
                            $childFlowIds = [regex]::Matches($flowJsonString, '"workflow":\s*"([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
                            foreach ($childId in $childFlowIds | Select-Object -Unique) {
                                $childItem = $AllItems | Where-Object { $_.Id -eq $childId }
                                if ($childItem) {
                                    $dependencies += [PSCustomObject]@{
                                        Type = "Calls Child Flow"
                                        TargetType = $childItem.Type
                                        TargetName = $childItem.Name
                                        TargetId = $childId
                                    }
                                }
                            }
                        }

                        # Find connector references
                        if ($flowJson.properties.connectionReferences) {
                            $connRefs = $flowJson.properties.connectionReferences
                            $connRefNames = $connRefs.PSObject.Properties.Name

                            foreach ($connRef in $connRefNames) {
                                $connData = $connRefs.$connRef
                                if ($connData.displayName) {
                                    $dependencies += [PSCustomObject]@{
                                        Type = "Uses Connector"
                                        TargetType = "Connector"
                                        TargetName = $connData.displayName
                                        TargetId = $connRef
                                    }
                                }
                            }
                        }
                    }
                }
            } elseif ($ItemType -eq "Classic Workflow" -or $ItemType -eq "Custom Action" -or $ItemType -eq "Business Process Flow") {
                # Get workflow XAML
                $workflowFetch = @"
<fetch top='1'>
  <entity name='workflow'>
    <attribute name='xaml' />
    <filter>
      <condition attribute='workflowid' operator='eq' value='$ItemId' />
    </filter>
  </entity>
</fetch>
"@
                $workflowData = Get-CrmRecordsByFetch -conn $Connection -Fetch $workflowFetch -ErrorAction SilentlyContinue

                if ($workflowData -and $workflowData.CrmRecords -and $workflowData.CrmRecords[0].xaml) {
                    $xaml = $workflowData.CrmRecords[0].xaml

                    # Find child workflow references in XAML
                    $childWorkflows = [regex]::Matches($xaml, 'ChildWorkflowId="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
                    foreach ($childId in $childWorkflows | Select-Object -Unique) {
                        $childItem = $AllItems | Where-Object { $_.Id -eq $childId }
                        if ($childItem) {
                            $dependencies += [PSCustomObject]@{
                                Type = "Calls"
                                TargetType = $childItem.Type
                                TargetName = $childItem.Name
                                TargetId = $childId
                            }
                        }
                    }
                }
            } elseif ($ItemType -eq "Canvas App") {
                # Canvas apps already analyzed during scan - look at the Reasons field
                $item = $AllItems | Where-Object { $_.Id -eq $ItemId }
                if ($item -and $item.Reasons) {
                    # Parse reasons for connector and flow references
                    $reasons = $item.Reasons -split "; "
                    foreach ($reason in $reasons) {
                        if ($reason -match "Uses connector: (.+)") {
                            $dependencies += [PSCustomObject]@{
                                Type = "Uses Connector"
                                TargetType = "Connector"
                                TargetName = $matches[1]
                                TargetId = ""
                            }
                        } elseif ($reason -match "Contains flow: (.+)") {
                            $dependencies += [PSCustomObject]@{
                                Type = "Contains Flow"
                                TargetType = "Embedded Flow"
                                TargetName = $matches[1]
                                TargetId = ""
                            }
                        }
                    }
                }
            }
        } catch {
            # Ignore errors in dependency extraction
        }

        return $dependencies
    }

    # Create dependency graph window
    $dependencyXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Dependency Graph"
        Height="700" Width="1000"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5">
    <Window.Resources>
        <SolidColorBrush x:Key="PrimaryPurple" Color="#3C1053"/>
        <SolidColorBrush x:Key="SecondaryOrange" Color="#E87722"/>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <TextBlock Grid.Row="0" Text="Dependency Graph" FontSize="18" FontWeight="Bold"
                   Foreground="{StaticResource PrimaryPurple}" Margin="0,0,0,15"/>

        <!-- Item Selection -->
        <Grid Grid.Row="1" Margin="0,0,0,15">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="20"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <StackPanel Grid.Column="0">
                <TextBlock Text="Select item to analyze:" FontWeight="SemiBold" Margin="0,0,0,5"/>
                <ComboBox x:Name="cmbItem" Height="30" DisplayMemberPath="DisplayName"/>
            </StackPanel>

            <Button x:Name="btnAnalyze" Grid.Column="2" Content="Analyze Dependencies"
                    Background="{StaticResource SecondaryOrange}" Foreground="White"
                    FontWeight="SemiBold" Padding="20,8" VerticalAlignment="Bottom"/>
        </Grid>

        <!-- Dependency Tree -->
        <Border Grid.Row="2" BorderBrush="#CCCCCC" BorderThickness="1" Background="White">
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                <TreeView x:Name="tvDependencies" Padding="10">
                    <TreeView.Resources>
                        <Style TargetType="TreeViewItem">
                            <Setter Property="IsExpanded" Value="True"/>
                            <Setter Property="FontSize" Value="13"/>
                            <Setter Property="Padding" Value="5"/>
                        </Style>
                    </TreeView.Resources>
                </TreeView>
            </ScrollViewer>
        </Border>
    </Grid>
</Window>
"@

    $depReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($dependencyXaml))
    $depWindow = [Windows.Markup.XamlReader]::Load($depReader)
    $depReader.Close()

    $cmbItem = $depWindow.FindName("cmbItem")
    $btnAnalyze = $depWindow.FindName("btnAnalyze")
    $tvDependencies = $depWindow.FindName("tvDependencies")

    # Populate combo box with all items
    $itemList = $script:SyncHash.Results | ForEach-Object {
        [PSCustomObject]@{
            DisplayName = "$($_.Type): $($_.Name)"
            Item = $_
        }
    }

    $cmbItem.ItemsSource = $itemList
    if ($itemList.Count -gt 0) {
        $cmbItem.SelectedIndex = 0
    }

    # Analyze button click
    $btnAnalyze.Add_Click({
        $selectedItem = $cmbItem.SelectedItem
        if (-not $selectedItem) {
            [System.Windows.MessageBox]::Show("Please select an item to analyze.", "Selection Required", "OK", "Warning")
            return
        }

        $item = $selectedItem.Item
        $tvDependencies.Items.Clear()

        # Create root node
        $rootNode = New-Object System.Windows.Controls.TreeViewItem
        $rootNode.Header = "$($item.Type): $($item.Name)"
        $rootNode.FontWeight = "Bold"
        $rootNode.Foreground = [System.Windows.Media.Brushes]::DarkBlue

        # Get dependencies
        try {
            $dependencies = Get-ItemDependencies -ItemId $item.Id -ItemType $item.Type -Connection $script:SyncHash.Connection -AllItems $script:SyncHash.Results

            if ($dependencies.Count -eq 0) {
                $noDepNode = New-Object System.Windows.Controls.TreeViewItem
                $noDepNode.Header = "No dependencies found"
                $noDepNode.Foreground = [System.Windows.Media.Brushes]::Gray
                $rootNode.Items.Add($noDepNode)
            } else {
                # Group dependencies by type
                $groupedDeps = $dependencies | Group-Object Type

                foreach ($group in $groupedDeps) {
                    $groupNode = New-Object System.Windows.Controls.TreeViewItem
                    $groupNode.Header = "$($group.Name) ($($group.Count))"
                    $groupNode.FontWeight = "SemiBold"

                    foreach ($dep in $group.Group) {
                        $depNode = New-Object System.Windows.Controls.TreeViewItem
                        $depNode.Header = "$($dep.TargetType): $($dep.TargetName)"

                        # Color code by type
                        if ($dep.TargetType -eq "Connector") {
                            $depNode.Foreground = [System.Windows.Media.Brushes]::DarkGreen
                        } elseif ($dep.Type -match "Calls") {
                            $depNode.Foreground = [System.Windows.Media.Brushes]::DarkOrange
                        } elseif ($dep.Type -eq "Trigger") {
                            $depNode.Foreground = [System.Windows.Media.Brushes]::Purple
                        } elseif ($dep.Type -eq "Flow Steps") {
                            $depNode.Foreground = [System.Windows.Media.Brushes]::DarkSlateBlue
                        } else {
                            $depNode.Foreground = [System.Windows.Media.Brushes]::Black
                        }

                        $groupNode.Items.Add($depNode)
                    }

                    $rootNode.Items.Add($groupNode)
                }
            }
        } catch {
            $errorNode = New-Object System.Windows.Controls.TreeViewItem
            $errorNode.Header = "Error analyzing dependencies: $($_.Exception.Message)"
            $errorNode.Foreground = [System.Windows.Media.Brushes]::Red
            $rootNode.Items.Add($errorNode)
        }

        $tvDependencies.Items.Add($rootNode)
    })

    # Auto-analyze first item
    if ($itemList.Count -gt 0) {
        $btnAnalyze.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    }

    $depWindow.ShowDialog() | Out-Null
})

# Clear results button click handler
$btnClearResults.Add_Click({
    $dgResults.ItemsSource = $null
    $script:SyncHash.Results = @()
    $txtResultCount.Text = " (0 items)"
    $txtStatus.Text = "Results cleared"
    $btnAISummary.IsEnabled = $false
    $btnExportCsv.IsEnabled = $false
    $btnSaveSnapshot.IsEnabled = $false
    $btnViewDependencies.IsEnabled = $false
    $btnClearResults.IsEnabled = $false
    $btnFlowHistory.IsEnabled = $false
})

# Window closing event handler - clean up cached token
$window.Add_Closing({
    param($windowSender, $e)

    try {
        $tokenPath = "$env:USERPROFILE\.powersearch\flowapi_token.json"

        if (Test-Path $tokenPath) {
            Remove-Item $tokenPath -Force
            Write-Host "Cached Flow API token deleted on exit" -ForegroundColor Cyan
        }

        # Clear in-memory token
        if ($script:SyncHash) {
            $script:SyncHash.FlowApiToken = $null
            $script:SyncHash.FlowApiTokenExpiry = $null
        }
    } catch {
        Write-Host "Error cleaning up token: $($_.Exception.Message)" -ForegroundColor Yellow
    }
})

# Try to authenticate on startup
Write-Host "Checking for authentication credentials..." -ForegroundColor Cyan

# Try authentication in priority order:
# 1. Saved client configurations (new multi-client approach)
# 2. Old single app config (backward compatibility)
# 3. Cached device code token (fallback)

$authenticated = $false

# First, check for saved client configurations
$savedClients = Get-SavedClientConfigs
if ($savedClients.Count -gt 0) {
    # If there are saved clients, try to authenticate with the first one (most recently used)
    Write-Host "Found $($savedClients.Count) saved client configuration(s)" -ForegroundColor Cyan
    Write-Host "Attempting to authenticate with '$($savedClients[0].ClientName)'..." -ForegroundColor Cyan

    $clientConfig = Load-ClientConfig -FileName $savedClients[0].FileName
    if ($clientConfig) {
        $token = Get-TokenWithClientCredentials -TenantId $clientConfig.TenantId -ClientId $clientConfig.ClientId -ClientSecret $clientConfig.ClientSecret
        if ($token) {
            $script:SyncHash.FlowApiToken = $token
            Write-Host "Successfully authenticated as '$($clientConfig.ClientName)'!" -ForegroundColor Green
            $authenticated = $true
        }
    }
}

# Fallback to old single app registration config (backward compatibility)
if (-not $authenticated) {
    $appConfig = Load-AppConfig
    if ($appConfig) {
        Write-Host "Found legacy app registration config, attempting to authenticate..." -ForegroundColor Cyan
        $token = Get-TokenWithClientCredentials -TenantId $appConfig.TenantId -ClientId $appConfig.ClientId -ClientSecret $appConfig.ClientSecret

        if ($token) {
            $script:SyncHash.FlowApiToken = $token
            Write-Host "Successfully authenticated using legacy app registration!" -ForegroundColor Green
            $authenticated = $true
        }
    }
}

# Final fallback: try cached device code token
if (-not $authenticated) {
    Write-Host "No app registration found, checking for cached device code token..." -ForegroundColor Cyan
    $cachedToken = Load-FlowApiToken
    if ($cachedToken) {
        $script:SyncHash.FlowApiToken = $cachedToken
        Write-Host "Flow API token loaded from cache" -ForegroundColor Green
        $authenticated = $true
    }
}

if (-not $authenticated) {
    Write-Host "No authentication found. Click 'Manage Credentials' to set up authentication." -ForegroundColor Yellow
}

# Update Flow API token status
Update-FlowApiTokenStatus

# Show the window
Write-Host "About to show window..." -ForegroundColor Cyan
try {
    $window.ShowDialog() | Out-Null
    Write-Host "Window closed." -ForegroundColor Cyan
} catch {
    Write-Host "ERROR showing window: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
