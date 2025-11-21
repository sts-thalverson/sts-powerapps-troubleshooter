#requires -Version 5.1
# PowerApps Troubleshooter GUI Tool
# Select Technology Branded - Purple (#3C1053) and Orange (#E87722)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

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
        Title="PowerApps Troubleshooter - Select Technology"
        Height="650" Width="950"
        MinHeight="500" MinWidth="700"
        WindowStartupLocation="CenterScreen"
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
                    <TextBlock Text="PowerApps Troubleshooter"
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
                                     Text="https://yourorg.crm.dynamics.com"
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
                                <CheckBox x:Name="chkCustomActions" Content="Custom Actions" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkServiceEndpoints" Content="Service Endpoints" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkSLAs" Content="SLAs" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                                <CheckBox x:Name="chkBPFs" Content="Business Process Flows" IsChecked="True" Style="{StaticResource ScanCheckBox}"/>
                            </StackPanel>
                        </Grid>
                    </StackPanel>

                    <!-- Right: Action Buttons -->
                    <StackPanel Grid.Column="2" VerticalAlignment="Center">
                        <Button x:Name="btnConnect" Content="Connect" Style="{StaticResource SecondaryButton}" Margin="0,0,0,8" Width="100"/>
                        <Button x:Name="btnScan" Content="Start Scan" Style="{StaticResource PrimaryButton}" IsEnabled="False" Width="100"/>
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
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="Discovery Results" FontWeight="SemiBold" FontSize="13" Foreground="{StaticResource PrimaryPurple}"/>
                    <TextBlock x:Name="txtResultCount" Text=" (0 items)" FontSize="11" Foreground="#666666" VerticalAlignment="Bottom" Margin="5,0,0,1"/>
                </StackPanel>

                <Button x:Name="btnExportCsv" Grid.Column="1" Content="Export CSV" Style="{StaticResource SecondaryButton}" IsEnabled="False" Margin="0,0,8,0"/>
                <Button x:Name="btnClearResults" Grid.Column="2" Content="Clear" Style="{StaticResource SecondaryButton}" IsEnabled="False"/>
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
                                <Button Content="Open" Tag="{Binding Url}"
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
                    <DataGridTextColumn Header="State" Binding="{Binding State}" Width="55">
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="VerticalAlignment" Value="Center"/>
                                <Setter Property="Padding" Value="6,0"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding State}" Value="On">
                                        <Setter Property="Foreground" Value="Green"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding State}" Value="Off">
                                        <Setter Property="Foreground" Value="Red"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding State}" Value="Active">
                                        <Setter Property="Foreground" Value="Green"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding State}" Value="Inactive">
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
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <Ellipse x:Name="statusIndicator" Width="8" Height="8" Fill="#FF6666" Margin="0,0,6,0"/>
                    <TextBlock x:Name="txtConnectionStatus" Text="Not Connected" Foreground="White" FontSize="10"/>
                </StackPanel>

                <TextBlock x:Name="txtStatus" Grid.Column="1" Text="Ready" Foreground="#CCBBDD" FontSize="10" Margin="15,0" TextAlignment="Center"/>

                <ProgressBar x:Name="progressBar" Grid.Column="2" Width="150" Height="5" IsIndeterminate="False" Value="0"
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
$txtEnvironmentUrl = $window.FindName("txtEnvironmentUrl")
$txtKeywords = $window.FindName("txtKeywords")
$dpBreakageDate = $window.FindName("dpBreakageDate")
$btnConnect = $window.FindName("btnConnect")
$btnScan = $window.FindName("btnScan")
$btnExportCsv = $window.FindName("btnExportCsv")
$btnClearResults = $window.FindName("btnClearResults")
$dgResults = $window.FindName("dgResults")
$txtResultCount = $window.FindName("txtResultCount")
$txtStatus = $window.FindName("txtStatus")
$txtConnectionStatus = $window.FindName("txtConnectionStatus")
$statusIndicator = $window.FindName("statusIndicator")
$progressBar = $window.FindName("progressBar")

# Checkboxes
$chkCloudFlows = $window.FindName("chkCloudFlows")
$chkClassicWorkflows = $window.FindName("chkClassicWorkflows")
$chkPlugins = $window.FindName("chkPlugins")
$chkWebResources = $window.FindName("chkWebResources")
$chkCustomActions = $window.FindName("chkCustomActions")
$chkServiceEndpoints = $window.FindName("chkServiceEndpoints")
$chkSLAs = $window.FindName("chkSLAs")
$chkBPFs = $window.FindName("chkBPFs")

# Global variables - use synchronized hashtable for cross-thread access
$script:SyncHash = [hashtable]::Synchronized(@{
    Connection = $null
    Results = @()
    EnvironmentId = $null
})

# Add click handler for Open buttons in the DataGrid
$dgResults.AddHandler(
    [System.Windows.Controls.Button]::ClickEvent,
    [System.Windows.RoutedEventHandler]{
        param($sender, $e)
        if ($e.OriginalSource -is [System.Windows.Controls.Button]) {
            $button = $e.OriginalSource
            $url = $button.Tag
            if ($url -and $url -ne "") {
                Start-Process $url
            }
        }
    }
)

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
        [bool]$ScanCustomActions,
        [bool]$ScanServiceEndpoints,
        [bool]$ScanSLAs,
        [bool]$ScanBPFs,
        [string]$OrgUrl
    )

    $conn = $Connection
    $discoveredItems = [System.Collections.ArrayList]::new()
    $totalSteps = 8
    $currentStep = 0

    # Debug: Check if connection is valid
    if (-not $conn -or -not $conn.IsReady) {
        $txtStatus.Text = "Error: Connection is not ready"
        return $discoveredItems
    }

    # Debug: Show keywords being used
    $keywordList = $Keywords -join ", "
    $txtStatus.Text = "Scanning with keywords: $keywordList"
    [System.Windows.Forms.Application]::DoEvents()

    # Helper function to generate URL for an item
    function Get-ItemUrl {
        param(
            [string]$Type,
            [string]$Id,
            [string]$OrgUrl
        )

        if (-not $Id -or -not $OrgUrl) { return "" }

        $baseUrl = "https://$OrgUrl"

        switch -Wildcard ($Type) {
            "Cloud Flow" {
                # Power Automate flow URL
                return "https://make.powerautomate.com/environments/Default-*/flows/$Id/details"
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
            default {
                return ""
            }
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
    <attribute name='description' />
    <attribute name='statecode' />
    <attribute name='statuscode' />
    <attribute name='category' />
    <attribute name='primaryentity' />
    <attribute name='createdon' />
    <attribute name='modifiedon' />
    <attribute name='clientdata' />
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
                    if ($flowName -like "*$keyword*") {
                        $matchScore += 2
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                if ($flow.description) {
                    $flowDesc = $flow.description.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($flowDesc -like "*$keyword*") {
                            $matchScore += 1
                            $matchReasons += "Description contains '$keyword'"
                        }
                    }
                }

                if ($flow.clientdata) {
                    $clientData = $flow.clientdata.ToLower()

                    if ($clientData -match 'send.*email|office365|smtp|mail\.send') {
                        $matchScore += 3
                        $matchReasons += "Contains email sending action"
                    }

                    if ($clientData -match 'msdyn_project|project') {
                        $matchScore += 2
                        $matchReasons += "References project entity"
                    }

                    if ($clientData -match 'approval|approve|reject') {
                        $matchScore += 2
                        $matchReasons += "Contains approval logic"
                    }

                    foreach ($keyword in $Keywords) {
                        if ($clientData -like "*$keyword*") {
                            $matchScore += 1
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
                    $state = if ($flow.statecode -eq 1) { "On" } else { "Off" }
                    $itemType = "Cloud Flow"
                    $itemId = $flow.workflowid

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $flow.name
                        Id = $itemId
                        State = $state
                        ModifiedOn = if ($flow.modifiedon) { $flow.modifiedon.ToString() } else { "" }
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = $flow.primaryentity
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl
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
                    if ($wfName -like "*$keyword*") {
                        $matchScore += 2
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                if ($wf.xaml) {
                    $xaml = $wf.xaml.ToLower()

                    if ($xaml -match 'sendemail|createemail') {
                        $matchScore += 3
                        $matchReasons += "Contains email action"
                    }

                    if ($xaml -match 'msdyn_project') {
                        $matchScore += 2
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
                    $state = if ($wf.statecode -eq 1) { "On" } else { "Off" }
                    $itemType = "Classic Workflow"
                    $itemId = $wf.workflowid

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $wf.name
                        Id = $itemId
                        State = $state
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

                $entityCode = $plugin.'filter.primaryobjecttypecode'
                if ($entityCode -eq 10054) {
                    $matchScore += 3
                    $matchReasons += "Registered on Project entity"
                }

                if ($plugin.'step.name') {
                    $stepName = $plugin.'step.name'.ToLower()
                    foreach ($keyword in $Keywords) {
                        if ($stepName -like "*$keyword*") {
                            $matchScore += 2
                            $matchReasons += "Step name contains '$keyword'"
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

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = "$($plugin.typename) - $($plugin.'step.name')"
                        Id = $itemId
                        State = $state
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
                    if ($wrName -like "*$keyword*" -or $wrDisplayName -like "*$keyword*") {
                        $matchScore += 2
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

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $wr.displayname
                        Id = $itemId
                        State = "N/A"
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

    # 5. Custom Actions
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
                    if ($actionName -like "*$keyword*") {
                        $matchScore += 2
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                if ($action.xaml) {
                    $xaml = $action.xaml.ToLower()

                    if ($xaml -match 'sendemail|createemail') {
                        $matchScore += 3
                        $matchReasons += "Contains email action"
                    }

                    if ($xaml -match 'msdyn_project') {
                        $matchScore += 2
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
                    $state = if ($action.statecode -eq 1) { "On" } else { "Off" }
                    $itemType = "Custom Action"
                    $itemId = $action.workflowid

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $action.name
                        Id = $itemId
                        State = $state
                        ModifiedOn = if ($action.modifiedon) { $action.modifiedon.ToString() } else { "" }
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = ""
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning Custom Actions: $($_.Exception.Message)"
        }
    }

    # 6. Service Endpoints
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
                    $matchScore += 3
                    $matchReasons += "Registered on Project entity"
                }

                $endpointName = if ($endpoint.name) { $endpoint.name.ToLower() } else { "" }
                foreach ($keyword in $Keywords) {
                    if ($endpointName -like "*$keyword*") {
                        $matchScore += 2
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                if ($endpoint.url) {
                    $url = $endpoint.url.ToLower()
                    if ($url -match 'logic.*app|function.*app|servicebus|webhook') {
                        $matchScore += 2
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

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = "$($endpoint.name) - $($endpoint.'step.name')"
                        Id = $itemId
                        State = $state
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

    # 7. SLAs
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
  </entity>
</fetch>
"@

        try {
            $slas = Get-CrmRecordsByFetch -conn $conn -Fetch $slasFetch

            foreach ($sla in $slas.CrmRecords) {
                $matchScore = 0
                $matchReasons = @()

                if ($sla.objecttypecode -eq 10054) {
                    $matchScore += 2
                    $matchReasons += "Applied to Project entity"
                }

                $slaName = if ($sla.name) { $sla.name.ToLower() } else { "" }
                foreach ($keyword in $Keywords) {
                    if ($slaName -like "*$keyword*") {
                        $matchScore += 2
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                if ($matchScore -gt 0) {
                    $state = if ($sla.statecode -eq 0) { "Active" } else { "Inactive" }
                    $itemType = "SLA"
                    $itemId = $sla.slaid

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $sla.name
                        Id = $itemId
                        State = $state
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

    # 8. Business Process Flows
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
                    if ($bpfName -like "*$keyword*") {
                        $matchScore += 2
                        $matchReasons += "Name contains '$keyword'"
                    }
                }

                if ($matchScore -gt 0) {
                    $state = if ($bpf.statecode -eq 1) { "On" } else { "Off" }
                    $itemType = "Business Process Flow"
                    $itemId = $bpf.workflowid

                    [void]$discoveredItems.Add([PSCustomObject]@{
                        Type = $itemType
                        Name = $bpf.name
                        Id = $itemId
                        State = $state
                        ModifiedOn = if ($bpf.modifiedon) { $bpf.modifiedon.ToString() } else { "" }
                        Score = $matchScore
                        Reasons = ($matchReasons -join "; ")
                        PrimaryEntity = ""
                        Url = Get-ItemUrl -Type $itemType -Id $itemId -OrgUrl $OrgUrl
                    })
                }
            }
        } catch {
            $txtStatus.Text = "Error scanning BPFs: $($_.Exception.Message)"
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
            -ScanCustomActions ($chkCustomActions.IsChecked -eq $true) `
            -ScanServiceEndpoints ($chkServiceEndpoints.IsChecked -eq $true) `
            -ScanSLAs ($chkSLAs.IsChecked -eq $true) `
            -ScanBPFs ($chkBPFs.IsChecked -eq $true) `
            -OrgUrl $script:SyncHash.OrgUrl

        # Sort and display results
        $sorted = $results | Sort-Object -Property Score -Descending
        $script:SyncHash.Results = $sorted

        $dgResults.ItemsSource = $sorted
        $txtResultCount.Text = " ($($sorted.Count) items)"
        $txtStatus.Text = "Scan complete - Found $($sorted.Count) matching items"
        $progressBar.Value = 100
        $btnExportCsv.IsEnabled = $sorted.Count -gt 0
        $btnClearResults.IsEnabled = $sorted.Count -gt 0

    } catch {
        $errorMsg = $_.Exception.Message
        $txtStatus.Text = "Scan failed: $errorMsg"
        [System.Windows.MessageBox]::Show("Scan failed: $errorMsg", "Scan Error", "OK", "Error")
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

# Clear results button click handler
$btnClearResults.Add_Click({
    $dgResults.ItemsSource = $null
    $script:SyncHash.Results = @()
    $txtResultCount.Text = " (0 items)"
    $txtStatus.Text = "Results cleared"
    $btnExportCsv.IsEnabled = $false
    $btnClearResults.IsEnabled = $false
})

# Show the window
$window.ShowDialog() | Out-Null
