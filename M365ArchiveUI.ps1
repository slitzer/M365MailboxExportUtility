#requires -version 5.1
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$ErrorActionPreference = "Stop"

# ============================================
# CONFIG
# ============================================
$script:EnginePath = "C:\TRANSFERSCRIPT\M365ArchiveEngine.ps1"

# State
$script:LastWould    = $null
$script:LastExported = $null
$script:LastSkipped  = $null
$script:LastLog      = $null

# ============================================
# THEME
# ============================================
$Themes = @{
  Dark = @{ Bg="#0F1115"; Text="#E6E6E6"; Muted="#A6ABB5"; Panel="#131720"; Card="#171A21"; Border="#2E3550"; Input="#0E1016"; Out="#0B0D12"; Accent="#4CC2FF"; Warn="#FFB020"; Danger="#FF5C7A"; Ok="#37D67A" }
  Light= @{ Bg="#F4F6FA"; Text="#1E2430"; Muted="#5B6578"; Panel="#FFFFFF"; Card="#FFFFFF"; Border="#CBD3E1"; Input="#FFFFFF"; Out="#FFFFFF"; Accent="#0B74FF"; Warn="#B25A00"; Danger="#B00020"; Ok="#0A7A35" }
}
$script:Theme = "Light"  # default light

function New-Brush([string]$hex) {
    $bc = New-Object System.Windows.Media.BrushConverter
    $b = $bc.ConvertFromString($hex)
    if ($b -is [System.Windows.Media.SolidColorBrush]) { $b.Freeze() }
    return $b
}
function ApplyTheme($w, $name) {
    $t = $Themes[$name]
    $script:Theme = $name
    $w.Background = New-Brush $t.Bg
    $w.Foreground = New-Brush $t.Text
    $w.Resources["Bg"] = New-Brush $t.Bg
    $w.Resources["Text"] = New-Brush $t.Text
    $w.Resources["Muted"] = New-Brush $t.Muted
    $w.Resources["Panel"] = New-Brush $t.Panel
    $w.Resources["Card"] = New-Brush $t.Card
    $w.Resources["Border"] = New-Brush $t.Border
    $w.Resources["Input"] = New-Brush $t.Input
    $w.Resources["Out"] = New-Brush $t.Out
    $w.Resources["Accent"] = New-Brush $t.Accent
    $w.Resources["Warn"] = New-Brush $t.Warn
    $w.Resources["Danger"] = New-Brush $t.Danger
    $w.Resources["Ok"] = New-Brush $t.Ok
}

# ============================================
# UI HELPERS
# ============================================
function Select-Folder {
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.ShowNewFolderButton = $true
    if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $fb.SelectedPath }
    return $null
}
function Select-File([string]$filter) {
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = $filter
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $ofd.FileName }
    return $null
}
function AppendOut([string]$s) {
    $txtOutput.AppendText($s + "`r`n")
    $txtOutput.ScrollToEnd()
}

function RefreshLogs {
    $lstLogs.Items.Clear()
    $dir = $txtLogDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($dir)) { return }
    if (!(Test-Path $dir)) { return }

    Get-ChildItem $dir -File -ErrorAction SilentlyContinue |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 300 |
      ForEach-Object { [void]$lstLogs.Items.Add($_.FullName) }
}

function LoadCsvToGrid([string]$path, $grid) {
    if (!(Test-Path $path)) { $grid.ItemsSource = @(); return }
    $grid.ItemsSource = @(Import-Csv $path)
}

function LoadLatestCsvs {
    $dir = $txtLogDir.Text.Trim()
    if (!(Test-Path $dir)) { AppendOut "Logs folder not found."; return }

    $w = Get-ChildItem $dir -Filter "WouldExport-*.csv" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $e = Get-ChildItem $dir -Filter "Exported-*.csv"    -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $s = Get-ChildItem $dir -Filter "Skipped-*.csv"     -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $l = Get-ChildItem $dir -Filter "M365Export-*.log"  -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($w) { $script:LastWould = $w.FullName }
    if ($e) { $script:LastExported = $e.FullName }
    if ($s) { $script:LastSkipped = $s.FullName }
    if ($l) { $script:LastLog = $l.FullName }

    AppendOut "Loaded latest outputs:"
    AppendOut ("  Would:    " + $(if ($script:LastWould) { $script:LastWould } else { "<none>" }))
    AppendOut ("  Exported: " + $(if ($script:LastExported) { $script:LastExported } else { "<none>" }))
    AppendOut ("  Skipped:  " + $(if ($script:LastSkipped) { $script:LastSkipped } else { "<none>" }))
    AppendOut ("  Log:      " + $(if ($script:LastLog) { $script:LastLog } else { "<none>" }))

    # Update summary cards from files if present (no engine parsing needed)
    if ($script:LastWould -and (Test-Path $script:LastWould)) { $cardWould.Text = @((Import-Csv $script:LastWould)).Count }
    if ($script:LastExported -and (Test-Path $script:LastExported)) { $cardExported.Text = @((Import-Csv $script:LastExported)).Count }
    if ($script:LastSkipped -and (Test-Path $script:LastSkipped)) {
        $sk = @((Import-Csv $script:LastSkipped))
        $cardSkipped.Text = $sk.Count
        $cardAmb.Text = @($sk | Where-Object { $_.SkipReason -match "Ambiguous" }).Count
        $cardFailed.Text = @($sk | Where-Object { $_.SkipReason -match "Export failed" }).Count
    }

    RefreshLogs
}

# ============================================
# RUN ENGINE (UPDATED)
# - Launches a VISIBLE PowerShell window so Graph auth prompts show.
# - Non-blocking: UI won't freeze.
# ============================================
function RunEngine([bool]$dry) {
    $txtOutput.Clear()

    if (!(Test-Path $script:EnginePath)) {
        AppendOut "ERROR: Engine not found:"
        AppendOut $script:EnginePath
        return
    }

    $mailbox = $txtMailbox.Text.Trim()
    $root    = $txtRootFolder.Text.Trim()
    $csv     = $txtCsv.Text.Trim()
    $map     = $txtMap.Text.Trim()
    $outRoot = $txtOutRoot.Text.Trim()
    $logDir  = $txtLogDir.Text.Trim()
    $tenant  = $txtTenant.Text.Trim()

    $useDev  = ($chkDevice.IsChecked -eq $true)
    $recurse = ($chkRecurse.IsChecked -eq $true)

    AppendOut ("=== Launching run: " + (Get-Date))
    AppendOut ("MailboxUPN: " + $mailbox)
    AppendOut ("RootFolderPath: " + $root)
    AppendOut ("CsvPath: " + $csv)
    AppendOut ("ClientMapPath: " + $map)
    AppendOut ("OutRoot: " + $outRoot)
    AppendOut ("LogDir: " + $logDir)
    AppendOut ("IncludeSubfolders: " + $recurse)
    AppendOut ("DryRun: " + $dry)
    AppendOut ("TenantId: " + $tenant)
    AppendOut ("UseDeviceCode: " + $useDev)

    if (-not $mailbox -or -not $root -or -not $csv -or -not $outRoot -or -not $logDir) {
        AppendOut "ERROR: Missing required fields (Mailbox UPN, Root Folder Path, Projects CSV, OutRoot, LogDir)."
        return
    }
    if (!(Test-Path $csv)) { AppendOut "ERROR: Projects CSV not found."; return }

    if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    if (!(Test-Path $outRoot)) { New-Item -ItemType Directory -Path $outRoot -Force | Out-Null }

    # Build a properly quoted argument string (handles spaces safely)
$argStr = @()
$argStr += "-NoExit -NoProfile -ExecutionPolicy Bypass"
$argStr += ('-File "{0}"' -f $script:EnginePath)
$argStr += ('-MailboxUPN "{0}"' -f $mailbox)
$argStr += ('-RootFolderPath "{0}"' -f $root)
$argStr += ('-CsvPath "{0}"' -f $csv)
$argStr += ('-OutRoot "{0}"' -f $outRoot)
$argStr += ('-LogDir "{0}"' -f $logDir)

if ($map)    { $argStr += ('-ClientMapPath "{0}"' -f $map) }
if ($tenant) { $argStr += ('-TenantId "{0}"' -f $tenant) }
if ($useDev) { $argStr += "-UseDeviceCode" }
if ($recurse){ $argStr += "-IncludeSubfolders" }
if ($dry)    { $argStr += "-DryRun" }

Start-Process -FilePath "powershell.exe" -ArgumentList ($argStr -join ' ')

    AppendOut ""
    AppendOut "Started export in a separate PowerShell window."
    AppendOut "When it finishes, click: Preview > Load Latest"
    AppendOut ("Logs: " + $logDir)
}

# ============================================
# XAML
# ============================================
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="M365 Mailbox Export Utility" Height="900" Width="1280" WindowStartupLocation="CenterScreen">
  <Window.Resources>
    <SolidColorBrush x:Key="Bg" Color="#F4F6FA"/>
    <SolidColorBrush x:Key="Text" Color="#1E2430"/>
    <SolidColorBrush x:Key="Muted" Color="#5B6578"/>
    <SolidColorBrush x:Key="Panel" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="Card" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="Border" Color="#CBD3E1"/>
    <SolidColorBrush x:Key="Input" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="Out" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="Accent" Color="#0B74FF"/>
    <SolidColorBrush x:Key="Warn" Color="#B25A00"/>
    <SolidColorBrush x:Key="Danger" Color="#B00020"/>
    <SolidColorBrush x:Key="Ok" Color="#0A7A35"/>

    <Style TargetType="Button">
      <Setter Property="Padding" Value="14,8"/>
      <Setter Property="Margin" Value="0,0,10,8"/>
      <Setter Property="Background" Value="{DynamicResource Panel}"/>
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor" Value="Hand"/>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Padding" Value="10,6"/>
      <Setter Property="Margin" Value="0,0,10,8"/>
      <Setter Property="Background" Value="{DynamicResource Input}"/>
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>

    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="{DynamicResource Input}"/>
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="RowBackground" Value="{DynamicResource Input}"/>
      <Setter Property="AlternatingRowBackground" Value="{DynamicResource Panel}"/>
      <Setter Property="HeadersVisibility" Value="All"/>
    </Style>
  </Window.Resources>

  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <DockPanel Grid.Row="0" Margin="0,0,0,12">
      <StackPanel Orientation="Horizontal">
        <TextBlock FontSize="22" FontWeight="Bold" Text="M365 Mailbox Export Utility"/>
        <TextBlock Margin="12,6,0,0" Foreground="{DynamicResource Muted}" Text="Exports mail folders to local disk (EML + attachments)"/>
      </StackPanel>
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
        <TextBlock Foreground="{DynamicResource Muted}" Margin="0,0,8,0" Text="Theme:"/>
        <ToggleButton Name="tglTheme" Width="84" Height="26" Content="Light" />
      </StackPanel>
    </DockPanel>

    <TabControl Grid.Row="1" Name="tabs" Background="{DynamicResource Panel}">
      <TabItem Header=" Run ">
        <Grid Margin="10">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <UniformGrid Grid.Row="0" Columns="5" Margin="0,0,0,12">
            <Border Background="{DynamicResource Card}" CornerRadius="14" Margin="0,0,12,0" Padding="14">
              <StackPanel><TextBlock FontWeight="SemiBold" Text="Would Export" Foreground="{DynamicResource Accent}"/>
                <TextBlock Name="cardWould" FontSize="26" FontWeight="Bold" Text="0" Margin="0,6,0,0"/></StackPanel>
            </Border>
            <Border Background="{DynamicResource Card}" CornerRadius="14" Margin="0,0,12,0" Padding="14">
              <StackPanel><TextBlock FontWeight="SemiBold" Text="Exported" Foreground="{DynamicResource Ok}"/>
                <TextBlock Name="cardExported" FontSize="26" FontWeight="Bold" Text="0" Margin="0,6,0,0"/></StackPanel>
            </Border>
            <Border Background="{DynamicResource Card}" CornerRadius="14" Margin="0,0,12,0" Padding="14">
              <StackPanel><TextBlock FontWeight="SemiBold" Text="Skipped" Foreground="{DynamicResource Warn}"/>
                <TextBlock Name="cardSkipped" FontSize="26" FontWeight="Bold" Text="0" Margin="0,6,0,0"/></StackPanel>
            </Border>
            <Border Background="{DynamicResource Card}" CornerRadius="14" Margin="0,0,12,0" Padding="14">
              <StackPanel><TextBlock FontWeight="SemiBold" Text="Ambiguous" Foreground="{DynamicResource Warn}"/>
                <TextBlock Name="cardAmb" FontSize="26" FontWeight="Bold" Text="0" Margin="0,6,0,0"/></StackPanel>
            </Border>
            <Border Background="{DynamicResource Card}" CornerRadius="14" Padding="14">
              <StackPanel><TextBlock FontWeight="SemiBold" Text="Failed" Foreground="{DynamicResource Danger}"/>
                <TextBlock Name="cardFailed" FontSize="26" FontWeight="Bold" Text="0" Margin="0,6,0,0"/></StackPanel>
            </Border>
          </UniformGrid>

          <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <TextBlock Grid.Row="0" Grid.Column="0" Margin="0,0,10,8" VerticalAlignment="Center" Text="Mailbox UPN:"/>
            <TextBox  Name="txtMailbox" Grid.Row="0" Grid.Column="1" Text="clientdata@yourdomain.com"/>
            <TextBlock Grid.Row="1" Grid.Column="0" Margin="0,0,10,8" VerticalAlignment="Center" Text="Root Folder Path:"/>
            <TextBox  Name="txtRootFolder" Grid.Row="1" Grid.Column="1" Text="ClientDATA Emails"/>

            <TextBlock Grid.Row="2" Grid.Column="0" Margin="0,0,10,8" VerticalAlignment="Center" Text="Projects CSV:"/>
            <TextBox  Name="txtCsv" Grid.Row="2" Grid.Column="1" Text="C:\TRANSFERSCRIPT\ProjectsToArchive.csv"/>
            <Button   Name="btnCsv" Grid.Row="2" Grid.Column="2" Content="Browse"/>

            <TextBlock Grid.Row="3" Grid.Column="0" Margin="0,0,10,8" VerticalAlignment="Center" Text="Client Map CSV:"/>
            <TextBox  Name="txtMap" Grid.Row="3" Grid.Column="1" Text="C:\TRANSFERSCRIPT\ClientFolderMap.csv"/>
            <Button   Name="btnMap" Grid.Row="3" Grid.Column="2" Content="Browse"/>

            <TextBlock Grid.Row="4" Grid.Column="0" Margin="0,0,10,8" VerticalAlignment="Center" Text="Export OutRoot:"/>
            <TextBox  Name="txtOutRoot" Grid.Row="4" Grid.Column="1" Text="C:\Temp\M365Export"/>
            <Button   Name="btnOutRoot" Grid.Row="4" Grid.Column="2" Content="Browse"/>

            <TextBlock Grid.Row="5" Grid.Column="0" Margin="0,0,10,8" VerticalAlignment="Center" Text="Logs Folder:"/>
            <TextBox  Name="txtLogDir" Grid.Row="5" Grid.Column="1" Text="C:\TRANSFERSCRIPT\Logs"/>
            <Button   Name="btnLogDir" Grid.Row="5" Grid.Column="2" Content="Browse"/>

            <TextBlock Grid.Row="6" Grid.Column="0" Margin="0,0,10,8" VerticalAlignment="Center" Text="Tenant ID (optional):"/>
            <TextBox  Name="txtTenant" Grid.Row="6" Grid.Column="1" Text=""/>

            <StackPanel Grid.Row="7" Grid.Column="1" Orientation="Horizontal" Margin="0,6,0,0">
              <CheckBox Name="chkDevice" Content="Use Device Code login" Margin="0,0,18,0" IsChecked="False"/>
              <CheckBox Name="chkRecurse" Content="Include subfolders" Margin="0,0,18,0"/>
              <CheckBox Name="chkDry" Content="Dry Run" IsChecked="True"/>
            </StackPanel>

            <StackPanel Grid.Row="7" Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Right">
              <Button Name="btnRunDry" Content="Dry Run"/>
              <Button Name="btnRun" Content="Run Export" Margin="0,0,0,8"/>
              <Button Name="btnOpenLogs" Content="Open Logs" Margin="0,0,0,8"/>
            </StackPanel>
          </Grid>

          <TextBox Name="txtOutput" Grid.Row="2"
                   FontFamily="Consolas" FontSize="12"
                   Background="{DynamicResource Out}" BorderBrush="{DynamicResource Border}"
                   VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                   IsReadOnly="True" TextWrapping="NoWrap"/>
        </Grid>
      </TabItem>

      <TabItem Header=" Preview ">
        <Grid Margin="10">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <Grid.ColumnDefinitions><ColumnDefinition Width="2*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>

          <DockPanel Grid.Row="0" Grid.ColumnSpan="2" Margin="0,0,0,10">
            <TextBlock FontWeight="Bold" Text="Preview latest CSV outputs"/>
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
              <Button Name="btnPrevWould" Content="Would"/>
              <Button Name="btnPrevExported" Content="Exported"/>
              <Button Name="btnPrevSkipped" Content="Skipped"/>
              <Button Name="btnLoadLatest" Content="Load Latest" Margin="0,0,0,8"/>
            </StackPanel>
          </DockPanel>

          <DataGrid Name="gridMain" Grid.Row="1" Grid.Column="0" IsReadOnly="True" AutoGenerateColumns="True"/>
          <StackPanel Grid.Row="1" Grid.Column="1" Margin="12,0,0,0">
            <TextBlock FontWeight="Bold" Text="Notes" Margin="0,0,0,6"/>
            <TextBlock Foreground="{DynamicResource Muted}" TextWrapping="Wrap">
- Dry Run launches a separate PowerShell window for Graph sign-in.
- When the run finishes, click "Load Latest" then preview the CSVs.
- Use ClientFolderMap.csv to fix ambiguous client folder names.
            </TextBlock>
          </StackPanel>
        </Grid>
      </TabItem>

      <TabItem Header=" Logs ">
        <Grid Margin="10">
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
          <ListBox Name="lstLogs" Grid.Column="0" Background="{DynamicResource Out}" BorderBrush="{DynamicResource Border}" BorderThickness="1"/>
          <StackPanel Grid.Column="1" Margin="12,0,0,0">
            <Button Name="btnRefreshLogs" Content="Refresh"/>
            <Button Name="btnOpenSelected" Content="Open Selected"/>
            <Button Name="btnOpenLogFolder" Content="Open Folder" Margin="0,0,0,8"/>
          </StackPanel>
        </Grid>
      </TabItem>

    </TabControl>

    <TextBlock Grid.Row="2" Foreground="{DynamicResource Muted}" Margin="0,12,0,0"
               Text="Delegated Graph sign-in required. Engine runs in a separate PowerShell window to show prompts."/>
  </Grid>
</Window>
"@

$window = [Windows.Markup.XamlReader]::Parse($xaml)

# Controls
$txtMailbox    = $window.FindName("txtMailbox")
$txtRootFolder = $window.FindName("txtRootFolder")
$txtCsv        = $window.FindName("txtCsv")
$txtMap        = $window.FindName("txtMap")
$txtOutRoot    = $window.FindName("txtOutRoot")
$txtLogDir     = $window.FindName("txtLogDir")
$txtTenant     = $window.FindName("txtTenant")

$chkDevice     = $window.FindName("chkDevice")
$chkRecurse    = $window.FindName("chkRecurse")
$chkDry        = $window.FindName("chkDry")

$btnCsv        = $window.FindName("btnCsv")
$btnMap        = $window.FindName("btnMap")
$btnOutRoot    = $window.FindName("btnOutRoot")
$btnLogDir     = $window.FindName("btnLogDir")

$btnRunDry     = $window.FindName("btnRunDry")
$btnRun        = $window.FindName("btnRun")
$btnOpenLogs   = $window.FindName("btnOpenLogs")

$btnPrevWould    = $window.FindName("btnPrevWould")
$btnPrevExported = $window.FindName("btnPrevExported")
$btnPrevSkipped  = $window.FindName("btnPrevSkipped")
$btnLoadLatest   = $window.FindName("btnLoadLatest")

$gridMain      = $window.FindName("gridMain")
$txtOutput     = $window.FindName("txtOutput")

$lstLogs       = $window.FindName("lstLogs")
$btnRefreshLogs  = $window.FindName("btnRefreshLogs")
$btnOpenSelected = $window.FindName("btnOpenSelected")
$btnOpenLogFolder= $window.FindName("btnOpenLogFolder")

$cardWould     = $window.FindName("cardWould")
$cardExported  = $window.FindName("cardExported")
$cardSkipped   = $window.FindName("cardSkipped")
$cardAmb       = $window.FindName("cardAmb")
$cardFailed    = $window.FindName("cardFailed")

$tglTheme      = $window.FindName("tglTheme")

# Apply theme (default Light)
ApplyTheme $window $script:Theme
$tglTheme.Content = $script:Theme

# Browse
$btnCsv.Add_Click({ $p = Select-File "CSV files (*.csv)|*.csv"; if ($p) { $txtCsv.Text = $p } })
$btnMap.Add_Click({ $p = Select-File "CSV files (*.csv)|*.csv"; if ($p) { $txtMap.Text = $p } })
$btnOutRoot.Add_Click({ $p = Select-Folder; if ($p) { $txtOutRoot.Text = $p } })
$btnLogDir.Add_Click({ $p = Select-Folder; if ($p) { $txtLogDir.Text = $p; RefreshLogs } })

# Run
$btnRunDry.Add_Click({ $chkDry.IsChecked = $true; RunEngine $true })
$btnRun.Add_Click({ $chkDry.IsChecked = $false; RunEngine $false })

# Open Logs
$btnOpenLogs.Add_Click({ if (Test-Path $txtLogDir.Text) { Start-Process explorer.exe $txtLogDir.Text } })

# Load latest
$btnLoadLatest.Add_Click({ LoadLatestCsvs })

# Preview
$btnPrevWould.Add_Click({ if ($script:LastWould) { LoadCsvToGrid $script:LastWould $gridMain } else { AppendOut "No Would CSV loaded (Load Latest first)." } })
$btnPrevExported.Add_Click({ if ($script:LastExported) { LoadCsvToGrid $script:LastExported $gridMain } else { AppendOut "No Exported CSV loaded (Load Latest first)." } })
$btnPrevSkipped.Add_Click({ if ($script:LastSkipped) { LoadCsvToGrid $script:LastSkipped $gridMain } else { AppendOut "No Skipped CSV loaded (Load Latest first)." } })

# Logs
$btnRefreshLogs.Add_Click({ RefreshLogs })
$btnOpenLogFolder.Add_Click({ if (Test-Path $txtLogDir.Text) { Start-Process explorer.exe $txtLogDir.Text } })
$btnOpenSelected.Add_Click({ $sel = $lstLogs.SelectedItem; if ($sel -and (Test-Path $sel)) { Start-Process $sel } })

# Theme toggle
$tglTheme.Add_Click({
    if ($script:Theme -eq "Light") { ApplyTheme $window "Dark" } else { ApplyTheme $window "Light" }
    $tglTheme.Content = $script:Theme
})

# Init
$txtOutput.Clear()
AppendOut "Ready."
AppendOut "Dry Run / Run Export launches a separate PowerShell window so Graph auth prompts are visible."
AppendOut "After the run finishes, click Preview > Load Latest."
RefreshLogs

$null = $window.ShowDialog()
