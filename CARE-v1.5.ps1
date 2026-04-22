# CARE v1.5 - Call AI for Report
# System Diagnostics Bridge Between Your PC and AI
# =============================================================================

# =============================================================================
# [CONFIGURANTION] 
# =============================================================================
$global:windowTitle = "CARE v1.5 - Call AI for Report"
$global:windowHeight = 720
$global:windowWidth = 1100
$global:aiDialogMode = $true            # AI в режиме диалога
# =============================================================================

# ================= ПРОВЕРКА ПРАВ АДМИНИСТРАТОРА =================
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║         ERROR: Administrator rights required!              ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host ""
    Write-Host "Press Enter to restart with administrator rights..." -ForegroundColor Green
    Read-Host
    
    $scriptPath = $MyInvocation.MyCommand.Path
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
    exit
}

# ================= ПОДКЛЮЧЕНИЕ СБОРОК =================
Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
Add-Type -AssemblyName PresentationCore -ErrorAction SilentlyContinue
Add-Type -AssemblyName WindowsBase -ErrorAction SilentlyContinue

# ================= СВОРАЧИВАНИЕ КОНСОЛИ =================
if (-not ([System.Management.Automation.PSTypeName]'HideConsole').Type) {
    Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class HideConsole {
            [DllImport("kernel32.dll")]
            static extern IntPtr GetConsoleWindow();
            [DllImport("user32.dll")]
            static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            public static void Minimize() {
                IntPtr hWnd = GetConsoleWindow();
                if (hWnd != IntPtr.Zero) ShowWindow(hWnd, 2);
            }
            public static void Show() {
                IntPtr hWnd = GetConsoleWindow();
                if (hWnd != IntPtr.Zero) ShowWindow(hWnd, 5);
            }
        }
"@
}

# ================= ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ =================
$global:window = $null
$global:autoRefreshTimer = $null
$global:scriptPID = $PID
$global:processName = (Get-Process -Id $global:scriptPID).ProcessName
$global:selfInducedCount = 0
$global:outputFolder = "C:\WindowsDiagnostics"
$global:isScanning = $false
$global:lastTopErrors = @()
$global:allDisks = @()

# ================= ФУНКЦИИ ЛОГИРОВАНИЯ =================
function Write-Log {
    param($Message, $Color = "White")
    try {
        if ($global:window -and $txtLog -and $txtLog.Dispatcher) {
            $txtLog.Dispatcher.Invoke([Action]{
                $timestamp = Get-Date -Format "HH:mm:ss"
                $icon = switch($Color) {
                    "Green" { "✓" }
                    "Red" { "✗" }
                    "Yellow" { "⚠️" }
                    "Cyan" { "→" }
                    default { "•" }
                }
                $txtLog.AppendText("[$timestamp] $icon $Message`r`n")
                $txtLog.ScrollToEnd()
            })
        }
    } catch { }
}

function Update-Progress {
    param($Value, $Text)
    try {
        if ($global:window -and $progressBar -and $progressBar.Dispatcher) {
            $progressBar.Dispatcher.Invoke([Action]{
                $progressBar.Value = $Value
                $txtProgress.Text = $Text
            })
        }
        if ($global:window -and $lblGlobalStatus -and $lblGlobalStatus.Dispatcher) {
            $lblGlobalStatus.Dispatcher.Invoke([Action]{ $lblGlobalStatus.Text = $Text })
        }
    } catch { }
}

# ================= АНОНИМИЗАЦИЯ =================
function Get-AnonymizedComputerName {
    param([string]$ComputerName)
    if ([string]::IsNullOrEmpty($ComputerName)) { return "COMP_ANONYMIZED" }
    if ($ComputerName.Length -gt 4) { return $ComputerName.Substring(0, 4) + ("*" * ($ComputerName.Length - 4)) }
    else { return $ComputerName + ("*" * (5 - $ComputerName.Length)) }
}

# ================= ФУНКЦИИ ДЛЯ УСТРОЙСТВ И СЛУЖБ =================
function Get-ProblemDevices {
    try {
        $problemDevices = @()
        $allDevices = Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object { 
            $_.Status -eq "Error" -or ($_.Problem -ne $null -and $_.Problem -ne 0)
        }
        
        $problemDescriptions = @{
            1 = "Driver not installed"
            3 = "Driver corrupted"
            10 = "Device not started"
            12 = "Not enough resources"
            14 = "Device cannot work properly"
            18 = "Driver reinstall needed"
            22 = "Device disabled"
            24 = "Device not present"
            28 = "Driver not installed"
            31 = "Driver not working"
            32 = "Driver disabled"
            33 = "Driver not ready"
            34 = "Device not working"
            35 = "Driver not working properly"
            36 = "Driver not working"
            37 = "Driver not working"
            38 = "Driver not working"
            39 = "Driver corrupted"
            40 = "Driver not working"
            41 = "Driver loaded but device not found"
            42 = "Driver problem"
            43 = "Device reported problem"
            44 = "Device stopped working"
            45 = "Device not connected"
            46 = "Device not available"
            47 = "Device not working"
            48 = "Device disabled"
            49 = "Device not working"
            52 = "Driver not verified"
        }
        
        foreach ($device in $allDevices) {
            $problemCode = $device.Problem
            $problemDesc = if ($problemCode -and $problemDescriptions.ContainsKey($problemCode)) {
                $problemDescriptions[$problemCode]
            } elseif ($problemCode) {
                "Problem code: $problemCode"
            } else {
                "Device error"
            }
            
            $severity = if ($problemCode -in @(10, 28)) { "Critical" }
                        elseif ($problemCode -in @(22, 24)) { "Warning" }
                        else { "Info" }
            
            $problemDevices += [PSCustomObject]@{
                Name = if ($device.FriendlyName) { $device.FriendlyName } else { $device.DeviceID }
                Status = $device.Status
                ProblemCode = if ($problemCode) { $problemCode.ToString() } else { "-" }
                Class = $device.Class
                ProblemDescription = $problemDesc
                Severity = $severity
            }
        }
        
        return $problemDevices | Sort-Object Severity, Name
    } catch {
        Write-Log "Error getting problem devices: $_" "Red"
        return @()
    }
}

function Get-ProblemServices {
    try {
        $problemServices = @()
        $allServices = Get-Service | Where-Object { 
            $_.Status -eq 'Stopped' -and $_.StartType -eq 'Automatic' 
        }
        
        foreach ($service in $allServices) {
            $problemServices += [PSCustomObject]@{
                DisplayName = $service.DisplayName
                Name = $service.Name
                Status = $service.Status
                StartType = $service.StartType
                Description = if ($service.Description) { $service.Description } else { "Stopped service" }
                Severity = "Warning"
            }
        }
        
        return $problemServices | Sort-Object DisplayName
    } catch {
        Write-Log "Error getting problem services: $_" "Red"
        return @()
    }
}

function Update-ProblemDevicesList {
    try {
        Write-Log "Checking problem devices..." "Cyan"
        $devices = Get-ProblemDevices
        $deviceCount = $devices.Count
        
        if ($lvProblemDevices) {
            $lvProblemDevices.Dispatcher.Invoke([Action]{
                $lvProblemDevices.ItemsSource = $null
                $lvProblemDevices.ItemsSource = $devices
                if ($txtDeviceCount) { $txtDeviceCount.Text = "Total: $deviceCount devices" }
            })
        }
        
        if ($deviceCount -gt 0) {
            Write-Log "  • Found $deviceCount problem devices" $(if($deviceCount -gt 20){"Yellow"}else{"White"})
        } else {
            Write-Log "  • No problem devices found" "Green"
        }
        
        return $deviceCount
    } catch {
        Write-Log "Error updating problem devices: $_" "Red"
        return 0
    }
}

function Update-ProblemServicesList {
    try {
        Write-Log "Checking problem services..." "Cyan"
        $services = Get-ProblemServices
        $serviceCount = $services.Count
        
        if ($lvProblemServices) {
            $lvProblemServices.Dispatcher.Invoke([Action]{
                $lvProblemServices.ItemsSource = $null
                $lvProblemServices.ItemsSource = $services
                if ($txtServiceCount) { $txtServiceCount.Text = "Total: $serviceCount services" }
            })
        }
        
        if ($serviceCount -gt 0) {
            Write-Log "  • Found $serviceCount problematic services" $(if($serviceCount -gt 5){"Yellow"}else{"White"})
        } else {
            Write-Log "  • No problematic services found" "Green"
        }
        
        return $serviceCount
    } catch {
        Write-Log "Error updating problem services: $_" "Red"
        return 0
    }
}

function Start-SelectedService {
    if ($lvProblemServices -and $lvProblemServices.SelectedItem) {
        $serviceName = $lvProblemServices.SelectedItem.Name
        $displayName = $lvProblemServices.SelectedItem.DisplayName
        
        $result = [System.Windows.MessageBox]::Show(
            "Start service '$displayName' ($serviceName)?",
            "Confirm",
            "YesNo",
            "Question"
        )
        
        if ($result -eq "Yes") {
            try {
                Start-Service -Name $serviceName -ErrorAction Stop
                Write-Log "✓ Service '$displayName' started successfully" "Green"
                Update-ProblemServicesList
            } catch {
                Write-Log "✗ Failed to start service: $_" "Red"
                [System.Windows.MessageBox]::Show(
                    "Failed to start service: $_",
                    "Error",
                    "OK",
                    "Error"
                )
            }
        }
    } else {
        [System.Windows.MessageBox]::Show(
            "Please select a service first.",
            "Information",
            "OK",
            "Information"
        )
    }
}

# ================= ФУНКЦИЯ ДЛЯ ПОЛУЧЕНИЯ ВСЕХ ДИСКОВ =================
function Get-AllDisks {
    try {
        $allDisks = @()
        # Фильтруем только диски >1 ГБ (исключаем ESP, Recovery)
        $logicalDisks = Get-CimInstance Win32_LogicalDisk -ErrorAction SilentlyContinue | 
                        Where-Object { $_.DriveType -in (2, 3) -and $_.Size / 1GB -gt 1 }
        
        foreach ($disk in $logicalDisks) {
            $size = [math]::Round($disk.Size / 1GB, 1)
            $free = [math]::Round($disk.FreeSpace / 1GB, 1)
            $percent = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1)
            
            $status = if ($percent -lt 10) { "Critical" }
                      elseif ($percent -lt 20) { "Warning" }
                      else { "OK" }
            
            $allDisks += [PSCustomObject]@{
                Drive = $disk.DeviceID
                Size = $size
                Free = $free
                Percent = $percent
                Status = $status
                VolumeName = $disk.VolumeName
                FileSystem = $disk.FileSystem
            }
        }
        return $allDisks
    } catch {
        Write-Log "Error getting disks: $_" "Red"
        return @()
    }
}

# ================= ОТКРЫТИЕ EVENT VIEWER =================
function Open-EventViewerWithHint {
    param($topErrors)
    
    Start-Process "eventvwr.msc"
    
    $hintText = "Event Viewer is now open.`n`n"
    $hintText += "HOW TO FIND THE LOG YOU NEED:`n"
    $hintText += "1. Click on 'Windows Logs' → 'System'`n"
    $hintText += "2. Click 'Filter Current Log' in the right panel`n"
    $hintText += "3. Enter Event ID (see below) and press OK`n"
    $hintText += "4. Double-click the event to see full description`n`n"
    
    if ($topErrors.Count -gt 0) {
        $hintText += "TOP EVENT IDs FROM YOUR REPORT:`n"
        foreach ($err in $topErrors) {
            $hintText += "  • ID $($err.Id) - $($err.Source) ($($err.Count)x)`n"
        }
        $hintText += "`n"
    }
    
    $hintText += "Then copy the event description and paste to your AI assistant.`n`n"
    $hintText += "⚠️ Note: Review the description for personal data before sharing."
    
    [System.Windows.MessageBox]::Show($hintText, "CARE - Event Viewer", "OK", "Information")
}

# ================= ГЕНЕРАЦИЯ ПОДСКАЗОК ДЛЯ AI =================
function Get-AIContextHints {
    param(
        $deviceCount,
        $criticalCount,
        $criticalEventId,
        $criticalEventSource,
        $errorCount,
        $topErrorIds,
        $tempCPU,
        $defenderEnabled,
        $serviceCount,
        $disks,
        $diskErrorCount
    )
    
    $hints = @()
    
    # Problem Devices
    if ($deviceCount -gt 20) {
        $hints += "• $deviceCount problem devices → likely USB ghost devices (flash drives removed without safe ejection). NOT driver corruption."
    } elseif ($deviceCount -gt 0 -and $deviceCount -le 20) {
        $hints += "• $deviceCount problem devices → moderate number. Could be real driver issues or disconnected peripherals."
    }
    
    # Critical events
    if ($criticalCount -gt 0) {
        if ($criticalEventId -eq 41 -and $criticalEventSource -match "Kernel-Power") {
            $hints += "• $criticalCount critical event(s) (Event ID 41 - Kernel-Power) → unexpected shutdown (power loss/battery). NOT a blue screen."
        } elseif ($criticalEventId -eq 1001 -or $criticalEventSource -match "BugCheck") {
            $hints += "• $criticalCount critical event(s) → blue screen (BSOD). Ask for error code if available."
        } else {
            $hints += "• $criticalCount critical event(s) (ID $criticalEventId). Ask user what happened before the crash."
        }
    }
    
    # Disk errors hint
    if ($diskErrorCount -gt 0) {
        $hints += "• $diskErrorCount disk error(s) detected. If errors persist, ask user: 'Do you have a secondary HDD (Caddy) or external drive? Some disk errors on secondary drives are normal during boot.'"
    }
    
    # Temperature
    if ($tempCPU -ne "N/A") {
        if ($tempCPU -gt 75) {
            $hints += "• CPU temperature $tempCPU°C → overheating possible. Ask about cooling."
        } elseif ($tempCPU -gt 65) {
            $hints += "• CPU temperature $tempCPU°C → warm but acceptable."
        } else {
            $hints += "• CPU temperature $tempCPU°C → normal."
        }
    } else {
        $hints += "• CPU temperature N/A → sensor unavailable. Not a thermal issue unless user reports overheating."
    }
    
    # Security
    if (-not $defenderEnabled) {
        $hints += "• Windows Defender: Inactive → security risk. Recommend enabling."
    }
    
    # Services
    if ($serviceCount -gt 0) {
        $hints += "• $serviceCount stopped services → likely disabled by optimizer tools. NOT an error."
    }
    
    # Disks - find the most critical one
    if ($disks.Count -gt 0) {
        $lowDisk = $disks | Where-Object { $_.Status -eq "Warning" -or $_.Status -eq "Critical" } | Sort-Object Percent | Select-Object -First 1
        if ($lowDisk) {
            $hints += "• Disk $($lowDisk.Drive): $($lowDisk.Percent)% free ($($lowDisk.Status))"
        } else {
            $totalFree = [math]::Round(($disks | Measure-Object -Property Free -Sum).Sum, 1)
            $totalSize = [math]::Round(($disks | Measure-Object -Property Size -Sum).Sum, 1)
            $totalPercent = [math]::Round(($totalFree / $totalSize) * 100, 1)
            $hints += "• Disks: $($disks.Count) drives, $totalFree/$totalSize GB free ($totalPercent%) overall"
        }
    }
    
    # Top errors
    if ($topErrorIds.Count -gt 0 -and $errorCount -gt 0) {
        $errorSummary = ($topErrorIds | ForEach-Object { "$($_.Id)($($_.Count)x)" }) -join ", "
        $hints += "• Top errors: $errorSummary"
    }
    
    return $hints
}

# ================= XAML ИНТЕРФЕЙС =================
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$($global:windowTitle)"
        Height="$($global:windowHeight)" 
        Width="$($global:windowWidth)"
        MinHeight="600" 
        MinWidth="900"
        WindowStartupLocation="CenterScreen"
        Background="#1E1E1E">
    
    <Window.Resources>
        <Style x:Key="TabHeaderStyle" TargetType="TabItem">
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Background" Value="#2D2D2D"/>
            
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border Name="Border" 
                                Background="{TemplateBinding Background}" 
                                BorderThickness="0"
                                Padding="{TemplateBinding Padding}">
                            <TextBlock Name="HeaderText" 
                                       Text="{TemplateBinding Header}"
                                       VerticalAlignment="Center"
                                       HorizontalAlignment="Center"
                                       Foreground="#E0E0E0"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#FFD93D"/>
                                <Setter TargetName="HeaderText" Property="Foreground" Value="#1E1E1E"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#FFD93D"/>
                                <Setter TargetName="HeaderText" Property="Foreground" Value="#1E1E1E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style x:Key="ActionButtonStyle" TargetType="Button">
            <Setter Property="Margin" Value="2"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Background" Value="#0078D7"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FFD93D"/>
                    <Setter Property="Foreground" Value="#1E1E1E"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="SuccessButtonStyle" TargetType="Button" BasedOn="{StaticResource ActionButtonStyle}">
            <Setter Property="Background" Value="#10B981"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FFD93D"/>
                    <Setter Property="Foreground" Value="#1E1E1E"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="DangerButtonStyle" TargetType="Button" BasedOn="{StaticResource ActionButtonStyle}">
            <Setter Property="Background" Value="#EF4444"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FFD93D"/>
                    <Setter Property="Foreground" Value="#1E1E1E"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="GridViewColumnHeaderStyle" TargetType="GridViewColumnHeader">
            <Setter Property="Background" Value="#2D2D2D"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="BorderBrush" Value="#404040"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FFD93D"/>
                    <Setter Property="Foreground" Value="#1E1E1E"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="ProblemItemStyle" TargetType="ListViewItem">
            <Setter Property="Padding" Value="5"/>
            <Style.Triggers>
                <DataTrigger Binding="{Binding Severity}" Value="Critical">
                    <Setter Property="Background" Value="#442222"/>
                    <Setter Property="Foreground" Value="#FFAAAA"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Severity}" Value="Warning">
                    <Setter Property="Background" Value="#444422"/>
                    <Setter Property="Foreground" Value="#FFFFAA"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Severity}" Value="Info">
                    <Setter Property="Background" Value="#224444"/>
                    <Setter Property="Foreground" Value="#AAFFFF"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
        
        <Style TargetType="TextBlock" x:Key="ResultTextBlock">
            <Setter Property="Foreground" Value="#E0E0E0"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Верхняя панель -->
        <Border Grid.Row="0" Background="#0078D7" Padding="10">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <CheckBox x:Name="chkAutoRefresh" 
                         Content="🔄 Auto-refresh (10s)" 
                         Foreground="White"
                         Margin="10,0"
                         FontSize="12"
                         VerticalAlignment="Center"/>
                <Button x:Name="btnStartScan" 
                        Content="▶️ START DIAGNOSTICS" 
                        Style="{StaticResource SuccessButtonStyle}"/>
                <Button x:Name="btnStopScan" 
                        Content="⏹️ Stop" 
                        Style="{StaticResource DangerButtonStyle}" 
                        Margin="5,0"
                        IsEnabled="False"/>
                <Button x:Name="btnShowConsole" 
                        Content="📟 Show Console" 
                        Style="{StaticResource ActionButtonStyle}" 
                        Margin="5,0,0,0"/>
            </StackPanel>
        </Border>

        <!-- Основной контент - TabControl -->
        <TabControl Grid.Row="1" Margin="10" x:Name="tabControl" Background="#1E1E1E">
            
            <!-- Вкладка 1: Диагностика -->
            <TabItem Header="🔍 DIAGNOSTICS" Style="{StaticResource TabHeaderStyle}">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button x:Name="btnClearLog" Content="🗑️ Clear Log" Style="{StaticResource ActionButtonStyle}"/>
                        <Button x:Name="btnOpenReports" Content="📂 Open Reports Folder" Style="{StaticResource ActionButtonStyle}" Margin="5,0"/>
                    </StackPanel>

                    <TextBox x:Name="txtLog" Grid.Row="1" 
                             Background="#2D2D2D" 
                             Foreground="#E0E0E0"
                             FontFamily="Consolas"
                             FontSize="10"
                             TextWrapping="Wrap"
                             AcceptsReturn="True"
                             IsReadOnly="True"
                             VerticalScrollBarVisibility="Auto"/>
                    
                    <ProgressBar x:Name="progressBar" Grid.Row="2" Height="20" Margin="0,10,0,5" 
                                 Foreground="#10B981" Background="#404040"/>
                    <TextBlock x:Name="txtProgress" Grid.Row="2" 
                               HorizontalAlignment="Center" 
                               VerticalAlignment="Center"
                               FontSize="11"
                               Foreground="White"
                               Text="Ready to start"/>
                </Grid>
            </TabItem>

            <!-- Вкладка 2: Результаты -->
            <TabItem Header="📊 RESULTS" Style="{StaticResource TabHeaderStyle}">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button x:Name="btnExportResults" Content="💾 Export to CSV" Style="{StaticResource ActionButtonStyle}"/>
                    </StackPanel>
                    
                    <ListView x:Name="lvResults" Grid.Row="1" Background="#2D2D2D" Foreground="#E0E0E0">
                        <ListView.View>
                            <GridView>
                                <GridView.ColumnHeaderContainerStyle>
                                    <Style TargetType="GridViewColumnHeader" BasedOn="{StaticResource GridViewColumnHeaderStyle}"/>
                                </GridView.ColumnHeaderContainerStyle>
                                <GridViewColumn Header="Metric" Width="200">
                                    <GridViewColumn.CellTemplate>
                                        <DataTemplate>
                                            <TextBlock Text="{Binding Metric}" Foreground="#E0E0E0"/>
                                        </DataTemplate>
                                    </GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header="Value" Width="300">
                                    <GridViewColumn.CellTemplate>
                                        <DataTemplate>
                                            <TextBlock Text="{Binding Value}" Foreground="#E0E0E0"/>
                                        </DataTemplate>
                                    </GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header="Status" Width="150">
                                    <GridViewColumn.CellTemplate>
                                        <DataTemplate>
                                            <TextBlock Text="{Binding Status}" Foreground="#E0E0E0"/>
                                        </DataTemplate>
                                    </GridViewColumn.CellTemplate>
                                </GridViewColumn>
                            </GridView>
                        </ListView.View>
                    </ListView>
                </Grid>
            </TabItem>

            <!-- Вкладка 3: Проблемные устройства -->
            <TabItem Header="🔧 PROBLEM DEVICES" Style="{StaticResource TabHeaderStyle}">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button x:Name="btnOpenDeviceManager" Content="🖥️ Open Device Manager" Style="{StaticResource ActionButtonStyle}"/>
                        <Button x:Name="btnRefreshDevices" Content="🔄 Refresh" Style="{StaticResource ActionButtonStyle}" Margin="5,0"/>
                        <TextBlock x:Name="txtDeviceCount" VerticalAlignment="Center" Margin="10,0" Foreground="#E0E0E0"/>
                    </StackPanel>
                    
                    <ListView x:Name="lvProblemDevices" Grid.Row="1" Background="#2D2D2D"
                              ItemContainerStyle="{StaticResource ProblemItemStyle}">
                        <ListView.View>
                            <GridView>
                                <GridView.ColumnHeaderContainerStyle>
                                    <Style TargetType="GridViewColumnHeader" BasedOn="{StaticResource GridViewColumnHeaderStyle}"/>
                                </GridView.ColumnHeaderContainerStyle>
                                <GridViewColumn Header="Device Name" Width="250" DisplayMemberBinding="{Binding Name}"/>
                                <GridViewColumn Header="Status" Width="120" DisplayMemberBinding="{Binding Status}"/>
                                <GridViewColumn Header="Problem Code" Width="100" DisplayMemberBinding="{Binding ProblemCode}"/>
                                <GridViewColumn Header="Class" Width="150" DisplayMemberBinding="{Binding Class}"/>
                                <GridViewColumn Header="Description" Width="350" DisplayMemberBinding="{Binding ProblemDescription}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                </Grid>
            </TabItem>

            <!-- Вкладка 4: Проблемные службы -->
            <TabItem Header="⚙️ PROBLEM SERVICES" Style="{StaticResource TabHeaderStyle}">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button x:Name="btnOpenServices" Content="📋 Open Services" Style="{StaticResource ActionButtonStyle}"/>
                        <Button x:Name="btnRefreshServices" Content="🔄 Refresh" Style="{StaticResource ActionButtonStyle}" Margin="5,0"/>
                        <Button x:Name="btnStartSelectedService" Content="▶️ Start Selected" Style="{StaticResource SuccessButtonStyle}" Margin="5,0"/>
                        <TextBlock x:Name="txtServiceCount" VerticalAlignment="Center" Margin="10,0" Foreground="#E0E0E0"/>
                    </StackPanel>
                    
                    <ListView x:Name="lvProblemServices" Grid.Row="1" Background="#2D2D2D"
                              SelectionMode="Single"
                              ItemContainerStyle="{StaticResource ProblemItemStyle}">
                        <ListView.View>
                            <GridView>
                                <GridView.ColumnHeaderContainerStyle>
                                    <Style TargetType="GridViewColumnHeader" BasedOn="{StaticResource GridViewColumnHeaderStyle}"/>
                                </GridView.ColumnHeaderContainerStyle>
                                <GridViewColumn Header="Display Name" Width="200" DisplayMemberBinding="{Binding DisplayName}"/>
                                <GridViewColumn Header="Service Name" Width="150" DisplayMemberBinding="{Binding Name}"/>
                                <GridViewColumn Header="Status" Width="80" DisplayMemberBinding="{Binding Status}"/>
                                <GridViewColumn Header="Start Type" Width="100" DisplayMemberBinding="{Binding StartType}"/>
                                <GridViewColumn Header="Description" Width="350" DisplayMemberBinding="{Binding Description}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                </Grid>
            </TabItem>

            <!-- Вкладка 5: Здоровье -->
            <TabItem Header="❤️ HEALTH" Style="{StaticResource TabHeaderStyle}">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button x:Name="btnOpenEventViewer" 
                                Content="🔍 Open Event Viewer with filter by top errors" 
                                Style="{StaticResource ActionButtonStyle}"
                                ToolTip="Opens Windows Event Viewer and shows how to find relevant errors"/>
                    </StackPanel>
                    
                    <TextBlock Grid.Row="0" x:Name="txtHealthScore" 
                               FontSize="24" 
                               FontWeight="Bold" 
                               HorizontalAlignment="Center" 
                               Margin="10,0,10,10"
                               Foreground="#10B981"/>
                    
                    <TextBox x:Name="txtRecommendations" Grid.Row="1"
                             Background="#2D2D2D"
                             Foreground="#E0E0E0"
                             FontFamily="Segoe UI"
                             FontSize="12"
                             TextWrapping="Wrap"
                             IsReadOnly="True"
                             VerticalScrollBarVisibility="Auto"
                             Margin="10"/>
                </Grid>
            </TabItem>
        </TabControl>

        <!-- Нижняя статусная строка -->
        <StatusBar Grid.Row="2" Background="#0078D7" Height="30">
            <StatusBarItem>
                <TextBlock x:Name="lblGlobalStatus" 
                          Text="✓ CARE v1.5 ready — Copy report to your AI" 
                          Foreground="White"
                          FontWeight="Bold"/>
            </StatusBarItem>
            <StatusBarItem HorizontalAlignment="Right">
                <TextBlock x:Name="lblAdminStatus" 
                          Text="👑 ADMIN" 
                          Foreground="#FFD700"
                          FontWeight="Bold"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
"@

# ================= ЗАГРУЗКА ИНТЕРФЕЙСА =================
try {
    $reader = New-Object System.Xml.XmlNodeReader([xml]$xaml)
    $global:window = [Windows.Markup.XamlReader]::Load($reader)
    $global:window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
} catch {
    Write-Host "Error loading XAML: $_" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

# ================= ПОЛУЧЕНИЕ ССЫЛОК НА ЭЛЕМЕНТЫ =================
$txtLog = $global:window.FindName("txtLog")
$lvResults = $global:window.FindName("lvResults")
$lvProblemDevices = $global:window.FindName("lvProblemDevices")
$lvProblemServices = $global:window.FindName("lvProblemServices")
$txtHealthScore = $global:window.FindName("txtHealthScore")
$txtRecommendations = $global:window.FindName("txtRecommendations")
$txtDeviceCount = $global:window.FindName("txtDeviceCount")
$txtServiceCount = $global:window.FindName("txtServiceCount")
$progressBar = $global:window.FindName("progressBar")
$txtProgress = $global:window.FindName("txtProgress")
$lblGlobalStatus = $global:window.FindName("lblGlobalStatus")
$btnStartScan = $global:window.FindName("btnStartScan")
$btnStopScan = $global:window.FindName("btnStopScan")
$btnClearLog = $global:window.FindName("btnClearLog")
$btnOpenReports = $global:window.FindName("btnOpenReports")
$btnExportResults = $global:window.FindName("btnExportResults")
$btnShowConsole = $global:window.FindName("btnShowConsole")
$btnOpenDeviceManager = $global:window.FindName("btnOpenDeviceManager")
$btnRefreshDevices = $global:window.FindName("btnRefreshDevices")
$btnOpenServices = $global:window.FindName("btnOpenServices")
$btnRefreshServices = $global:window.FindName("btnRefreshServices")
$btnStartSelectedService = $global:window.FindName("btnStartSelectedService")
$btnOpenEventViewer = $global:window.FindName("btnOpenEventViewer")
$chkAutoRefresh = $global:window.FindName("chkAutoRefresh")
$tabControl = $global:window.FindName("tabControl")

# ================= ОСНОВНАЯ ФУНКЦИЯ ДИАГНОСТИКИ =================
function Start-Diagnostics {
    $global:isScanning = $true
    $global:selfInducedCount = 0
    
    $btnStartScan.IsEnabled = $false
    $btnStopScan.IsEnabled = $true
    $txtLog.Clear()
    $progressBar.Value = 0
    
    Write-Log "=" * 60 "Cyan"
    Write-Log "CARE v1.5 - Starting diagnostics" "Green"
    Write-Log "=" * 60 "Cyan"
    
    try {
        # Шаг 1: Системная информация
        Update-Progress -Value 5 -Text "Collecting system information..."
        Write-Log "Collecting system information..." "Cyan"
        
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $comp = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $cpu = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
        
        $totalRAM = [math]::Round($comp.TotalPhysicalMemory / 1GB, 1)
        $freeRAM = [math]::Round($os.FreePhysicalMemory * 1024 / 1GB, 2)
        $usedRAM = [math]::Round($totalRAM - $freeRAM, 1)
        $ramPercent = [math]::Round(($usedRAM / $totalRAM) * 100, 1)
        
        $computerName = Get-AnonymizedComputerName $env:COMPUTERNAME
        
        Write-Log "  • Computer: $computerName" "White"
        Write-Log "  • RAM: $usedRAM/$totalRAM GB ($ramPercent%)" "White"
        Write-Log "  • CPU: $($cpu.Name)" "White"
        Write-Log "✓ System information collected" "Green"
        
        # Шаг 2: Все диски
        Update-Progress -Value 15 -Text "Checking all disks..."
        Write-Log "Checking all disks..." "Cyan"
        
        $global:allDisks = Get-AllDisks
        $diskLines = @()
        $lowSpaceWarning = $false
        foreach ($disk in $global:allDisks) {
            $statusIcon = if ($disk.Status -eq "Warning") { "⚠️" } elseif ($disk.Status -eq "Critical") { "🔴" } else { "•" }
            $diskLines += "  $statusIcon $($disk.Drive): $($disk.Free)/$($disk.Size) GB ($($disk.Percent)%)"
            if ($disk.Status -eq "Warning" -or $disk.Status -eq "Critical") { $lowSpaceWarning = $true }
            Write-Log "  • $($disk.Drive): $($disk.Free)/$($disk.Size) GB ($($disk.Percent)%)" $(if($disk.Status -eq "Warning"){"Yellow"}elseif($disk.Status -eq "Critical"){"Red"}else{"White"})
        }
        Write-Log "✓ Disk check completed ($($global:allDisks.Count) drives found)" "Green"
        
        # Шаг 3: Температура
        Update-Progress -Value 25 -Text "Checking temperatures..."
        Write-Log "Checking temperatures..." "Cyan"
        
        $temps = Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction SilentlyContinue
        $tempCPU = "N/A"
        if ($temps) {
            $rawTemp = ($temps[0].CurrentTemperature - 2732) / 10
            if ($rawTemp -gt -50 -and $rawTemp -lt 120) {
                $tempCPU = [math]::Round($rawTemp, 1)
                Write-Log "  • CPU: $tempCPU°C" $(if($tempCPU -gt 75){"Red"}elseif($tempCPU -gt 65){"Yellow"}else{"Green"})
            } else {
                Write-Log "  • CPU temperature: N/A (invalid sensor reading)" "Gray"
            }
        } else {
            Write-Log "  • Temperature information not available" "Gray"
        }
        Write-Log "✓ Temperature check completed" "Green"
        
        # Шаг 4: Проблемные устройства
        Update-Progress -Value 40 -Text "Checking problem devices..."
        Write-Log "Checking problem devices..." "Cyan"
        $deviceCount = Update-ProblemDevicesList
        
        # Шаг 5: Проблемные службы
        Update-Progress -Value 50 -Text "Checking problem services..."
        Write-Log "Checking problem services..." "Cyan"
        $serviceCount = Update-ProblemServicesList
        
        # Шаг 6: События сбора и анализа
        Update-Progress -Value 65 -Text "Collecting events..."
        Write-Log "Collecting events from last 7 days..." "Cyan"
        
        $startDate = (Get-Date).AddDays(-7)
        $criticalCount = 0
        $errorCount = 0
        $diskErrorCount = 0
        $wmiErrorCount = 0
        
        $criticalEventId = $null
        $criticalEventSource = $null
        $errorIdCounts = @{}
        
        $logsToCheck = @("System", "Application")
        foreach ($logName in $logsToCheck) {
            try {
                $logEvents = Get-WinEvent -LogName $logName -MaxEvents 500 -ErrorAction SilentlyContinue | 
                            Where-Object { $_.TimeCreated -ge $startDate -and $_.Level -in @(1,2,3) }
                
                if ($logEvents) {
                    foreach ($evt in $logEvents) {
                        $isSelf = $false
                        if ($evt.Message -match $global:processName -or $evt.Message -match $global:scriptPID) {
                            $global:selfInducedCount++
                            $isSelf = $true
                        }
                        
                        if (-not $isSelf) {
                            if ($evt.Level -eq 1) { 
                                $criticalCount++
                                if (-not $criticalEventId) {
                                    $criticalEventId = $evt.Id
                                    $criticalEventSource = $evt.ProviderName
                                }
                            }
                            if ($evt.Level -eq 2) { 
                                $errorCount++
                                $key = "$($evt.Id)|$($evt.ProviderName)"
                                if ($errorIdCounts.ContainsKey($key)) {
                                    $errorIdCounts[$key]++
                                } else {
                                    $errorIdCounts[$key] = 1
                                }
                            }
                            if ($evt.Message -match "disk|drive|storport") { $diskErrorCount++ }
                            if ($evt.Message -match "WMI|WinMgmt") { $wmiErrorCount++ }
                        }
                    }
                    Write-Log "  • $logName : $($logEvents.Count) events" "White"
                }
            } catch { Write-Log "  • $logName : unavailable" "Gray" }
        }
        
        # Топ-3 ошибок по частоте
        $topErrors = $errorIdCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 3 | ForEach-Object { 
            $parts = $_.Key -split '\|'
            [PSCustomObject]@{ Id = $parts[0]; Source = $parts[1]; Count = $_.Value }
        }
        $global:lastTopErrors = $topErrors
        
        Write-Log "  • Critical: $criticalCount" $(if($criticalCount -gt 0){"Red"}else{"White"})
        Write-Log "  • Errors: $errorCount" $(if($errorCount -gt 0){"Yellow"}else{"White"})
        Write-Log "  • Disk errors: $diskErrorCount" $(if($diskErrorCount -gt 0){"Yellow"}else{"White"})
        if ($global:selfInducedCount -gt 0) {
            Write-Log "  • Self-induced (ignored): $global:selfInducedCount" "Gray"
        }
        Write-Log "✓ Event collection completed" "Green"
        
        # Шаг 7: Безопасность
        Update-Progress -Value 85 -Text "Checking security..."
        Write-Log "Checking security..." "Cyan"
        
        try {
            $defender = Get-MpComputerStatus -ErrorAction SilentlyContinue
            $defenderEnabled = $defender.AntivirusEnabled
            Write-Log "  • Windows Defender: $(if($defenderEnabled){'Active ✅'}else{'Inactive ⚠️'})" $(if(-not $defenderEnabled){"Yellow"}else{"Green"})
        } catch {
            Write-Log "  • Security information unavailable" "Gray"
            $defenderEnabled = $false
        }
        Write-Log "✓ Security check completed" "Green"
        
        # Шаг 8: Создание отчета
        Update-Progress -Value 100 -Text "Creating report..."
        Write-Log "Creating diagnostic report..." "Cyan"
        
        if (-not (Test-Path $global:outputFolder)) {
            New-Item -ItemType Directory -Path $global:outputFolder -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmm"
        
        # Расчет оценки здоровья (с учётом дисков)
        $score = 100
        $score -= $criticalCount * 20
        $score -= [Math]::Min($errorCount / 10, 30)
        $score -= $diskErrorCount * 5
        $score -= $serviceCount * 3
        $score -= [Math]::Min($deviceCount / 5, 20)
        if ($tempCPU -ne "N/A" -and $tempCPU -gt 75) { $score -= 20 }
        
        # Штраф за диски с низким свободным местом
        foreach ($disk in $global:allDisks) {
            if ($disk.Percent -lt 10) { $score -= 15 }
            elseif ($disk.Percent -lt 20) { $score -= 8 }
            elseif ($disk.Percent -lt 30) { $score -= 3 }
        }
        
        $score = [Math]::Max(0, [Math]::Min(100, $score))
        
        $healthStatus = if ($score -ge 80) { "EXCELLENT" }
                        elseif ($score -ge 60) { "GOOD" }
                        elseif ($score -ge 40) { "FAIR" }
                        else { "CRITICAL" }
        
        # Обновление UI
        $txtHealthScore.Dispatcher.Invoke([Action]{
            $txtHealthScore.Text = "Health Score: $score/100`r`nStatus: $healthStatus"
            $txtHealthScore.Foreground = if ($score -ge 80) { [System.Windows.Media.Brushes]::LightGreen }
                                         elseif ($score -ge 60) { [System.Windows.Media.Brushes]::Yellow }
                                         else { [System.Windows.Media.Brushes]::OrangeRed }
        })
        
        # Рекомендации для пользователя
        $recommendations = ""
        if ($criticalCount -gt 0) { $recommendations += "⚠️ CRITICAL EVENTS: $criticalCount critical events require attention`r`n`r`n" }
        if ($diskErrorCount -gt 0) { $recommendations += "💾 DISK ERRORS: $diskErrorCount disk errors detected - check drive health`r`n`r`n" }
        if ($serviceCount -gt 0) { $recommendations += "⚙️ SERVICE ISSUES: $serviceCount services are stopped but should be running - check the Problem Services tab`r`n`r`n" }
        if ($deviceCount -gt 0) { $recommendations += "🔌 DEVICE ISSUES: $deviceCount problem devices found - check the Problem Devices tab for details`r`n`r`n" }
        if ($tempCPU -ne "N/A" -and $tempCPU -gt 75) { $recommendations += "🌡️ HIGH TEMPERATURE: CPU at $tempCPU°C - clean cooling system`r`n`r`n" }
        if (-not $defenderEnabled) { $recommendations += "🛡️ SECURITY: Windows Defender is disabled - recommended to enable`r`n`r`n" }
        
        # Добавляем предупреждение о дисках с низким свободным местом
        $lowDisks = $global:allDisks | Where-Object { $_.Status -eq "Warning" -or $_.Status -eq "Critical" }
        if ($lowDisks.Count -gt 0) {
            $recommendations += "💿 LOW DISK SPACE: "
            foreach ($disk in $lowDisks) {
                $recommendations += "$($disk.Drive) has only $($disk.Percent)% free ($($disk.Status)). "
            }
            $recommendations += "Consider cleaning or moving files.`r`n`r`n"
        }
        
        if ($recommendations -eq "") {
            $recommendations = "✅ System is in excellent condition! Continue regular maintenance.`r`n`r`n"
        }
        
        $recommendations += "📁 Report saved to: $global:outputFolder`r`n"
        $recommendations += "🤖 Copy the report content and send to AI assistant (ChatGPT, Gemini, Copilot) for detailed analysis`r`n"
        $recommendations += "🔍 Use 'Open Event Viewer' button to examine specific errors if AI needs more details`r`n"
        
        $txtRecommendations.Dispatcher.Invoke([Action]{ $txtRecommendations.Text = $recommendations })
        
        # Создание CSV для результатов
        $results = New-Object System.Collections.Generic.List[PSObject]
        $results.Add([PSCustomObject]@{ Metric = "Computer"; Value = $computerName; Status = "OK" })
        $results.Add([PSCustomObject]@{ Metric = "RAM Usage"; Value = "$usedRAM/$totalRAM GB ($ramPercent%)"; Status = if($ramPercent -gt 85){"Critical"}elseif($ramPercent -gt 70){"Warning"}else{"Normal"} })
        $results.Add([PSCustomObject]@{ Metric = "Critical Events"; Value = $criticalCount; Status = if($criticalCount -gt 0){"Critical"}else{"Normal"} })
        $results.Add([PSCustomObject]@{ Metric = "Errors"; Value = $errorCount; Status = if($errorCount -gt 20){"Warning"}else{"Normal"} })
        $results.Add([PSCustomObject]@{ Metric = "Disk Errors"; Value = $diskErrorCount; Status = if($diskErrorCount -gt 0){"Warning"}else{"Normal"} })
        $results.Add([PSCustomObject]@{ Metric = "Problematic Services"; Value = $serviceCount; Status = if($serviceCount -gt 0){"Warning"}else{"Normal"} })
        $results.Add([PSCustomObject]@{ Metric = "Problem Devices"; Value = $deviceCount; Status = if($deviceCount -gt 0){"Warning"}else{"Normal"} })
        $results.Add([PSCustomObject]@{ Metric = "CPU Temperature"; Value = if($tempCPU -ne "N/A"){"$tempCPU°C"}else{"N/A"}; Status = if($tempCPU -ne "N/A" -and $tempCPU -gt 75){"Critical"}elseif($tempCPU -ne "N/A" -and $tempCPU -gt 65){"Warning"}else{"Normal"} })
        $results.Add([PSCustomObject]@{ Metric = "Windows Defender"; Value = if($defenderEnabled){"Active"}else{"Inactive"}; Status = if($defenderEnabled){"OK"}else{"Warning"} })
        $results.Add([PSCustomObject]@{ Metric = "Disks Found"; Value = $global:allDisks.Count; Status = "OK" })
        
        $lvResults.Dispatcher.Invoke([Action]{ $lvResults.ItemsSource = $results })
        
        # Генерация подсказок для AI
        $topErrorList = $topErrors | ForEach-Object { @{ Id = $_.Id; Source = $_.Source; Count = $_.Count } }
        $aiHints = Get-AIContextHints -deviceCount $deviceCount `
                                       -criticalCount $criticalCount `
                                       -criticalEventId $criticalEventId `
                                       -criticalEventSource $criticalEventSource `
                                       -errorCount $errorCount `
                                       -topErrorIds $topErrorList `
                                       -tempCPU $tempCPU `
                                       -defenderEnabled $defenderEnabled `
                                       -serviceCount $serviceCount `
                                       -disks $global:allDisks `
                                       -diskErrorCount $diskErrorCount
        
        $hintsText = if ($aiHints.Count -gt 0) { $aiHints -join "`n" } else { "• No significant issues detected." }
        
        # Формирование строки дисков для отчёта (без лишних двоеточий)
        $diskReport = ""
        foreach ($disk in $global:allDisks) {
            $warning = if ($disk.Status -eq "Warning") { " ⚠️" } elseif ($disk.Status -eq "Critical") { " 🔴" } else { "" }
            # Убираем лишнее двоеточие — выводим только букву диска
            $driveLetter = $disk.Drive -replace ':', ''
            $diskReport += "$driveLetter $($disk.Free)/$($disk.Size) GB ($($disk.Percent)%)$warning`n"
        }
        
        # Формирование строки с дисками для инструкции AI
        $diskWarningText = ""
        if ($lowDisks.Count -gt 0) {
            $driveList = ($lowDisks | ForEach-Object { $_.Drive -replace ':', '' }) -join ', '
            $verb = if ($lowDisks.Count -eq 1) { "is" } else { "are" }
            $diskWarningText = "   - If low disk space: `"Your $driveList drive(s) $verb getting full. Is this affecting performance?`"`n"
        }
        
        $diskSummary = if ($lowDisks.Count -gt 0) { "⚠️ $($lowDisks.Count) drive(s) with low space" } else { "all drives have adequate space" }
        
        # Сохранение отчета
        $reportFile = "$global:outputFolder\CARE_Report_$timestamp.txt"
        
        $report = @"
╔══════════════════════════════════════════════════════════════════╗
║                    CARE v1.5 - Diagnostic Report                 ║
║                    Call AI for Report                            ║
╚══════════════════════════════════════════════════════════════════╝

SYSTEM INFORMATION
═══════════════════════════════════════════════════════════════════
Computer: $computerName
RAM: $usedRAM/$totalRAM GB ($ramPercent%)
CPU: $($cpu.Name)

STORAGE DEVICES
═══════════════════════════════════════════════════════════════════
$diskReport
Total: $($global:allDisks.Count) drives, $diskSummary

TEMPERATURES
═══════════════════════════════════════════════════════════════════
CPU: $(if($tempCPU -ne "N/A"){"$tempCPU°C"}else{"N/A"})

EVENT LOGS (last 7 days)
═══════════════════════════════════════════════════════════════════
Critical: $criticalCount$(if($criticalEventId){" (ID $criticalEventId - $criticalEventSource)"})
Errors: $errorCount
Disk Errors: $diskErrorCount
WMI Errors: $wmiErrorCount
Self-induced (ignored): $global:selfInducedCount

SERVICES & DEVICES (summary only - details hidden for privacy)
═══════════════════════════════════════════════════════════════════
Problematic Services Count: $serviceCount
Problem Devices Count: $deviceCount

SECURITY
═══════════════════════════════════════════════════════════════════
Windows Defender: $(if($defenderEnabled){"Active"}else{"Inactive"})

HEALTH ASSESSMENT
═══════════════════════════════════════════════════════════════════
Overall Score: $score/100 ($healthStatus)

🔍 QUICK CONTEXT FOR AI (token-efficient)
═══════════════════════════════════════════════════════════════════
$hintsText

🤖 INSTRUCTIONS FOR AI ASSISTANT — DIALOG MODE
═══════════════════════════════════════════════════════════════════
DO NOT give final answers immediately. Instead:

1. Start with empathy: "I see your system is $healthStatus ($score/100)."

2. Based on the Quick Context above, ask ONE focused question. Examples:
   - If USB ghost devices suspected: "Do you use many USB flash drives and remove them without 'safe ejection'?"
   - If unexpected shutdown: "Did the PC shut down unexpectedly or blue-screen?"
   - If services disabled: "Did you disable any Windows services for optimization?"
   - If temperature high: "Is the fan loud or does the PC feel hot?"
   - If disk errors present: "I see disk errors in the log. Do you have a secondary HDD (Caddy) or external drive? Some disk errors on boot are normal for secondary drives."
$diskWarningText
3. WAIT for user's answer. Do NOT ask multiple questions at once.

4. After response, provide ONE actionable step, then ask for confirmation.

5. Keep responses short to save token budget.

USER'S CAPABILITY — LOG EXPORT
═══════════════════════════════════════════════════════════════════
If Quick Context is insufficient, user CAN provide full event details:
• Open Event Viewer via CARE's "🔍 Open Event Viewer" button
• Find events by ID (see top errors in report)
• Copy/paste event descriptions to you

How to ask: "Could you open Event Viewer (CARE has a button) and copy 
the description of Event ID X? That would help me give a better answer."

Note: User may decline. Respect their choice. Do not demand logs.

DEEP DIVE — OPTIONAL (ask only if user reports specific symptoms)
═══════════════════════════════════════════════════════════════════
• Graphics issues: "Filter Event Viewer by source 'Display' or 'Kernel-PnP'"
• Network problems: "Filter by source 'Microsoft-Windows-NetworkProfile'"
• USB/Thunderbolt: "Check Device Manager for yellow exclamation marks"
• Virtual machines: "Do you use Hyper-V, VMware, or VirtualBox?"

GOAL: User should feel guided, not overwhelmed. One question at a time.

═══════════════════════════════════════════════════════════════════
Report Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Version: CARE v1.5
Data Privacy: All personal information has been masked
═══════════════════════════════════════════════════════════════════
"@
        $report | Out-File $reportFile -Encoding UTF8
        Write-Log "✓ Report saved: $reportFile" "Green"
        
        Write-Log ""
        Write-Log "=" * 60 "Green"
        Write-Log "✓ DIAGNOSTICS COMPLETED SUCCESSFULLY!" "Green"
        Write-Log "=" * 60 "Green"
        Write-Log ""
        Write-Log "📁 Report saved to: $global:outputFolder" "Cyan"
        Write-Log "💿 Disks found: $($global:allDisks.Count) | Low space: $(($lowDisks | Measure-Object).Count)" $(if($lowDisks.Count -gt 0){"Yellow"}else{"White"})
        Write-Log "🔧 Problem devices: $deviceCount | Problem services: $serviceCount" "Yellow"
        Write-Log "   Check the 'Problem Devices' and 'Problem Services' tabs for details" "Cyan"
        Write-Log "🤖 Copy report content and send to AI for analysis" "Yellow"
        Write-Log "💬 AI will ask ONE question at a time — just answer naturally" "Cyan"
        Write-Log "🔍 Use 'Open Event Viewer' button if AI needs more details" "Cyan"
        
        $btnExportResults.IsEnabled = $true
        
    } catch {
        Write-Log "Error: $($_.Exception.Message)" "Red"
    } finally {
        $global:isScanning = $false
        $btnStartScan.IsEnabled = $true
        $btnStopScan.IsEnabled = $false
        Update-Progress -Value 0 -Text "Ready to start"
    }
}

# ================= ОБРАБОТЧИКИ СОБЫТИЙ =================
$btnStartScan.Add_Click({ if (-not $global:isScanning) { Start-Diagnostics } })
$btnStopScan.Add_Click({ $global:scanCancelled = $true; Write-Log "⏹️ Stopping diagnostics..." "Yellow" })
$btnClearLog.Add_Click({ $txtLog.Clear() })
$btnOpenReports.Add_Click({ if (Test-Path $global:outputFolder) { Start-Process explorer $global:outputFolder } else { [System.Windows.MessageBox]::Show("Reports folder not found. Run diagnostics first.", "Information", "OK", "Information") } })
$btnExportResults.Add_Click({ 
    if ($lvResults.ItemsSource -eq $null) { 
        [System.Windows.MessageBox]::Show("No data to export. Run diagnostics first.", "Information", "OK", "Information")
        return
    }
    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveDialog.FileName = "CARE_Results_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    if ($saveDialog.ShowDialog()) {
        try {
            $lvResults.ItemsSource | Export-Csv -Path $saveDialog.FileName -Encoding UTF8 -NoTypeInformation
            [System.Windows.MessageBox]::Show("Results exported successfully!", "Success", "OK", "Information")
        } catch { [System.Windows.MessageBox]::Show("Export error: $_", "Error", "OK", "Error") }
    }
})
$btnShowConsole.Add_Click({ try { [HideConsole]::Show() } catch {} })
$btnOpenDeviceManager.Add_Click({ Start-Process "devmgmt.msc" })
$btnRefreshDevices.Add_Click({ Update-ProblemDevicesList })
$btnOpenServices.Add_Click({ Start-Process "services.msc" })
$btnRefreshServices.Add_Click({ Update-ProblemServicesList })
$btnStartSelectedService.Add_Click({ Start-SelectedService })
$btnOpenEventViewer.Add_Click({ Open-EventViewerWithHint -topErrors $global:lastTopErrors })

# Автообновление
$chkAutoRefresh.Add_Checked({
    if ($global:autoRefreshTimer) { $global:autoRefreshTimer.Stop() }
    $global:autoRefreshTimer = New-Object System.Windows.Threading.DispatcherTimer
    $global:autoRefreshTimer.Interval = [TimeSpan]::FromSeconds(10)
    $global:autoRefreshTimer.Add_Tick({
        $selectedTab = $tabControl.SelectedIndex
        switch ($selectedTab) {
            2 { Update-ProblemDevicesList }
            3 { Update-ProblemServicesList }
        }
    })
    $global:autoRefreshTimer.Start()
})
$chkAutoRefresh.Add_Unchecked({ if ($global:autoRefreshTimer) { $global:autoRefreshTimer.Stop(); $global:autoRefreshTimer = $null } })

$global:window.Add_Closed({ if ($global:autoRefreshTimer) { $global:autoRefreshTimer.Stop() } })

# ================= ЗАПУСК =================
Write-Host "✓ Starting CARE v1.5 - Call AI for Report..." -ForegroundColor Green

# Первоначальная загрузка
Update-ProblemDevicesList
Update-ProblemServicesList

# Сворачиваем консоль через 2 секунды
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(2)
$timer.Add_Tick({ $timer.Stop(); try { [HideConsole]::Minimize() } catch {} })
$timer.Start()

$global:window.ShowDialog() | Out-Null