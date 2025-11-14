<#
.SYNOPSIS
    A modern, WPF-based Serial Port Monitor and Real-time Data Plotter application.
.DESCRIPTION
    This script launches a GUI for monitoring and communicating with serial COM ports.
    It includes two tabs: one for standard text-based monitoring and another for
    plotting incoming numerical data on a live graph.
.NOTES
    Requires .NET Framework (included with modern Windows versions).
    The plotter expects numerical data, typically one value per line, from the serial port.
    
    Version 1.1.0: 
                    * Added Live-Typing Mode for sending data instantly after each keypress.
                    * Fixed line ending character
                    * Fixed Version number

    Version 1.1.1:
                    * Use Append instead of AppendLine to add text to the text box when not using timestamp, to avoid line breaks.

    Version 1.2.0:
                    * Change the behaviour of Live-Typing Mode, where:
                        - Input Textbox section is hidden in Live-Typing Mode, and uses the output textbox as it's input.
                        - Added some special key and character support.

                    TODO Use timestamp for X-axis in plotter instead of point index.
                    TODO Add option to save log to file.
                    TODO Add some more special key/character support to the list.
#>

# --- VERSION INFO ---
$Version = "1.2.0"

# --- Welcome Message ---
function Show-WelcomeMessage {
    param(
        [string]$Version
    )
    $boxContent = @(
        "",
        "   ┏━━━━━┓              WPF Serial Monitor & Plotter",
        "   ┃>~~~ ┃╮╭╮╭╮         Version: $Version",
        "   ┃>~~~ ┃╰╯╰╯╰         Author: Arif Rachmat",
        "   ┗━━━━━┻━━━━━         License: MIT",
        ""
    )

    # Find the longest line to determine the required width, then add padding
    $maxWidth = ($boxContent | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    $padding = 4 # 2 spaces of padding on each side
    $innerWidth = $maxWidth + $padding

    $topBorder = "╔" + ("═" * $innerWidth) + "╗"
    $bottomBorder = "╚" + ("═" * $innerWidth) + "╝"
    $fullBoxWidth = $innerWidth + 2

    Write-Host $topBorder -ForegroundColor Green
    foreach ($line in $boxContent) {
        # Pad the current line with spaces to match the full inner width
        $paddedLine = $line.PadRight($innerWidth)
        Write-Host ("║" + $paddedLine + "║") -ForegroundColor Green
    }

    Write-Host $bottomBorder -ForegroundColor Green
    Write-Host " A modern GUI for monitoring serial COM ports with real-time data plotting. 📈`n"
    Write-Host " ## USAGE" -ForegroundColor Yellow
    Write-Host " * Select a COM Port and Baud Rate from the dropdown menus."
    Write-Host " * Click 'Connect' to start receiving data."
    Write-Host " * The 'Plotter' tab will graph any numerical data sent over the serial line.`n"
    Write-Host ("-" * $fullBoxWidth)
    Write-Host "Initializing GUI..." -ForegroundColor Cyan
}
Show-WelcomeMessage -Version $Version

# --- ASSEMBLY LOADING ---
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms #for MessageBox

# --- XAML UI DEFINITION ---
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="WPF Serial Monitor Plotter" Height="600" Width="800"
        MinHeight="500" MinWidth="650"
        Background="#2D2D30">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#3E3E42"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#4F4F53"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Opacity" Value="0.5"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="ConnectButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#007ACC"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#108EE2"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="#1E1E1E"/>
            <Setter Property="Foreground" Value="#000000"/>
            <Setter Property="BorderBrush" Value="#555555"/>
        </Style>
         <Style TargetType="TextBox">
            <Setter Property="Background" Value="#3E3E42"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="CaretBrush" Value="#F1F1F1"/>
            <Setter Property="Padding" Value="3"/>
        </Style>
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Background" Value="#3E3E42"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border Name="Border" BorderThickness="1,1,1,0" BorderBrush="#555555" CornerRadius="3,3,0,0" Margin="0,0,2,0">
                            <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center" HorizontalAlignment="Center" ContentSource="Header" Margin="10,2"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#1E1E1E" />
                            </Trigger>
                            <Trigger Property="IsSelected" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#333337" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="Margin" Value="0,0,10,0"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        
        <Border Grid.Row="0" Background="#333337" CornerRadius="3" Padding="10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="COM Port:" Foreground="#F1F1F1" VerticalAlignment="Center" Margin="0,0,5,0"/>
                    <ComboBox x:Name="ComPortComboBox" Width="100" VerticalContentAlignment="Center"/>
                    <Button x:Name="RefreshButton" Content="&#x27F3;" Width="30" Margin="5,0,0,0" ToolTip="Refresh COM Ports" FontWeight="Bold"/>
                    <TextBlock Text="Baud Rate:" Foreground="#F1F1F1" VerticalAlignment="Center" Margin="15,0,5,0"/>
                    <ComboBox x:Name="BaudRateComboBox" Width="100" VerticalContentAlignment="Center"/>
                    <Button x:Name="ConnectButton" Content="Connect" Width="90" Margin="20,0,0,0" Style="{StaticResource ConnectButton}"/>
                </StackPanel>

                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                    <Ellipse x:Name="LiveTypingIndicator" Width="12" Height="12" Fill="LimeGreen" Margin="0,0,5,0"
                             ToolTip="Live Typing is ON" Visibility="Collapsed"/>
                    <TextBlock x:Name="LiveTypingIndicatorLabel" Text="Live Typing ON" Foreground="LimeGreen" 
                               FontWeight="Bold" VerticalAlignment="Center" Visibility="Collapsed"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <TabControl Grid.Row="1" Margin="0,10,0,0">
            <TabItem Header="Serial Monitor">
                 <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <ScrollViewer x:Name="OutputScrollViewer" Grid.Row="0" Margin="0,10,0,0" VerticalScrollBarVisibility="Auto">
                        <TextBox x:Name="OutputTextBox" IsReadOnly="True" TextWrapping="Wrap"
                                 FontFamily="Consolas" FontSize="12"
                                 Background="#1E1E1E" Foreground="#DCDCDC"
                                 BorderThickness="0" VerticalScrollBarVisibility="Auto"
                                 CaretBrush = "Transparent"/>
                    </ScrollViewer>
                    
                    <Grid x:Name="SendingGrid" Grid.Row="1" Margin="0,10,0,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="SendTextBox" Grid.Column="0" IsEnabled="False" VerticalContentAlignment="Center"/>
                        <ComboBox x:Name="LineEndingComboBox" Grid.Column="1" Width="130" Margin="10,0,0,0" VerticalContentAlignment="Center"/>
                        <Button x:Name="SendButton" Grid.Column="2" Content="Send" Width="80" Margin="10,0,0,0" IsEnabled="False"/>
                    </Grid>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,10,0,0">
                        <Button x:Name="ClearButton" Content="Clear Output" Width="100"/>
                        <CheckBox x:Name="EnableTimestampCheckBox" Content="Enable Timestamps" Margin="20,0,0,0"/>
                        <CheckBox x:Name="UseRelativeTimeCheckBox" Content="Use Relative Time" Margin="20,0,0,0"/>
                        <CheckBox x:Name="LiveTypingCheckBox" Content="Live Typing" Margin="20,0,0,0" IsEnabled="False"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            <TabItem Header="Plotter">
                <Border Margin="0,10,0,0" Background="#FF1E1E1E" CornerRadius="5">
                    <Canvas x:Name="GraphCanvas" ClipToBounds="True"/>
                </Border>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@

# --- WPF INITIALIZATION ---
# Use the WPF XamlReader to parse the string, which correctly handles XML entities.
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# --- GLOBAL VARIABLES & THREADING OBJECTS ---
$global:serialPort = $null
$global:textDataQueue = New-Object System.Collections.Concurrent.ConcurrentQueue[string]
$global:plotDataQueue = New-Object System.Collections.Concurrent.ConcurrentQueue[double]
$global:plotPoints = New-Object System.Collections.Generic.List[double]
$maxPlotPoints = 500 # Adjust this to show more or less history on the graph
$global:uiTimer = New-Object System.Windows.Threading.DispatcherTimer
$global:eventSubscriber = $null
$global:enableTimestamps = $false
$global:useRelativeTime = $false
$global:liveTypingEnabled = $false

# --- FIND UI ELEMENTS ---
$controls = @{
    ComPortComboBox     = $window.FindName("ComPortComboBox")
    BaudRateComboBox    = $window.FindName("BaudRateComboBox")
    ConnectButton       = $window.FindName("ConnectButton")
    OutputTextBox       = $window.FindName("OutputTextBox")
    SendingGrid         = $window.FindName("SendingGrid")
    SendTextBox         = $window.FindName("SendTextBox")
    SendButton          = $window.FindName("SendButton")
    ClearButton         = $window.FindName("ClearButton")
    OutputScrollViewer  = $window.FindName("OutputScrollViewer")
    RefreshButton       = $window.FindName("RefreshButton")
    LineEndingComboBox  = $window.FindName("LineEndingComboBox")
    GraphCanvas         = $window.FindName("GraphCanvas")
    EnableTimestampCheckBox = $window.FindName("EnableTimestampCheckBox")
    UseRelativeTimeCheckBox = $window.FindName("UseRelativeTimeCheckBox")
    LiveTypingCheckBox      = $window.FindName("LiveTypingCheckBox")
    LiveTypingIndicator     = $window.FindName("LiveTypingIndicator")
    LiveTypingIndicatorLabel = $window.FindName("LiveTypingIndicatorLabel")
}

# --- INITIAL CONTROL SETUP ---
$populateComPorts = {
    $selectedPort = $controls.ComPortComboBox.SelectedItem
    try {
        $availablePorts = [System.IO.Ports.SerialPort]::GetPortNames()
        $controls.ComPortComboBox.ItemsSource = $availablePorts
        if ($availablePorts.Count -gt 0) {
            if ($selectedPort -and $availablePorts -contains $selectedPort) { $controls.ComPortComboBox.SelectedItem = $selectedPort }
            else { $controls.ComPortComboBox.SelectedIndex = 0 }
        }
    } catch {
        # Handle cases where GetPortNames might fail (e.g., driver issues)
        $controls.ComPortComboBox.ItemsSource = @()
    }
}
$populateComPorts.Invoke()

$controls.BaudRateComboBox.ItemsSource = @("300", "1200", "2400", "4800", "9600", "19200", "38400", "57600", "74880", "115200", "230400", "250000")
$controls.BaudRateComboBox.SelectedItem = "9600"
$controls.LineEndingComboBox.ItemsSource = @("No line ending", "Newline (\n)", "Carriage return (\r)", "Both NL & CR (\r\n)")
$controls.LineEndingComboBox.SelectedItem = "Newline (\n)"

# --- EVENT HANDLERS ---
$controls.RefreshButton.Add_Click({
    $populateComPorts.Invoke()
    $controls.OutputTextBox.AppendText("[INFO] COM Port list refreshed.`n")
})

# Setup a timer that will run on the UI thread to process the queues
$global:uiTimer.Interval = [TimeSpan]::FromMilliseconds(10) # ~100 FPS
$global:uiTimer.Add_Tick({
    # --- 1. Process Text Queue for Serial Monitor ---
    $output = New-Object System.Text.StringBuilder
    $result = ""
    while ($global:textDataQueue.TryDequeue([ref]$result)) {
        [void]$output.Append($result)
    }
    if($output.Length -gt 0) {
        $controls.OutputTextBox.AppendText($output.ToString())
        $controls.OutputScrollViewer.ScrollToEnd()
    }

    # --- 2. Process Data Queue for Plotter ---
    $newDataPoint = 0.0
    $newDataAvailable = $false # Flag to track if we need to redraw
    while($global:plotDataQueue.TryDequeue([ref]$newDataPoint)) {
        $global:plotPoints.Add($newDataPoint)
        if ($global:plotPoints.Count -gt $maxPlotPoints) {
            $global:plotPoints.RemoveAt(0)
        }
        $newDataAvailable = $true # Set flag if we process at least one point
    }

    # --- 3. Redraw the Graph ---
    # Only redraw if new data was actually added and the canvas is ready to be drawn on
    if ($newDataAvailable -and $controls.GraphCanvas.ActualWidth -gt 0 -and $global:plotPoints.Count -ge 2) {
        $controls.GraphCanvas.Children.Clear()

        $canvasWidth = $controls.GraphCanvas.ActualWidth
        $canvasHeight = $controls.GraphCanvas.ActualHeight
        $padding = 40 # Reserve space for labels
        
        $plotAreaWidth = $canvasWidth - (2 * $padding)
        $plotAreaHeight = $canvasHeight - (2 * $padding)

        $min = $global:plotPoints[0]; $max = $global:plotPoints[0]
        foreach($point in $global:plotPoints) {
            if($point -lt $min) {$min = $point}
            if($point -gt $max) {$max = $point}
        }
        $range = $max - $min
        if ($range -eq 0) { $range = 1 }

        $yTickCount = 5
        $xTickCount = 5
        $gridlineBrush = ([System.Windows.Media.BrushConverter]::new()).ConvertFromString("#404040")
        $labelBrush = [System.Windows.Media.Brushes]::LightGray

        for($i = 0; $i -lt $yTickCount; $i++) {
            $value = $min + ($i * ($range / ($yTickCount - 1)))
            $y = ($canvasHeight - $padding) - ($i * ($plotAreaHeight / ($yTickCount - 1)))
            $gridLineY = New-Object System.Windows.Shapes.Line -Property @{ X1=$padding; Y1=$y; X2=$canvasWidth - $padding; Y2=$y; Stroke=$gridlineBrush; StrokeThickness=1 }
            $controls.GraphCanvas.Children.Add($gridLineY)
            $labelY = New-Object System.Windows.Controls.TextBlock -Property @{ Text=$value.ToString("F2"); Foreground=$labelBrush }
            [System.Windows.Controls.Canvas]::SetLeft($labelY, 5); [System.Windows.Controls.Canvas]::SetTop($labelY, $y - 10)
            $controls.GraphCanvas.Children.Add($labelY)
        }
        
        for($i = 0; $i -lt $xTickCount; $i++) {
            $pointIndex = [math]::Round($i * (($global:plotPoints.Count - 1) / ($xTickCount - 1)))
            $x = $padding + ($i * ($plotAreaWidth / ($xTickCount - 1)))
            $gridLineX = New-Object System.Windows.Shapes.Line -Property @{ X1=$x; Y1=$padding; X2=$x; Y2=$canvasHeight - $padding; Stroke=$gridlineBrush; StrokeThickness=1 }
            $controls.GraphCanvas.Children.Add($gridLineX)
            $labelX = New-Object System.Windows.Controls.TextBlock -Property @{ Text=$pointIndex; Foreground=$labelBrush; TextAlignment="Center"; Width=50 }
            [System.Windows.Controls.Canvas]::SetLeft($labelX, $x - 25); [System.Windows.Controls.Canvas]::SetTop($labelX, $canvasHeight - $padding + 5)
            $controls.GraphCanvas.Children.Add($labelX)
        }

        $polyline = New-Object System.Windows.Shapes.Polyline
        $polyline.Stroke = [System.Windows.Media.Brushes]::LimeGreen
        $polyline.StrokeThickness = 2
        $points = New-Object System.Windows.Media.PointCollection
        for ($i = 0; $i -lt $global:plotPoints.Count; $i++) {
            $x = $padding + ($i / ($global:plotPoints.Count - 1)) * $plotAreaWidth
            $y = ($canvasHeight - $padding) - (($global:plotPoints[$i] - $min) / $range * $plotAreaHeight)
            $points.Add([System.Windows.Point]::new($x, $y))
        }
        $polyline.Points = $points
        $controls.GraphCanvas.Children.Add($polyline)
    }
})

# Initialize timestamp checkboxes
$controls.EnableTimestampCheckBox.IsChecked = $global:enableTimestamps
$controls.EnableTimestampCheckBox.Add_Checked({ $global:enableTimestamps = $true })
$controls.EnableTimestampCheckBox.Add_Unchecked({ $global:enableTimestamps = $false })
$controls.UseRelativeTimeCheckBox.IsChecked = $global:useRelativeTime
$controls.UseRelativeTimeCheckBox.Add_Checked({ $global:useRelativeTime = $true })
$controls.UseRelativeTimeCheckBox.Add_Unchecked({ $global:useRelativeTime = $false })

# Live Typing Mode Checkbox Handler
$controls.LiveTypingCheckBox.Add_Click({
    $global:liveTypingEnabled = $controls.LiveTypingCheckBox.IsChecked
    if ($global:liveTypingEnabled) {
        $controls.LiveTypingIndicator.Visibility = 'Visible'
        $controls.LiveTypingIndicatorLabel.Visibility = 'Visible'
        $controls.SendTextBox.Text = "" # Clear text on mode switch
        $controls.SendingGrid.IsEnabled = $false
        $controls.SendingGrid.Visibility = 'Collapsed'
        $controls.OutputTextBox.IsReadOnly = $false
        $controls.OutputTextBox.focus()
    } else {
        $controls.LiveTypingIndicator.Visibility = 'Collapsed'
        $controls.LiveTypingIndicatorLabel.Visibility = 'Collapsed'
        $controls.OutputTextBox.IsReadOnly = $true
        $controls.SendingGrid.Visibility = 'Visible'
        $controls.SendingGrid.IsEnabled = $true
        $controls.SendTextBox.focus()
    }
})

# Toggle connection state
$controls.ConnectButton.Add_Click({
    if ($global:serialPort -and $global:serialPort.IsOpen) {
        # --- DISCONNECT ---
        $global:uiTimer.Stop()
        if ($global:eventSubscriber) { Unregister-Event -SubscriptionId $global:eventSubscriber.Id -ErrorAction SilentlyContinue; $global:eventSubscriber = $null }
        $global:serialPort.Close(); $global:serialPort.Dispose(); $global:serialPort = $null

        while ($global:textDataQueue.TryDequeue([ref]$null)) {}
        while ($global:plotDataQueue.TryDequeue([ref]$null)) {}
        
        $controls.ConnectButton.Content = "Connect"
        $controls.ComPortComboBox.IsEnabled = $true
        $controls.RefreshButton.IsEnabled = $true
        $controls.BaudRateComboBox.IsEnabled = $true
        $controls.SendTextBox.IsEnabled = $false
        $controls.SendButton.IsEnabled = $false
        $controls.LiveTypingCheckBox.IsChecked = $false
        $controls.LiveTypingCheckBox.IsEnabled = $false
        $controls.LiveTypingIndicator.Visibility = 'Collapsed' # Ensure indicator is off
        $controls.LiveTypingIndicatorLabel.Visibility = 'Collapsed'
        $controls.OutputTextBox.AppendText("`n[INFO] Disconnected.`n")
    }
    else {
        # --- CONNECT ---
        if (-not $controls.ComPortComboBox.SelectedItem) { [System.Windows.Forms.MessageBox]::Show("No COM Port selected.", "Error", "OK", "Error"); return }
        
        $portName = $controls.ComPortComboBox.SelectedItem.ToString()
        $baudRate = [int]$controls.BaudRateComboBox.SelectedItem.ToString()
        $global:serialPort = New-Object System.IO.Ports.SerialPort($portName, $baudRate)
        
        try {
            $global:serialPort.Open()
            $global:plotPoints.Clear()

            $global:startTime = Get-Date
            $global:eventSubscriber = Register-ObjectEvent -InputObject $global:serialPort -EventName DataReceived -Action {
                try {
                    $port = $event.Sender
                    if ($port.IsOpen) {
                        $receivedText = $port.ReadExisting()
                        if (-not [string]::IsNullOrEmpty($receivedText)) {
                            $lines = $receivedText -split "(`r?`n)"
                            $processedText = New-Object System.Text.StringBuilder

                            if($global:enableTimestamps) {
                                foreach ($line in $lines) {
                                    if ($line.Trim().Length -gt 0) {
                                        $timestamp = ""
                                        $now = Get-Date
                                        if ($global:useRelativeTime) {
                                            $elapsed = $now - $global:startTime
                                            $timestamp = "[{0:hh\:mm\:ss\.fff}] " -f $elapsed
                                        } else {
                                            $timestamp = "[{0:HH:mm:ss.fff}] " -f $now
                                        }
                                        $processedText.Append("`n$timestamp$line")
                                    }
                                }
                            }
                            else {
                                $processedText.Append("$receivedText")
                            }

                            $global:textDataQueue.Enqueue($processedText.ToString())

                            foreach ($line in $lines) {
                                $number = 0.0
                                if ([double]::TryParse($line.Trim(), [ref]$number)) {
                                    $global:plotDataQueue.Enqueue($number)
                                }
                            }
                        }
                    }
                } catch {}
            }

            $global:uiTimer.Start()
            
            $controls.ConnectButton.Content = "Disconnect"
            $controls.ComPortComboBox.IsEnabled = $false
            $controls.RefreshButton.IsEnabled = $false
            $controls.BaudRateComboBox.IsEnabled = $false
            $controls.SendTextBox.IsEnabled = $true
            $controls.SendButton.IsEnabled = $true
            $controls.LiveTypingCheckBox.IsEnabled = $true
            $controls.OutputTextBox.AppendText("[INFO] Connected to $portName at $baudRate baud.`n")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to open port '$portName'.`nIt might be in use by another application.", "Connection Error", "OK", "Error")
            if ($global:serialPort) { $global:serialPort.Dispose(); $global:serialPort = $null }
        }
    }
})

# Action to send data
$sendAction = {
    if (-not $global:liveTypingEnabled -and $global:serialPort -and $global:serialPort.IsOpen -and $controls.SendTextBox.Text.Length -gt 0) {
        $textToSend = $controls.SendTextBox.Text
        switch ($controls.LineEndingComboBox.SelectedItem) {
            "Newline (\n)"        { $textToSend += "`n" }
            "Carriage return (\r)"{ $textToSend += "`r" }
            "Both NL & CR (\r\n)" { $textToSend += "`r`n" }
        }
        $global:serialPort.Write($textToSend)
        $controls.SendTextBox.Clear()
    }
}
$controls.SendButton.Add_Click($sendAction)

# Send on Enter key, only if not in live typing mode
$controls.SendTextBox.Add_KeyDown({ param($senderID, $e)
    if ($e.Key -eq 'Enter' -and -not $global:liveTypingEnabled) { $sendAction.Invoke() }
})

# --- Live Typing Mode ---
$controls.OutputTextBox.add_PreviewTextInput({
    param($senderID, $e)
    if ($global:serialPort.IsOpen) {
        try {
            $global:serialPort.Write($e.Text)
            $e.Handled = $true
        }
        catch {
            Write-Warning "Error Writing to Port: $($_.Exception.Message)"
        }
    }
})

# Special keys Handler
$controls.OutputTextBox.add_PreviewKeyDown({
    param($senderID, $e)
    if ($global:serialPort.IsOpen) {
        try {
            switch ($e.Key) {
                'Space'     { 
                    $global:serialPort.Write([char]32) 
                    $e.Handled = $true
                }
                'Back'      { 
                    $global:serialPort.Write([char]8) 
                    $e.Handled = $true 
                }
                'Tab'       { 
                    $global:serialPort.Write([char]9) 
                    $e.Handled = $true 
                }
                'Delete'    { 
                    $global:serialPort.Write([char]127) 
                    $e.Handled = $true 
                }
                'Left'      { 
                    $global:serialPort.Write([char]27) 
                    $global:serialPort.Write([char]91) 
                    $global:serialPort.Write([char]68) 
                    $e.Handled = $true 
                }
                'Right'     { 
                    $global:serialPort.Write([char]27) 
                    $global:serialPort.Write([char]91) 
                    $global:serialPort.Write([char]67) 
                    $e.Handled = $true 
                }
                'Down'      { 
                    $global:serialPort.Write([char]27) 
                    $global:serialPort.Write([char]91) 
                    $global:serialPort.Write([char]66) 
                    $e.Handled = $true 
                }
                'Up'        { 
                    $global:serialPort.Write([char]27) 
                    $global:serialPort.Write([char]91) 
                    $global:serialPort.Write([char]65) 
                    $e.Handled = $true 
                }
                Default {}
            }
        }
        catch {
            Write-Warning "Error Writing to Port: $($_.Exception.Message)"
        }
    }
})

# Clear the output window
$controls.ClearButton.Add_Click({
    $controls.OutputTextBox.Clear()
})

# Cleanup on form closing
$window.Add_Closing({
    if ($global:serialPort -and $global:serialPort.IsOpen) {
        $global:uiTimer.Stop()
        if ($global:eventSubscriber) { Unregister-Event -SubscriptionId $global:eventSubscriber.Id -ErrorAction SilentlyContinue }
        $global:serialPort.Close()
        $global:serialPort.Dispose()
    }
})


# --- SHOW THE WINDOW ---
[void]$window.ShowDialog()