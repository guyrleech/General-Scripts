#requires -version 3

<#
.SYNOPSIS

Display digitial clock with granularity in seconds or a stopwatch or countdown timer in a window with ability to stop/start/reset

.PARAMETER stopwatch

Run a stopwatch rather than showing current time

.PARAMETER start

Start the stopwatch immediately

.PARAMETER notOnTop

Do not place the window on top of all other windows

.PARAMETER expiredMessage

Text to display in a popup when the countdown timer expires

.PARAMETER markerFile

A file to look for which will be then seen in a SysInternals Process Monitor trace as a CreateFile operation to allow cross referencing to that

.PARAMETER countdown

Run a countdown timer starting at the value specified as hh:mm:ss

.PARAMETER tenths

Show tenths of a second for the clock

.PARAMETER beep

Emit a beep of the duration specified in milliseconds when the countdown timer expires

.EXAMPLE

& '.\Digital Clock.ps1'

Display an updating digital clock in a window

.EXAMPLE

& '.\Digital Clock.ps1' -tenths

Display an updating digital clock in a window with tenths of a second granularity

.EXAMPLE

& '.\Digital Clock.ps1' -stopwatch

Display a stopwatch in a window but do not start it until the Run checkbox is checked

.EXAMPLE

& '.\Digital Clock.ps1' -stopwatch -start

Display a stopwatch in a window and start it immediately

.EXAMPLE

& '.\Digital Clock.ps1' -countdown 00:03:00 -beep 2000

Display a countdown timer starting at 3 minutes in a window but do not start it until the Run checkbox is checked. When the timer expires, sound a beep for 2 seconds

.NOTES

    Modification History:

    @guyrleech 14/05/2020  Initial release
                           Rewrote to use WPF DispatcherTimer rather than runspaces
                           Added marker functionality
    @guyrleech 15/05/2020  Pressing C puts existing marker items onto the Windows clipboard
    @guyrleech 21/05/2020  Added Clear button, other GUI adjustments
    @guyrleech 22/05/2020  Forced date to 24 hour clock as problem reported with Am/PM indicators when using date format "T"
    @guyrleech 25/05/2020  Added edit and delete context menu items for markers
                           Fixed resizing regression
                           Added countdown timer with -beep and -countdown
    @guyrleech 27/05/2020  Fixed bug with 01:00:00 countdown & added validation to countdown string passed/entered
    @guyrleech 02/06/2020  Added tenths seconds option to clock
    @guyrleech 07/06/2020  Added insert above/below and save options for markers
    @guyrleech 12/10/2020  Added message box for timer expiry text
                           Fixed countdown timer re-run bug where timer not reset so expires immediately
#>

[CmdletBinding()]

Param
(
    [switch]$stopWatch ,
    [switch]$start ,
    [string]$markerFile ,
    [string]$countdown ,
    [string]$expiredMessage ,
    [switch]$notOnTop ,
    [switch]$tenths ,
    [int]$beep
)

[int]$exitCode = 0

[string]$mainwindowXAML = @'
<Window x:Class="Timer.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Timer"
        mc:Ignorable="d"
        Title="Guy's Clock" Height="272.694" Width="636.432">
    <Grid>
        <TextBox x:Name="txtClock" HorizontalAlignment="Left" Height="111" Margin="24,29,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="358" FontSize="72" IsReadOnly="True" FontWeight="Bold" BorderThickness="0"/>
        <Grid Margin="24,200,244,10">
            <CheckBox x:Name="checkboxRun" Content="_Run" HorizontalAlignment="Left" Height="18" Margin="-11,5,0,0" VerticalAlignment="Top" Width="52" IsChecked="True"/>
            <Button x:Name="btnReset" Content="Re_set" HorizontalAlignment="Left" Height="23" Margin="112,-1,0,0" VerticalAlignment="Top" Width="72"/>
            <Button x:Name="btnMark" Content="_Mark" HorizontalAlignment="Left" Height="23" Margin="35,-1,0,0" VerticalAlignment="Top" Width="72"/>
            <Button x:Name="btnClear" Content="_Clear" HorizontalAlignment="Left" Height="23" Margin="189,-1,0,0" VerticalAlignment="Top" Width="72"/>
            <Button x:Name="btnCountdown" Content="Count _Down" HorizontalAlignment="Left" Height="23" Margin="268,-1,0,0" VerticalAlignment="Top" Width="72"/>

        </Grid>
        <TextBox x:Name="txtMarkerFile" HorizontalAlignment="Left" Height="24" Margin="92,144,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="268"/>
        <Label Content="Marker File" HorizontalAlignment="Left" Height="23" Margin="10,145,0,0" VerticalAlignment="Top" Width="72"/>
        <ListView x:Name="listMarkings" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="382,13,10,10" >
            <ListView.ContextMenu>
                <ContextMenu>
                    <MenuItem Header="Edit" Name="EditContextMenu" />
                    <MenuItem Header="Delete" Name="DeleteContextMenu" />
                    <MenuItem Header="Save" Name="SaveContextMenu" />
                    <MenuItem Header="Insert Above" Name="InsertAboveContextMenu" />
                    <MenuItem Header="Insert Below" Name="InsertBelowContextMenu" />
                </ContextMenu>
            </ListView.ContextMenu>
            <ListView.View>
                <GridView>
                    <GridView.ColumnHeaderContextMenu>
                        <ContextMenu/>
                    </GridView.ColumnHeaderContextMenu>
                    <GridViewColumn Header="Timestamp" DisplayMemberBinding="{Binding Timestamp}"/>
                    <GridViewColumn Header="Notes" DisplayMemberBinding="{Binding Notes}"/>
                </GridView>
            </ListView.View>
        </ListView>
        <CheckBox x:Name="checkboxBeep" Content="Beep" HorizontalAlignment="Left" Height="15" Margin="13,180,0,0" VerticalAlignment="Top" Width="93"/>
    </Grid>
</Window>
'@

[string]$markerTextXAML = @'
<Window x:Class="Timer.Test"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Timer"
        mc:Ignorable="d"
        Title="Marker Text" Height="299.878" Width="589.034" Name="Marker">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="31*"/>
            <ColumnDefinition Width="217*"/>
            <ColumnDefinition Width="544*"/>
        </Grid.ColumnDefinitions>
        <Grid Grid.ColumnSpan="3" Margin="10,20,10,71">
            <TextBox x:Name="textboxTimestamp" HorizontalAlignment="Stretch"  Margin="82,10,10,122" VerticalAlignment="Stretch" AllowDrop="False" />
            <TextBox x:Name="textBoxMarkerText" HorizontalAlignment="Stretch"  Margin="82,62,10,10" TextWrapping="Wrap" VerticalAlignment="Stretch" SpellCheck.IsEnabled="True"/>
            <Label x:Name="labelTimestamp" Content="Timestamp" HorizontalAlignment="Left" Height="28" VerticalAlignment="Top" Width="76" Margin="0,17,0,0"/>
            <Label x:Name="labelMarkText" Content="Marker Text" HorizontalAlignment="Left" Height="28" Margin="0,82,0,0" VerticalAlignment="Top" Width="76" RenderTransformOrigin="0.464,1.5"/>
        </Grid>
        <Grid Grid.ColumnSpan="3" Margin="22,196,252,0">
            <Button x:Name="btnMarkerTextOk" Content="OK" HorizontalAlignment="Left" Height="48" VerticalAlignment="Top" Width="120" IsDefault="True"/>
            <Button x:Name="btnMarkerTextOk_Copy" Content="Cancel" HorizontalAlignment="Left" Height="48" Margin="148,0,10,10" VerticalAlignment="Top" Width="120" IsCancel="True"/>
        </Grid>
    </Grid>
</Window>
'@

[string]$saveXAML = @'
<Window x:Class="Timer.Save"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Timer"
        mc:Ignorable="d"
        Title="Save" Height="220" Width="618.913">
    <Grid>
        <TextBox x:Name="textboxFilename" HorizontalAlignment="Left" Height="34" Margin="88,35,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="415"/>
        <Label Content="File name" HorizontalAlignment="Left" Height="34" Margin="10,35,0,0" VerticalAlignment="Top" Width="65"/>
        <Grid Margin="88,86,218,72">
            <RadioButton x:Name="radiobuttonOverwrite" Content="Overwrite" HorizontalAlignment="Left" Height="21" VerticalAlignment="Top" Width="181" GroupName="OverwriteAppend"/>
            <RadioButton x:Name="radiobuttonAppend" Content="Append" HorizontalAlignment="Left" Height="21" Margin="124,0,0,0" VerticalAlignment="Top" Width="181" GroupName="OverwriteAppend" IsChecked="True"/>
        </Grid>
        <Grid Margin="10,117,20,19">
            <Button x:Name="buttonSaveOk" Content="OK" HorizontalAlignment="Left" Height="43" VerticalAlignment="Top" Width="140" IsDefault="True"/>
            <Button x:Name="buttonSaveCancel" Content="Cancel" HorizontalAlignment="Left" Height="43" Margin="171,0,0,0" VerticalAlignment="Top" Width="140" IsCancel="True"/>
        </Grid>
        <Button x:Name="buttonSaveBrowser" Content="..." HorizontalAlignment="Left" Height="34" Margin="524,35,0,0" VerticalAlignment="Top" Width="50"/>

    </Grid>
</Window>
'@

Function Set-MarkerText
{
    [CmdletBinding()]
    Param
    (
        $item , ## if not passed then a new item otherwise editing
        $timestamp ,
        [int]$selected = -1 ,
        [ValidateSet('Above','Below','None')]
        [string]$insert = 'None'
    )

    if( $markerTextForm = New-Form -inputXaml $markerTextXAML )
    {
        $markerTextForm.TopMost = $true ## got to be on top of the clock itself
        $WPFbtnMarkerTextOk.Add_Click({
            $_.Handled = $true          
            $markerTextForm.DialogResult = $true 
            $markerTextForm.Close()
            })
        
        if( $item )
        {
            $WPFtextBoxMarkerText.Text = $item.Notes
        }
        $WPFtextboxTimestamp.IsReadOnly = ( ! $insert -or $insert -eq 'None' ) ## only enabled for insertions
        $WPFtextboxTimestamp.Text = $( if( $item ) { $item.timestamp } else { $timestamp })
        if( $insert -eq 'None' )
        {
            $WPFtextBoxMarkerText.Focus()
        }
        else
        {
            $WPFtextboxTimestamp.Focus()
        }
        $WPFMarker.Title = "$(if( $item ) { 'Edit' } else { 'Set' }) text for marker @ $(if( $item ) { $item.Timestamp } else { $Timestamp})"

        if( $markerTextForm.ShowDialog() )
        {
            if( $item )
            {
                $item.Notes = $WPFtextBoxMarkerText.Text.ToString()
                $WPFlistMarkings.Items.Refresh()
            }
            elseif( ! $insert -or $insert -eq 'None' ) ## new item
            {
                $null = $WPFlistMarkings.Items.Add( ([pscustomobject]@{ 'Timestamp' = $timestamp ; 'Notes' = $WPFtextBoxMarkerText.Text.ToString() }) )  
            }
            elseif( $insert -eq 'Above' -or $insert -eq 'Below' ) ## insert so need to get currently selected item so we know where to insert it
            {
                ## verify date format entered
                if( [string]::IsNullOrEmpty( $WPFtextboxTimestamp.Text ) )
                {
                    [void][Windows.MessageBox]::Show( "Text `"$($WPFtextboxTimestamp.Text)`" not in correct format or invalid value" , 'Marker Error' , 'Ok' ,'Exclamation' )
                }
                else
                {
                    [string]$dateText = $WPFtextboxTimestamp.Text
                    [string[]]$datestampParts = $dateText -split '\s'
                    if( [string]::IsNullOrEmpty( $countdown ) -and ( ! $datestampParts -or $datestampParts.Count -ne 2 -or $datestampParts[0] -notmatch '^\d\d/\d\d/\d\d(\d\d)?$' -or $datestampParts[1] -notmatch '^[012]\d:[0-5]\d:[0-5]\d(\.\d{1-6})?' ) )
                    {
                        ## see if just a time so we fill in date of selected item
                        if( $datestampParts.Count -eq 1 -and $datestampParts[0] -match '^[012]\d:[0-5]\d:[0-5]\d(\.\d{1-6})?' )
                        {
                            $dateText = "{0} {1}" -f $(if( $selected -ge 0 ) { $WPFlistMarkings.SelectedItem.Timestamp.Split( ' ' )[0] } else { Get-Date -Format d } ), $datestampParts[0]
                        }
                        else
                        {
                            [void][Windows.MessageBox]::Show( "Text `"$($WPFtextboxTimestamp.Text)`" not in correct format or invalid value" , 'Marker Error' , 'Ok' ,'Exclamation' )
                        }
                    }
                    
                    if( $dateText )
                    {
                        [int]$position = $(if( $insert -eq 'Above' ) { $selected } else { $selected + 1 } )
                        if( $position -lt 0 )
                        {
                            $position = 0
                        }

                        $null = $WPFlistMarkings.Items.Insert( $position , ([pscustomobject]@{ 'Timestamp' = $dateText ; 'Notes' = $WPFtextBoxMarkerText.Text.ToString() }) )
                    }
                }
            }
        }
    }
}

Function New-Form
{
    Param
    (
        [Parameter(Mandatory=$true)]
        $inputXaml
    )

    $form = $null
    if( ( $inputXML = $inputXaml -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window' ) `
        -and ( [xml]$xaml = $inputXML ) `
            -and ($reader = New-Object -TypeName Xml.XmlNodeReader -ArgumentList $xaml ) )
    {
        try
        {
            $form = [Windows.Markup.XamlReader]::Load( $reader )
        }
        catch
        {
            Throw "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
        }
 
        $xaml.SelectNodes( '//*[@Name]' ) | ForEach-Object `
        {
            if( $value = $Form.FindName($_.Name) )
            {
                Set-Variable -Name "WPF$($_.Name)" -Value $value -Scope Script
            }
        }
    }
    else
    {
        Throw 'Failed to convert input XAML to WPF XML'
    }

    $form
}

Function Set-ForegroundWindow
{
    Param
    (
        [Parameter(Mandatory)]
        [int]$thePid = $pid
    )

    ## Windows may not be visible so make it so
    $pinvokeCode = @'
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow); 
        [DllImport("user32.dll", SetLastError=true)]
        public static extern int SetForegroundWindow(IntPtr hwnd);
'@

    if( ! ([System.Management.Automation.PSTypeName]'Win32.User32').Type )
    {
        Add-Type -MemberDefinition $pinvokeCode -Name 'User32' -Namespace 'Win32' -UsingNamespace System.Text -ErrorAction Stop
    }

    if( $process = Get-Process -Id $thePid )
    {
        if( ( [intptr]$windowHandle = $process.MainWindowHandle ) -and $windowHandle -ne [intptr]::Zero )
        {
            [int]$operation = 9
            [bool]$setForegroundWindow = [win32.user32]::ShowWindowAsync( $windowHandle , $operation ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if( ! $setForegroundWindow )
            {
                Write-Warning -Message "Failed to set window to foreground for process $($process.Name) (pid $($process.Id)): $lastError"
            }
            else
            {
                [int]$foregroundWindow = [win32.user32]::SetForegroundWindow( $windowHandle ); $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                Write-Verbose -Message "Operation $operation on $($process.Name) (pid $($process.Id)) succeeded"
            }
        }
        else
        {
            Write-Warning -Message "No main window handle for process $($process.Name) (pid $($process.Id))"
        }
    }
    ## else ## no process for pid but will have errored
}

Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Windows.Forms

if( $PSBoundParameters[ 'stopwatch' ] -and $PSBoundParameters[ 'countdown' ] )
{
    Throw "Cannot have both -stopwatch and -countdown"
}

if( ! ( $Form = New-Form -inputXaml $mainwindowXAML ) )
{
    Exit 1
}

[string]$dateFormat = $(if( $tenths )
{
    'HH:mm:ss.f'
}
else
{
    'HH:mm:ss'
})

$WPFcheckboxBeep.IsEnabled = $null -ne $PSBoundParameters[ 'countdown' ]
$WPFcheckboxBeep.IsChecked = $null -ne $PSBoundParameters[ 'beep' ]

$form.TopMost = ! $notOnTop
$form.Title = $(if( $stopWatch ) { 'Guy''s Stopwatch' } elseif( $countdown ) { 'Guy''s Countdown Timer' } else { 'Guy''s Clock' })

[int]$countdownSeconds = 0
[string]$global:lastfilename = $null

$WPFtxtClock.Text = $(if( $stopWatch )
    {
        '00:00:00.0'
    }
    elseif( $countdown )
    {
        if( $countdown -match '^(\d{1,2}):(\d{1,2}):(\d{1,2})$' )
        {
            if( [int]$Matches[2] -ge 60 -or [int]$Matches[3] -ge 60 )
            {
                Throw "Bad countdown time `"$countdown`""
            }
            $countdownSeconds = [int]$Matches[1] * 3600 + [int]$Matches[2] * 60 + [int]$Matches[3]
            ## reconstitute in case didn't have leading zeroes e.g. 0:3:0
            [timespan]$timespan = [timespan]::FromSeconds( $script:countdownSeconds )
            '{0:d2}:{1:d2}:{2:d2}' -f [int][math]::Floor( $timespan.TotalHours ) , $timespan.Minutes , $timespan.Seconds
        }
        else
        {
            Throw "Countdown period must be specified in hh:mm:ss"
        }
    }
    else
    {
        Get-Date -Format $dateFormat
    })

$WPFbtnReset.Add_Click({ 
    $_.Handled = $true
    $timer.Reset()
    if( $WPFcheckboxRun.IsChecked )
    {
        $timer.Start() 
    }
    elseif( $stopWatch )
    {
        $WPFtxtClock.Text = '00:00:00.0'
    }
    else
    {
        $WPFtxtClock.Text = $countdown
    }})

$WPFbtnReset.IsEnabled = $stopWatch -or $countdown

$WPFbtnClear.Add_Click({
    $WPFlistMarkings.Items.Clear()
})

$WPFbtnMark.Add_Click({
    $_.Handled = $true

    [string]$timestamp = $(if( $stopWatch )
    {
        '{0:d2}:{1:d2}:{2:d2}.{3:d3}' -f $timer.Elapsed.Hours , $timer.Elapsed.Minutes , $timer.Elapsed.Seconds, $timer.Elapsed.Milliseconds
    }
    elseif( $countdown )
    {
        if( ( [int]$secondsLeft = $countdownSeconds - $timer.Elapsed.TotalSeconds) -le 0 )
        {
            $secondsLeft = 0
        }
        [timespan]$timespan = [timespan]::FromSeconds( $secondsLeft )
        '{0:d2}:{1:d2}:{2:d2}' -f [int][math]::Floor( $timespan.TotalHours ) , $timespan.Minutes , $timespan.Seconds
    }
    else
    {
        "{0} {1}" -f (Get-Date -Format d) , (Get-Date -Format 'HH:mm:ss.ffffff')
    } )
    
    Write-Verbose -Message "Mark button pressed, timestamp $timestamp"

    ## if file exists then read else write it
    if( ! [string]::IsNullOrEmpty( $WPFtxtMarkerFile.Text ) )
    {
        ## SysInternals Process Monitor will see this so can be cross referenced to here
        Test-Path -Path (([Environment]::ExpandEnvironmentVariables( $WPFtxtMarkerFile.Text ))) -ErrorAction SilentlyContinue
    }
    ## add current time/stopwatch to gridview
    Set-MarkerText -timestamp $timestamp
})

$WPFcheckboxRun.Add_Click({
    $_.Handled = $true
    if( $stopWatch -or $countdown )
    {
        if( $WPFcheckboxRun.IsChecked )
        {
            if( $countdown )
            {
                $WPFbtnCountdown.IsEnabled = $false
            }
            $timer.Start() 
        }
        else
        {
            if( $countdown )
            {
                $WPFbtnCountdown.IsEnabled = $true
            }
            $timer.Stop() 
        }
    }})

$form.add_KeyDown({
    Param
    (
        [Parameter(Mandatory)][Object]$sender,
        [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$event
    )
    if( $event -and $event.Key -eq 'Space' )
    {
        $_.Handled = $true
        $WPFcheckboxRun.IsChecked = ! $WPFcheckboxRun.IsChecked
        if( $stopWatch -or $countdown )
        {
            if( $WPFcheckboxRun.IsChecked )
            {
                $timer.Start()
            }
            else
            {
                $timer.Stop()
            }
        }
    }
    elseif( $event -and $event.Key -eq 'C' -and $WPFlistMarkings.Items -and $WPFlistMarkings.Items.Count )
    {
        $_.Handled = $true
        $WPFlistMarkings.Items | Out-String | Set-Clipboard
    }
})

$WPFEditContextMenu.Add_Click({
    $_.Handled = $true
    ForEach( $item in $WPFlistMarkings.SelectedItems )
    {
        Set-MarkerText -item $item
    }
})

$WPFlistMarkings.add_MouseDoubleClick({
    $_.Handled = $true
    ForEach( $item in $WPFlistMarkings.SelectedItems )
    {
        Set-MarkerText -item $item
    }
})

$WPFbtnCountdown.Add_Click({
    $_.Handled = $true
    ## unclick the run box, prompt for time and display ready for time run to be clicked
    $WPFcheckboxRun.IsChecked = $false
    ## borrowing the marker text dialogue to input countdown timer time
    if( $markerTextForm = New-Form -inputXaml $markerTextXAML )
    {
        $markerTextForm.TopMost = $true ## got to be on top of the clock itself
        $WPFbtnMarkerTextOk.Add_Click({
            $_.Handled = $true
            
            if( $WPFtextboxTimestamp.Text -notmatch '^(\d{1,2}):(\d{1,2}):(\d{1,2})$' -or [int]$Matches[2] -ge 60 -or [int]$Matches[3] -ge 60)
            {
                [void][Windows.MessageBox]::Show( "Text `"$($WPFtextboxTimestamp.Text)`" not in hh:mm:ss format or invalid values" , 'Countdown Timer Error' , 'Ok' ,'Exclamation' )
            }
            else
            {
                $markerTextForm.DialogResult = $true 
                $markerTextForm.Close()
            }})
        
        $WPFtextboxTimestamp.Text = $countdown
        $WPFtextboxTimestamp.Focus()
        ##$WPFtextBoxMarkerText.IsEnabled = $false
        $WPFMarker.Title = "Enter countdown time in hh:mm:ss"
        $WPFtextBoxMarkerText.Text = $script:expiredMessage
        $WPFlabelTimestamp.Content = 'Countdown'
        $WPFlabelMarkText.Content = 'Expired Text'

        if( $markerTextForm.ShowDialog() )
        {
            if( $WPFtextboxTimestamp.Text -match '^(\d{1,2}):(\d{1,2}):(\d{1,2})$' )
            {
                if( [int]$Matches[2] -ge 60 -or [int]$Matches[3] -ge 60 )
                {
                    [void][Windows.MessageBox]::Show( "Text `"$($WPFtextboxTimestamp.Text)`" contains invalid ours or minutes" , 'Countdown Timer Error' , 'Ok' ,'Exclamation' )
                }
                else
                {
                    $script:countdownSeconds = [int]$Matches[1] * 3600 + [int]$Matches[2] * 60 + [int]$Matches[3]
                    $script:expiredMessage = $WPFtextBoxMarkerText.Text
                    ## reconstitute in case didn't have leading zeroes e.g. 0:3:0
                    [timespan]$timespan = [timespan]::FromSeconds( $script:countdownSeconds )
                    $WPFtxtClock.Text = $script:countdown = '{0:d2}:{1:d2}:{2:d2}' -f [int][math]::Floor( $timespan.TotalHours ) , $timespan.Minutes , $timespan.Seconds
                }
            }
            $WPFbtnReset.IsEnabled = $true
            $WPFcheckboxBeep.IsEnabled = $true
        }
    }
})

$WPFSaveContextMenu.Add_Click({
    $_.Handled = $true
    
    if( ! $WPFlistMarkings.Items.Count )
    {
        [void][Windows.MessageBox]::Show( "Nothing to save" , 'Save Error' , 'Ok' ,'Exclamation' )
    }
    elseif( $saveForm = New-Form -inputXaml $saveXAML )
    {
        $WPFbuttonSaveBrowser.Add_Click({
            $_.Handled = $true
            if( $fileBrowser = New-Object -TypeName System.Windows.Forms.OpenFileDialog )
            {
                $fileBrowser.InitialDirectory = Get-Location -PSProvider FileSystem | Select-Object -ExpandProperty Path
                if( ( $file = $fileBrowser.ShowDialog() ) -eq 'OK' )
                {
                    $WPFtextboxFilename.Text = $fileBrowser.FileName
                }
            }
        })
        $WPFbuttonSaveOk.Add_Click({
            if( [string]::IsNullOrEmpty( $WPFtextboxFilename.Text.Trim() ) )
            {
                [void][Windows.MessageBox]::Show( "Must specify a file name" , 'Save Error' , 'Ok' ,'Exclamation' )
            }
            else
            {
                $saveForm.DialogResult = $true 
                $saveForm.Close()
            }
        })

        $WPFtextboxFilename.Text = $global:lastfilename
        $saveForm.Topmost = $true

        if( $saveForm.ShowDialog() )
        {
            $saveError = $null
            $WPFlistMarkings.Items | Out-File -FilePath ($WPFtextboxFilename.Text -replace '"') -Append:$wpfradiobuttonAppend.IsChecked -ErrorVariable saveError
            $global:lastfilename = $WPFtextboxFilename.Text
            if( $saveError )
            {
                [void][Windows.MessageBox]::Show( "Must specify a file name" , 'Save Error' , 'Ok' ,'Exclamation' )
            }
        }
    }
})

$WPFInsertAboveContextMenu.Add_Click({
    $_.Handled = $true
    Set-MarkerText -insert Above -selected $WPFlistMarkings.SelectedIndex
})

$WPFInsertBelowContextMenu.Add_Click({
    $_.Handled = $true
    Set-MarkerText -insert Below -selected $WPFlistMarkings.SelectedIndex
})

$WPFDeleteContextMenu.Add_Click({
    $_.Handled = $true
    [array]$removals = @( ForEach( $item  in $WPFlistMarkings.SelectedItems )
    {
        $item ## can't remove items whilst enumerating so put in an array
    })
    ForEach( $removal in $removals )
    {
        $WPFlistMarkings.Items.Remove( $removal ) 
    }
})

[scriptblock]$timerBlock = `
{
    if( $WPFcheckboxRun.IsChecked )
    {
        $newTime = $(if( $stopWatch )
        {
            '{0:d2}:{1:d2}:{2:d2}.{3:d1}' -f $timer.Elapsed.Hours , $timer.Elapsed.Minutes , $timer.Elapsed.Seconds, $( [int]$tenths = $timer.Elapsed.Milliseconds / 100 ; if( $tenths -ge 10 ) { 0 } else { $tenths } )
        }
        elseif( $countdown )
        {
            [double]$secondsLeft = $countdownSeconds - $timer.Elapsed.TotalSeconds
            [timespan]$timespan = [timespan]::FromSeconds( $secondsLeft )
            [string]$display = '{0:d2}:{1:d2}:{2:d2}' -f [int][math]::Floor( $timespan.TotalHours ) , $timespan.Minutes , $timespan.Seconds
            if( $secondsLeft -le 0 )
            {
                $timer.Stop()
                $WPFcheckboxRun.IsChecked = $false
                $WPFbtnCountdown.IsEnabled = $true
                if( $WPFcheckboxBeep.IsChecked -and $display -ne $script:lastTime )
                {
                    [console]::Beep( 1000 , [int]$(if( $script:beep -gt 0 ) { $script:beep } else { 500 } ))
                }
                Set-ForegroundWindow -thePid $pid
                if( ! [string]::IsNullOrEmpty( $script:expiredMessage ) )
                {
                    [void][Windows.MessageBox]::Show(  $script:expiredMessage , 'Countdown Timer Expired' , 'Ok' ,'Information' )
                }
                if( $secondsLeft -lt 0 )
                {
                    $secondsLeft = 0
                }
                $display = $script:countdown
                $timer.Reset()
            }
            $display
        }
        else
        {
            Get-Date -Format $dateFormat
        })
        if( $newTime -ne $script:lastTime )
        {
            Write-Debug -Message "New time is $newTime, lasttime was $script:lasttime"
            $script:lastTime = $newTime
            $WPFtxtClock.Text = $newTime
        }
    }
}

## https://richardspowershellblog.wordpress.com/2011/07/07/a-powershell-clock/
$form.Add_SourceInitialized({
    if( $formTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer )
    {
        ## need 0.1s granularity for the stopwatch but only just sub-second for the clock
        $formTimer.Interval = $(if( $stopWatch -or $tenths ) { [Timespan]'00:00:00.100' } else { [Timespan]'00:00:00.5' })
        $formTimer.Add_Tick( $timerBlock )
        $formTimer.Start()
    }
})

$WPFtxtMarkerFile.Text = $markerFile

$timer = New-Object -TypeName Diagnostics.Stopwatch

$script:lastTime = $null

if( $stopWatch -or $countdown )
{
    if( $WPFcheckboxRun.IsChecked = $start )
    {
        $timer.Start()
    }
}
else
{
    $WPFcheckboxRun.IsChecked = $true
}

$null = $Form.ShowDialog()

## put marker items onto the pipeline so can be copy'n'pasted into notes
if( $WPFlistMarkings.Items -and $WPFlistMarkings.Items.Count )
{
    $WPFlistMarkings.Items
}

# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUqDyqgk7YQe5uiXfFN4xP2gEy
# ajigggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
# AQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQg
# Q29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# +NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6kkPApfmJ
# 1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJUM1B0
# sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5wMWYzcT6s
# cKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4Tz
# rGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH5DiLanMg
# 0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIB
# ADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYIKwYBBQUH
# AQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYI
# KwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaG
# NGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0
# dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYE
# FFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2qB1dHC06
# GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5j
# DhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEpKBo6cSgC
# PC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIy
# sjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9CBoYs4Gb
# T8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIFTzCC
# BDegAwIBAgIQBP3jqtvdtaueQfTZ1SF1TjANBgkqhkiG9w0BAQsFADByMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29k
# ZSBTaWduaW5nIENBMB4XDTIwMDcyMDAwMDAwMFoXDTIzMDcyNTEyMDAwMFowgYsx
# CzAJBgNVBAYTAkdCMRIwEAYDVQQHEwlXYWtlZmllbGQxJjAkBgNVBAoTHVNlY3Vy
# ZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMRgwFgYDVQQLEw9TY3JpcHRpbmdIZWF2
# ZW4xJjAkBgNVBAMTHVNlY3VyZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAr20nXdaAALva07XZykpRlijxfIPk
# TUQFAxQgXTW2G5Jc1YQfIYjIePC6oaD+3Zc2WN2Jrsc7bj5Qe5Nj4QHHHf3jopLy
# g8jXl7Emt1mlyzUrtygoQ1XpBBXnv70dvZibro6dXmK8/M37w5pEAj/69+AYM7IO
# Fz2CrTIrQjvwjELSOkZ2o+z+iqfax9Z1Tv82+yg9iDHnUxZWhaiEXk9BFRv9WYsz
# qTXQTEhv8fmUI2aZX48so4mJhNGu7Vp1TGeCik1G959Qk7sFh3yvRugjY0IIXBXu
# A+LRT00yjkgMe8XoDdaBoIn5y3ZrQ7bCVDjoTrcn/SqfHvhEEMj1a1f0zQIDAQAB
# o4IBxTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0O
# BBYEFE16ovlqIk5uX2JQy6og0OCPrsnJMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUE
# DDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUw
# QzA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNl
# cnQuY29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNp
# Z25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAU9zO
# 9UpTkPL8DNrcbIaf1w736CgWB5KRQsmp1mhXbGECUCCpOCzlYFCSeiwH9MT0je3W
# aYxWqIpUMvAI8ndFPVDp5RF+IJNifs+YuLBcSv1tilNY+kfa2OS20nFrbFfl9QbR
# 4oacz8sBhhOXrYeUOU4sTHSPQjd3lpyhhZGNd3COvc2csk55JG/h2hR2fK+m4p7z
# sszK+vfqEX9Ab/7gYMgSo65hhFMSWcvtNO325mAxHJYJ1k9XEUTmq828ZmfEeyMq
# K9FlN5ykYJMWp/vK8w4c6WXbYCBXWL43jnPyKT4tpiOjWOI6g18JMdUxCG41Hawp
# hH44QHzE1NPeC+1UjTGCAigwggIkAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EAT946rb3bWrnkH02dUhdU4wCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFD7KurDAvwB9kPzCsmu1
# TUkurNbZMA0GCSqGSIb3DQEBAQUABIIBAGmjiN5kT3Ztvoc2dI5PHtWigi8ma8My
# EPUSIsbwcslHOd3H+E9f06hbF9Mbu83YCNTmHGfjBtOvoTMpggS72bmyIY6MJqtc
# kLEKEDAAZCIE+8yOsP8Aok9u/7bEssWAYBOWYO94tGGip9y5yAjkl1s7YRTIyDK2
# GtC+ueMyUe74hAn+nqrY2T13dvnOx1pEL3c62xdl8oBXedcZLf7MAViRMl9HTYfA
# Wakc/bikxU6sY4P8RqS4xNSVQ2fuvl5ZPP0uISZYto5xDRKUx7W61Q3n1weXwly6
# UytN3tMo0K6Z2ksDQf6Mlwi/j+zVPmUnTqY5EbIw4wqD6NoE+0ByxdA=
# SIG # End signature block
