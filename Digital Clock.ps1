#requires -version 3

<#
.SYNOPSIS

Display digitial clock with granularity in seconds or a stopwatch or countdown timer in a window with ability to stop/start/reset

.DESCRIPTION

Includes a "Mark" button which allows a comment to be placed in the window, along with the time the button was pressed,
eg to mark when something occurred so that it can be cross referenced to log files, event logs, etc

.PARAMETER stopwatch

Run a stopwatch rather than showing current time

.PARAMETER start

Start the stopwatch immediately

.PARAMETER notOnTop

Do not place the window on top of all other windows

.PARAMETER title

Title of the clock window. A default will be used if not specified.

.PARAMETER autosaveFile

Name/path of csv file to save the marked items to at the frequency set by -autosaveMinutes
Pseudo environment variables are allowed for date components which will be expanded - %day%, %month%, %monthname%, %dayname% , %year%
Special folders can be specified with, e.g. ^desktop which will be expanded to the location of the user's desktop folder (see example 2)

.PARAMETER autosaveMinutes

Frequency in minutes to save the new marked items to. If not specified, defaults to 1 minute

.PARAMETER noappend

Remove any existing autosave file before writing new marker items for this clock instance.
If not specified, new marker items will be appended to the autosave file if it already exists.

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

& '.\Digital Clock.ps1' -autosaveFile '^desktop\clock.notes.%monthname%.%year%'

Display an updating digital clock in a window and automatically write any marker items to a file containing the month name and year on the user's desktop

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
                           Fixed countdown timer re-run bug where timer not reset so expires 
    @guyrleech 30/11/2020  Added -title argument
    @guyrleech 18/12/2020  Added code to stop and delete timer on exit
    @guyrleech 26/01/2022  Added autosave options, support for special folders and pseudo environment variables %day%, %month%, %monthname%, %dayname% , %year%
#>

[CmdletBinding()]

Param
(
    [switch]$stopWatch ,
    [switch]$start ,
    [string]$autosaveFile ,
    [decimal]$autosaveMinutes = 1 ,
    [switch]$noappend ,
    [string]$markerFile ,
    [string]$countdown ,
    [string]$expiredMessage ,
    [string]$title ,
    [switch]$notOnTop ,
    [switch]$tenths ,
    [int]$beep
)

if( -Not $PSBoundParameters[ 'autosavefile' ] -and $PSBoundParameters[ 'autosaveMinutes' ] )
{
    Throw "Must not specify -autosaveMinutes without -autosavefile"
}

if( -Not $PSBoundParameters[ 'autosavefile' ] -and $PSBoundParameters[ 'noappend' ] )
{
    Throw "Must not specify -noappend without -autosavefile"
}

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
        Title="Marker Text" Height="300" Width="600" Name="Marker">
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
if( $PSBoundParameters[ 'title' ] )
{
    $form.Title = $title
}
else
{
    $form.Title = $(if( $stopWatch ) { 'Guy''s Stopwatch' } elseif( $countdown ) { 'Guy''s Countdown Timer' } else { 'Guy''s Clock' })
}

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

$script:nextAutoSaveTime = $null
$script:lastAutoSaveTime = $null
[array]$script:outstandingAutoSaveItems = @()

if( -Not [string]::IsNullOrEmpty( $autosaveFile ) )
{
    $script:nextAutoSaveTime = [datetime]::Now.AddMinutes( $autosaveMinutes )
    ## set pseudo environment variables for date components 
    $env:Day   = [datetime]::Now.Day
    $env:Month = [datetime]::Now.Month
    $env:Year  = [datetime]::Now.Year
    $env:MonthName = (Get-Culture).DateTimeFormat.GetMonthName( [datetime]::Now.Month )
    $env:DayName   = (Get-Culture).DateTimeFormat.GetDayName(   [datetime]::Now.DayOfWeek )

    $autosaveFile = [System.Environment]::ExpandEnvironmentVariables( $autosaveFile )
    
    ## expand ^specialfolder to the specialfolder
    ## '^desktop\clock.notes.csv
    if( $autosaveFile -match '\^([^\\]+)' )
    {
        [string]$specialFolder = $null
        $specialFolder = [System.Environment]::GetFolderPath( $matches[1] )
        if( $specialFolder )
        {
            ## $matches[0] will be the whole string that matched, eg ^desktop , so we have to escape since ^ means anchor to start of line
            [string]$expandedAutoSaveFile = $autosaveFile -replace ([regex]::Escape( $matches[0] ) ) , $specialFolder
            if( $expandedAutoSaveFile -ne $autosaveFile )
            {
                Write-Verbose -Message "Autosave file changed from `"$autosaveFile`" to `"$expandedAutoSaveFile`""
                $autosaveFile = $expandedAutoSaveFile
            }
        }
        else
        {
            Throw "Failed to resolve special folder `"$($matches[1])`""
        }
    }
   
    if( $noappend -and ( Test-Path -Path $autosaveFile -ErrorAction SilentlyContinue ) )
    {
        Remove-Item -Path $autosaveFile -ErrorAction Stop ## fatal error if can't remove file as probably then can't write to it either so force caller to fix issue
    }
}

[scriptblock]$timerBlock = `
{
    if( $script:nextAutoSaveTime -and [datetime]::Now -ge $nextAutoSaveTime )
    {
        ## must only save notes added since last save unless we had a save failure
        $newItems = New-Object -TypeName System.Collections.Generic.List[object]
        if( $script:outstandingAutoSaveItems.Count )
        {
            $newItems += $script:outstandingAutoSaveItems
        }
        $newItems += @( $WPFlistMarkings.Items | Where-Object { -Not $script:lastAutoSaveTime -or ($_.Timestamp -as [datetime]) -ge $script:lastAutoSaveTime } )
     
        if( $newItems -and $newItems.Count )
        {
            Write-Verbose -Message "$(Get-Date -Format G): autosaving $($newItems.Count) new items to $autosaveFile"
            $newItems | Export-Csv -NoTypeInformation -Path $autosaveFile -Append
                
            if( $? )
            {
                $script:outstandingAutoSaveItems = @()
            }
            else ## if error saving then record it so will try again next time regardless of whether there are new items or not
            {
                $script:outstandingAutoSaveItems = $newItems
            }
        }
        $script:lastAutoSaveTime = $nextAutoSaveTime
        $script:nextAutoSaveTime = [datetime]::Now.AddMinutes( $autosaveMinutes )
    }
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
$script:formTimer = $null
$form.Add_SourceInitialized({
    if( $script:formTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer )
    {
        ## need 0.1s granularity for the stopwatch but only just sub-second for the clock
        $script:formTimer.Interval = $(if( $stopWatch -or $tenths ) { [Timespan]'00:00:00.100' } else { [Timespan]'00:00:00.5' })
        $script:formTimer.Add_Tick( $timerBlock )
        $script:formTimer.Start()
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

if( $script:formTimer )
{
    $script:formTimer.Stop()
    $script:formTimer.remove_Tick( $timerBlock )
    $script:formTimer = $null
    $timerBlock = $null
    Remove-Variable -Name timerBlock -Force -Confirm:$false
}

## put marker items onto the pipeline so can be copy'n'pasted into notes
if( $WPFlistMarkings.Items -and $WPFlistMarkings.Items.Count )
{
    $WPFlistMarkings.Items
}

# SIG # Begin signature block
# MIIZsAYJKoZIhvcNAQcCoIIZoTCCGZ0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUyAhf8cqKDvKqsiDNnJtsjFAj
# HDmgghS+MIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMB4XDTIxMDEwMTAwMDAwMFoXDTMxMDEw
# NjAwMDAwMFowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAMLmYYRnxYr1DQikRcpja1HXOhFCvQp1dU2UtAxQ
# tSYQ/h3Ib5FrDJbnGlxI70Tlv5thzRWRYlq4/2cLnGP9NmqB+in43Stwhd4CGPN4
# bbx9+cdtCT2+anaH6Yq9+IRdHnbJ5MZ2djpT0dHTWjaPxqPhLxs6t2HWc+xObTOK
# fF1FLUuxUOZBOjdWhtyTI433UCXoZObd048vV7WHIOsOjizVI9r0TXhG4wODMSlK
# XAwxikqMiMX3MFr5FK8VX2xDSQn9JiNT9o1j6BqrW7EdMMKbaYK02/xWVLwfoYer
# vnpbCiAvSwnJlaeNsvrWY4tOpXIc7p96AXP4Gdb+DUmEvQECAwEAAaOCAbgwggG0
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEEGA1UdIAQ6MDgwNgYJYIZIAYb9bAcBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAfBgNVHSMEGDAWgBT0tuEgHf4prtLk
# YaWyoiWyyBc1bjAdBgNVHQ4EFgQUNkSGjqS6sGa+vCgtHUQ23eNqerwwcQYDVR0f
# BGowaDAyoDCgLoYsaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC10cy5jcmwwMqAwoC6GLGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtdHMuY3JsMIGFBggrBgEFBQcBAQR5MHcwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBPBggrBgEFBQcwAoZDaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRFRpbWVzdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEASBzctemaI7znGucgDo5nRv1CclF0CiNH
# o6uS0iXEcFm+FKDlJ4GlTRQVGQd58NEEw4bZO73+RAJmTe1ppA/2uHDPYuj1UUp4
# eTZ6J7fz51Kfk6ftQ55757TdQSKJ+4eiRgNO/PT+t2R3Y18jUmmDgvoaU+2QzI2h
# F3MN9PNlOXBL85zWenvaDLw9MtAby/Vh/HUIAHa8gQ74wOFcz8QRcucbZEnYIpp1
# FUL1LTI4gdr0YKK6tFL7XOBhJCVPst/JKahzQ1HavWPWH1ub9y4bTxMd90oNcX6X
# t/Q/hOvB46NJofrOp79Wz7pZdmGJX36ntI5nePk2mOHLKNpbh6aKLzCCBTAwggQY
# oAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsx
# SRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2qvCchqXYJawO
# eSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJ
# RdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/6XzLkqHlOzEc
# z+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE94zRICUj6whk
# PlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8l
# k9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQD
# AgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0wazAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsME8GA1UdIARI
# MEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdp
# Y2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg
# +S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG
# 9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh134LYP3DPQ/E
# r4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63XX0R58zYUBor3
# nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPAJRHinBRHoXpo
# aK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW
# 6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ
# 92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBTEwggQZoAMCAQICEAqhJdbW
# Mht+QeQF2jaXwhUwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIG
# A1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTE2MDEwNzEyMDAw
# MFoXDTMxMDEwNzEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGln
# aUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAL3QMu5LzY9/3am6gpnFOVQoV7YjSsQOB0Uz
# URB90Pl9TWh+57ag9I2ziOSXv2MhkJi/E7xX08PhfgjWahQAOPcuHjvuzKb2Mln+
# X2U/4Jvr40ZHBhpVfgsnfsCi9aDg3iI/Dv9+lfvzo7oiPhisEeTwmQNtO4V8CdPu
# XciaC1TjqAlxa+DPIhAPdc9xck4Krd9AOly3UeGheRTGTSQjMF287DxgaqwvB8z9
# 8OpH2YhQXv1mblZhJymJhFHmgudGUP2UKiyn5HU+upgPhH+fMRTWrdXyZMt7HgXQ
# hBlyF/EXBu89zdZN7wZC/aJTKk+FHcQdPK/P2qwQ9d2srOlW/5MCAwEAAaOCAc4w
# ggHKMB0GA1UdDgQWBBT0tuEgHf4prtLkYaWyoiWyyBc1bjAfBgNVHSMEGDAWgBRF
# 66Kv9JLLgjEtUYunpyGd823IDzASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB5BggrBgEFBQcBAQRtMGswJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBQBgNV
# HSAESTBHMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cu
# ZGlnaWNlcnQuY29tL0NQUzALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggEB
# AHGVEulRh1Zpze/d2nyqY3qzeM8GN0CE70uEv8rPAwL9xafDDiBCLK938ysfDCFa
# KrcFNB1qrpn4J6JmvwmqYN92pDqTD/iy0dh8GWLoXoIlHsS6HHssIeLWWywUNUME
# aLLbdQLgcseY1jxk5R9IEBhfiThhTWJGJIdjjJFSLK8pieV4H9YLFKWA1xJHcLN1
# 1ZOFk362kmf7U2GJqPVrlsD0WGkNfMgBsbkodbeZY4UijGHKeZR+WfyMD+NvtQEm
# tmyl7odRIeRYYJu6DC0rbaLEfrvEJStHAgh8Sa4TtuF8QkIoxhhWz0E0tmZdtnR7
# 9VYzIi8iNrJLokqV2PWmjlIwggVPMIIEN6ADAgECAhAE/eOq2921q55B9NnVIXVO
# MA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwHhcNMjAwNzIwMDAw
# MDAwWhcNMjMwNzI1MTIwMDAwWjCBizELMAkGA1UEBhMCR0IxEjAQBgNVBAcTCVdh
# a2VmaWVsZDEmMCQGA1UEChMdU2VjdXJlIFBsYXRmb3JtIFNvbHV0aW9ucyBMdGQx
# GDAWBgNVBAsTD1NjcmlwdGluZ0hlYXZlbjEmMCQGA1UEAxMdU2VjdXJlIFBsYXRm
# b3JtIFNvbHV0aW9ucyBMdGQwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQCvbSdd1oAAu9rTtdnKSlGWKPF8g+RNRAUDFCBdNbYbklzVhB8hiMh48LqhoP7d
# lzZY3YmuxztuPlB7k2PhAccd/eOikvKDyNeXsSa3WaXLNSu3KChDVekEFee/vR29
# mJuujp1eYrz8zfvDmkQCP/r34Bgzsg4XPYKtMitCO/CMQtI6Rnaj7P6Kp9rH1nVO
# /zb7KD2IMedTFlaFqIReT0EVG/1ZizOpNdBMSG/x+ZQjZplfjyyjiYmE0a7tWnVM
# Z4KKTUb3n1CTuwWHfK9G6CNjQghcFe4D4tFPTTKOSAx7xegN1oGgifnLdmtDtsJU
# OOhOtyf9Kp8e+EQQyPVrV/TNAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7
# KgqjpepxA8Bg+S32ZXUOWDAdBgNVHQ4EFgQUTXqi+WoiTm5fYlDLqiDQ4I+uyckw
# DgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4w
# NaAzoDGGL2h0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3Mt
# ZzEuY3JsMDWgM6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1
# cmVkLWNzLWcxLmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUF
# BwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYI
# KwYBBQUHAQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wTgYIKwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFNIQTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAA
# MA0GCSqGSIb3DQEBCwUAA4IBAQBT3M71SlOQ8vwM2txshp/XDvfoKBYHkpFCyanW
# aFdsYQJQIKk4LOVgUJJ6LAf0xPSN7dZpjFaoilQy8Ajyd0U9UOnlEX4gk2J+z5i4
# sFxK/W2KU1j6R9rY5LbScWtsV+X1BtHihpzPywGGE5eth5Q5TixMdI9CN3eWnKGF
# kY13cI69zZyyTnkkb+HaFHZ8r6binvOyzMr69+oRf0Bv/uBgyBKjrmGEUxJZy+00
# 7fbmYDEclgnWT1cRROarzbxmZ8R7Iyor0WU3nKRgkxan+8rzDhzpZdtgIFdYvjeO
# c/IpPi2mI6NY4jqDXwkx1TEIbjUdrCmEfjhAfMTU094L7VSNMYIEXDCCBFgCAQEw
# gYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1
# cmVkIElEIENvZGUgU2lnbmluZyBDQQIQBP3jqtvdtaueQfTZ1SF1TjAJBgUrDgMC
# GgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYK
# KwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG
# 9w0BCQQxFgQUGZ3AkgxZiON3IqIlIQ47rWGGDawwDQYJKoZIhvcNAQEBBQAEggEA
# Zww9rgLQNyYqvYdgLakb2Sy0KQVJZiM++b23A8PKijwE1M55Y8YIp+dQXC5pVbMc
# jigDvAdCaXFG4Ia2qIEWegopVXX3o0j14EtUe0Z97HF+sFm6ObW9Kup0JiRql5Au
# LIclVAYNJHFWspNP6qtgQ/LAO3JbeQ8j+cW1D+OrgceY5xV8SuCwqlSAYMKmmFGl
# 7Ed3+Thvk0IVagPMVLuIK93tvuBipmrdaR+sYP2CM1qjR12vj7Y7xEhtSe47vmKt
# 8SUtVvAfHBC9isQfsR0x/jta2ZNfeDdBTK9v8/v1MqmLqCtCYx4fxmlx+ztTLBSZ
# tKByL6D8LIZJHMuutte8TaGCAjAwggIsBgkqhkiG9w0BCQYxggIdMIICGQIBATCB
# hjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3Vy
# ZWQgSUQgVGltZXN0YW1waW5nIENBAhANQkrgvjqI/2BAIc4UAPDdMA0GCWCGSAFl
# AwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMjIwMTI3MDEyMTI4WjAvBgkqhkiG9w0BCQQxIgQgsbwn5VlCR+/VN8SBKr+U
# JyxYjgc+26E7PPem+4hEypQwDQYJKoZIhvcNAQEBBQAEggEAdPBl2S4FKlt7tmFW
# XKRRJ4PR0QQKlPr6jDA/lVkUkYYedWP2FR2PQyKbEYcSO4wHg7cx+QIuQNHH9SfO
# QoFsX/GsDf1CddnZfIAvQmfktgGeYhZLF+WRPn2eU6+PdUZTiQfiYW8s636Wx+bP
# wVQoFzO5crF89XEdGob2tQNwrEbfIVczPeAfkDbnqqpZTu7LBx/GDIpes4OlU9xv
# gWh5XRan6/kx4sJV0JD9yFajYOTSiSf8OTAThmdPgNk4r5nkPAhjfV+iZJG368zK
# sBI48Nvk4dAdjNQ62cuLFt88kJjO2jpAVI+cmcRvM//iBfheCo37cKg8sP904QBG
# lssQbA==
# SIG # End signature block
