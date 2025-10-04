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

.PARAMETER noDate

Do not display the date

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
    @guyrleech 03/05/2025  Added date & parameter to not show it
    @guyrleech 04/10/2025  Added Hide/Show items
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
    [switch]$noDate ,
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
        <TextBox x:Name="txtDate"  HorizontalAlignment="Left" Height="30" Margin="24,10,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="300" FontSize="16" IsReadOnly="True" FontWeight="Normal" BorderThickness="0"/>
  
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
                    <Separator />
                    <MenuItem Header="Hide All Items" Name="ToggleItemsContextMenu" />
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
[int]$script:previousDayOfYear = -1

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

$WPFToggleItemsContextMenu.Add_Click({
    $_.Handled = $true
    
    if ($script:itemsAreHidden) {
        # Show all items - restore from hidden items array
        $WPFlistMarkings.Items.Clear()
        ForEach ($item in $script:hiddenItems) {
            $null = $WPFlistMarkings.Items.Add($item)
        }
        $script:hiddenItems = @()
        $script:itemsAreHidden = $false
        $WPFToggleItemsContextMenu.Header = "Hide All Items"
    } else {
        # Hide all items - store current items and clear the list
        $script:hiddenItems = @($WPFlistMarkings.Items)
        $WPFlistMarkings.Items.Clear()
        # Add dummy line with asterisks
        $null = $WPFlistMarkings.Items.Add(([pscustomobject]@{ 'Timestamp' = '*****' ; 'Notes' = '*****' }))
        $script:itemsAreHidden = $true
        $WPFToggleItemsContextMenu.Header = "Show All Items"
    }
})

$script:nextAutoSaveTime = $null
$script:lastAutoSaveTime = $null
[array]$script:outstandingAutoSaveItems = @()
[array]$script:hiddenItems = @()
[bool]$script:itemsAreHidden = $false

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
        $now = [datetime]::Now

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
            $now.ToString( $dateFormat )
        })
        if( $newTime -ne $script:lastTime )
        {
            Write-Debug -Message "New time is $newTime, lasttime was $script:lasttime"
            $script:lastTime = $newTime
            $WPFtxtClock.Text = $newTime
            if( -Not $noDate )
            {
                if( ( $script:previousDayOfYear -lt 0 -or $script:previousDayOfYear -ne $now.DayOfYear ) )
                {
                    ## https://ss64.com/ps/syntax-dateformats.html
                    Write-Verbose "$($now.ToString('G')) date change"
                    $wpftxtDate.Text = "$($now.ToString( 'dddd' ) ) $($now.ToString( 'D' ) )"
                    $script:previousDayOfYear = $now.DayOfYear
                }
            }
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
# MIIktwYJKoZIhvcNAQcCoIIkqDCCJKQCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCg8At+51gZcQFQ
# bCUb8YRF5SEGUT6e/RRFiuPDokxs0qCCH2AwggV9MIIDZaADAgECAhAB1rN1Nl8g
# zZEd1y/l+ZNkMA0GCSqGSIb3DQEBCwUAMFoxCzAJBgNVBAYTAkxWMRkwFwYDVQQK
# ExBFblZlcnMgR3JvdXAgU0lBMTAwLgYDVQQDEydHb0dldFNTTCBHNCBDUyBSU0E0
# MDk2IFNIQTI1NiAyMDIyIENBLTEwHhcNMjUwNzIxMDAwMDAwWhcNMjYwNzIwMjM1
# OTU5WjBxMQswCQYDVQQGEwJHQjESMBAGA1UEBxMJV2FrZWZpZWxkMSYwJAYDVQQK
# Ex1TZWN1cmUgUGxhdGZvcm0gU29sdXRpb25zIEx0ZDEmMCQGA1UEAxMdU2VjdXJl
# IFBsYXRmb3JtIFNvbHV0aW9ucyBMdGQwdjAQBgcqhkjOPQIBBgUrgQQAIgNiAARE
# VushBxmaLDZJys/h4fGHMe+gEacCcTcalje+NTkKlUboku0+BdNDPxotbsh0aHHv
# HhwndrrL7f/pD45f5VUVKK5F3rQY7bZjZ6gxwGa/BzuFZsRhO12MTMC7zawyaQCj
# ggHUMIIB0DAfBgNVHSMEGDAWgBTJ/BDvUMjLa3+9CETvOmKT7VtemjAdBgNVHQ4E
# FgQU+7W5w1B8mWyXaJebkcnMJkcSB+cwPgYDVR0gBDcwNTAzBgZngQwBBAEwKTAn
# BggrBgEFBQcCARYbaHR0cDovL3d3dy5kaWdpY2VydC5jb20vQ1BTMA4GA1UdDwEB
# /wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzCBlwYDVR0fBIGPMIGMMESgQqBA
# hj5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vR29HZXRTU0xHNENTUlNBNDA5NlNI
# QTI1NjIwMjJDQS0xLmNybDBEoEKgQIY+aHR0cDovL2NybDQuZGlnaWNlcnQuY29t
# L0dvR2V0U1NMRzRDU1JTQTQwOTZTSEEyNTYyMDIyQ0EtMS5jcmwwgYMGCCsGAQUF
# BwEBBHcwdTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME0G
# CCsGAQUFBzAChkFodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vR29HZXRTU0xH
# NENTUlNBNDA5NlNIQTI1NjIwMjJDQS0xLmNydDAJBgNVHRMEAjAAMA0GCSqGSIb3
# DQEBCwUAA4ICAQALHZsdOuMeT/e1fsdQfhIz/2wS19UWlG1lxXieYOmPAju0DA5I
# ZheTgMtWMkUm96gWNtixny+q5nX8ckzuuD47esI2bM4G9RcVVKN0vdLZHv6QXZE5
# Ht2qTX8E1bqfejtDcGY0aqOjVYeOi/o98BsR98ItkjWNP3xE2oKEx6xYyZBL6d/z
# HB2ySd7hdk4VfmH9rTRftAsAn5L9s6m2ILRK8QRrkUJY9RxXswvQy2gNzccg+eYw
# y5gvLnzp4kdsTleV8SyZZQ2Tcp+HHPGxekB1NIM55vlCb9ocYw5j7noae3/PF+u/
# Zt/E+copm1c+MDju2bz1EelqXxuVsICRV9ikpJ7QEU+LwUiT7Ne+mgBmQ3IIyb8d
# QwR0xu1E/sKoWZjPRha6JLe65RaoBnXOX6fWQglPx467qjTUQpLxKGKMdQjS+LGJ
# uI2/BMWBHJfdhz/3GR9XVaDOWLhk+ChkjoXBgF2uFEXSiv4LNgigQ1R9RiojukEz
# mSRe5LK0UPdJch5I/HXg1lJFPORx05Ila0uSMusisgrPNvl5fEuf+DGYl2ywHsZQ
# pkeKT5wHQ5QJEocTKwGPfiGe9drO5DoMos5AXL5xnrPh/aQB4XKrttZTFy0+YXrU
# WSYa9v2rp7cAwuDsmd9gLoQzfW7jagbHfxvmLQ0CJTO9Y4BsJZML4bjimDCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBqEwggSJoAMCAQICEAeEPa0B
# wRXCdO5BpygiRnkwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDYyMzAwMDAwMFoX
# DTMyMDYyMjIzNTk1OVowWjELMAkGA1UEBhMCTFYxGTAXBgNVBAoTEEVuVmVycyBH
# cm91cCBTSUExMDAuBgNVBAMTJ0dvR2V0U1NMIEc0IENTIFJTQTQwOTYgU0hBMjU2
# IDIwMjIgQ0EtMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAK0e9Aey
# Q2aKomd3JZUKpfgW1inkV8ks71KHQG7Be68F41i3bXF/yH+ksn/tjN4pw2r3PjMS
# F5yq0PTGFu9IHySwEB9YExk7Q0t85PSPtbI24Puu/5kXNr6bEhDv2zV0KLBQzAai
# dqgMruapl8OkoQTFHpTIoGHdpq1PvdTYibH/H59hOZAWr43wWsuzoWHpgQZYlOCz
# HLDV8AKEJ+C0RxmR21yAruq6qyQe1bo8n2XlU0ntPdZenOew47GvPerHQLNaPArz
# 7cq/ZfqmJa93xhF8A7JxKtPj88zRwcsVznGz/ib96TgUKdhJ6bjd2gV+0HFkFCQc
# qg9bcG3pUiuFUjC87uGtSOFyEkllh1KV3dsA0O+Inn9Og/sQ2/3tT0y0oW+YLl3N
# 3WngfEVHSnnBaZhBtb7LdEWbSenof2bnsxQw2nTKyZ0mNvR/v51Utfc4QFRvof6v
# UtEtlP/EQ5O7A7EaDZLjDbkoYv1IIFRbieQGWM8d4lOhT5Me3Q/xlB/gH0gWUbG0
# srSDe44CfBIWAq2Y2OGROXxxosBDBuAQg0KquFzRvqTZnz5DCBUAvSci7Mz1yQo/
# zG0hxaDzkvf+XRxOQdB176MoYMbYJ3EEe1+UshS3CoPskDa4LFaUjNE2DBPSQfZ6
# fT5BXwrOJIZqo7avTqTBuirb/HLh73wIR9JjAgMBAAGjggFZMIIBVTASBgNVHRMB
# Af8ECDAGAQH/AgEAMB0GA1UdDgQWBBTJ/BDvUMjLa3+9CETvOmKT7VtemjAfBgNV
# HSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8BAf8EBAMCAYYwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhho
# dHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNl
# cnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3J0MEMGA1Ud
# HwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRy
# dXN0ZWRSb290RzQuY3JsMBwGA1UdIAQVMBMwBwYFZ4EMAQMwCAYGZ4EMAQQBMA0G
# CSqGSIb3DQEBCwUAA4ICAQAL2wrXsh2YpMJRq0Szv7J7CEmcni3KsvA0Sd+Xoesb
# w+bsdnRv7klz4YaolPyRFzuaKG5Wt2xg0d2Jy54Mv2KaG0K6wj+tSaPCG1+Wn5eA
# uSaAsauawSGjVv6UaJGnssL/XR2LxIA6KUOQePlnHXjbFG8GutaP1451IYfBi5sC
# wR3oIOiPBxpXP2l9czNg7eRzQ7o9eDVORyCRiUJQC4Me6T+xnHrxbQVWPU/aIyH5
# VSr2UvWm6ugDJ2h5ZVSH/5wxd6p+CkGqUBb6vyZrkXrIovSy1VAfy9hvURLSYlIg
# /Ih+QiMLVXSltlLenRBawppp8Sivx8t8s19LGe1Uj+6C63T7p6SW49ZGkRcf4kCI
# 11GNsttlyEJER7f9eXRiXCQD55BUl8zQDteK4W1j+Y6krYA34Tbm3mZBgWGl3Flf
# ktMMpaAOdv4DpCcS3iI3K6Rwtom5Lwg+A/PQTYAtku3dGqz6VeJ8r8bCc0hZA1t/
# sWYsMH2mHyLx2+xHWGyNzYo8RbhsCxu8tzPyE3XMTX9Bs5X3a8MagWPBk6LaShD5
# TISHR7yMMUcAol5Mj6nwQ8T+aq+iMsUCe3fXJcgDa2O3QRG2yOSEE2ZljpIQ5+ci
# g7C/Kpq8s/JrVS3f/Zw4Uu41DxCXodpmw1ASuefOf5kQBpROQ9lp5XKgciQ4MQIv
# OTCCBrQwggScoAMCAQICEA3HrFcF/yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAw
# YjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290
# IEc0MB4XDTI1MDUwNzAwMDAwMFoXDTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMC
# VVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBU
# cnVzdGVkIEc0IFRpbWVTdGFtcGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALR4MdMKmEFyvjxGwBysdduj
# Rmh0tFEXnU2tjQ2UtZmWgyxU7UNqEY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S
# 9SLrC6Kbltqn7SWCWgzbNfiR+2fkHUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+
# 42DFUF0mR/vtLa4+gKPsYfwEu7EEbkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg6
# 2IVwxKSpO0XaF9DPfNBKS7Zazch8NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21
# Qomb+zzQWKhxKTVVgtmUPAW35xUUFREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8
# y9IaaGBpPNXKFifinT7zL2gdFpBP9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQ
# NfVmUB5KlCX3ZA4x5HHKS+rqBvKWxdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gao
# u30yZ46t4Y9F20HHfIY4/6vHespYMQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6g
# qztiT96Fv/9bH7mQyogxG9QEPHrPV6/7umw052AkyiLA6tQbZl1KhBtTasySkuJD
# psZGKdlsjg4u70EwgWbVRSX1Wd4+zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D
# 8bpfm4CLKczsG7ZrIGNTAgMBAAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEA
# MB0GA1UdDgQWBBTvb1NK6eQGfHrK4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC
# 0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYB
# BQUHAwgwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5k
# aWdpY2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSG
# Mmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQu
# Y3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0B
# AQsFAAOCAgEAF877FoAc/gc9EXZxML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6F
# TGNpoV2V4wzSUGvI9NAzaoQk97frPBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mC
# efSG+tXqGpYZ3essBS3q8nL2UwM+NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57m
# QfQXwcAEGCvRR2qKtntujB71WPYAgwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9
# ydOal95CHfmTnM4I+ZI2rVQfjXQA1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dB
# wp9nEC8EAqoxW6q17r0z0noDjs6+BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdq
# fMTCW/NmKLJ9M+MtucVGyOxiDf06VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2
# puE6FndlENSmE+9JGYxOGLS/D284NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAO
# k5eCkhSxZON3rGlHqhpB/8MluDezooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL
# 0Q4ssd8xHZnIn/7GELH3IdvG2XlM9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBun
# vAZapsiI5YKdvlarEvf8EA+8hcpSM9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE
# 1aADAgECAhAKgO8YS43xBYLRxHanlXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNV
# BAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNl
# cnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBD
# QTEwHhcNMjUwNjA0MDAwMDAwWhcNMzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJV
# UzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNI
# QTI1NiBSU0E0MDk2IFRpbWVzdGFtcCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHf
# yjfMGUIwYzKomd8U1nH7C8Dr0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPx
# NyFPJIDZHhAqlUPt281mHrBbZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpk
# BaMUNg7MOLxI6E9RaUueHTQKWXymOtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFv
# ZSjKs3SKO1QNUdFd2adw44wDcKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1zn
# OM8odbkqoK+lJ25LCHBSai25CFyD23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8f
# cpK40uhktzUd/Yk0xUvhDU6lvJukx7jphx40DQt82yepyekl4i0r8OEps/FNO4ah
# fvAk12hE5FVs9HVVWcO5J4dVmVzix4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUD
# y9Z2hSgctaepZTd0ILIUbWuhKuAeNIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9
# w6CtjuuVHJOVoIJ/DtpJRE7Ce7vMRHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTn
# nkrT3pXWETTJkhd76CIDBbTRofOsNyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKa
# cJ+A9/z7eacCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7
# /PIx7f391/ORcWMZUEPPYYzoMB8GA1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ
# 7MtOMA4GA1UdDwEB/wQEAwIHgDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYI
# KwYBBQUHAQEEgYgwgYUwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0
# LmNvbTBdBggrBgEFBQcwAoZRaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0VHJ1c3RlZEc0VGltZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEu
# Y3J0MF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0Ex
# LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcN
# AQELBQADggIBAGUqrfEcJwS5rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF
# 0RkP2AGr181o2YWPoSHz9iZEN/FPsLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKq
# dT8wv2UV+Kbz/3ImZlJ7YXwBD9R0oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbU
# UO75ZSpbh1oipOhcUT8lD8QAGB9lctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTe
# HihsQyfFg5fxUFEp7W42fNBVN4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG
# 7aEQJmmrJTV3Qhtfparz+BW60OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NB
# qycz0BZwhB9WOfOu/CIJnzkQTwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6
# +iX8MmB10nfldPF9SVD7weCC3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaA
# yBjFBtXVLcKtapnMG3VH3EmAp/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyP
# ehwJVxwC+UpX2MSey2ueIu9THFVkT+um1vshETaWyQo8gmBto/m3acaP9QsuLj3F
# NwFlTxq25+T4QwX9xa6ILs84ZPvmpovq90K8eWyG2N01c4IhSOxqt81nMYIErTCC
# BKkCAQEwbjBaMQswCQYDVQQGEwJMVjEZMBcGA1UEChMQRW5WZXJzIEdyb3VwIFNJ
# QTEwMC4GA1UEAxMnR29HZXRTU0wgRzQgQ1MgUlNBNDA5NiBTSEEyNTYgMjAyMiBD
# QS0xAhAB1rN1Nl8gzZEd1y/l+ZNkMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQB
# gjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIDuUJAwR
# LIX+aevJvg/JY1MQwM8keiBp1G11AWrWR1ndMAsGByqGSM49AgEFAARnMGUCMQDv
# 1S2lIX2BGHGOOxCGqkdo1J/QnxtUC/BD3t4+L/c3bpWinlsSgv+AtevkeRaubdIC
# MGSYCGS0xWT07crBZ+lV1RrUfFxFQUJ5tLeiau+7SMCGJQ3KBdFPb8MvSsXfPtsD
# 0KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVT
# MRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1
# c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA
# 7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJ
# KoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNTEwMDQxMjAyMDdaMC8GCSqGSIb3
# DQEJBDEiBCCM6AXxPDCoq4t92LcGtWsANeqKCP8fa5t1+o+Zdu9qvTANBgkqhkiG
# 9w0BAQEFAASCAgCJRVn8uWAEHz2qCaFvlGoQDaDYucOVZmb8SgWY8Uy9xKuH+P1X
# TYQEFhxoq+xtJDqtaUYo0F+/pxKebJoB5xs+8UGD6ZxU8EH+WJ/huqkTxGsRPOwa
# W03UbPlpT9hqNov9u4pWgW8s3KFu/O0EXv/qfutZ2Ju0cpxCUbOGUPxiFyGFB7/+
# VhTJaPTvGtRzyX1ng82b3P/NgHZF5lGyiNvvU6rNy2y7qWo960dNIKvnPnUUdORj
# C3f2Bjyb0s9VVk7slAoXMdRHF2W0TIwhKIt56K3bokjqaWAmiKLdtSQ9xdK/SBYf
# JVsHGRgm3/l+ZrZVqeHVwM+j+oXwU9Ry0AesIlL8YbL5PppYXqCEUsStTXXIlQvq
# 0GYV8hzYKWxEoZ++NeXK00hPko97JPGyDz78liUFymMlm92gti9Mywj4vy8QiP11
# RqJQjzwTGbcB088bxh6YGl78ZnSBUu3xNzDt1awpOWPH3cz3TmCHmL4yGwmrOu6t
# DqkRVTCLzldbH7T+/MmZtiCmykF3xzPeuzLcZ4cE8ucQvYYXbwf3aeYJ159xgJdI
# PqOj5YZGXPhucBKshFGpxsWRChhbhcVBjU03Rt2jdJIGDfxkCs1FOfUBut1XXy26
# 1UYS22RArbGtBOyO52ddntFfZeHLE/Q1RTWFZUlAZASTI1Af2O3gYl+GIA==
# SIG # End signature block
