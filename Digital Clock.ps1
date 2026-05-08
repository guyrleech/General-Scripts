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

.PARAMETER timezone

Display the time in the specified timezone rather than local time. Accepts common abbreviations (e.g. EDT, PST, GMT, JST) or a Windows timezone ID (e.g. 'Eastern Standard Time').
The full timezone name and UTC offset are shown in a label beneath the date.

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
    @guyrleech 08/05/2026  Added timezone parameter and calculation of time difference
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
    [int]$beep ,
    [string]$timezone
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
        Title="Guy's Clock" Height="313.694" Width="636.432">
    <Grid>
        <TextBox x:Name="txtClock" HorizontalAlignment="Left" Height="111" Margin="24,70,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="358" FontSize="72" IsReadOnly="True" FontWeight="Bold" BorderThickness="0"/>
        <TextBox x:Name="txtDate"  HorizontalAlignment="Left" Height="30" Margin="24,10,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="300" FontSize="16" IsReadOnly="True" FontWeight="Normal" BorderThickness="0"/>
        <Label x:Name="lblTimezone" HorizontalAlignment="Left" Height="28" Margin="24,40,0,0" VerticalAlignment="Top" Width="400" FontSize="16" Content=""/>
  
        <Grid Margin="24,241,244,10">
            <CheckBox x:Name="checkboxRun" Content="_Run" HorizontalAlignment="Left" Height="18" Margin="-11,5,0,0" VerticalAlignment="Top" Width="52" IsChecked="True"/>
            <Button x:Name="btnReset" Content="Re_set" HorizontalAlignment="Left" Height="23" Margin="112,-1,0,0" VerticalAlignment="Top" Width="72"/>
            <Button x:Name="btnMark" Content="_Mark" HorizontalAlignment="Left" Height="23" Margin="35,-1,0,0" VerticalAlignment="Top" Width="72"/>
            <Button x:Name="btnClear" Content="_Clear" HorizontalAlignment="Left" Height="23" Margin="189,-1,0,0" VerticalAlignment="Top" Width="72"/>
            <Button x:Name="btnCountdown" Content="Count _Down" HorizontalAlignment="Left" Height="23" Margin="268,-1,0,0" VerticalAlignment="Top" Width="72"/>

        </Grid>
        <TextBox x:Name="txtMarkerFile" HorizontalAlignment="Left" Height="24" Margin="92,185,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="268"/>
        <Label Content="Marker File" HorizontalAlignment="Left" Height="23" Margin="10,186,0,0" VerticalAlignment="Top" Width="72"/>
        <ListView x:Name="listMarkings" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="382,13,10,10" >
            <ListView.ContextMenu>
                <ContextMenu>
                    <MenuItem Header="Edit" Name="EditContextMenu" />
                    <MenuItem Header="Delete" Name="DeleteContextMenu" />
                    <MenuItem Header="Save" Name="SaveContextMenu" />
                    <MenuItem Header="Insert Above" Name="InsertAboveContextMenu" />
                    <MenuItem Header="Insert Below" Name="InsertBelowContextMenu" />
                    <MenuItem Header="Time Difference" Name="TimeDiffContextMenu" />
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
        <CheckBox x:Name="checkboxBeep" Content="Beep" HorizontalAlignment="Left" Height="15" Margin="13,221,0,0" VerticalAlignment="Top" Width="93"/>
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

## Resolve -timezone abbreviation or ID to a TimeZoneInfo object
$script:timezoneInfo = $null
if( -Not [string]::IsNullOrEmpty( $timezone ) )
{
    [hashtable]$tzAbbreviations = @{
        'EDT'  = 'Eastern Standard Time'    ; 'EST'  = 'Eastern Standard Time'
        'CDT'  = 'Central Standard Time'    ; 'CST'  = 'Central Standard Time'
        'MDT'  = 'Mountain Standard Time'   ; 'MST'  = 'Mountain Standard Time'
        'PDT'  = 'Pacific Standard Time'    ; 'PST'  = 'Pacific Standard Time'
        'AKDT' = 'Alaskan Standard Time'    ; 'AKST' = 'Alaskan Standard Time'
        'HDT'  = 'Hawaiian Standard Time'   ; 'HST'  = 'Hawaiian Standard Time'
        'GMT'  = 'GMT Standard Time'        ; 'UTC'  = 'UTC'
        'BST'  = 'GMT Standard Time'        ; 'WET'  = 'GMT Standard Time'
        'CET'  = 'Central Europe Standard Time' ; 'CEST' = 'Central Europe Standard Time'
        'EET'  = 'E. Europe Standard Time'  ; 'EEST' = 'E. Europe Standard Time'
        'MSK'  = 'Russian Standard Time'
        'IST'  = 'India Standard Time'      ; 'PKT'  = 'Pakistan Standard Time'
        'ICT'  = 'SE Asia Standard Time'
        'HKT'  = 'China Standard Time'      ; 'SGT'  = 'Singapore Standard Time'
        'JST'  = 'Tokyo Standard Time'      ; 'KST'  = 'Korea Standard Time'
        'AEST' = 'AUS Eastern Standard Time'; 'AEDT' = 'AUS Eastern Standard Time'
        'NZST' = 'New Zealand Standard Time'; 'NZDT' = 'New Zealand Standard Time'
    }

    [string]$resolvedTzId = $null
    [string]$tzUpper = $timezone.ToUpper()

    if( $tzAbbreviations.ContainsKey( $tzUpper ) )
    {
        $resolvedTzId = $tzAbbreviations[ $tzUpper ]
    }
    else
    {
        ## Try direct Windows timezone ID, display name, standard name or daylight name
        $matchedZone = [TimeZoneInfo]::GetSystemTimeZones() | Where-Object {
            $_.Id -eq $timezone -or $_.StandardName -eq $timezone -or $_.DaylightName -eq $timezone -or $_.DisplayName -eq $timezone
        } | Select-Object -First 1
        if( $matchedZone )
        {
            $resolvedTzId = $matchedZone.Id
        }
    }

    if( $resolvedTzId )
    {
        try
        {
            $script:timezoneInfo = [TimeZoneInfo]::FindSystemTimeZoneById( $resolvedTzId )
        }
        catch
        {
            Throw "Cannot find timezone ID `"$resolvedTzId`" for `"$timezone`": $_"
        }
    }
    else
    {
        Throw "Unknown timezone `"$timezone`". Use a common abbreviation (e.g. EDT, PST, GMT) or a Windows timezone ID (e.g. 'Eastern Standard Time')."
    }
}

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

## Set timezone label content; if no -timezone was specified, use the local timezone
if( -Not $script:timezoneInfo )
{
    $script:timezoneInfo = [TimeZoneInfo]::Local
}

if( $script:timezoneInfo )
{
    [scriptblock]$script:updateTimezoneLabel = {
        [timespan]$tzOffset    = $script:timezoneInfo.GetUtcOffset( [datetime]::UtcNow )
        [int]$tzHours          = [math]::Abs( [math]::Truncate( $tzOffset.TotalHours ) )
        [int]$tzMins           = [math]::Abs( $tzOffset.Minutes )
        [string]$tzSign        = if( $tzOffset.TotalMinutes -ge 0 ) { '+' } else { '-' }
        [string]$tzOffsetStr   = if( $tzMins -eq 0 ) { "UTC{0}{1}" -f $tzSign,$tzHours } else { "UTC{0}{1}:{2:00}" -f $tzSign,$tzHours,$tzMins }
        [string]$tzDisplayName = if( $script:timezoneInfo.IsDaylightSavingTime( [datetime]::UtcNow ) ) { $script:timezoneInfo.DaylightName } else { $script:timezoneInfo.StandardName }
        $WPFlblTimezone.Content = "$tzDisplayName ($tzOffsetStr)"
    }
    & $script:updateTimezoneLabel
}

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

$WPFTimeDiffContextMenu.Add_Click({
    $_.Handled = $true
    if( $WPFlistMarkings.SelectedItems.Count -ne 2 )
    {
        [void][Windows.MessageBox]::Show( "Please select exactly 2 items to calculate the time difference" , 'Time Difference' , 'Ok' , 'Exclamation' )
    }
    else
    {
        $item1 = $WPFlistMarkings.SelectedItems[0]
        $item2 = $WPFlistMarkings.SelectedItems[1]
        [timespan]$diff    = [timespan]::Zero
        [bool]$parsed      = $false
        [datetime]$dt1     = [datetime]::MinValue
        [datetime]$dt2     = [datetime]::MinValue
        if( [datetime]::TryParse( $item1.Timestamp , [ref]$dt1 ) -and [datetime]::TryParse( $item2.Timestamp , [ref]$dt2 ) )
        {
            $diff   = $dt2 - $dt1
            if( $diff -lt [timespan]::Zero ) { $diff = $diff.Negate() }
            $parsed = $true
        }
        else
        {
            [timespan]$ts1 = [timespan]::Zero
            [timespan]$ts2 = [timespan]::Zero
            if( [timespan]::TryParse( $item1.Timestamp , [ref]$ts1 ) -and [timespan]::TryParse( $item2.Timestamp , [ref]$ts2 ) )
            {
                $diff   = $ts2 - $ts1
                if( $diff -lt [timespan]::Zero ) { $diff = $diff.Negate() }
                $parsed = $true
            }
        }
        if( $parsed )
        {
            [string]$diffText = '{0:d2}:{1:d2}:{2:d2}.{3:d3}' -f [int][math]::Floor( $diff.TotalHours ) , $diff.Minutes , $diff.Seconds , $diff.Milliseconds
            [void][Windows.MessageBox]::Show( $diffText , 'Time Difference' , 'Ok' , 'Information' )
        }
        else
        {
            [void][Windows.MessageBox]::Show( "Unable to parse timestamps:`n  $($item1.Timestamp)`n  $($item2.Timestamp)" , 'Time Difference' , 'Ok' , 'Exclamation' )
        }
    }
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
        $now = $(if( $script:timezoneInfo ) { [TimeZoneInfo]::ConvertTimeFromUtc( [datetime]::UtcNow, $script:timezoneInfo ) } else { [datetime]::Now })

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
                    ## refresh timezone label in case DST changed at midnight
                    if( $script:timezoneInfo ) { & $script:updateTimezoneLabel }
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
# MIIkkQYJKoZIhvcNAQcCoIIkgjCCJH4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUTZIZU7LgpuAoqZO5/c6Pm6hB
# AmOggh9gMIIFfTCCA2WgAwIBAgIQAdazdTZfIM2RHdcv5fmTZDANBgkqhkiG9w0B
# AQsFADBaMQswCQYDVQQGEwJMVjEZMBcGA1UEChMQRW5WZXJzIEdyb3VwIFNJQTEw
# MC4GA1UEAxMnR29HZXRTU0wgRzQgQ1MgUlNBNDA5NiBTSEEyNTYgMjAyMiBDQS0x
# MB4XDTI1MDcyMTAwMDAwMFoXDTI2MDcyMDIzNTk1OVowcTELMAkGA1UEBhMCR0Ix
# EjAQBgNVBAcTCVdha2VmaWVsZDEmMCQGA1UEChMdU2VjdXJlIFBsYXRmb3JtIFNv
# bHV0aW9ucyBMdGQxJjAkBgNVBAMTHVNlY3VyZSBQbGF0Zm9ybSBTb2x1dGlvbnMg
# THRkMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAERFbrIQcZmiw2ScrP4eHxhzHvoBGn
# AnE3GpY3vjU5CpVG6JLtPgXTQz8aLW7IdGhx7x4cJ3a6y+3/6Q+OX+VVFSiuRd60
# GO22Y2eoMcBmvwc7hWbEYTtdjEzAu82sMmkAo4IB1DCCAdAwHwYDVR0jBBgwFoAU
# yfwQ71DIy2t/vQhE7zpik+1bXpowHQYDVR0OBBYEFPu1ucNQfJlsl2iXm5HJzCZH
# EgfnMD4GA1UdIAQ3MDUwMwYGZ4EMAQQBMCkwJwYIKwYBBQUHAgEWG2h0dHA6Ly93
# d3cuZGlnaWNlcnQuY29tL0NQUzAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYI
# KwYBBQUHAwMwgZcGA1UdHwSBjzCBjDBEoEKgQIY+aHR0cDovL2NybDMuZGlnaWNl
# cnQuY29tL0dvR2V0U1NMRzRDU1JTQTQwOTZTSEEyNTYyMDIyQ0EtMS5jcmwwRKBC
# oECGPmh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9Hb0dldFNTTEc0Q1NSU0E0MDk2
# U0hBMjU2MjAyMkNBLTEuY3JsMIGDBggrBgEFBQcBAQR3MHUwJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBNBggrBgEFBQcwAoZBaHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0dvR2V0U1NMRzRDU1JTQTQwOTZTSEEyNTYyMDIy
# Q0EtMS5jcnQwCQYDVR0TBAIwADANBgkqhkiG9w0BAQsFAAOCAgEACx2bHTrjHk/3
# tX7HUH4SM/9sEtfVFpRtZcV4nmDpjwI7tAwOSGYXk4DLVjJFJveoFjbYsZ8vquZ1
# /HJM7rg+O3rCNmzOBvUXFVSjdL3S2R7+kF2ROR7dqk1/BNW6n3o7Q3BmNGqjo1WH
# jov6PfAbEffCLZI1jT98RNqChMesWMmQS+nf8xwdskne4XZOFX5h/a00X7QLAJ+S
# /bOptiC0SvEEa5FCWPUcV7ML0MtoDc3HIPnmMMuYLy586eJHbE5XlfEsmWUNk3Kf
# hxzxsXpAdTSDOeb5Qm/aHGMOY+56Gnt/zxfrv2bfxPnKKZtXPjA47tm89RHpal8b
# lbCAkVfYpKSe0BFPi8FIk+zXvpoAZkNyCMm/HUMEdMbtRP7CqFmYz0YWuiS3uuUW
# qAZ1zl+n1kIJT8eOu6o01EKS8ShijHUI0vixibiNvwTFgRyX3Yc/9xkfV1Wgzli4
# ZPgoZI6FwYBdrhRF0or+CzYIoENUfUYqI7pBM5kkXuSytFD3SXIeSPx14NZSRTzk
# cdOSJWtLkjLrIrIKzzb5eXxLn/gxmJdssB7GUKZHik+cB0OUCRKHEysBj34hnvXa
# zuQ6DKLOQFy+cZ6z4f2kAeFyq7bWUxctPmF61FkmGvb9q6e3AMLg7JnfYC6EM31u
# 42oGx38b5i0NAiUzvWOAbCWTC+G44pgwggWNMIIEdaADAgECAhAOmxiO+dAt5+/b
# UOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMT
# G0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBaFw0z
# MTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# AL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/z
# G6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZ
# anMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7s
# Wxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL
# 2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfb
# BHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3
# JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3c
# AORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqx
# YxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0
# viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aL
# T8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIBNjAP
# BgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzAf
# BgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMCAYYw
# eQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0aHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# bDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0NcVec4
# X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnovLbc4
# 7/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65ZyoUi0
# mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFWjuyk
# 1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPFmCLB
# sln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9ztwGp
# n1eqXijiuZQwggahMIIEiaADAgECAhAHhD2tAcEVwnTuQacoIkZ5MA0GCSqGSIb3
# DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAX
# BgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0
# ZWQgUm9vdCBHNDAeFw0yMjA2MjMwMDAwMDBaFw0zMjA2MjIyMzU5NTlaMFoxCzAJ
# BgNVBAYTAkxWMRkwFwYDVQQKExBFblZlcnMgR3JvdXAgU0lBMTAwLgYDVQQDEydH
# b0dldFNTTCBHNCBDUyBSU0E0MDk2IFNIQTI1NiAyMDIyIENBLTEwggIiMA0GCSqG
# SIb3DQEBAQUAA4ICDwAwggIKAoICAQCtHvQHskNmiqJndyWVCqX4FtYp5FfJLO9S
# h0BuwXuvBeNYt21xf8h/pLJ/7YzeKcNq9z4zEhecqtD0xhbvSB8ksBAfWBMZO0NL
# fOT0j7WyNuD7rv+ZFza+mxIQ79s1dCiwUMwGonaoDK7mqZfDpKEExR6UyKBh3aat
# T73U2Imx/x+fYTmQFq+N8FrLs6Fh6YEGWJTgsxyw1fAChCfgtEcZkdtcgK7quqsk
# HtW6PJ9l5VNJ7T3WXpznsOOxrz3qx0CzWjwK8+3Kv2X6piWvd8YRfAOycSrT4/PM
# 0cHLFc5xs/4m/ek4FCnYSem43doFftBxZBQkHKoPW3Bt6VIrhVIwvO7hrUjhchJJ
# ZYdSld3bANDviJ5/ToP7ENv97U9MtKFvmC5dzd1p4HxFR0p5wWmYQbW+y3RFm0np
# 6H9m57MUMNp0ysmdJjb0f7+dVLX3OEBUb6H+r1LRLZT/xEOTuwOxGg2S4w25KGL9
# SCBUW4nkBljPHeJToU+THt0P8ZQf4B9IFlGxtLK0g3uOAnwSFgKtmNjhkTl8caLA
# QwbgEINCqrhc0b6k2Z8+QwgVAL0nIuzM9ckKP8xtIcWg85L3/l0cTkHQde+jKGDG
# 2CdxBHtflLIUtwqD7JA2uCxWlIzRNgwT0kH2en0+QV8KziSGaqO2r06kwboq2/xy
# 4e98CEfSYwIDAQABo4IBWTCCAVUwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4E
# FgQUyfwQ71DIy2t/vQhE7zpik+1bXpowHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5n
# P+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcG
# CCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQu
# Y29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGln
# aUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAcBgNV
# HSAEFTATMAcGBWeBDAEDMAgGBmeBDAEEATANBgkqhkiG9w0BAQsFAAOCAgEAC9sK
# 17IdmKTCUatEs7+yewhJnJ4tyrLwNEnfl6HrG8Pm7HZ0b+5Jc+GGqJT8kRc7mihu
# VrdsYNHdicueDL9imhtCusI/rUmjwhtflp+XgLkmgLGrmsEho1b+lGiRp7LC/10d
# i8SAOilDkHj5Zx142xRvBrrWj9eOdSGHwYubAsEd6CDojwcaVz9pfXMzYO3kc0O6
# PXg1TkcgkYlCUAuDHuk/sZx68W0FVj1P2iMh+VUq9lL1puroAydoeWVUh/+cMXeq
# fgpBqlAW+r8ma5F6yKL0stVQH8vYb1ES0mJSIPyIfkIjC1V0pbZS3p0QWsKaafEo
# r8fLfLNfSxntVI/ugut0+6ekluPWRpEXH+JAiNdRjbLbZchCREe3/Xl0YlwkA+eQ
# VJfM0A7XiuFtY/mOpK2AN+E25t5mQYFhpdxZX5LTDKWgDnb+A6QnEt4iNyukcLaJ
# uS8IPgPz0E2ALZLt3Rqs+lXifK/GwnNIWQNbf7FmLDB9ph8i8dvsR1hsjc2KPEW4
# bAsbvLcz8hN1zE1/QbOV92vDGoFjwZOi2koQ+UyEh0e8jDFHAKJeTI+p8EPE/mqv
# ojLFAnt31yXIA2tjt0ERtsjkhBNmZY6SEOfnIoOwvyqavLPya1Ut3/2cOFLuNQ8Q
# l6HaZsNQErnnzn+ZEAaUTkPZaeVyoHIkODECLzkwgga0MIIEnKADAgECAhANx6xX
# Bf8hmS5AQyIMOkmGMA0GCSqGSIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yNTA1MDcwMDAwMDBa
# Fw0zODAxMTQyMzU5NTlaMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBp
# bmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4IC
# DwAwggIKAoICAQC0eDHTCphBcr48RsAcrHXbo0ZodLRRF51NrY0NlLWZloMsVO1D
# ahGPNRcybEKq+RuwOnPhof6pvF4uGjwjqNjfEvUi6wuim5bap+0lgloM2zX4kftn
# 5B1IpYzTqpyFQ/4Bt0mAxAHeHYNnQxqXmRinvuNgxVBdJkf77S2uPoCj7GH8BLux
# BG5AvftBdsOECS1UkxBvMgEdgkFiDNYiOTx4OtiFcMSkqTtF2hfQz3zQSku2Ws3I
# fDReb6e3mmdglTcaarps0wjUjsZvkgFkriK9tUKJm/s80FiocSk1VYLZlDwFt+cV
# FBURJg6zMUjZa/zbCclF83bRVFLeGkuAhHiGPMvSGmhgaTzVyhYn4p0+8y9oHRaQ
# T/aofEnS5xLrfxnGpTXiUOeSLsJygoLPp66bkDX1ZlAeSpQl92QOMeRxykvq6gby
# lsXQskBBBnGy3tW/AMOMCZIVNSaz7BX8VtYGqLt9MmeOreGPRdtBx3yGOP+rx3rK
# WDEJlIqLXvJWnY0v5ydPpOjL6s36czwzsucuoKs7Yk/ehb//Wx+5kMqIMRvUBDx6
# z1ev+7psNOdgJMoiwOrUG2ZdSoQbU2rMkpLiQ6bGRinZbI4OLu9BMIFm1UUl9Vne
# Ps6BaaeEWvjJSjNm2qA+sdFUeEY0qVjPKOWug/G6X5uAiynM7Bu2ayBjUwIDAQAB
# o4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU729TSunkBnx6
# yuKQVvYv1Ensy04wHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYD
# VR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUFBwEBBGsw
# aTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUF
# BzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVk
# Um9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBABfO+xaAHP4HPRF2
# cTC9vgvItTSmf83Qh8WIGjB/T8ObXAZz8OjuhUxjaaFdleMM0lBryPTQM2qEJPe3
# 6zwbSI/mS83afsl3YTj+IQhQE7jU/kXjjytJgnn0hvrV6hqWGd3rLAUt6vJy9lMD
# PjTLxLgXf9r5nWMQwr8Myb9rEVKChHyfpzee5kH0F8HABBgr0UdqirZ7bowe9Vj2
# AIMD8liyrukZ2iA/wdG2th9y1IsA0QF8dTXqvcnTmpfeQh35k5zOCPmSNq1UH410
# ANVko43+Cdmu4y81hjajV/gxdEkMx1NKU4uHQcKfZxAvBAKqMVuqte69M9J6A47O
# vgRaPs+2ykgcGV00TYr2Lr3ty9qIijanrUR3anzEwlvzZiiyfTPjLbnFRsjsYg39
# OlV8cipDoq7+qNNjqFzeGxcytL5TTLL4ZaoBdqbhOhZ3ZRDUphPvSRmMThi0vw9v
# ODRzW6AxnJll38F0cuJG7uEBYTptMSbhdhGQDpOXgpIUsWTjd6xpR6oaQf/DJbg3
# s6KCLPAlZ66RzIg9sC+NJpud/v4+7RWsWCiKi9EOLLHfMR2ZyJ/+xhCx9yHbxtl5
# TPau1j/1MIDpMPx0LckTetiSuEtQvLsNz3Qbp7wGWqbIiOWCnb5WqxL3/BAPvIXK
# UjPSxyZsq8WhbaM2tszWkPZPubdcMIIG7TCCBNWgAwIBAgIQCoDvGEuN8QWC0cR2
# p5V0aDANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGln
# aUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0
# YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExMB4XDTI1MDYwNDAwMDAwMFoX
# DTM2MDkwMzIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEyNTYgUlNBNDA5NiBUaW1lc3Rh
# bXAgUmVzcG9uZGVyIDIwMjUgMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBANBGrC0Sxp7Q6q5gVrMrV7pvUf+GcAoB38o3zBlCMGMyqJnfFNZx+wvA69HF
# TBdwbHwBSOeLpvPnZ8ZN+vo8dE2/pPvOx/Vj8TchTySA2R4QKpVD7dvNZh6wW2R6
# kSu9RJt/4QhguSssp3qome7MrxVyfQO9sMx6ZAWjFDYOzDi8SOhPUWlLnh00Cll8
# pjrUcCV3K3E0zz09ldQ//nBZZREr4h/GI6Dxb2UoyrN0ijtUDVHRXdmncOOMA3Co
# B/iUSROUINDT98oksouTMYFOnHoRh6+86Ltc5zjPKHW5KqCvpSduSwhwUmotuQhc
# g9tw2YD3w6ySSSu+3qU8DD+nigNJFmt6LAHvH3KSuNLoZLc1Hf2JNMVL4Q1Opbyb
# pMe46YceNA0LfNsnqcnpJeItK/DhKbPxTTuGoX7wJNdoRORVbPR1VVnDuSeHVZlc
# 4seAO+6d2sC26/PQPdP51ho1zBp+xUIZkpSFA8vWdoUoHLWnqWU3dCCyFG1roSrg
# HjSHlq8xymLnjCbSLZ49kPmk8iyyizNDIXj//cOgrY7rlRyTlaCCfw7aSUROwnu7
# zER6EaJ+AliL7ojTdS5PWPsWeupWs7NpChUk555K096V1hE0yZIXe+giAwW00aHz
# rDchIc2bQhpp0IoKRR7YufAkprxMiXAJQ1XCmnCfgPf8+3mnAgMBAAGjggGVMIIB
# kTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTkO/zyMe39/dfzkXFjGVBDz2GM6DAf
# BgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezLTjAOBgNVHQ8BAf8EBAMCB4Aw
# FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsGAQUFBwEBBIGIMIGFMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXQYIKwYBBQUHMAKGUWh0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNydDBfBgNVHR8EWDBWMFSgUqBQ
# hk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRUaW1l
# U3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5jcmwwIAYDVR0gBBkwFzAIBgZn
# gQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQBlKq3xHCcEua5g
# QezRCESeY0ByIfjk9iJP2zWLpQq1b4URGnwWBdEZD9gBq9fNaNmFj6Eh8/YmRDfx
# T7C0k8FUFqNh+tshgb4O6Lgjg8K8elC4+oWCqnU/ML9lFfim8/9yJmZSe2F8AQ/U
# dKFOtj7YMTmqPO9mzskgiC3QYIUP2S3HQvHG1FDu+WUqW4daIqToXFE/JQ/EABgf
# ZXLWU0ziTN6R3ygQBHMUBaB5bdrPbF6MRYs03h4obEMnxYOX8VBRKe1uNnzQVTeL
# ni2nHkX/QqvXnNb+YkDFkxUGtMTaiLR9wjxUxu2hECZpqyU1d0IbX6Wq8/gVutDo
# jBIFeRlqAcuEVT0cKsb+zJNEsuEB7O7/cuvTQasnM9AWcIQfVjnzrvwiCZ85EE8L
# UkqRhoS3Y50OHgaY7T/lwd6UArb+BOVAkg2oOvol/DJgddJ35XTxfUlQ+8Hggt8l
# 2Yv7roancJIFcbojBcxlRcGG0LIhp6GvReQGgMgYxQbV1S3CrWqZzBt1R9xJgKf4
# 7CdxVRd/ndUlQ05oxYy2zRWVFjF7mcr4C34Mj3ocCVccAvlKV9jEnstrniLvUxxV
# ZE/rptb7IRE2lskKPIJgbaP5t2nGj/ULLi49xTcBZU8atufk+EMF/cWuiC7POGT7
# 5qaL6vdCvHlshtjdNXOCIUjsarfNZzGCBJswggSXAgEBMG4wWjELMAkGA1UEBhMC
# TFYxGTAXBgNVBAoTEEVuVmVycyBHcm91cCBTSUExMDAuBgNVBAMTJ0dvR2V0U1NM
# IEc0IENTIFJTQTQwOTYgU0hBMjU2IDIwMjIgQ0EtMQIQAdazdTZfIM2RHdcv5fmT
# ZDAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG
# 9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIB
# FTAjBgkqhkiG9w0BCQQxFgQUoo3w7KjjjxEBHZ5wZIyIU3iAIKswCwYHKoZIzj0C
# AQUABGYwZAIwDxs8/dBASE4XoEw0aBd7+hFZfUcBQBYdpD2xCvPhM6veSF90r9y6
# mOgQWCThzm+GAjB5qFGf6u3lT/pfOEzrY2dA+nlr/5K85aAHRhtMAPrYEjyPOoJ+
# g/gqF2hnli3zfTShggMmMIIDIgYJKoZIhvcNAQkGMYIDEzCCAw8CAQEwfTBpMQsw
# CQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERp
# Z2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIw
# MjUgQ0ExAhAKgO8YS43xBYLRxHanlXRoMA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZI
# hvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjYwNTA4MTcxOTU0
# WjAvBgkqhkiG9w0BCQQxIgQgrL6Onlc7Z6zqu6ADJw2ueAd/nzruJQWNW/+Mk3m1
# o+owDQYJKoZIhvcNAQEBBQAEggIAzsGdN3Wo8sdsjCrr7ZyXyD1E711GOFvwhjMN
# IZ0byncfVKNuW7YFCI2sgYONpUL1MPy2MYX33fb6naArtvKyRA/m1vN/oYvFPksh
# EAKGeWCAqhho0muBwlFkfkl2wdo7DGTDeVv3vksKTHQ1+JZf1xnRViw7rYexFP/k
# 1PEI4R8tbEkbfckI+y97tMbLpSk+cEJ3lmbac1qCa2wP1Dr7sTaOwwpskxIMqFVA
# xVjqhvlXXeUW/SR10TREZ3KEOSPVFYkb1zSLuvD6V8Y3JwLrqakSixtbHzlFMjyL
# LviJkyH3csi0SnmTnUvE8adipvssm1b+r/GC8nxssE+7sIm9XDQ/z1Qc2W4o0t+P
# UpG5QuPrNvzvcJgV+WpbhHZ74k0XHLcThFKbtG1qj4h3qT+fF9DxeWWvTcRH3O2a
# H69eSevyOy2oFG8c+OEKuvxSAL2fZZIfA6kgE0e8eqaSfUwBJ9p6Fzm31MbZIVq3
# Ne958CnFXUSYkj7f8Rp4zRTXUAxMANEF/eI2BWIy1JKFDmTGzC2CKXhFFG1d3cr7
# waNQBNt+AVXB0B4sJKhafASq3CdC+QRhDI7oqorXdB0S6ogY6+Bpdhwx8538yIWv
# 9K1d0Wj8wpGLUmUOXRT4GPiMsy0zRAnBLOBIoYXvI1FXxBWj4vwidav2IHko+kMs
# TO9S/K4=
# SIG # End signature block
