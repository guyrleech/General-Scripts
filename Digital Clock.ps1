#requires -version 3

<#
.SYNOPSIS

Display digitial clock with granularity in seconds or a stopwatch in a window with ability to stop/start

.PARAMETER stopwatch

Run a stopwatch rather than showing current time

.PARAMETER start

Start the stopwatch immediately

.PARAMETER notOnTop

Do not place the window on top of all other windows

.PARAMETER markerFile

A file to look for which will be then seen in a SysInternals Process Monitor trace as a CreateFile operation to allow cross referencing to that

.EXAMPLE

& '.\Digital Clock.ps1'

Display an updating digital clock in a window

.EXAMPLE

& '.\Digital Clock.ps1' -stopwatch

Display a stopwatch in a window but do not start it until the Run checkbox is checked

.EXAMPLE

& '.\Digital Clock.ps1' -stopwatch -start

Display a stopwatch in a window and start it immediately

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
#>

[CmdletBinding()]

Param
(
    [switch]$stopWatch ,
    [switch]$start ,
    [string]$markerFile ,
    [switch]$notOnTop
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
        Title="Guy's Clock" Height="272.694" Width="621.765">
    <Grid>
        <TextBox x:Name="txtClock" HorizontalAlignment="Left" Height="126" Margin="24,29,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="358" FontSize="72" IsReadOnly="True" FontWeight="Bold" BorderThickness="0"/>
        <Grid Margin="24,200,270,10">
            <CheckBox x:Name="checkboxRun" Content="_Run" HorizontalAlignment="Left" Height="18" Margin="-11,5,0,0" VerticalAlignment="Top" Width="76" IsChecked="True"/>
            <Button x:Name="btnReset" Content="Re_set" HorizontalAlignment="Left" Height="23" Margin="147,-1,0,0" VerticalAlignment="Top" Width="65"/>
            <Button x:Name="btnMark" Content="_Mark" HorizontalAlignment="Left" Height="23" Margin="65,0,0,0" VerticalAlignment="Top" Width="65"/>
            <Button x:Name="btnClear" Content="_Clear" HorizontalAlignment="Left" Height="23" Margin="231,-1,0,0" VerticalAlignment="Top" Width="65"/>

        </Grid>
        <TextBox x:Name="txtMarkerFile" HorizontalAlignment="Left" Height="24" Margin="92,154,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="268"/>
        <Label Content="Marker File" HorizontalAlignment="Left" Height="23" Margin="10,155,0,0" VerticalAlignment="Top" Width="72"/>
            <ListView x:Name="listMarkings" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="375,13,10,10" >
                 <ListView.ContextMenu>
                    <ContextMenu>
                        <MenuItem Header="Edit" Name="EditContextMenu" />
                        <MenuItem Header="Delete" Name="DeleteContextMenu" />
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
        Title="Marker Text" Height="285.211" Width="589.034" Name="Marker">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="31*"/>
            <ColumnDefinition Width="217*"/>
            <ColumnDefinition Width="544*"/>
        </Grid.ColumnDefinitions>
        <TextBox x:Name="textBoxMarkerText" Grid.ColumnSpan="2" Grid.Column="1" HorizontalAlignment="Left" Height="97" Margin="0,31,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="533"/>
        <Button x:Name="btnMarkerTextOk" Content="OK" Grid.Column="1" HorizontalAlignment="Left" Height="48" Margin="0,160,0,0" VerticalAlignment="Top" Width="120" IsDefault="True"/>
        <Button x:Name="btnMarkerTextOk_Copy" Content="Cancel" Grid.Column="1" HorizontalAlignment="Left" Height="48" Margin="148,160,0,0" VerticalAlignment="Top" Width="120" Grid.ColumnSpan="2" IsCancel="True"/>
    </Grid>
</Window>
'@

Function Set-MarkerText
{
    Param
    (
        $item , ## if not passed then a new item otherwise editing
        $timestamp
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
        $WPFtextBoxMarkerText.Focus()
        $WPFMarker.Title = "$(if( $item ) { 'Edit' } else { 'Set' }) text for marker @ $(if( $item ) { $item.Timestamp } else { $Timestamp})"

        if( $markerTextForm.ShowDialog() )
        {
            if( $item )
            {
                $item.Notes = $WPFtextBoxMarkerText.Text.ToString()
                $WPFlistMarkings.Items.Refresh()
            }
            else ## new item
            {
                $null = $WPFlistMarkings.Items.Add( ([pscustomobject]@{ 'Timestamp' = $timestamp ; 'Notes' = $WPFtextBoxMarkerText.Text.ToString() }) )  
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

Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Windows.Forms

if( ! ( $Form = New-Form -inputXaml $mainwindowXAML ) )
{
    Exit 1
}

$form.TopMost = ! $notOnTop
$form.Title = $(if( $stopWatch ) { 'Guy''s Stopwatch' } else { 'Guy''s Clock' })
$WPFtxtClock.Text = $(if( $stopWatch ) { '00:00:00.0' } else { Get-Date -Format 'HH:mm:ss' })

$WPFbtnReset.Add_Click({ 
    $_.Handled = $true
    $timer.Reset()
    if( $WPFcheckboxRun.IsChecked )
    {
        $timer.Start() 
    }
    else
    {
        $WPFtxtClock.Text = '00:00:00.0'
    }})

$WPFbtnReset.IsEnabled = $stopWatch

$WPFbtnClear.Add_Click({
    $WPFlistMarkings.Items.Clear()
})

$WPFbtnMark.Add_Click({
    $_.Handled = $true

    [string]$timestamp = $(if( $stopWatch ) { '{0:d2}:{1:d2}:{2:d2}.{3:d3}' -f $timer.Elapsed.Hours , $timer.Elapsed.Minutes , $timer.Elapsed.Seconds, $timer.Elapsed.Milliseconds } else { Get-Date -Format 'HH:mm:ss.ffffff' } )
    
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
    if( $stopWatch )
    {
        if( $WPFcheckboxRun.IsChecked )
        {
            $timer.Start() 
        }
        else
        {
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
        if( $stopWatch )
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
    if( $WPFcheckboxRun.IsChecked -and ` 
        ( $newTime = $(if( $stopWatch ) { '{0:d2}:{1:d2}:{2:d2}.{3:d1}' -f $timer.Elapsed.Hours , $timer.Elapsed.Minutes , $timer.Elapsed.Seconds, $( [int]$tenths = $timer.Elapsed.Milliseconds / 100 ; if( $tenths -ge 10 ) { 0 } else { $tenths } ) } else { Get-Date -Format 'HH:mm:ss' })) -ne $script:lastTime )
    {
        Write-Debug -Message "New time is $newTime, lasttime was $script:lasttime"
        $script:lastTime = $newTime
        $WPFtxtClock.Text = $newTime
    }
}

## https://richardspowershellblog.wordpress.com/2011/07/07/a-powershell-clock/
$form.Add_SourceInitialized({
    if( $formTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer )
    {
        ## need 0.1s granularity for the stopwatch but only just sub-second for the clock
        $formTimer.Interval = $(if( $stopWatch ) { [Timespan]'00:00:00.100' } else { [Timespan]'00:00:00.5' })
        $formTimer.Add_Tick( $timerBlock )
        $formTimer.Start()
    }
})

$WPFtxtMarkerFile.Text = $markerFile

$timer = New-Object -TypeName Diagnostics.Stopwatch

$script:lastTime = $null

if( $stopWatch )
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
