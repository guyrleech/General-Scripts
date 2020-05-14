#requires -version 3

<#
.SYNOPSIS

Display digitial clock with granularity in seconds or a stopwatch in a window with ability to stop/start

.PARAMETER stopwatch

Run a stopwatch rather than showing current time

.PARAMETER start

Start the stopwatch immediately

.PARAMETER timeoutSeconds

The time in seconds to wait for the user interface to appear before declaring an error

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
#>

[CmdletBinding()]

Param
(
    [switch]$stopWatch ,
    [switch]$start ,
    [int]$timeoutSeconds = 30
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
        Title="Guy's Clock" Height="225.194" Width="404.265">
    <Grid>
        <TextBox x:Name="txtClock" HorizontalAlignment="Left" Height="126" Margin="24,29,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="358" FontSize="72" IsReadOnly="True" FontWeight="Bold" BorderThickness="0"/>
        <CheckBox x:Name="checkboxRun" Content="Run" HorizontalAlignment="Left" Height="18" Margin="24,160,0,0" VerticalAlignment="Top" Width="151" IsChecked="True"/>
        <Button x:Name="btnReset" Content="Reset" HorizontalAlignment="Left" Height="23" Margin="162,155,0,0" VerticalAlignment="Top" Width="147"/>
    </Grid>
</Window>
'@

Function Load-GUI
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

##Start-Transcript -Path (Join-Path -Path $env:temp -ChildPath "clock.thread.log")
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Windows.Forms

if( ! ( $Form = Load-GUI -inputXaml $mainwindowXAML ) )
{
    Exit 1
}

$form.TopMost = $true
$form.Title = $(if( $stopWatch ) { 'Guy''s Stopwatch' } else { 'Guy''s Clock' })
$WPFtxtClock.Text = $(if( $stopWatch ) { '00:00:00.0' } else { Get-Date -Format T })

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
})

[scriptblock]$timerBlock = `
{
    if( $WPFcheckboxRun.IsChecked -and ` 
        ( $newTime = $(if( $stopWatch ) { '{0:d2}:{1:d2}:{2:d2}.{3:d1}' -f $timer.Elapsed.Hours , $timer.Elapsed.Minutes , $timer.Elapsed.Seconds, $( [int]$tenths = $timer.Elapsed.Milliseconds / 100 ; if( $tenths -ge 10 ) { 0 } else { $tenths } ) } else { Get-Date -Format T })) -ne $script:lastTime )
    {
        Write-Debug -Message "New time is $newTime, lasttime was $script:lasttime"
        $script:lastTime = $newTime
        $WPFtxtClock.Text = $newTime
    }
}

## https://richardspowershellblog.wordpress.com/2011/07/07/a-powershell-clock/
$form.Add_SourceInitialized({
    $formTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer
    $formTimer.Interval = $(if( $stopWatch ) { [Timespan]'00:00:00.100' } else { [Timespan]'00:00:00.750' })
    $formTimer.Add_Tick( $timerBlock )
    $formTimer.Start()
})

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
