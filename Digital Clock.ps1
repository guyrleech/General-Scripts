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
#>

[CmdletBinding()]

Param
(
    [switch]$stopWatch ,
    [switch]$start ,
    [int]$timeoutSeconds = 30
)

[int]$exitCode = 0

## code from TTYE
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
#$newRunspace.ThreadOptions = "ReuseThread"         
$newRunspace.Open()
$syncHash = [hashtable]::Synchronized(@{})
$newRunspace.SessionStateProxy.SetVariable( 'syncHash' , $syncHash )

$powerShellScript = [PowerShell]::Create().AddScript({
    try
    {
        [string]$mainwindowXAML = @'
<Window x:Class="Timer.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Timer"
        mc:Ignorable="d"
        Title="Guy's Clock" Height="224.694" Width="373.265">
    <Grid>
        <TextBox x:Name="txtClock" HorizontalAlignment="Left" Height="126" Margin="24,29,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="325" FontSize="72" IsReadOnly="True" FontWeight="Bold" BorderThickness="0"/>
        <CheckBox x:Name="checkboxRun" Content="Run" HorizontalAlignment="Left" Height="18" Margin="24,160,0,0" VerticalAlignment="Top" Width="151" IsChecked="True"/>
        <Button x:Name="btnReset" Content="Reset" HorizontalAlignment="Left" Height="23" Margin="156,155,0,0" VerticalAlignment="Top" Width="147"/>

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
            $inputXML = $inputXaml -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
 
            [xml]$xaml = $inputXML

            if( $xaml )
            {
                $reader = New-Object -TypeName Xml.XmlNodeReader -ArgumentList $xaml

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

        if( $syncHash.Form = Load-GUI -inputXaml $mainwindowXAML )
        {
            $syncHash.Form.TopMost = $true
            $WPFtxtClock.Text = ''
            $syncHash.Clock = $WPFtxtClock
            $syncHash.RunCheckBox = $WPFcheckboxRun
            $syncHash.ResetButton = $WPFbtnReset
            $syncHash.Reset = $false
            $WPFbtnReset.Add_Click( { $_.Handled = $true ; $syncHash.Reset = ! $syncHash.Reset } )
            $synchash.Form.add_KeyDown({
                Param
                (
                  [Parameter(Mandatory)][Object]$sender,
                  [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$event
                )
                if( $event -and $event.Key -eq 'Space' )
                {
                    $_.Handled = $true
                    $WPFcheckboxRun.IsChecked = ! $WPFcheckboxRun.IsChecked
                }    
            })
            $null = $syncHash.Form.ShowDialog()
            $syncHash.Form = $null
        }
        else
        {
            Throw $_
        }
    }
    catch
    {
        $syncHash.Form = $null
        Throw $_
    }
    finally
    {
        ##Stop-Transcript
    }
})
  
$powerShellScript.Runspace = $newRunspace
$invocation = $powerShellScript.BeginInvoke()

## Wait for form to be visible
$timer = [Diagnostics.Stopwatch]::StartNew()
  
## Start-Sleep doesn't return anything so we are just sleeping part way through the while statement :-)
do
{
    try
    {
        $notDone = (  ! $invocation.IsCompleted -and ! $syncHash.Contains( 'Form' ) -and ! (Start-Sleep -Milliseconds 333) -and ! $syncHash.Form.PSObject.Properties[ 'Handle' ]  )
    }
    catch
    {
        $notDone = $True
        Write-Debug -Message $_
    }
    if( ! $notDone -and $timer.Elapsed.TotalSeconds -gt $timeoutSeconds )
    {
        Throw "Timeout waiting for form to appear"
    }
} while( $notDone )
    
$timer.Stop()

if( $syncHash.Contains( 'Form' ) -and $syncHash.Form -and $syncHash.Clock )
{
    $notStarted = $false
    $lastTime = $null

    if( $stopWatch )
    {
        $timer.Reset()
        if( $start )
        {
            $timer.Start()
        }
        else
        {
            $syncHash.Form.Dispatcher.Invoke( 'Normal' , [action]{ $syncHash.Clock.Text = '00:00:00' } )
            $notStarted = $true
        }
        $syncHash.Form.Dispatcher.Invoke( 'Normal' , [action]{ $syncHash.RunCheckBox.IsChecked = $start } )   
    }
    else
    {
        $syncHash.Form.Dispatcher.Invoke( 'Normal' , [action]{ $syncHash.ResetButton.IsEnabled = $false } )      
    }
    While( $syncHash.Form -and $syncHash.Clock -and $syncHash.Form.IsVisible )
    {
        ## for reasons unknown the IsChecked property doesn't work so this parses the text "System.Windows.Controls.CheckBox Content:Run IsChecked:False"
        if( ! $syncHash.RunCheckBox -or $syncHash.RunCheckBox.ToString() -match 'IsChecked:True' )
        {
            if( $stopWatch )
            {
                if( $syncHash.Reset -or $notStarted )
                {
                    $timer.Reset()
                    $timer.Start()
                    $syncHash.Reset = $notStarted = $false
                }
                elseif( ! $timer.IsRunning )
                {
                    $timer.Start()
                }
                $timeNow = "{0:d2}:{1:d2}:{2:d2}" -f $timer.Elapsed.Hours , $timer.Elapsed.Minutes , $timer.Elapsed.Seconds
            }
            else
            {
                $timeNow = Get-Date -Format T
            }
            if( $timeNow -ne $lastTime )
            {
                $syncHash.Form.Dispatcher.Invoke( 'Normal' , [action]{ $syncHash.Clock.Text = $timeNow } )
                $lastTime = $timeNow
            }
        }
        elseif( $syncHash.Reset )
        {
            $timer.Reset()
            $syncHash.Reset = $false
            $syncHash.Form.Dispatcher.Invoke( 'Normal' , [action]{ $syncHash.Clock.Text = '00:00:00' } )
        }
        elseif( $stopWatch )
        {
            $timer.Stop()
        }
        Start-Sleep -Milliseconds 333 ## yuck
    }
}
elseif( ! $invocation.IsCompleted )
{
    ## Terminate thread
    $powerShellScript.Stop()
    $powerShellScript.Dispose()
}
else
{
    $result = $powerShellScript.EndInvoke( $invocation )
    if( $result )
    {
        Write-Error -Message "Failed to create Windows form: $result"
        $exitCode = 1
    }
}

Exit $exitCode
