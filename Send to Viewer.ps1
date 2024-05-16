#requires -version 4

<#
    Simple text viewer so won't add history anywhere. Takes multiple files, from explorer send to, or if no files, uses clipboard

    @guyrleech 04/12/2020

    Modification History

    06/12/2020  GRL  Tidy up & optimisation
    07/12/2020  GRL  Stopped windows being topmost
    07/12/2020  GRL  Check type of data in clipboard if not text
    05/11/2022  GRL  Added dark mode and search
    15/05/2024  GRL  Changed ReadAllLines to ReadLines, added file list to output when clipboard contains file drop list
    16/05/2024  GRL  Change window title if parent is explorer to help identify window if left open

    TODO colour parameters, font, size via reg value
#>

[CmdletBinding()]

[string]$mainwindowXAML = @'
<Window x:Class="TextViewer.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:TextViewer"
        mc:Ignorable="d"
        Title="Text Viewer" Height="450" Width="800" Background="Black">
    <Grid>
        <RichTextBox x:Name="richtextboxMain" HorizontalAlignment="Left" Margin="0" VerticalAlignment="Top" IsReadOnly="False" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto" BorderThickness="0" FontFamily="Consolas" FontSize="14" Foreground="#FD7609" Background="Black">
            <FlowDocument>
                <Paragraph x:Name="paragraph">
                    <Run x:Name="run"/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
    </Grid>
</Window>
'@

[string]$searchWindowXAML = @'
<Window x:Class="Viewer.Window1"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Viewer"
        mc:Ignorable="d"
        Title="Search" Height="465" Width="828">
    <Grid Margin="0,0,439,308">
        <Label Content="Regex" HorizontalAlignment="Left" Height="34" Margin="16,13,0,0" VerticalAlignment="Top" Width="64" Grid.ColumnSpan="2"/>
        <RadioButton x:Name="radioButtonUp" Content="Up" GroupName="Direction" HorizontalAlignment="Left" Height="22" Margin="16,53,0,0" VerticalAlignment="Top" Width="99" IsChecked="False" Grid.ColumnSpan="2"/>
        <RadioButton x:Name="radioButtonDown" Content="Down" GroupName="Direction" HorizontalAlignment="Left" Height="22" Margin="82,53,0,0" VerticalAlignment="Top" Width="98" IsChecked="True"/>
        <Button x:Name="buttonFind" Content="_Find" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="36" Margin="20,88,0,0" VerticalAlignment="Top" Width="95" IsDefault="True"/>
        <Button x:Name="buttonCancel" Content="_Cancel" HorizontalAlignment="Left" Height="36" Margin="237,88,0,0" VerticalAlignment="Top" Width="94" IsCancel="True"/>
        <TextBox x:Name="textBoxSearchText" Grid.Column="1" HorizontalAlignment="Center" Height="34" Margin="0,10,0,95" TextWrapping="Wrap" Text="" Width="252"/>
        <Button x:Name="buttonFindAll" Content="Find _All" HorizontalAlignment="Left" Height="36" Margin="126,88,0,0" VerticalAlignment="Top" Width="95" IsDefault="True"/>
        <CheckBox x:Name="checkBoxCaseSensitive" Content="Case Sensitive" HorizontalAlignment="Left" Height="18" Margin="174,53,0,0" VerticalAlignment="Top" Width="184"/>

    </Grid>
</Window>   
'@

Function New-GUI( $inputXAML )
{
    $form = $NULL
    [xml]$XAML = $inputXAML -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
  
    if( $reader = New-Object -TypeName Xml.XmlNodeReader -ArgumentList $xaml )
    {
        try
        {
            if( $Form = [Windows.Markup.XamlReader]::Load( $reader ) )
            {
                $xaml.SelectNodes( '//*[@Name]' ) | . { Process `
                {
                    Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName( $_.Name ) -Scope Script
                }}
            }
        }
        catch
        {
            Write-Error "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$($_.Exception.InnerException)"
            $form = $null
        }
    }

    $form ## return
}

Function Set-HighlightedText
{
    Param
    (
        [string]$searchText ,
        [switch]$caseSensitive ,
        [switch]$findAll
    )
    
    [int]$offset = -1

    if( -Not [string]::IsNullOrEmpty( $searchText ))
    {
        $startPoint = $(if( $findAll ) { $WPFrichtextboxMain.Document.ContentStart } else { $WPFrichtextboxMain.CaretPosition })
        $textRange = New-Object -TypeName System.Windows.Documents.TextRange -ArgumentList $startPoint , $WPFrichtextboxMain.Document.ContentEnd

        ## IndexOf is case sensitive so if case-insenistive requested, convert search and text to same case
        [string]$textToSearch = $null
        if( $caseSensitive )
        {
            $textToSearch = $textRange.Text
        }
        else
        {
            $textToSearch = $textRange.Text.ToLower()
            $searchText = $searchText.ToLower()
        }

        if( ($offset = $textToSearch.IndexOf( $searchText )) -ge 0 )
        {
            $currentPosition = $textRange.Start.GetInsertionPosition( [System.Windows.Documents.LogicalDirection]::Forward )
            $selectionStart = $currentPosition.GetPositionAtOffset( $offset , [System.Windows.Documents.LogicalDirection]::Forward)
            $selectionEnd = $currentPosition.GetPositionAtOffset( $offset + $searchText.Length , [System.Windows.Documents.LogicalDirection]::Forward)
            if( $selection = New-Object -TypeName System.Windows.Documents.TextRange -ArgumentList $selectionStart , $selectionend )
            {
                ## need to move the caret to this position so we can scroll to it/make it visible and then make selection
                $WPFrichtextboxMain.CaretPosition = $selectionStart
                $WPFrichtextboxMain.Selection.Select( $selection.Start , $selection.End )
                if( $frameworkContentElement =  [System.Windows.FrameworkContentElement]$selectionStart.Parent )
                {
                    $frameworkContentElement.BringIntoView()                                                                                                                                                                                          
                }
            }
            $WPFrichtextboxMain.Focus()    
        }
        else
        {
            [void][Windows.MessageBox]::Show( $searchText , 'Text Not Found' , 'OK' , 'Information' )
        }
    }
}

Add-Type -AssemblyName PresentationCore , PresentationFramework , System.Windows.Forms

if( $args -and $args.Count )
{
    if( $host -and $host.name -imatch 'console' )
    {
        $parent = $null
        $parent = Get-Process -Id (Get-CimInstance -ClassName win32_process -Filter "ProcessId = '$pid'").ParentProcessId

        if( $parent -and $parent.Name -ieq 'explorer' )
        {
            ## replace double letters with single to make shorter to fit into terminal tab, wg dd/MM/yyyy to d/M/YY
            ## should also work with the illogical US date format
            [string]$dateFormat= (Get-Culture).DateTimeFormat.ShortDatePattern -replace '(\w)\1' , '$1'
            [console]::Title = "$(Split-Path -Path $PSCommandPath -Leaf) $([datetime]::Now.ToString('t')) $([datetime]::Now.ToString( $dateFormat ))"
        }
    }

    ## as GUI is built in runspace so is harder to see errors, build a dummy now to check it is ok
    if( ! ( New-GUI -inputXAML $mainwindowXAML ) )
    {
        Throw 'Failed to create WPF from XAML'
    }

    if( ! ( $functionDefinition = Get-Content -Path Function:\New-GUI ) )
    {
        Throw 'Failed to get definition for function New-GUI'
    }

    ## use runspaces so we can have multiple files open at once, eg to compare
    $jobs = New-Object -TypeName System.Collections.Generic.List[object]
    if( $SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault() )
    {
        $SessionState.ApartmentState = 'STA' ## otherwise get exception making GUI - "The calling thread must be STA, because many UI components require this"
    }
    else
    {
        Throw 'Failed to create a default runspaces session'
    }

    $sessionState.Commands.Add( (New-Object -TypeName System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'New-GUI' , $functionDefinition ) )
    
    [hashtable]$parameters = @{ 'file' = $null ; 'mainwindowXAML' = $mainwindowXAML }

    ForEach( $file in $args )
    {
        ## see if we can read the file because if not there's no point spawning a runspace to do it
        try
        {
            $fileHandle = [System.IO.File]::OpenRead( $file ) 
        }
        catch
        {
            $fileHandle = $null
        }

        if( ! $fileHandle )
        {
            [void][Windows.MessageBox]::Show( $error[0].Exception.InnerException.Message , 'Viewer Error' , 'OK' , 'Exclamation' )
        }
        elseif( $fileHandle.Length -eq 0 )
        {
            [void][Windows.MessageBox]::Show( "Zero length file `"$file`"" , 'Viewer Error' , 'OK' , 'Exclamation' )
        }
        else
        {
            $fileHandle.Close()
            $fileHandle.Dispose()
            $fileHandle = $null

            $parameters.file = $file

            $runspace = [System.Management.Automation.PowerShell]::Create( $sessionState )

            [void]$runspace.AddScript({
                Param( $file , $mainwindowXAML )

                Add-Type -AssemblyName PresentationCore , PresentationFramework , System.Windows.Forms

                if( $mainForm = New-GUI -inputXAML $mainwindowXAML )
                {
                    [System.IO.File]::ReadLines( $file ) | . { Process `
                    {
                        $WPFrichtextboxMain.AppendText( "$($_)`r" )
                    }}

                    if( ( $textRange = New-Object -TypeName System.Windows.Documents.TextRange( $WPFrichtextboxMain.Document.ContentStart , $WPFrichtextboxMain.Document.ContentEnd  ) ) -and $textRange.Text.Length -gt 0 )
                    {
                        $mainForm.Title = $file
                        ## if launched from shortcut set to run minimised, we must restore the window
                        $mainForm.Add_Loaded( {
                            $_.Handled = $true
                            $mainForm.WindowState = 'Normal'
                            $mainForm.Focus()
                        })

                        $mainForm.ShowDialog()
                    }
                    else
                    {
                        [void][Windows.MessageBox]::Show( "No data from `"$file`"" , 'Viewer Error' , 'OK' , 'Exclamation' )
                    }
                }
            })
            [void]$runspace.AddParameters( $parameters )
            [void]$jobs.Add( [pscustomobject]@{ 'Runspace' = $runspace ; 'Handle' = $runspace.BeginInvoke() } )
        }
    }
    
    Write-Verbose -Message "$(Get-Date -Format G): waiting on $($jobs.Count) runspaces to finish"

    ## Wait for dialogs to be closed because if we are running in PowerShell process just for this script, exiting PowerShell will destroy the windows

    $jobs | ForEach-Object `
    {
        if( $_.Runspace.HadErrors )
        {
           $_.Runspace.Streams.Error |  Write-Error
        }
        [void]$_.Runspace.EndInvoke( $_.handle )
        [void]$_.Runspace.Dispose()
    }
    $jobs.clear()

}
else ## no file names so put clipboard contents, if text, into a window
{
    if( ! [string]::IsNullOrEmpty( ( [string]$content = (Get-Clipboard -Format Text -TextFormatType Text -Raw) -replace '\n' ) ) )
    {
        if( $mainForm = New-GUI -inputXAML $mainwindowXAML )
        {
            $script:lastSearch = $null
            $mainForm.Title = "<Contents of clipboard ($($content.Length) characters)>"
            ## if launched from shortcut set to run minimised, we most restore the window
            $mainForm.Add_Loaded( {
                $_.Handled = $true
                $mainForm.WindowState = 'Normal'
                $mainForm.Focus()
            })
            $mainForm.add_KeyDown({
                Param
                (
                  [Parameter(Mandatory)][Object]$sender,
                  [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$event
                )
                if( $event -and ($modifiers = [System.Windows.Input.KeyBoard]::Modifiers) -and $modifiers -eq [System.Windows.Input.ModifierKeys]::Control )
                {
                    if( $event.Key -ieq 'F' )
                    {
                        if( $searchForm = New-GUI -inputXAML $searchWindowXAML )
                        {
                            $WPFtextBoxSearchText.Text = $script:lastSearch
                            $WPFtextBoxSearchText.Focus()

                            $WPFbuttonFind.Add_Click({
                                $searchForm.DialogResult = $true 
                                $searchForm.Close()
                                Set-HighlightedText -caseSensitive:($wpfcheckBoxCaseSensitive.IsChecked) -searchText $WPFtextBoxSearchText.Text
                                $_.Handled = $true
                            })
                            $WPFbuttonFindAll.Add_Click({
                                $searchForm.DialogResult = $true 
                                $searchForm.Close()
                                Set-HighlightedText -findAll -caseSensitive:($wpfcheckBoxCaseSensitive.IsChecked) -searchText $WPFtextBoxSearchText.Text
                                $_.Handled = $true
                            })
                            if( $searchDialogResult = $searchForm.ShowDialog() )
                            {
                                $script:lastSearch = $WPFtextBoxSearchText.Text
                            }
                        }
                        $_.Handled = $true
                    }
                    elseif( $event.Key -ieq 'i' )
                    {
                        $WPFrichtextboxMain.FontSifize += 2
                        $_.Handled = $true
                    }
                    elseif( $event.Key -ieq 'd' )
                    {
                        if( $WPFrichtextboxMain.FontSize -ge 4 )
                        {
                            $WPFrichtextboxMain.FontSize -= 2
                        }
                        $_.Handled = $true
                    }
                }
                elseif( $event.Key -ieq 'F3' )
                {
                    Set-HighlightedText -caseSensitive:($wpfcheckBoxCaseSensitive.IsChecked) -searchText $WPFtextBoxSearchText.Text
                }
            })
            $WPFrichtextboxMain.AppendText( $content )
            [void]$mainForm.ShowDialog()
        }
    }
    else
    {
        if( $clipboard = Get-Clipboard -Format Image -ErrorAction SilentlyContinue )
        {
            [void][Windows.MessageBox]::Show( "Clipboard contains an image $($clipboard.Width) x $($clipboard.Height)" , 'Viewer Error' , 'OK' , 'Exclamation' )
        }
        elseif( $clipboard = Get-Clipboard -Format Audio -ErrorAction SilentlyContinue )
        {
            [void][Windows.MessageBox]::Show( "Clipboard contains audio" , 'Viewer Error' , 'OK' , 'Exclamation' )
        }
        elseif( $clipboard = Get-Clipboard -Format FileDropList -ErrorAction SilentlyContinue )
        {
            [string]$messageText = "Clipboard contains a file drop list of $($clipboard.Count) items`r`n`r`n$($clipboard.FullName -join "`r`n")"
            [void][Windows.MessageBox]::Show( $messageText , 'Viewer Error' , 'OK' , 'Exclamation' )
        }
        else
        {
            [void][Windows.MessageBox]::Show( "No text in clipboard" , 'Viewer Error' , 'OK' , 'Exclamation' )
        }
    }
}
