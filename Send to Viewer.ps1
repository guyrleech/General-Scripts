#requires -version 4

<#
    Simple text viewer so won't add history anywhere. Takes multiple files, from explorer send to, or if no files, uses clipboard

    @guyrleech 04/12/2020

    Modification History

    06/12/2020  GRL     Tidy up & optimisation
    07/12/2020  GRL     Stopped windows being topmost
    07/12/2020  GRL     Check type of data in clipboard if not text
    05/11/2022  GRL     Added dark mode and search
    15/05/2024  GRL     Changed ReadAllLines to ReadLines, added file list to output when clipboard contains file drop list
    16/05/2024  GRL     Change window title if parent is explorer to help identify window if left open
    18/09/2024  GRL     Added info on data displayed to parent window
    08/08/2026  Claude  Added ctrl + / ctrl - font size zoom and ctrl 0 to reset to starting size, for file and clipboard windows
    23/09/2026  Claude  Added ctrl + mouse wheel font size zoom, for file and clipboard windows
    23/09/2026  Claude  Replaced Get-Clipboard -Format, which is Windows PowerShell 5.1 only, with the WPF clipboard so non-text clipboard contents are reported in pwsh 7 too
    23/09/2026  Claude  Fixed multi line clipboard text all being shown on one line - the array Get-Clipboard returns was being flattened to a space separated string
    23/09/2026  Claude  Laid the search dialogue out in a grid so the controls fill it and resize with it, and removed the duplicate IsDefault so enter runs Find rather than Find All
    24/09/2026  Claude  Made Find All actually highlight every match - it only ever found the first one from the top of the document
    24/09/2026  Claude  Made the search an actual regex, as the dialogue has always claimed, rather than a literal IndexOf

    TODO colour parameters, font, size via reg value
#>

## must be loaded before anything at script level references a WPF type, eg the highlight brushes below
Add-Type -AssemblyName PresentationCore , PresentationFramework , System.Windows.Forms

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
        Title="Search" Height="180" Width="440" MinHeight="170" MinWidth="330">
    <!-- rows: search box, options, spacer so the buttons stay on the bottom edge, buttons -->
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <Label Grid.Row="0" Grid.Column="0" Content="Regex" VerticalAlignment="Center" Margin="0,0,8,0"/>
        <!-- the only control that should soak up extra width -->
        <TextBox x:Name="textBoxSearchText" Grid.Row="0" Grid.Column="1" HorizontalAlignment="Stretch" VerticalAlignment="Center" MinHeight="26" VerticalContentAlignment="Center" Padding="2,0" Text=""/>

        <!-- wraps onto more lines rather than clipping when the dialogue is made narrow -->
        <WrapPanel Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,12,0,0">
            <RadioButton x:Name="radioButtonUp" Content="Up" GroupName="Direction" IsChecked="False" VerticalAlignment="Center" Margin="0,4,20,4"/>
            <RadioButton x:Name="radioButtonDown" Content="Down" GroupName="Direction" IsChecked="True" VerticalAlignment="Center" Margin="0,4,20,4"/>
            <CheckBox x:Name="checkBoxCaseSensitive" Content="Case Sensitive" VerticalAlignment="Center" Margin="0,4,0,4"/>
        </WrapPanel>

        <StackPanel Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,12,0,0">
            <Button x:Name="buttonFind" Content="_Find" MinWidth="88" Height="28" Margin="0,0,8,0" IsDefault="True"/>
            <Button x:Name="buttonFindAll" Content="Find _All" MinWidth="88" Height="28" Margin="0,0,8,0"/>
            <Button x:Name="buttonCancel" Content="_Cancel" MinWidth="88" Height="28" IsCancel="True"/>
        </StackPanel>
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

Function Set-ViewerFontSize
{
    [CmdletBinding()]

    Param
    (
        [Parameter(Mandatory)]
        [ValidateSet( 'Increase' , 'Decrease' , 'Reset' )]
        [string]$action ,
        [double]$step = 2 ,
        [double]$minimum = 4 ,
        [double]$maximum = 96
    )

    if( -Not $script:WPFrichtextboxMain )
    {
        return
    }

    ## capture the size the viewer started at, if not already done, so that reset has something to go back to
    if( $null -eq $script:originalFontSize )
    {
        $script:originalFontSize = $script:WPFrichtextboxMain.FontSize
    }

    [double]$newFontSize = $( switch( $action )
    {
        'Increase' { $script:WPFrichtextboxMain.FontSize + $step }
        'Decrease' { $script:WPFrichtextboxMain.FontSize - $step }
        'Reset'    { $script:originalFontSize }
    })

    ## clamp rather than ignore so that a large change, eg a fast mouse wheel scroll, still takes it as far as it can go
    if( $newFontSize -lt $minimum )
    {
        $newFontSize = $minimum
    }
    elseif( $newFontSize -gt $maximum )
    {
        $newFontSize = $maximum
    }

    if( $newFontSize -ne $script:WPFrichtextboxMain.FontSize )
    {
        $script:WPFrichtextboxMain.FontSize = $newFontSize
        ## the flow document has its own font size so set that too otherwise the displayed text may not follow the control
        if( $script:WPFrichtextboxMain.Document )
        {
            $script:WPFrichtextboxMain.Document.FontSize = $newFontSize
        }
    }
}

Function Add-ViewerFontSizeKeyHandler
{
    [CmdletBinding()]

    Param
    (
        [Parameter(Mandatory)]$form
    )

    ## remember what we started at so that ctrl 0 can restore it
    $script:originalFontSize = $script:WPFrichtextboxMain.FontSize

    ## preview so that the rich text box doesn't get to consume the key first
    $form.Add_PreviewKeyDown( {
        Param
        (
          [Parameter(Mandatory)][Object]$sender,
          [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$event
        )

        ## test with -band so that ctrl shift + (which is how + is typed on most layouts) works too
        if( -Not $event -or ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne [System.Windows.Input.ModifierKeys]::Control )
        {
            return
        }

        ## keys can come from the main keyboard or the numeric keypad
        switch -regex ( [string]$event.Key )
        {
            '^(OemPlus|Add)$'       { Set-ViewerFontSize -action Increase ; $event.Handled = $true }
            '^(OemMinus|Subtract)$' { Set-ViewerFontSize -action Decrease ; $event.Handled = $true }
            '^(D0|NumPad0)$'        { Set-ViewerFontSize -action Reset    ; $event.Handled = $true }
        }
    })
}

Function Add-ViewerFontSizeMouseWheelHandler
{
    [CmdletBinding()]

    Param
    (
        [Parameter(Mandatory)]$form ,
        ## smaller than the keyboard step since the wheel is easier to over do
        [double]$step = 1
    )

    ## remember what we started at so that ctrl 0 can restore it, in case the key handler hasn't already done it
    if( $null -eq $script:originalFontSize )
    {
        $script:originalFontSize = $script:WPFrichtextboxMain.FontSize
    }

    $script:mouseWheelFontSizeStep = $step

    ## preview so that we see the wheel before the rich text box's scroll viewer scrolls the text instead
    $form.Add_PreviewMouseWheel( {
        Param
        (
          [Parameter(Mandatory)][Object]$sender,
          [Parameter(Mandatory)][Windows.Input.MouseWheelEventArgs]$event
        )

        ## test with -band so that ctrl shift wheel, etc, works too
        if( -Not $event -or $event.Delta -eq 0 -or ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne [System.Windows.Input.ModifierKeys]::Control )
        {
            return
        }

        ## one wheel notch is 120 but high resolution wheels and touchpads can send less so always change by at least one step
        [double]$change = $script:mouseWheelFontSizeStep * [math]::Max( 1 , [math]::Round( [math]::Abs( $event.Delta ) / [System.Windows.Input.Mouse]::MouseWheelDeltaForOneLine ) )

        Set-ViewerFontSize -action $(if( $event.Delta -gt 0 ) { 'Increase' } else { 'Decrease' }) -step $change

        ## stop the text scrolling as well as zooming
        $event.Handled = $true
    })
}

## colours used to mark every match that Find All finds - must stand out against the black background and orange text
[System.Windows.Media.Brush]$script:highlightBackground = [System.Windows.Media.Brushes]::Gold
[System.Windows.Media.Brush]$script:highlightForeground = [System.Windows.Media.Brushes]::Black

Function Get-ViewerText
{
    <#
        The text of the whole document. TextRange.Text always terminates the last paragraph with a newline that isn't
        really in the content, so drop it, otherwise the document grows by two characters every time it is rebuilt
    #>

    [CmdletBinding()]

    Param()

    $textRange = New-Object -TypeName System.Windows.Documents.TextRange -ArgumentList $script:WPFrichtextboxMain.Document.ContentStart , $script:WPFrichtextboxMain.Document.ContentEnd
    [string]$text = $textRange.Text

    if( $text.EndsWith( "`r`n" ) )
    {
        $text = $text.Substring( 0 , $text.Length - 2 )
    }
    elseif( $text.EndsWith( "`r" ) -or $text.EndsWith( "`n" ) )
    {
        $text = $text.Substring( 0 , $text.Length - 1 )
    }

    $text ## return
}

Function Get-SearchRegex
{
    <#
        Compile the search text as a regular expression, which is what the dialogue has always said it takes. Returns
        nothing if the pattern won't compile, having told the user why.

        A match timeout is set because a pattern that backtracks badly, eg (a+)+$, would otherwise hang the window with no
        way out - the search runs on the UI thread.
    #>

    [CmdletBinding()]

    Param
    (
        [Parameter(Mandatory)][string]$searchText ,
        [switch]$caseSensitive ,
        [int]$timeoutSeconds = 5
    )

    $options = $(if( $caseSensitive ) { [System.Text.RegularExpressions.RegexOptions]::None } else { [System.Text.RegularExpressions.RegexOptions]::IgnoreCase })

    try
    {
        New-Object -TypeName System.Text.RegularExpressions.Regex -ArgumentList $searchText , $options , ([timespan]::FromSeconds( $timeoutSeconds ))
    }
    catch
    {
        [void][Windows.MessageBox]::Show( "$searchText`r`n`r`n$($_.Exception.Message)" , 'Invalid Regular Expression' , 'OK' , 'Exclamation' )
    }
}

Function Set-AllHighlightedText
{
    <#
        Mark every occurrence of the pattern. A RichTextBox selection is a single contiguous range so it can only ever
        show one match, but any number of matches can be coloured by giving them their own runs.

        The document is rebuilt in a single pass rather than applying a background to each match in turn, because doing it
        one match at a time splits the single run over and over and gets quadratically slower - 5000 matches took over a
        minute that way, against well under a second like this.

        Pass no regex to strip the highlighting and put the document back to a single plain run.

        Returns the number of matches found, or -1 if the pattern timed out.
    #>

    [CmdletBinding()]

    Param
    (
        [System.Text.RegularExpressions.Regex]$regex
    )

    [string]$text = Get-ViewerText

    $inlines = New-Object -TypeName System.Collections.Generic.List[System.Windows.Documents.Inline]
    $firstMatch = $null
    [int]$position = 0
    [int]$matchCount = 0

    if( $null -ne $regex )
    {
        try
        {
            ## Matches returns them in order and non overlapping, and deals with advancing past a zero length match itself
            ForEach( $match in $regex.Matches( $text ) )
            {
                ## a zero width match, eg from ^ or \b, has nothing to colour so skip it
                if( $match.Length -le 0 )
                {
                    continue
                }

                ## the unmatched text since the previous match, left with the control's own colours
                if( $match.Index -gt $position )
                {
                    $inlines.Add( ( New-Object -TypeName System.Windows.Documents.Run -ArgumentList $text.Substring( $position , $match.Index - $position ) ) )
                }

                $matchRun = New-Object -TypeName System.Windows.Documents.Run -ArgumentList $match.Value
                $matchRun.Background = $script:highlightBackground
                $matchRun.Foreground = $script:highlightForeground
                $inlines.Add( $matchRun )

                if( $null -eq $firstMatch )
                {
                    $firstMatch = $matchRun
                }

                $position = $match.Index + $match.Length
                $matchCount++
            }
        }
        catch [System.Text.RegularExpressions.RegexMatchTimeoutException]
        {
            [void][Windows.MessageBox]::Show( "$($regex.ToString())`r`n`r`nThe pattern took too long against this text so the search was abandoned. Try making it less ambiguous." , 'Search Timed Out' , 'OK' , 'Exclamation' )
            return -1
        }
    }

    ## whatever is left after the last match, or the entire text if there were none
    if( $position -lt $text.Length )
    {
        $inlines.Add( ( New-Object -TypeName System.Windows.Documents.Run -ArgumentList $text.Substring( $position ) ) )
    }

    $paragraph = New-Object -TypeName System.Windows.Documents.Paragraph

    ## keep the existing paragraph's layout so replacing it doesn't shift the text about
    if( $existingParagraph = $script:WPFrichtextboxMain.Document.Blocks.FirstBlock )
    {
        $paragraph.Margin = $existingParagraph.Margin
        $paragraph.TextAlignment = $existingParagraph.TextAlignment
        $paragraph.LineHeight = $existingParagraph.LineHeight
    }

    $paragraph.Inlines.AddRange( $inlines )

    [double]$verticalOffset = $script:WPFrichtextboxMain.VerticalOffset

    $script:WPFrichtextboxMain.Document.Blocks.Clear()
    $script:WPFrichtextboxMain.Document.Blocks.Add( $paragraph )

    $script:highlightsApplied = ( $matchCount -gt 0 )

    if( $null -ne $firstMatch )
    {
        ## put the caret at the first match so that a subsequent Find carries on from there, and scroll it into view
        $script:WPFrichtextboxMain.CaretPosition = $firstMatch.ContentStart
        $firstMatch.BringIntoView()
    }
    else
    {
        ## nothing to scroll to so stay where we were
        $script:WPFrichtextboxMain.ScrollToVerticalOffset( $verticalOffset )
    }

    $matchCount ## return
}

Function Clear-AllHighlightedText
{
    ## put the document back to a single unformatted run, keeping the scroll position

    [CmdletBinding()]

    Param()

    if( $script:highlightsApplied )
    {
        $null = Set-AllHighlightedText
        $script:highlightsApplied = $false
    }
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
        ## the dialogue asks for a regex so treat it as one - \$ finds a literal $, \d+ finds digits and so on
        if( -Not ( $regex = Get-SearchRegex -searchText $searchText -caseSensitive:$caseSensitive ) )
        {
            return
        }

        ## a new search starts from a clean slate, otherwise the previous Find All's marks are left behind
        Clear-AllHighlightedText

        if( $findAll )
        {
            if( ( [int]$matchCount = Set-AllHighlightedText -regex $regex ) -eq 0 )
            {
                [void][Windows.MessageBox]::Show( $searchText , 'Text Not Found' , 'OK' , 'Information' )
            }
            elseif( $matchCount -gt 0 )
            {
                ## Focus returns a bool which would otherwise end up in the output
                $null = $WPFrichtextboxMain.Focus()
            }
            ## -1 means the pattern timed out and the user has already been told
            return
        }

        ## find the next match from wherever the caret is - Find All is dealt with above
        $startPoint = $WPFrichtextboxMain.CaretPosition
        $textRange = New-Object -TypeName System.Windows.Documents.TextRange -ArgumentList $startPoint , $WPFrichtextboxMain.Document.ContentEnd

        [string]$textToSearch = $textRange.Text
        $match = $null

        try
        {
            ## skip any zero width match, eg from \b, since there would be nothing to select
            $match = $regex.Match( $textToSearch )
            while( $match.Success -and $match.Length -le 0 )
            {
                $match = $match.NextMatch()
            }
        }
        catch [System.Text.RegularExpressions.RegexMatchTimeoutException]
        {
            [void][Windows.MessageBox]::Show( "$searchText`r`n`r`nThe pattern took too long against this text so the search was abandoned. Try making it less ambiguous." , 'Search Timed Out' , 'OK' , 'Exclamation' )
            return
        }

        if( $match -and $match.Success )
        {
            $offset = $match.Index
            $currentPosition = $textRange.Start.GetInsertionPosition( [System.Windows.Documents.LogicalDirection]::Forward )
            $selectionStart = $currentPosition.GetPositionAtOffset( $offset , [System.Windows.Documents.LogicalDirection]::Forward)
            ## the match length, not the pattern length - they are rarely the same for a regex
            $selectionEnd = $currentPosition.GetPositionAtOffset( $offset + $match.Length , [System.Windows.Documents.LogicalDirection]::Forward)
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
            $null = $WPFrichtextboxMain.Focus()
        }
        else
        {
            [void][Windows.MessageBox]::Show( $searchText , 'Text Not Found' , 'OK' , 'Information' )
        }
    }
}

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

    ## functions which must be available inside each runspace
    [string[]]$functionsForRunspaces = @( 'New-GUI' , 'Set-ViewerFontSize' , 'Add-ViewerFontSizeKeyHandler' , 'Add-ViewerFontSizeMouseWheelHandler' )

    [hashtable]$functionDefinitions = @{}

    ForEach( $functionName in $functionsForRunspaces )
    {
        if( ! ( $functionDefinition = Get-Content -Path "Function:\$functionName" ) )
        {
            Throw "Failed to get definition for function $functionName"
        }
        $functionDefinitions.Add( $functionName , $functionDefinition )
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

    ForEach( $functionName in $functionsForRunspaces )
    {
        $sessionState.Commands.Add( (New-Object -TypeName System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $functionName , $functionDefinitions[ $functionName ] ) )
    }
    
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
            
            Write-Output "$([datetime]::Now.ToString()): $file"

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
                        ## ctrl + / ctrl - to change text size, ctrl 0 to go back to the size it started at
                        Add-ViewerFontSizeKeyHandler -form $mainForm
                        ## ctrl mouse wheel to change text size too
                        Add-ViewerFontSizeMouseWheelHandler -form $mainForm
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
    ## Multi line text comes back as an array of lines and casting that to [string] would join them with spaces, so join with carriage
    ## returns instead, which is what the rich text box wants for a new line. Any line feeds within an element, eg text copied with
    ## unix line endings which doesn't get split into an array, are converted rather than removed so those lines are kept separate too
    [string]$content = @( Get-Clipboard ) -replace '\r?\n' , "`r" -join "`r"

    if( ! [string]::IsNullOrEmpty( $content ) )
    {
        if( $mainForm = New-GUI -inputXAML $mainwindowXAML )
        {
            $script:lastSearch = $null
            $mainForm.Title = "<Contents of clipboard ($($content.Length) characters)>"
            ## output to parent window to help identify what other window will contain
            Write-Output "$([datetime]::Now.ToString()): $($mainForm.Title)"
            ## if launched from shortcut set to run minimised, we most restore the window
            $mainForm.Add_Loaded( {
                $_.Handled = $true
                $mainForm.WindowState = 'Normal'
                $mainForm.Focus()
            })
            ## ctrl + / ctrl - to change text size, ctrl 0 to go back to the size it started at
            Add-ViewerFontSizeKeyHandler -form $mainForm
            ## ctrl mouse wheel to change text size too
            Add-ViewerFontSizeMouseWheelHandler -form $mainForm
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
                        Set-ViewerFontSize -action Increase
                        $event.Handled = $true
                    }
                    elseif( $event.Key -ieq 'd' )
                    {
                        Set-ViewerFontSize -action Decrease
                        $event.Handled = $true
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
        ## Get-Clipboard -Format is Windows PowerShell 5.1 only - it doesn't exist in pwsh 7 and the parameter binding error can't be
        ## silenced with -ErrorAction - so use the WPF clipboard which behaves the same in both

        [string]$messageText = 'No text in clipboard'

        try
        {
            if( [Windows.Clipboard]::ContainsImage() )
            {
                if( $clipboard = [Windows.Clipboard]::GetImage() )
                {
                    ## a BitmapSource, so pixel dimensions rather than the .Width / .Height of the 5.1 System.Drawing.Bitmap
                    $messageText = "Clipboard contains an image $($clipboard.PixelWidth) x $($clipboard.PixelHeight)"
                }
                else
                {
                    $messageText = 'Clipboard contains an image but it could not be retrieved'
                }
            }
            elseif( [Windows.Clipboard]::ContainsAudio() )
            {
                $messageText = 'Clipboard contains audio'
            }
            elseif( [Windows.Clipboard]::ContainsFileDropList() )
            {
                ## a collection of full paths as strings, not the file objects that 5.1's -Format FileDropList returns
                [string[]]$clipboard = @( [Windows.Clipboard]::GetFileDropList() )
                $messageText = "Clipboard contains a file drop list of $($clipboard.Count) items`r`n`r`n$($clipboard -join "`r`n")"
            }
        }
        catch
        {
            ## another process can have the clipboard locked, and the WPF clipboard needs an STA thread
            $messageText = "Failed to read clipboard - $($_.Exception.Message)"
        }

        [void][Windows.MessageBox]::Show( $messageText , 'Viewer Error' , 'OK' , 'Exclamation' )
    }
}
