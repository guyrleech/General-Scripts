#requires -version 3.0
<#
    Designed to sit as an explorer right click menu item. Give a drop down list of algorithm and compute file hashes when OK is clicked for all files passed as arguments

    @guyrleech 2018

    Modification History:

    26/11/18  GRL   Put '<Folder>' in results when item is a folder as cannot checksum a folder, only files
    05/12/18  GRL   Don't show GUI if environment variable CHECKSUM_ALGORITHM is set as it uses that as the checksum algorithm
    27/09/20  GRL   Added support for folders
    07/10/20  GRL   Added file size, last write time and file version, or product version if file version empty, properties to output
    05/12/20  GRL   Minimise PowerShell window if parent is explorer so sendto shortcut doesn't have to be minimised but Out-Gridview window will not be minimised as no easy way to restore grid view windows
    12/12/20  GRL   Fixed bug showing wrong file size
    04/06/24  GRL   Add window title so can tell it's summer. Set verbose on so PS window shows what's occurring
    25/03/25  GRL   Output to stdout too for copy/paste
#>

$VerbosePreference = 'Continue'

[string]$mainwindowXAML = @'
<Window x:Class="Checksummer.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Checksummer"
        mc:Ignorable="d"
        Title="Compute Checksum" Height="182.115" Width="352.35">
    <Grid Margin="0,0,1.676,19.974">
        <ComboBox x:Name="comboAlgorithm" HorizontalAlignment="Left" Height="35.248" Margin="127.238,28.222,0,0" VerticalAlignment="Top" Width="197.389">
            <ComboBoxItem Content="MD5" IsSelected="True"/>
            <ComboBoxItem Content="SHA1"/>
            <ComboBoxItem Content="SHA256"/>
            <ComboBoxItem Content="SHA384"/>
            <ComboBoxItem Content="SHA512"/>
            <ComboBoxItem Content="RIPEMD160"/>
            <ComboBoxItem Content="MACTripleDES"/>
        </ComboBox>
        <Label Content="Algorithm" HorizontalAlignment="Left" Height="31.723" Margin="7.394,31.747,0,0" VerticalAlignment="Top" Width="119.844"/>
        <Button x:Name="btnOK" Content="OK" HorizontalAlignment="Left" Height="35.248" Margin="5.045,83.444,0,0" VerticalAlignment="Top" Width="119.844" IsDefault="True"/>
        <Button x:Name="btnCancel" Content="Cancel" HorizontalAlignment="Left" Height="35.248" Margin="209.484,83.444,0,0" VerticalAlignment="Top" Width="115.143" IsCancel="True"/>
    </Grid>
</Window>
'@

[string]$algorithm = 'MD5'

Function Get-HashAndProperties
{
    Param
    (
        [Parameter(Mandatory,HelpMessage='Full path to the file to checksum')]        
        [string]$path ,
        [ValidateSet( 'MACTripleDES' , 'MD5' , 'RIPEMD160' , ' SHA1' , 'SHA256' , 'SHA384' , 'SHA512' )]
        [string]$algorithm = 'MD5' ,
        [AllowNull()]
        $properties 
    )

    Write-Verbose -Message "$([datetime]::Now.ToString('G')): calculating $algorithm of `"$path`""

    $hash = Get-FileHash -Path $path -Algorithm $algorithm -ErrorAction Continue | Select-Object -ExpandProperty Hash

    if( ! $properties )
    {
        $properties = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
    }
    
    $modified = $fileVersion = $null
    [double]$size = 0

    if( $properties )
    {
        ## check fileversion not null
        if( ! ( $fileVersion = ($properties | Select-Object -ExpandProperty VersionInfo -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FileVersion)) -or [string]::IsNullOrEmpty( $fileVersion.Trim() ) )
        {
            if( ([string]$productVersion = ($properties | Select-Object -ExpandProperty VersionInfo -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ProductVersion)) -and ! [string]::IsNullOrEmpty( $productVersion.Trim() ) )
            {
                $fileVersion = $productVersion
            }
        }
        $modified = $properties.LastWriteTime
        $size = [math]::Round( $properties.Length / 1MB , 1 )
    }

    [pscustomobject][ordered]@{ 'File' = $path ; "$algorithm checksum" = $hash ; "Size (MB)" = $size ; 'File Version' = $fileVersion ; 'Modified' = $modified }
}

Function Load-GUI( $inputXml )
{
    $form = $NULL
    $inputXML = $inputXML -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
 
    [xml]$XAML = $inputXML
 
    $reader = New-Object Xml.XmlNodeReader $xaml

    try
    {
        $Form = [Windows.Markup.XamlReader]::Load( $reader )
    }
    catch
    {
        Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
        return $null
    }
 
    $xaml.SelectNodes('//*[@Name]') | ForEach-Object `
    {
        Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
    }

    return $form
}

if( $host -and $host.name -imatch 'console' )
{
    $parent = $null
    $parent = Get-Process -Id (Get-CimInstance -ClassName win32_process -Filter "ProcessId = '$pid'" -Verbose:$false).ParentProcessId

    if( $parent -and $parent.Name -ieq 'explorer' )
    {
        ## replace double letters with single to make shorter to fit into terminal tab, wg dd/MM/yyyy to d/M/YY
        ## should also work with the illogical US date format
        [string]$dateFormat= (Get-Culture).DateTimeFormat.ShortDatePattern -replace '(\w)\1' , '$1'
        [console]::Title = "$(Split-Path -Path $PSCommandPath -Leaf) $([datetime]::Now.ToString('t')) $([datetime]::Now.ToString( $dateFormat ))"
    }
}

if( -Not $args -or $args.Count -eq 0 )
{
    $null = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    [string]$errorMessage = "No file names passed as arguments"
    $null = [Microsoft.VisualBasic.Interaction]::MsgBox( $errorMessage , 'OKOnly,SystemModal,Exclamation' , (Split-Path -Leaf -Path (& { $myInvocation.ScriptName }) ) )
    Throw $errorMessage
}

[string]$algorithm = $env:CHECKSUM_ALGORITHM

Write-Verbose -Message "$($args.count) arguments passed"

## if algorithm not in %CHECKSUM_ALGORITHM% then prompt via GUI

if( [string]::IsNullOrEmpty( $algorithm ) )
{
    $null = [void][Reflection.Assembly]::LoadWithPartialName('Presentationframework')

    $mainForm = Load-GUI $mainwindowXAML

    if( -Not $mainForm )
    {
        return
    }

    if( $DebugPreference -eq 'Inquire' )
    {
        Get-Variable -Name WPF*
    }
    ## set up call backs

    $WPFbtnOk.add_Click({
        $_.Handled = $true
        $mainForm.DialogResult = $true
        $mainForm.Close()
    })

    $mainForm.add_Loaded({
        if( $_.Source.WindowState -eq 'Minimized' )
        {
            $_.Source.WindowState = 'Normal'
        }
        $_.Handled = $true
    })

    Write-Verbose -Message "Showing GUI to select hash algorithm - set environment variable CHECKSUM_ALGORITHM to always use that algorithm"

    if( $mainForm.ShowDialog() )
    {
        $algorithm = $WPFcomboAlgorithm.SelectedItem.Content
    }
}

if( -Not [string]::IsNullOrEmpty( $algorithm ) )
{
    ## can't easily explicitly make out-gridview window foreground/restored so if parent is explorer.exe we'll hide the PowerShell window

    if( ( [int]$parentProcessId = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = '$pid'" -Verbose:$false | Select-Object -ExpandProperty ParentProcessId ) `
        -and ($parentProcess = Get-Process -Id $parentProcessId -ErrorAction SilentlyContinue) -and $parentProcess.Name -eq 'explorer' )
    {
        ## Executing window may be visible so make it not so
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

        if( ( $process = Get-Process -Id $pid ) -and ( [intptr]$windowHandle = $process.MainWindowHandle ) -and $windowHandle -ne [intptr]::Zero )
        {
            [int]$operation = 7 ## SW_SHOWMINNOACTIVE
            [bool]$setForegroundWindow = [win32.user32]::ShowWindowAsync( $windowHandle , $operation )
            $setForegroundWindow = [win32.user32]::ShowWindowAsync( $windowHandle , $operation ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if( ! $setForegroundWindow )
            {
                Write-Warning -Message "Failed to set window to minimsed for process $($process.Name) (pid $($process.Id)): $lastError"
            }
            else
            {
                Write-Verbose -Message "No error from setting window state of pid $pid to $operation"
            }
        }
        else
        {
            Write-Warning -Message "No main window handle for process $($process.Name) (pid $($process.Id))"
        }
    }

    [array]$results = @( $args | ForEach-Object `
    {
        $item = $_
        try
        {
            if( Test-Path -Path $item -PathType Container -ErrorAction SilentlyContinue )
            {
                Get-ChildItem -Path $item -File -Force -Recurse | ForEach-Object -Process `
                {
                    Get-HashAndProperties -Path $_.FullName -Algorithm $algorithm -Properties $_
                }
            }
            else
            {
                Get-HashAndProperties -path $item -algorithm $algorithm
            }
        }
        catch
        {
            [pscustomobject][ordered]@{ 'File' = $item; "$algorithm checksum" = $_.ToString() }
        }
    })

    if( $results -and $results.Count )
    {
        Write-Verbose -Message "$([datetime]::Now.ToString('G')): got $($results.Count) hashes"
        ## also output results to calling window so easier to copy/paste
        $results ## don't format in case truncates on wrap or similar and can then use in pipeline
        if( $selected = $results | Out-GridView -Title "$algorithm checksums of $($results.Count) files" -PassThru)
        {
            $selected | Set-Clipboard
        }
    }
}
