#requires -version 3.0
<#
    Designed to sit as an explorer right click menu item. Give a drop down list of algorithm and compute file hashes when OK is clicked for all files passed as arguments

    @guyrleech 2018

    Modification History:

    26/11/18  GRL   Put '<Folder>' in results when item is a folder as cannot checksum a folder, only files

    05/12/18  GRL   Don't show GUI if environment variable CHECKSUM_ALGORITHM is set as it uses that as the checksum algorithm

    27/09/20  GRL   Added support for folders

    07/10/20  GRL   Added file size, last write time and file version properties to output
#>

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

if( ! $args -or ! $args.Count )
{
    $null = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    [string]$errorMessage = "No file names passed as arguments"
    $null = [Microsoft.VisualBasic.Interaction]::MsgBox( $errorMessage , 'OKOnly,SystemModal,Exclamation' , (Split-Path -Leaf -Path (& { $myInvocation.ScriptName }) ) )
    Throw $errorMessage
}

[string]$algorithm = $env:CHECKSUM_ALGORITHM

## if algorithm not in %CHECKSUM_ALGORITHM% then prompt via GUI

if( [string]::IsNullOrEmpty( $algorithm ) )
{
    $null = [void][Reflection.Assembly]::LoadWithPartialName('Presentationframework')

    $mainForm = Load-GUI $mainwindowXAML

    if( ! $mainForm )
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

    if( $mainForm.ShowDialog() )
    {
        $algorithm = $WPFcomboAlgorithm.SelectedItem.Content
    }
}

if( ! [string]::IsNullOrEmpty( $algorithm ) )
{
    [array]$results = @( $args | ForEach-Object `
    {
        $item = $_
        try
        {
            if( Test-Path -Path $item -PathType Container -ErrorAction SilentlyContinue )
            {
                Get-ChildItem -Path $item -File -Force -Recurse | ForEach-Object -Process `
                {
                    $hash = Get-FileHash -Path $_.FullName -Algorithm $algorithm -ErrorAction Stop | Select-Object -ExpandProperty Hash
                    $fileVersion = $_ | Select-Object -ExpandProperty VersionInfo -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FileVersion
                    [pscustomobject][ordered]@{ 'File' = $_.FullName ; "$algorithm checksum" = $hash ; "Size (MB)" = [math]::Round( $_.Length / 1MB , 1 ) ; 'File Version' = $fileVersion ; 'Modified' = $_.LastWriteTime }
                }
            }
            else
            {
                $hash = Get-FileHash -Path $item -Algorithm $algorithm -ErrorAction Stop | Select-Object -ExpandProperty Hash
                $lastModified = $fileVersion = $null
                if( $properties = Get-ItemProperty -Path $item -ErrorAction SilentlyContinue )
                {
                    $fileVersion = $properties | Select-Object -ExpandProperty VersionInfo -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FileVersion
                    $modified = $properties.LastWriteTime
                }
                [pscustomobject][ordered]@{ 'File' = $item; "$algorithm checksum" = $hash ; "Size (MB)" = [math]::Round( $_.Length / 1MB , 1 ) ; 'File Version' = $fileVersion ; 'Modified' = $lastModified }
            }
        }
        catch
        {
            [pscustomobject][ordered]@{ 'File' = $item; "$algorithm checksum" = $_.ToString() }
        }
    })

    if( $results -and $results.Count )
    {
        if( $selected = $results | Out-GridView -Title "$algorithm checksums of $($results.Count) files" -PassThru )
        {
            $selected | Set-Clipboard
        }
    }
}