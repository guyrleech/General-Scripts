#requires -version 3.0
<#
    Designed to sit as an explorer right click menu item. Give a drop down list of algorithm and compute file hashes when OK is clicked for all files passed as arguments

    @guyrleech 2018

    Modification History:

    26/11/18  GRL   Put '<Folder>' in results when item is a folder as cannot checksum a folder, only files
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

    [array]$results = @( $args | ForEach-Object `
    {
        [string]$result = `
            try
            {
                if( Test-Path -Path $_ -PathType Container -ErrorAction SilentlyContinue )
                {
                    '<Folder>'
                }
                else
                {
                    Get-FileHash -Path $_ -Algorithm $algorithm -ErrorAction Stop | Select -ExpandProperty Hash
                }
            }
            catch
            {
                $_.ToString()
            }
        [pscustomobject][ordered]@{ 'File' = $_ ; "$algorithm checksum" = $result }
    })

    if( $results -and $results.Count )
    {
        $selected = $results | Out-GridView -Title "$algorithm checksums of $($args.Count) files" -PassThru
        if( $selected )
        {
            $selected | Set-Clipboard
        }
    }
}