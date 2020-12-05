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

    $hash = Get-FileHash -Path $path -Algorithm $algorithm -ErrorAction Continue | Select-Object -ExpandProperty Hash

    if( ! $properties )
    {
        $properties = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
    }
    
    $modified = $fileVersion = $null

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
    }

    [pscustomobject][ordered]@{ 'File' = $path ; "$algorithm checksum" = $hash ; "Size (MB)" = [math]::Round( $_.Length / 1MB , 1 ) ; 'File Version' = $fileVersion ; 'Modified' = $modified }
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
    ## can't easily explicitly make out-gridview window foreground/restored so if parent is explorer.exe we'll hide the PowerShell window

    if( ( [int]$parentProcessId = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = '$pid'" | Select-Object -ExpandProperty ParentProcessId ) `
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
        if( $selected = $results | Out-GridView -Title "$algorithm checksums of $($results.Count) files" -PassThru)
        {
            $selected | Set-Clipboard
        }
    }
}
# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU36e4FaW0kXcjv9guiyOjYwky
# jYigggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
# AQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQg
# Q29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# +NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6kkPApfmJ
# 1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJUM1B0
# sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5wMWYzcT6s
# cKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4Tz
# rGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH5DiLanMg
# 0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIB
# ADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYIKwYBBQUH
# AQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYI
# KwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaG
# NGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0
# dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYE
# FFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2qB1dHC06
# GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5j
# DhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEpKBo6cSgC
# PC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIy
# sjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9CBoYs4Gb
# T8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIFTzCC
# BDegAwIBAgIQBP3jqtvdtaueQfTZ1SF1TjANBgkqhkiG9w0BAQsFADByMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29k
# ZSBTaWduaW5nIENBMB4XDTIwMDcyMDAwMDAwMFoXDTIzMDcyNTEyMDAwMFowgYsx
# CzAJBgNVBAYTAkdCMRIwEAYDVQQHEwlXYWtlZmllbGQxJjAkBgNVBAoTHVNlY3Vy
# ZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMRgwFgYDVQQLEw9TY3JpcHRpbmdIZWF2
# ZW4xJjAkBgNVBAMTHVNlY3VyZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAr20nXdaAALva07XZykpRlijxfIPk
# TUQFAxQgXTW2G5Jc1YQfIYjIePC6oaD+3Zc2WN2Jrsc7bj5Qe5Nj4QHHHf3jopLy
# g8jXl7Emt1mlyzUrtygoQ1XpBBXnv70dvZibro6dXmK8/M37w5pEAj/69+AYM7IO
# Fz2CrTIrQjvwjELSOkZ2o+z+iqfax9Z1Tv82+yg9iDHnUxZWhaiEXk9BFRv9WYsz
# qTXQTEhv8fmUI2aZX48so4mJhNGu7Vp1TGeCik1G959Qk7sFh3yvRugjY0IIXBXu
# A+LRT00yjkgMe8XoDdaBoIn5y3ZrQ7bCVDjoTrcn/SqfHvhEEMj1a1f0zQIDAQAB
# o4IBxTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0O
# BBYEFE16ovlqIk5uX2JQy6og0OCPrsnJMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUE
# DDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUw
# QzA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNl
# cnQuY29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNp
# Z25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAU9zO
# 9UpTkPL8DNrcbIaf1w736CgWB5KRQsmp1mhXbGECUCCpOCzlYFCSeiwH9MT0je3W
# aYxWqIpUMvAI8ndFPVDp5RF+IJNifs+YuLBcSv1tilNY+kfa2OS20nFrbFfl9QbR
# 4oacz8sBhhOXrYeUOU4sTHSPQjd3lpyhhZGNd3COvc2csk55JG/h2hR2fK+m4p7z
# sszK+vfqEX9Ab/7gYMgSo65hhFMSWcvtNO325mAxHJYJ1k9XEUTmq828ZmfEeyMq
# K9FlN5ykYJMWp/vK8w4c6WXbYCBXWL43jnPyKT4tpiOjWOI6g18JMdUxCG41Hawp
# hH44QHzE1NPeC+1UjTGCAigwggIkAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EAT946rb3bWrnkH02dUhdU4wCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFOX6XYRoddKdQrbKhJXK
# TUT2ePoWMA0GCSqGSIb3DQEBAQUABIIBAGOhq9dAb7sCc4QU/8npfqsJrUs3g4Ou
# sDe2bBhY25f09XAjl/fRu1rCT9X4hMM7LaNXEGaSo0DAUiScYYnJ4RdnYBbQFlx9
# TeOQL3ZbRQ5rMJWXlD57t2h7F7wSTXR/mByFgBtgDPH5k9vdsHIG5AqcQy4HbQHY
# pRI2llzrdCM9xAldjBPK27o6K5bsamREFDrUfspw66HE+mceI3UPhIA3W7m4Ruf7
# a2BXYzjj9OJUZPt/oWdC1zusH9lAqnNH9tcZmm9WjpaKBWuN0Hz/C4eSnQmBpyE7
# x63GHq9iPyYqf+CgxpODwEN3UDhhXCrf9qKnguNgzSXFYd2pwHWU62U=
# SIG # End signature block
