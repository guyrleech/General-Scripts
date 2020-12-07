#requires -version 4

<#
    Simple text viewer so won't add history anywhere. Takes multiple files, from explorer send to, or if no files, uses clipboard

    @guyrleech 04/12/2020

    Modification History

    06/12/2020  GRL  Tidy up & optimisation
    07/12/2020  GRL  Stopped windows being topmost
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
        Title="Text Viewer" Height="450" Width="800">
    <Grid>
        <RichTextBox x:Name="richtextboxMain" HorizontalAlignment="Left" Margin="0" VerticalAlignment="Top" IsReadOnly="False" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto" BorderThickness="0" FontFamily="Consolas" FontSize="14">
            <FlowDocument>
                <Paragraph>
                    <Run/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
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

Add-Type -AssemblyName PresentationCore , PresentationFramework , System.Windows.Forms

if( $args -and $args.Count )
{
    ## as GUI is built in runspace so is harder to see errors, build a dummy now to check it is ok
    if( ! ( New-GUI -inputXAML $mainwindowXAML ) )
    {
        Throw 'Failed to create WPF from XAML'
    }

    if( ! ( $functionDefinition = Get-Content -Path Function:\New-GUI ) )
    {
        Throw 'To get definition for function New-GUI'
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
                    [System.IO.File]::ReadAllLines( $file ) | . { Process `
                    {
                        $WPFrichtextboxMain.AppendText( "$($_)`r" )
                    }}

                    if( ( $textRange = New-Object -TypeName System.Windows.Documents.TextRange( $WPFrichtextboxMain.Document.ContentStart , $WPFrichtextboxMain.Document.ContentEnd  ) ) -and $textRange.Text.Length -gt 0 )
                    {
                        $mainForm.Title = $file
                        ## if launched from shortcut set to run minimised, we most restore the window
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
            $mainForm.Title = "<Contents of clipboard ($($content.Length) characters)>"
            ## if launched from shortcut set to run minimised, we most restore the window
            $mainForm.Add_Loaded( {
                $_.Handled = $true
                $mainForm.WindowState = 'Normal'
                $mainForm.Focus()
            })
            $WPFrichtextboxMain.AppendText( $content )
            [void]$mainForm.ShowDialog()
        }
    }
    else
    {     
        [void][Windows.MessageBox]::Show( "No text in clipboard" , 'Viewer Error' , 'OK' , 'Exclamation' )
    }
}

# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU0Jp6k5uh//kjWnXn9hVNwjpk
# VXigggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFLX3rE/IakDSKzsDtl0y
# JdUbXDArMA0GCSqGSIb3DQEBAQUABIIBABYBS5+2wZj/rLWwJXQgDNaA6BcaWPNp
# YsUOGM22ammjyYfMJNsDxsc1JFgRjnytmEWvDX21/8drdO7AegyhZNje1razet+W
# 9IE56noFEK2PVWDqYYcCT/RDCbVUoX3SRnWosqrJA4d54QscVU/ubOf2lAfdmsK7
# cMoHdoA7TUUiljs0l5Y0GvB7UCsCTewH52fJy2XUTDyjRFpNsv5VdTyTHYUKCY/J
# v2voGFmnGqMyezAth0H7VTJ0kxLLbTejyB9Dd0q3RT4Xgzcf1weFnWTga/1/Xn04
# P7v7N4oykZveQHeu5DD3oZA8qno0aQXrIJzEopFnooigT5DTPyklv4g=
# SIG # End signature block
