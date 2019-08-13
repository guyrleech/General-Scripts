<#
    Checks and dumps shortcuts to CSV for easier checking and optionally email as an attachment

    Guy Leech, HCL, 2016
#>

<#
.SYNOPSIS

Produce csv reports of the shortcuts in a given folder and sub-folders and optionally email the resulting csv file

.DESCRIPTION
Can check shortcuts locally (default) or on a remote server, e.g. for checking centralised Citrix XenApp shortcuts.
By default it will check that the target and working directory exist for a shortcut so the resulting csv file can be filtered on these columns to easily find bad shortcuts.

.PARAMETER folder

The top level folder where the shortcuts to be listed are located.

.PARAMETER allUsers

Use the All Users start menu as the folder to check.

.PARAMETER startMenu

User the start menu folder for the user running the script. Will work with folder redirection.

.PARAMETER registry

Read the 'PreferTemplateDirectory' value from the registry which is what Citrix Receiver uses to check for pre-defined shortcuts if using KEYWORDS:Prefer when defining apps in Studio

.PARAMETER csv

The output file name. Will be overwritten if it already exists.

.PARAMETER computerName

If specified, shortcut target paths and working directories will be checked on this computer via the administrative shares, e.g. a shortcut in "c:\program files" will be checked in "\\computername\c$\program files"

.PARAMETER nocheck
 
Target paths and working directories will not be checked.

.PARAMETER mailServer

The SMTP mail server used to send the csv file.

.PARAMETER recipients

A comma separated list of email addresses to which the csv file will be emailed.

.PARAMETER subject

The subject of the email.

.PARAMETER from

The email address from which the email address will be sent.

.PARAMETER useSSL

Communication with the email server will be over SSL.

.EXAMPLE

& '.\Shortcuts to csv.ps1' -folder 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs' -csv c:\temp\shortcuts.csv

Will produce a csv of validated shortcuts in the given folder and all sub-folders to the file "shortcuts.csv" in the "c:\temp" folder

.EXAMPLE

& '.\Shortcuts to csv.ps1' -folder 'D:\Central Store' -csv c:\temp\shortcuts.csv -computername xenapp001 -mailserver smtp -recipients guy.leech@somewhere.com

Will produce a csv of shortcuts validated on the computer "xenapp001" to the specified csv file and then email it to the given recipient via the smtp email server called "smtp"

.NOTES

Initially written to check shortcuts when using in Citrix Receiver to create shortcuts to local apps via "PreferTemplateDirectory" - https://www.citrix.com/blogs/2015/01/06/shortcut-creation-for-locally-installed-apps-via-citrix-receiver/

Due to limitations in the Microsoft API used, the script is currently unable to check internet shortcuts.

#>

[CmdletBinding()]

Param
(
    [string]$folder ,
    [switch]$allusers ,
    [switch]$startmenu ,
    [switch]$registry ,
    [string]$csv ,
    [string]$computerName ,
    [string]$mailserver ,
    [string[]]$recipients ,
    [string]$subject = "Shortcuts export from $env:COMPUTERNAME" ,
    [switch]$nocheck ,
    [string]$from  = "$($env:COMPUTERNAME)@$($env:USERDNSDOMAIN)" ,
    [switch]$useSSL
)

[array]$shortcuts = @()
[int]$count = 0
[int]$bad = 0
[string]$directoryExists = ''
[string]$targetExists = ''

if( ( ! [string]::IsNullOrEmpty( $folder ) -and $allusers ) -or ( $allusers -and $startmenu ) -or ( ! [string]::IsNullOrEmpty( $folder ) -and $startmenu ) )
{
    Write-Error "-registry, -folder, -startmenu and -allusers options are mutually exclusive so only specify one of them"
    Return
}

if( $allusers )
{
    $folder = [environment]::GetFolderPath('CommonStartMenu') 
}
elseif( $startmenu )
{
    $folder = [environment]::GetFolderPath('StartMenu')
}
elseif( $registry )
{
    $folder = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Citrix\Dazzle' -Name 'PreferTemplateDirectory').PreferTemplateDirectory
}

if( [string]::IsNullOrEmpty( $folder ) )
{
    Write-Error "Must specify a valid starting folder via one of the -folder, -registry, -startmenu or -allusers options"
    Return
}

Write-Verbose "Checking folder `"$folder`""

$ShellObject = New-Object -ComObject Wscript.Shell

Get-ChildItem -Path $folder -Include *.lnk -Recurse | ForEach-Object `
{
    $shortcut = $ShellObject.CreateShortcut($_.FullName)
    [string]$targetPath =  $shortcut.TargetPath 
    [string]$workingDirectory = [System.Environment]::ExpandEnvironmentVariables( $shortcut.WorkingDirectory ) ## note local expand, not on remote computer
        
    if( ! $nocheck )
    {
        if( ! [string]::IsNullOrEmpty( $computerName ) )
        {
            if( $shortcut.TargetPath -match '^[a-z]:\\' )
            {
                $targetPath = '\\' + $computerName + '\' + $shortcut.TargetPath.Substring(0,1) + '$' + $shortcut.TargetPath.Substring(2)
            }
            if( $shortcut.WorkingDirectory -match '^[a-z]:\\' )
            {
                $workingDirectory = '\\' + $computerName + '\' + $shortcut.WorkingDirectory.Substring(0,1) + '$' + $shortcut.WorkingDirectory.Substring(2)
            }
        }
        
        Write-Verbose "$count : checking `"$targetPath`" , working dir `"$workingDirectory`""
        if( ! [string]::IsNullOrEmpty( $targetPath ) )
        {
            $targetExists = ( Test-Path $targetPath ) 
        }
        else ## internet shortcuts have an empty target path but there is no way to detect this or find what the URL is AFAIK
        {
            $targetExists = ''
        }
        if( ! [string]::IsNullOrEmpty( $WorkingDirectory ) )
        {
            $directoryExists = ( Test-Path $WorkingDirectory.Trim( '"' ) )  
        }
        else
        {
            $directoryExists = ''
        }

        $shortcut | Add-Member -MemberType NoteProperty -Name 'Target Path Exists' -Value $targetExists
        $shortcut | Add-Member -MemberType NoteProperty -Name 'Working Directory Exists' -Value $directoryExists

        if( $targetExists -eq $false )
        {
            $bad++
            Write-Warning "Target `"$targetPath`" does not exist for shortcut `"$($_.FullName)`""
        }
        elseif( $directoryExists -eq $false )
        {
            Write-Warning "Working directory `"$workingDirectory`" does not exist for shortcut `"$($_.FullName)`""
        }
    }

    $shortcuts += , $shortcut 

    $count++
}

Write-Verbose "Found $count shortcuts, $bad bad"

if( ! [string]::IsNullOrEmpty( $csv ) -and $shortcuts.Count -gt 0 )
{
    $shortcuts | Select "FullName","Arguments","Description","Hotkey","IconLocation","RelativePath","TargetPath", "Target Path Exists","WindowStyle","WorkingDirectory","Working Directory Exists"| Export-Csv $csv -NoTypeInformation
    
    ## workaround for scheduled task not passing array through properly
    if( $recipients.Count -eq 1 -And $recipients[0].IndexOf(",") -ge 0 )
    {
        $recipients = $recipients[0] -split ","
    }

    if( ! [string]::IsNullOrEmpty( $mailserver ) -And $recipients.Count -gt 0 )
    {
        Send-MailMessage -SmtpServer $mailserver -Attachments $csv -To $recipients -Subject $subject -Body "Shortcuts from `"$folder`" exported to csv attached." -From $from -UseSsl:$useSSL
    }
}