#requires -version 3
<#
    Open a given number of handles to a process. Used to test Zombie process detection where the process has exited whilst this script's handles are still open

    @guyrleech 2018
#>

<#
.SYNOPSIS

Open a given number of handles to one or more processes and close them after a specified amount of time or upon user input.
Can be used to test leak detection tools.

.PARAMETER pids

A comma separated list of process ids to open handles for.

.PARAMETER numberOfHandles

The number of handles to open for each process id specified.

.PARAMETER seconds

The number of seconds to wait after opening all handles before closing them. If not specified then user input will be required before closure is performed.

.EXAMPLE

& '.\Zombie Handle Generator.ps1' -pids 1234,5678 -numberOfHandles 100

Opens 100 handles onto the processes with process ids of 1234 and 5678 and waits for the user to hit the enter key before closing all of the handles.

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true)]
    [int[]]$pids ,
    [int]$numberOfHandles = 1 ,
    [int]$seconds 
)

Add-Type @"
using System;

using System.Runtime.InteropServices;

namespace PInvoke.Win32
{
    public static class Process
    {    
     [DllImport( "kernel32.dll",SetLastError = true )]
        public static extern IntPtr OpenProcess( UInt32 dwDesiredAccess, bool bInheritHandle, UInt32 dwProcessId );
     [DllImport( "kernel32.dll",SetLastError = true )]
        public static extern bool CloseHandle( IntPtr hObject );

     public enum Access
     {
        PROCESS_QUERY_INFORMATION = 0x0400 ,
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000,
     };
    }
}
"@

[hashtable]$processInfo = @{} 

[array]$processHandles = @( Get-Process -Id $pids | Where-Object { $_.Id } | ForEach-Object `
{
    $process = $_
    1..$numberOfHandles | ForEach-Object `
    {
        [long]$handle = [Pinvoke.Win32.Process]::OpenProcess(  [Pinvoke.Win32.Process+Access]::PROCESS_QUERY_LIMITED_INFORMATION , $false ,$process.Id )
        if( $handle )
        {
            $handle
            $processInfo.Add( $handle , $process )
        }
        else
        {
            $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Warning "Unable to open handle for $($process.ProcessName) ($($process.id)) - $LastError"
        }
    }
} )

if( $processHandles -and $processHandles.Count )
{
    Write-Output "$(Get-Date -Format G): got $($processHandles.Count) handles to $($pids.Count) processes"

    if( $PSBoundParameters[ 'seconds' ] )
    {
        Write-Output "`tsleeping for $seconds seconds before closing them"
        Start-Sleep -Seconds $seconds
    }
    else
    {
        $null = Read-Host "Hit <enter> to close the handles"
    }

    $processHandles | ForEach-Object `
    {
        $handle = $_
        if( ! [Pinvoke.Win32.Process]::CloseHandle( $handle ) )
        {
            $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            $process = $processInfo[ $handle ]
            Write-Warning "Error closing handle for $($process.ProcessName) ($($process.id)) - $LastError"
        }
    }
}
else
{
    Write-Warning 'Got no process handles'
}

if( ! $PSBoundParameters[ 'seconds' ] )
{
    $null = Read-Host "Hit <enter> to exit the script"
}
