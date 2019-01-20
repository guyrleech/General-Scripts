#requires -version 3
<#
    Open a given number of handles to a process. Used to test Zombie process detection where the process has exited whilst this script's handles are still open

    @guyrleech 2018

    Modification History:

    18/01/19  GRL  Threads added
#>

<#
.SYNOPSIS

Open a given number of handles to one or more processes and close them after a specified amount of time or upon user input.
Can be used to test leak detection tools.

.PARAMETER pids

A comma separated list of process ids to open handles for.

.PARAMETER threads

Open handles to threads in the specified process rather than to processes

.PARAMETER numberOfHandles

The number of handles to open for each process id specified.

.PARAMETER seconds

The number of seconds to wait after opening all handles before closing them. If not specified then user input will be required before closure is performed.

.EXAMPLE

& '.\Zombie Handle Generator.ps1' -pids 1234,5678 -numberOfHandles 100

Opens 100 process handles each onto the processes with process ids of 1234 and 5678 and waits for the user to hit the enter key before closing all of the handles.

& '.\Zombie Handle Generator.ps1' -pids 1234,5678 -numberOfHandles 42 -threads -seconds 300

Opens 42 thread handles onto the threads in the processes with process ids of 1234 and 5678 and waits for 300 seconds before closing all of the handles.

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true)]
    [int[]]$pids ,
    [switch]$threads ,
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
        public static extern IntPtr OpenThread( UInt32 dwDesiredAccess, bool bInheritHandle, UInt32 dwThreadId );

     [DllImport( "kernel32.dll",SetLastError = true )]
        public static extern bool CloseHandle( IntPtr hObject );

     public enum ProcessAccess
     {
        PROCESS_QUERY_INFORMATION = 0x0400 ,
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000,
     };
     public enum ThreadAccess
     {
        THREAD_QUERY_INFORMATION = 0x40 ,
        THREAD_QUERY_LIMITED_INFORMATION = 0x800,
     };
    }
}
"@

[hashtable]$processInfo = @{} 
[string]$objectType = if( $threads ) { 'thread' } else { 'process' }
[string]$objectTypePlural = if( $threads ) { 'threads' } else { 'processes' }

[array]$processHandles = @( Get-Process -Id $pids | Where-Object { $_.Id } | ForEach-Object `
{
    $process = $_
    [int]$threadIndex = 0
    [int]$objectId = -1
    1..$numberOfHandles | ForEach-Object `
    {
        [long]$handle = $null
        if( $threads )
        {
            if( $threadIndex -ge $process.Threads.Count )
            {
                $threadIndex = 0
            }
            $objectId = $process.Threads[ $threadIndex ].Id 
            $handle = [Pinvoke.Win32.Process]::OpenThread(  [Pinvoke.Win32.Process+ThreadAccess]::THREAD_QUERY_LIMITED_INFORMATION , $false , $objectId );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            $threadIndex++
        }
        else
        {
            $objectId = $process.Id
            $handle = [Pinvoke.Win32.Process]::OpenProcess(  [Pinvoke.Win32.Process+ProcessAccess]::PROCESS_QUERY_LIMITED_INFORMATION , $false , $objectId );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
        }
       
        if( $handle )
        {
            $handle
            $processInfo.Add( $handle , $process )
        }
        else
        {
            
            Write-Warning "Unable to open handle for $objectType $objectId $($process.ProcessName) - $LastError"
        }
    }
} )

if( $processHandles -and $processHandles.Count )
{
    Write-Output "$(Get-Date -Format G): got $($processHandles.Count) handles to $($pids.Count) $objectTypePlural in process $pid"

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
