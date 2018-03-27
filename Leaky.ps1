<#
    Leak working set memory. For demonstration purposes only. Do not use in a production environment. Just don't. Ever.

    Guy Leech, 2018
#>


<#
.SYNOPSIS

Allocate working set memory to simulate a process that leaks memory.

.DESCRIPTION

Run with -verbose to show progress messages

.PARAMETER memoryChunk

Size of the memory chunk to allocate on each loop iteration

.PARAMETER frequency

How often in seconds to allocate the memory chunk

.PARAMETER duration

How long in seconds to run for

.PARAMETER waitToQuit

Require <enter> key to be pressed before the script finishes

.PARAMETER dontFree

Do not free any of the allocated memory. The parent PowerShell process will still be consuming the leaked memory

.EXAMPLE

.\leaky.ps1 -memoryChunk 250MB -frequency 10 -duration 60 -verbose

Allocate 250MB of extra working set every 10 seconds for 60 seconds in total, showing verbose information, and then exit, freeing the "leaked" memory first

.EXAMPLE 

.\leaky.ps1 -memoryChunk 100MB -frequency 15 -duraton 120 -verbose -waitToQuit

Allocate 100MB of extra working set every 15 seconds for 120 seconds in total, showing verbose information, and then exit, freeing the "leaked" memory first, when <Enter> has been pressed

.NOTES

Since it is the parent powershell.exe process that is consuming the memory, if the -dontFree option is used, that PowerShell process will still be consuming the total amount of leaked working set until it is exited

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true)]
    [long]$memoryChunk ,
    [int]$frequency = 30 ,
    [int]$duration = 300 ,
    [switch]$waitToQuit ,
    [switch]$dontFree
)

Add-Type @'
using System;
using System.Runtime.InteropServices;

namespace PInvoke.Win32
{
  
    public static class Memory
    {
        [DllImport("msvcrt.dll", SetLastError=true)]
        public static extern IntPtr malloc( int dwBytes );

        [DllImport("msvcrt.dll", SetLastError=true)]
        public static extern void free( IntPtr memBlock );

        [DllImport("ntoskrnl.exe", SetLastError=true)]
        public static extern void RtlZeroMemory( IntPtr destination , int length );
    }
}
'@

$memories = New-Object -TypeName System.Collections.ArrayList
[long]$allocated = 0
$timer = [Diagnostics.Stopwatch]::StartNew()

Write-Verbose "$(Get-Date) : script started, working set initially $([math]::Round( (Get-Process -Id $pid).WorkingSet64 / 1MB ))MB for process $pid "

While( $timer.Elapsed.TotalSeconds -le $duration )
{
    [long]$memory = [PInvoke.Win32.Memory]::malloc( $memoryChunk ) ; $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
    if( $memory )
    {
        [PInvoke.Win32.Memory]::RtlZeroMemory( $memory , $memoryChunk ) ## Need to use memory in order for it to actually get added to the working set
        $null = $memories.Add( $memory ) ## save the pointer lest we actually decide to free it later
        $allocated += $memoryChunk
    }
    else
    {
        Write-Error "$(Get-Date) : Failed to allocate $($memoryChunk / 1MB)MB - $LastError"
    }
    Write-Verbose "$(Get-Date) : total allocated $($allocated / 1MB)MB , working set now $([math]::Round( (Get-Process -Id $pid).WorkingSet64 / 1MB ))MB for process $pid - sleeping for $frequency seconds ..."

    Start-Sleep -Seconds $frequency
}

$timer.Stop()

if( $waitToQuit )
{
    $null = Read-Host "$(Get-Date) : hit <Enter> to exit "
}

if( ! $dontFree )
{
    ## We weren't really leaking! 
    Write-Verbose "Freeing $($memories.Count) allocations of $($memoryChunk / 1MB)MB each"
    $memories | ForEach-Object { [PInvoke.Win32.Memory]::free( $_ ) }
}

Write-Verbose "$(Get-Date) : script finished, working set now $([math]::Round( (Get-Process -Id $pid).WorkingSet64 / 1MB ))MB, peak $([math]::Round( (Get-Process -Id $pid).PeakWorkingSet64 / 1MB ))MB for process $pid "
