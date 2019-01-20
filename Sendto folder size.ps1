#requires -version 3
<#
    Show grid view of disk space used by folders passed as arguments. 
    For all folders selected in grid view, show the largest 50 files.
    For all selected files, prompt to delete them - default is to use recycle bin.
    Designed to be used as a right click explorer SendTo item.

    Script used at own risk. The author accepts no liability for any problems that arise from using this script.

    @guyrleech 2018
#>

[int]$topFiles = 50
[bool]$shellDelete = $true ## set to false to not put in recycle bin

$null = Add-Type -Assembly 'Microsoft.VisualBasic'

if( ! $args -or ! $args.Count )
{
    [string]$errorMessage = "No file names passed as arguments"
    $null = [Microsoft.VisualBasic.Interaction]::MsgBox( $errorMessage , 'OKOnly,SystemModal,Exclamation' , (Split-Path -Leaf -Path (& { $myInvocation.ScriptName }) ) )
    Throw $errorMessage
}

[long]$totalSize = 0

[array]$results = @($args | ForEach-Object `
{
    $size = $null
    if( Test-Path -Path $_ -ErrorAction SilentlyContinue )
    {
        $size = [math]::Round( ( Get-ChildItem -Path $_ -Recurse -Force -File | Measure-Object -Sum -Property Length | Select-Object -Expand Sum ) / 1MB )
        $totalSize += $size
    }
    [pscustomobject][ordered]@{
        'Folder' = $_
        'Size (MB)' = $size }
})

[array]$selected = @( $results | Out-GridView -Title "$($results.Count) folders consuming $($totalSize)MB in total" -PassThru )

if( $selected -and $selected.Count )
{
    ## Now show 50 largest files in these folders
    $top = @( Get-ChildItem -Path ($selected|Select -ExpandProperty 'Folder') -File -Force -Recurse|Sort 'Length' -Descending | Select -First $topFiles )
    $totalFileSize = [math]::Round( ( $top | Measure-Object -Sum -Property Length | Select-Object -Expand Sum ) / 1MB )
    [array]$selectedFiles = @( $top | Select -Property 'FullName',@{n='Size (MB)';e={[math]::Round( $_.Length / 1MB , 1 )}},'CreationTime','LastWriteTime' | Out-GridView -Title "Largest $topFiles files, consuming $($totalFileSize)MB out of $($totalSize)MB" -PassThru )
    if( $selectedFiles -and $selectedFiles.Count )
    {
        [int]$selectedSize = $selectedFiles | Measure-Object -Sum -Property 'Size (MB)' | Select-Object -Expand Sum 
        $answer = [Microsoft.VisualBasic.Interaction]::MsgBox( "Do you wish to delete these $($selectedFiles.Count) files, consuming $($selectedSize)MB?" , 'YesNo,SystemModal,Question' , (Split-Path -Leaf -Path (& { $myInvocation.ScriptName }) ) )
        if( $answer -eq 'Yes' )
        {
            if( $shellDelete )
            {
                $selectedFiles | ForEach-Object `
                {
                    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile( $_.FullName , 'OnlyErrorDialogs' , 'SendToRecycleBin' )
                }
            }
            else
            {
                $removeErrors = $null
                Remove-Item -Path ($selectedFiles|select -ExpandProperty FullName) -Force -Confirm:$false -ErrorVariable RemoveErrors
                if( ! $? -or $removeErrors )
                {
                    $null = [Microsoft.VisualBasic.Interaction]::MsgBox( "Error removing files: $removeErrors" , 'OKOnly,SystemModal,Exclamation' , (Split-Path -Leaf -Path (& { $myInvocation.ScriptName }) ) )
                }
            }
        }
    }
}
