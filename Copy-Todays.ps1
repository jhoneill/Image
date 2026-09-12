#PSUseSingularNouns
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns','',Justification='Posessive not plural ')]
param()
Function Copy-Todays {
    [CmdletBinding()]
    <#
    .SYNOPSIS
        Copies todays files from (by default) the DCIM folder on a memory, to (by default) the current folder, optionally renaming as we go
    .EXAMPLE
        Copy-Todays -Rename "^IM","Dive"
    #>
    param  (
        # Copy from
        [string]$Path        = "D:\DCIM",
        # Copy to
        [String]$Destination = $pwd,
        # Reg-ex to replace in the name and text to replace it with.
        [ValidateCount(2,2)]
        [String[]]$Replace,
        $AddHours = 0
    )
    $cutOff = [datetime]::Today.AddHours($AddHours)

    if ($Path -like "D:\*" -and -not (Test-path $Path -PathType Container) -and (Test-path ($Path -replace "^D:", "E:") -PathType Container)) {
        $Path = $Path -replace "^D:", "E:"
    }
    elseif (-not (Test-Path $Path -PathType Container) ) {
        Write-Error "Source path not found : $Path " ; return
    }
    elseif (-not (Test-Path $Destination -PathType Container) ) {
        Write-Error "Destination path not found : $Destination" ; return
    }
    # Only copy files from the right time if the number isn't in the destination. If we louse up renaming don't make a second copy.
    $filesToCopy = Get-childitem $Path -Recurse -File  | Where-Object  {$_.LastWriteTime -GT $cutOff -and -not (Test-Path (Join-path $Destination ($_.name -replace '^\D+','*' ) ))}  |

    if (-not $Replace) {
        $filesCopied = $filesToCopy  | Copy-Item -Destination $Destination -Verbose -PassThru
    }
    else  {
        $filesCopied = $filesToCopy  | Copy-Item -Destination {Join-Path $Destination ($_.Name -replace $replace)} -Verbose
    }
    $filesCopied | Measure-Object  -Sum -Property  length |
        Format-Table       -Property  Count, @{n="Bytes";e={$_.sum.tostring("N0")}}
}