#requires -version 2.0
#Original code James Brundage / Microsoft 2009 , published in PSIMageTools
function Set-ImageFilter {
    <#
        .Synopsis
            Applies an image filter to one or more images
        .Description
            Applies a Windows Image Acquisition filter to one or more Windows Image Acquisition images
        .Example
            $image = Get-Image .\Try.jpg
            $image = $image | Set-ImageFilter -filter (Add-RotateFlipFilter -flipHorizontal -passThru)
            $image.SaveFile("$pwd\Try2.jpg")
        .Parameter image
            The image or images the filter will be applied to
        .Parameter filter
            One or more Windows Image Acquisition filters to apply to the image
        .Parameter NewName
            The name under which the file should be saved. If multiple files are specified,
            a script block should be used to determine the name to be used,  {$_.fullname } will use the file's full name
        .Parameter NoClobber
            This only has an effect when a save name is specified, and ensures an image will not be overwritten
    #>
    [CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName='FileName')]
    param(
        [Parameter(ValueFromPipeline=$true)]
        $Image,
        [__ComObject[]]
        $Filter,
        [Parameter(ParameterSetName='FileName')]
        [Alias("FileName","FullName","SaveName")]
        $NewName ,
        [Parameter(ParameterSetName='Folder')]
        $Destination,
        [switch]$NoClobber,
        [switch]$PassThru,
        [Parameter(DontShow=$true)]
        $psc
    )
    process {
        if ( $null   -eq $psc)    { $psc = $pscmdlet }   ; if (-not $PSBoundParameters.psc) {$PSBoundParameters.add("psc",$psc)}
        if (($Image -is [System.IO.Fileinfo]) -or
            ($Image -is [string])) { $Image = Get-Image $Image}
        if  ($Image.count -gt 1)   { [Void]$PSBoundParameters.Remove("Image") ;  $Image | ForEach-object {Set-ImageFilter -Image $_ @PSBoundParameters} ; return }
        if (-not $image.LoadFile)  { return }
        $OriginalName = $Image.OriginalName
        $noExtension  = $Image.FullName -replace "\.$($Image.FileExtension)",""
        foreach ($f in $Filter) { $Image = $f.Apply($Image.PSObject.BaseObject) }
        $Image        = $Image | Add-Member -NotePropertyName FullName     -NotePropertyValue "$noExtension.$($Image.FileExtension)" -Force -PassThru |
                                 Add-Member -NotePropertyName OriginalName -NotePropertyValue  $OriginalName -Force -PassThru
        if ($Destination) {
            if (-not (Test-Path -PathType Container $Destination)) {New-Item -Type Directory -Path $Destination  -ErrorAction Stop | Out-Null}
            $NewName = Join-path -Path $Destination -ChildPath (Split-Path $image.FullName -Leaf)
        }
        if ($NewName) {
            Write-Verbose -Message "Saving image $NewName "
            $Image | Save-Image -Path $NewName -psc $psc -NoClobber:$NoClobber
            if ($PassThru) {$Image}
        }
        else               {$Image}
    }
}