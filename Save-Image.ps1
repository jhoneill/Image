function Save-Image {
    <#
      .Synopsis
        Saves a Windows Image Acquisition image
      .Description
        Saves a Windows Image Acquisition image
      .Example
        C:\ps> $image | Save-image -NoClobber -fileName {$_.FullName -replace ".jpg$","-small.jpg$"}
        Saves the JPG image(s) in $image, in the same folder as the source JPG(s), appending
        -small to the file name(s), so that "MyImage.JPG" becomes "MyImage-Small.JPG"
        Existing images will not be overwritten.
      .Parameter Image
        The image or images the filter will be applied to; images may be passed via the pipeline.
        If multiple images are passed either no filename must be included
        (so the image will be saved under its original name), or the fileName must be a code block,
        otherwise the images will all be written over the same file.
      .Parameter PassThru
        If set, the image or images will be emitted onto the pipeline
      .Parameter Filename
        If not set the existing file will be overwritten. The filename may be a string,
        or be a script block - as in the example
      .Parameter NoClobber
        specifies the target file should not be over written if it already exists
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
        $Image,
        [parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true)][ValidateNotNullOrEmpty()]
        $Path ,
        [switch]$PassThru,
        [switch]$NoClobber,
        [Parameter(DontShow=$true)]
        $psc = $PSCmdlet
    )
    process {
        if ( -not $psc -eq $null )          { $psc = $pscmdlet }   ; if (-not $PSBoundParameters.psc) {$PSBoundParameters.add("psc",$psc)}
        if ( $image.count -gt 1     )   { [Void]$PSBoundParameters.Remove("Image") ;  $image | ForEach-object {Save-Image -Image $_ @PSBoundParameters }  ; return}
        if ( $Path -is [scriptblock])   { $fname = Invoke-Expression ([string]$Path )}
        else                            { $fname = $Path }
        if ($Image.OriginalName -eq $fname) { Write-Warning -Message "You can't save over the top of the source file." ; return }
        if (Test-Path -path         $fname) {
                if     ($noclobber)          { Write-Warning -Message "$fName exists and WILL NOT be overwritten"; if ($passthru) {$image} ; return }
                elseIf ($psc.ShouldProcess($FName,"Delete file")) {Remove-Item  -Path $fname -Force -Confirm:$false }
                else   {return}
        }
        if ((Test-Path -Path $Fname -IsValid) -and ($pscmdlet.shouldProcess($FName,"Write image")))  { $image.SaveFile($FName) }
        if ($passthru) {
            $image.fullName = $fname
            $image
        }
   }
}
