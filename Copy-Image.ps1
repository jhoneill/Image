function Copy-Image {
    <#
        .Synopsis
            Copies an image, applying EXIF data from GPS data points
        .Description
            Copies an image, applying EXIF data from GPS data points
        .Example
            C:\PS>  Dir E:\dcim –inc IMG*.jpg –rec | Copy-Image -Keywords "Oxfordshire" -rotate -DestPath "$env:userprofile\pictures\oxford" -replace  "IMG","OX-"
            Copies IMG files from folders under E:\DCIM to the user's picture\Oxford folder, replacing IMG in the file name with OX-.
            The Keywords field is set to Oxfordshire, pictures are GeoTagged with the data in $points and rotated.
        .Parameter Image
            A WIA image object, a path to an image, or a file object representing an image file. It may be passed via the pipeline.
        .Parameter Destination
            The FOLDER to which the file should be saved.
        .Parameter Keywords
            If specified, sets the keywords Exif field.
        .Parameter Title
            If specified, sets the Title Exif field..
        .Parameter Replace
            If specified, this contains two values seperated by a comma specifying a replacement in the file name
        .Parameter Rotate
            If this switch is specified, the image will be auto-rotated based on its orientation filed
        .Parameter NoClobber
            Unless this switch is specified, a pre-existing image WILL be over-written
        .Parameter ReturnInfo
            If this switch is specified, the path to the saved image will be returned.
    #>
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true, Mandatory=$true)]
        [Alias("Path","FullName")]$Image ,
        [ValidateScript({Test-path -Path $_ })]
        [string]$Destination ,
        [String[]]$Keywords ,
        [string]$Title,
        [String[]]$Replace,
        [__ComObject]$Filter,
        [int]$Width,
        [int]$Height,
        [int]$Quality,
        [switch]$Rotate,
        [switch]$NoClobber,
        [switch]$ReturnInfo,
        [Parameter(DontShow=$true)]
        $psc = $pscmdlet
    )
    process {
        if ($null  -eq $psc )  {$psc = $pscmdlet} ; if (-not $PSBoundParameters.psc) {$PSBoundParameters.add("psc",$psc)}
        if ($Image -is [system.io.fileinfo] ) {$Image = $Image.FullName }
        if ($Image -is [String]             ) {[Void]$PSBoundParameters.Remove("Image")
                                               Get-Image $Image | Copy-Image @PSBoundParameters
                                               return
        }
        if ($Image.count -gt 1              ) {[Void]$PSBoundParameters.Remove("Image")
                                               $Image | ForEach-object {Copy-Image -image $_ @PSBoundParameters}
                                               return
        }
        if ($Image -is [__comObject]) {Write-Verbose             -Message ("Processing " + $Image.fullname)
           if (-not $Filter)          {$Filter = New-Imagefilter}
           if ($Rotate)               {$orient = (Read-Exif -path $Image).Orientation}  # Leave $orient unset if we aren't rotating
           if ($Keywords)             {Add-EXIFFilter            -Filter  $Filter      -Exiftag Keywords              -TypeID   1101 -String ($Keywords -join ";") }
           if ($Title)                {Add-EXIFFilter            -Filter  $Filter      -Exiftag Title                 -TypeID   1101 -String  $Title    }
           if ($orient -eq 8)         {Add-RotateFlipFilter      -Filter  $Filter      -Angle   270   # Orientation 8=90 degrees, 6=270 degrees, rotate round to 360
                                       Add-EXIFFilter            -Filter  $Filter      -Exiftag Orientation -TypeID   1003 -Value   1
                                       Write-Verbose             -Message "Rotating image counter-clockwise"}
           if ($orient -eq 6)         {Add-RotateFlipFilter      -Filter  $Filter      -Angle   90
                                       Add-EXIFFilter            -Filter  $Filter      -Exiftag Orientation -TypeID   1003 -Value   1
                                       Write-Verbose             -Message "Rotating image clockwise"}
           if ($Width -and $Height)   {Add-ScaleFilter           -Filter  $Filter      -Width   $Width               -Height   $Height}
           if ($Quality)              {Add-ConversionFilter      -Filter  $Filter      -Quality $Quality             -TypeName JPG}
           if (-not $Destination)     {$Destination = Split-Path -Path    $Image.FullName }
           if ($Replace)              {$SavePath    = Join-Path  -Path    (Resolve-Path -path $Destination) -ChildPath ((Split-Path -Path $Image.FullName -Leaf) -Replace $Replace)}
           else                       {$SavePath    = Join-Path  -Path    (Resolve-Path -path $Destination) -ChildPath  (Split-Path -Path $Image.FullName -Leaf)  }
           if ($Filter.Filters.Count) {$Image | Set-ImageFilter  -Filter  $Filter      -NewName $savePath -psc $psc  -noClobber:$NoClobber}
           else                       {$Image.SaveFile($SavePath)}
           $orient = $Image =  $filter = $null
           if ($returnInfo) {$SavePath}
        }
    }
}