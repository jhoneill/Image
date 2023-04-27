$Script:ExifToolPath = Split-Path -Path $MyInvocation.PSCommandPath -Parent
function Select-Face {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        $Path                     ,
        # Destination to write the file
        [Parameter(Mandatory=$true)]
        $Destination              ,
        # JPG Quality value 100 = best
        $Quality            = 60  ,
        # Size of final image in pixels
        $PixelSize          = 320 ,
        #How big a border as a proportion of the recognised "face square" size. 50% all round doubles width and height. 100% trebles etc.
        $BorderFactor      = 0.5 ,
        # Path to EXE for EXIFTool (from http://www.sno.phy.queensu.ca/~phil/exiftool/) used to strip EXIF data. get-
        $ExifToolPath       = $Script:ExifToolPath
    )
    process {
        foreach ($file in (Resolve-Path -Path $Path)) { #allow for wildcards or multiple strings in Path parameter
            #We want the Rectangle the Microsoft tagging draws round the picture: it's recorded in the "XAP" XML data Segment
            #It's multiple levels down, in a tag <MPReg:Rectangle xmlns:MPReg="http://ns.microsoft.com/photo/1.2/t/Region#">
            $XAPData                = Read-Exif $file -DumpAsXML  "http://ns.adobe.com/xap/1.0/"
            if ($XAPData -match "<MPReg:Rectangle.*?>([^<]*)</MPReg:Rectangle>") {
                  # If we found the tag, there will be 4 floating point numbers from 0 to 1 giving the top,left, height and width.
                  # We need to tell the crop filter how many pixels to trim off the top edge, bottom edge, left edge, right edge
                  # So do the sums for that and set up the filter.
                  # and add filters to scale it to fit to size and save with known JPG quality
                 $img = Get-Image -Path $file
                 [float[]]$Rectangle = $Matches[1] -split ",\s*"
                 [int]$left          = ($img.Width  * $Rectangle[0]) -  ($img.Width  * $Rectangle[2]) *  $BorderFactor
                 if  ($left -lt 0)     {$left = 0}
                 [int]$right         = ($img.Width  - $left        ) -  ($img.Width  * $Rectangle[2]  * ($BorderFactor * 2 + 1) )
                 if ($right -lt 0)     {$right = 0}
                 [int]$top           = ($img.Height * $Rectangle[1]) -  ($img.Height * $Rectangle[3]) *  $BorderFactor
                 if  ($top -lt 0)      {$top = 0}
                 [int]$bottom        =  $img.Height  - $top          -  ($img.Height * $Rectangle[3]  * ($BorderFactor * 2 + 1) )
                 if ($bottom -lt 0)    {$bottom = 0}
                 $filter             = Add-CropFilter       -Left    $left   -Right    $right       -Top     $top -Bottom $bottom -PassThru
                 $filter             = Add-ScaleFilter      -Filter  $filter -Width    $PixelSize   -Height  $PixelSize           -PassThru
                 $filter             = Add-ConversionFilter -Filter  $filter -TypeName jpg          -Quality $Quality             -PassThru
                 # Get the image, apply the filter to it and save
                 Write-Verbose -Message  "Cropping $file to region $left,$top - $right,$bottom resizing to $PixelSize and writing to $outFile"
                 $img = $img | Set-ImageFilter    -Filter $filter -passThru | Save-Image -Path  $Destination -PassThru
                 & "$exiftoolPath\exifTool.exe"  "-all=" "-overwrite_original" "$Destination" | Out-Null
                 #Make sure file handles are closed.
                 $img = $null
                 [gc]::Collect()
            }
            else {Write-Warning -Message "No Face Rectangle found in $file"}
        }
    }
}
