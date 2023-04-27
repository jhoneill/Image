
Param (
    [ValidateNotNullOrEmpty()]
    $path = $pwd,
    [ValidateRange(100,4000)]
    $width  = 1500,
    [ValidateRange(100,4000)]
    $height = 1500,
    $ExifToolPath = (Get-Command exiftool.exe -ErrorAction SilentlyContinue).source
)
if (-not $ExifToolPath -or -not (Test-path $ExifToolPath)) {throw "Can't find exiftool please pass it the path to it as a parameter"; return }

$f = New-Imagefilter | Add-ScaleFilter -Width $width -Height $height -PassThru
Get-Image -Path  (Join-Path -Path $path -ChildPath '*.jpg') |
    Where-Object {$_.width -gt 1500 -or $_.height -gt 1500} |
        Set-ImageFilter -Filter $f -PassThru |
            Save-image  -NoClobber -Path {$image.FullName -replace ".jpg$","-small.jpg"} -PassThru |
                ForEach-Object {& $ExifToolPath "-all=" "-overwrite_original" $_.fullName}
