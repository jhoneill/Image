function Resize-Image {
    param (
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
            $path   ,
        [int]$Width  ,
        [int]$Height ,
        [int]$Quality = 100,
        [string[]]$Replace
    )
    begin {
        $f = New-Imagefilter ;
        Add-ScaleFilter      -Filter $f -Width   $Width   -height   $Height
        Add-ConversionFilter -Filter $f -Quality $Quality -TypeName JPG
    }
    process {
        foreach ($P in $path) {
                Set-ImageFilter -Filter $f -Image $p -NewName {$_.fullname -replace $Replace}  -Verbose
        }
    }
}