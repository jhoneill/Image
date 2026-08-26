

function Get-Exif {
    param (
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [alias('FullName')]
        $Path
    )
    process {
        foreach ($p in $Path) {
            exiftool -json  -x XMP:* -x ICC_Profile:* -x app14:* -x photoshop:* -x Composite:* -G $p  |
                ForEach-Object {$_ -replace '"\w+:([\w-]+":)', '"$1'} |
                    ConvertFrom-Json | Add-Member -PassThru -TypeName 'ExifToolItem'
        }
    }
}