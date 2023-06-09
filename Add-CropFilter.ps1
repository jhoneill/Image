#requires -version 2.0
#Original code James Brundage / Microsoft 2009
function Add-CropFilter {
    <#
        .Synopsis
            Adds a Crop Filter to a list of filters, or creates a new filter
        .Description
            Adds a Crop Filter to a list of filters, or creates a new filter
        .Example
            $Image = Get-Image .\Try.jpg
            $Image = $Image | Set-ImageFilter -filter (Add-CropFilter -Image $Image -Left .1 -Right .1 -Top .1 -Bottom .1 -passThru) -passThru
            $Image.SaveFile("$pwd\Try2.jpg")
        .Parameter Image
            Optional.  If set, allows you to specify the crop in terms of a percentage
        .Parameter Left
            The number of pixels to crop from the left (if left is greater than 1) or the percentage of space to crop from the left (if image is provided)
        .Parameter Top
            The number of pixels to crop from the top (if top is greater than 1) or the percentage of space to crop from the top (if image is provided)
        .Parameter Right
            The number of pixels to crop from the right (if right is greater than 1) or the percentage of space to crop from the right (if image is provided)
        .Parameter Bottom
            The number of pixels to crop from the bottom (if bottom is greater than 1) or the percentage of space to crop from the bottom (if image is provided)
        .Parameter Passthru
            If set, the filter will be returned through the pipeline.  This should be set unless the filter is saved to a variable.
        .Parameter Filter
            The filter chain that the rotate filter will be added to.  If no chain exists, then the filter will be created
    #>
        param(
        [Parameter(ValueFromPipeline=$true)]
        [__ComObject]
        $Filter,

        [__ComObject]
        $Image,

        [Double]$Left,
        [Double]$Top,
        [Double]$Right,
        [Double]$Bottom,

        [switch]$PassThru
    )

    process {
        if (-not $Filter) {
            $Filter = New-Object -ComObject Wia.ImageProcess
        }
        $index = $Filter.Filters.Count + 1
        if (-not $Filter.Apply) { return }
        $crop = $Filter.FilterInfos.Item("Crop").FilterId
        if ($Left   -lt 0) { $Left   = 0 }
        if ($Top    -lt 0) { $Top    = 0 }
        if ($Right  -lt 0) { $Right  = 0 }
        if ($Bottom -lt 0) { $Bottom = 0 }
        $isPercent   =  $true
        if ($Left   -gt 1) { $isPercent = $false }
        if ($Top    -gt 1) { $isPercent = $false }
        if ($Right  -gt 1) { $isPercent = $false }
        if ($Bottom -gt 1) { $isPercent = $false }
        $Filter.Filters.Add($crop)
        if ($isPercent -and $Image) {
            $Filter.Filters.Item($index).Properties.Item("Left")   = $Image.Width  * $Left
            $Filter.Filters.Item($index).Properties.Item("Top")    = $Image.Height * $Top
            $Filter.Filters.Item($index).Properties.Item("Right")  = $Image.Width  * $Right
            $Filter.Filters.Item($index).Properties.Item("Bottom") = $Image.Height * $Bottom
        } else {
            $Filter.Filters.Item($index).Properties.Item("Left")   = $Left
            $Filter.Filters.Item($index).Properties.Item("Top")    = $Top
            $Filter.Filters.Item($index).Properties.Item("Right")  = $Right
            $Filter.Filters.Item($index).Properties.Item("Bottom") = $Bottom
        }
        if ($passthru) { return $Filter }
    }
}
