#requires -version 2.0
#Original code James Brundage / Microsoft 2009 , published in PSIMageTools
function Add-RotateFlipFilter {
    <#
        .Synopsis
            Adds a Rotate Filter to a list of filters, or creates a new filter
        .Description
            Adds a Rotate Filter to a list of filters, or creates a new filter
        .Example
            $image = Get-Image .\Try.jpg
            $image = $image | Set-ImageFilter -filter (Add-RotateFlipFilter -flipHorizontal -passThru) -passThru
            $image.SaveFile("$pwd\Try2.jpg")
        .Parameter angle
            The Angle of the rotation.  Can only be 0, 90, 180, 270, or 360
        .Parameter flipHorizontal
            If set, the filter will flip images horizontally
        .Parameter flipVertical
            If set, the filter will flip images vertically
        .Parameter passthru
            If set, the filter will be returned through the pipeline.  This should be set unless the filter is saved to a variable.
        .Parameter filter
            The filter chain that the rotate filter will be added to.  If no chain exists, then the filter will be created
    #>
    param(
    [Parameter(ValueFromPipeline=$true)]
    [__ComObject]
    $Filter,

    [ValidateSet(0, 90, 180, 270, 360)]
    [int]$Angle,
    [switch]$FlipHorizontal,
    [switch]$FlipVertical,

    [switch]$PassThru
    )

    process {
        if (-not $Filter) {
            $Filter = New-Object -ComObject Wia.ImageProcess
        }
        $index = $Filter.Filters.Count + 1
        if (-not $Filter.Apply) { return }
        $scale = $Filter.FilterInfos.Item("RotateFlip").FilterId
        $Filter.Filters.Add($scale)
        $Filter.Filters.Item($index).Properties.Item("FlipHorizontal") = "$FlipHorizontal"
        $Filter.Filters.Item($index).Properties.Item("FlipVertical")   = "$FlipVertical"
        $Filter.Filters.Item($index).Properties.Item("RotationAngle")  = $Angle
        if ($PassThru) { return $Filter }
    }
}