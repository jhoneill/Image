#requires -version 2.0
#Original code James Brundage / Microsoft 2009 , published in PSIMageTools
function Add-ScaleFilter {
    <#
        .Synopsis
            Adds a Scale  Filter to a list of filters, or creates a new filter
        .Description
            Adds a Scale Filter to a list of filters, or creates a new filter
        .Example
            $Image = Get-Image .\Try.jpg
            $Image = $Image | Set-ImageFilter -filter (Add-ScaleFilter -Width 200 -Height 200 -PassThru) -PassThru
            $Image.SaveFile("$pwd\Try2.jpg")
        .Parameter image
            Optional.  If set, allows you to specify the crop in terms of a percentage
        .Parameter Width
            The new Width of the image in pixels (if Width is greater than one) or in percent (if Width is less than one and image is provided)
        .Parameter Height
            The new Height of the image in pixels (if Height is greater than one) or in percent (if Height is less than one and image is provided)
        .Parameter doNotPreserveAspectRatio
            If set, the aspect ratio will not be conserved when resizing
        .Parameter PassThru
            If set, the filter will be returned through the pipeline.  This should be set unless the filter is saved to a variable.
        .Parameter filter
            The filter chain that the rotate filter will be added to.  If no chain exists, then the filter will be created
    #>

    param(
    [Parameter(ValueFromPipeline=$true)]
    [__ComObject]
    $Filter,

    [__ComObject]
    $Image,

    [Double]$Width,
    [Double]$Height,

    [switch]$DoNotPreserveAspectRatio,

    [switch]$PassThru
    )

    process {
        if (-not $Filter) {
            $Filter = New-Object -ComObject Wia.ImageProcess
        }
        $index      = $Filter.Filters.Count + 1
        if (-not      $Filter.Apply) { return }
        $scale      = $Filter.FilterInfos.Item("Scale").FilterId
        $isPercent  = $true
        if ($Width  -gt 1) { $isPercent = $false }
        if ($Height -gt 1) { $isPercent = $false }
        $Filter.Filters.Add($scale)
        $Filter.Filters.Item($index).Properties.Item("PreserveAspectRatio") = "$(-not $DoNotPreserveAspectRatio)"
        if ($isPercent -and $Image) {
            $Filter.Filters.Item($index).Properties.Item("MaximumWidth")  = $Image.Width  * $Width
            $Filter.Filters.Item($index).Properties.Item("MaximumHeight") = $Image.Height * $Height
        } else {
            $Filter.Filters.Item($index).Properties.Item("MaximumWidth")  = $Width
            $Filter.Filters.Item($index).Properties.Item("MaximumHeight") = $Height
        }
        if ($PassThru) { return $Filter }
    }
}