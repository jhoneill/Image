#requires -version 2.0
#Original code James Brundage / Microsoft 2009 , published in PSIMageTools
function Add-OverlayFilter {
    <#
        .Synopsis
            Adds an Overlay Filter to a list of filters, or creates a new filter
        .Description
            Adds an Overlay Filter to a list of filters, or creates a new filter
        .Example
            $image = Get-Image .\Try.jpg
            $otherImage = Get-Image .\OtherImage.jpg
            $image = $image | Set-ImageFilter -filter (Add-OverLayFilter -Image $otherImage -Left 10 -Top 10 -passThru) -passThru
            $image.SaveFile("$pwd\Try2.jpg")
        .Parameter image
            Optional.  If set, allows you to specify the crop in terms of a percentage
        .Parameter left
            The horizontal location within the image where the overlay should be added
        .Parameter top
            The vertical location within the image where the overlay should be added
        .Parameter passthru
            If set, the filter will be returned through the pipeline.  This should be set unless the filter is saved to a variable.
        .Parameter filter
            The filter chain that the rotate filter will be added to.  If no chain exists, then the filter will be created
    #>

    param(
        [Parameter(ValueFromPipeline=$true)]
        [__ComObject]$Filter,
        [__ComObject]$Image,
        [Double]$Left,
        [Double]$Top,
        [switch]$PassThru
    )

    process {
        if (-not $Filter) {
            $Filter = New-Object -ComObject Wia.ImageProcess
        }
        $index = $Filter.Filters.Count + 1
        if (-not $Filter.Apply) { return }
        $stamp = $Filter.FilterInfos.Item("Stamp").FilterId
        $Filter.Filters.Add($stamp)
        $Filter.Filters.Item($index).Properties.Item("ImageFile") = $Image.PSObject.BaseObject
        $Filter.Filters.Item($index).Properties.Item("Left")      = $Left
        $Filter.Filters.Item($index).Properties.Item("Top")       = $Top
        if ($PassThru) { return $Filter }
    }
}