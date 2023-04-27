function New-Imagefilter {
    <#
        .Synopsis
            Takes no parameters and returns a new Wia.ImageProcess (filter) object
        .Description
            Takes no parameters and returns a new Wia.ImageProcess (filter) object
        .Example
            $filter = New-ImageFilter
            Creates a new Wia.ImageProcess COM object and assigns it to $Filter
    #>
    New-Object -ComObject Wia.ImageProcess
}

function Add-ConversionFilter {
    <#
        .Synopsis
            Adds a Conversion  Filter to a list of filters, or creates a new filter
        .Description
            Adds a Conversion Filter to a list of filters, or creates a new filter.
        .Example
            $filter = add-ConversionFilter -typename jpg -quality 70 -passthru
            Creates a new filter to convert to JPG format with a quality of 70, and assigns it to $filter.
        .Parameter typeName
            The new file type: may be "BMP","GIF", "JPEG","JPG"(default),"PNG","TIF","TIFF"
        .Parameter Quality
            For JPEG images determines the quality from 1 to 100.
            Higher numbers result in bigger files. Defaults to 100
        .Parameter passthru
            If set, the filter will be returned through the pipeline.
            This should be set if filter is not passed as parameter.
            If it is created or passed throught the pipeline it will not be returned otherwise.
        .Parameter filter
            The filter chain that the filter will be added to.
            If no chain exists, then the filter will be created.
            The filter may be passed using the pipeline
    #>
    param(
        [Parameter(ValueFromPipeline=$true,ParameterSetName)]
        [AllowNull()]
        [__ComObject]$Filter = $null,

        [ValidateSet("BMP","GIF", "JPEG","JPG","PNG","TIF","TIFF")]
        [string]$TypeName="JPG",
        [ValidateRange(1,100)]
        [int]$Quality = 100,
        [switch]$PassThru
    )

    process {
        $typeID=@{"BMP"="{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}" ; "PNG"  = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
                "GIF"="{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}" ; "JPEG" = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
                "JPG"="{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}" ; "TIFF" = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
                "TIF"="{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}" }["$TypeName"]
        if  (-not $Filter)                           {$Filter = New-Imagefilter }
        if ((-not $Filter.Apply) -or (-not $typeID)) { return }
        $Filter.Filters.Add($Filter.FilterInfos.Item("Convert").FilterId)
        $Filter.filters.Item($Filter.Filters.Count).properties.item("FormatID")=$TypeID
        if ($TypeName -like "JP*")  { $Filter.filters.Item($filter.Filters.Count).properties.item("Quality") = $Quality }
        if ($PassThru)              { return $Filter }
    }
}