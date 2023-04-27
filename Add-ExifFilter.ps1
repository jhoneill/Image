

function Add-EXIFFilter {
<#
        .Synopsis
            Adds an Exif Filter to a list of filters, or creates a new filter
        .Description
            Adds an Exif Filter to a list of filters, or creates a new filter
        .Example
            Add-exifFilter -passThru -ExifID $ExifIDKeywords -typeid 1101 -string "Ocean,Bahamas"    |
            Adds a filter to set the keywords to Ocean; Bahamas, using the numeric type ID
            and getting the function to convert the string to the vector type required
         .Example
            Add-exifFilter -passThru -ExifID $ExifIDTitle -typeName "vectorofbyte" -string "fish"
            Add a filter to set the Title to "fish", using the name of the type
            and getting the function to convert the string to the vector type required
         .Example
            Add-exifFilter -passThru -ExifID $ExifidCopyright -typeName "String" -value "© James O'Neill 2009"
            Sets the copyright field (note this is a normal string, not a vector of bytes containing a unicode string)
         .Example
            Add-exifFilter -passThru -ExifID $ExifIDGPSAltitude -typeName "uRational" -Numerator 123 -denominator 10
            Add a filter to set the GPS Altitude to 12.3M
            getting the function to create the unsigned rational required
        .Parameter ExifID
            The ID of the field to be added / updated
        .Parameter TypeID
            The code representing the data type for this field (String, byte, integer, ratio, vector etc)
        .Parameter Value
            The new value for the field
        .Parameter TypeName
            Reserved Will allow the type to specified as a name rather than a numeric code
        .Parameter Numerator
            Reserved will ratios to be passed as numerator / denominator
        .Parameter Denominator
            Reserved will ratios to be passed as numerator / denominator
        .Parameter String
            Reserved will allow the value for Vectors which hold strings to be passed as a string
        .Parameter passthru
            If set, the filter will be returned through the pipeline.  This should be set unless the filter is saved to a variable.
        .Parameter filter
            The filter chain that the rotate filter will be added to.  If no chain exists, then the filter will be created
    #>

    param (
        [Parameter(ValueFromPipeline=$true)]
        [__ComObject]
        $Filter,

        [Parameter(Mandatory=$true)]$ExifTag,
        [alias("TypeID")]
        [ExifType]$TypeName ,
        $Value , $String , $Numerator, $Denominator ,
        [switch]$PassThru
    )

    process {
        if       ($ExifTag -is [int])     {$ExifID = $ExifTag}
        else                              {$ExifID = $ExifTagValues[$ExifTag] }
        if (-not ($ExifID -ge 0))         {Write-Warning -Message "$ExifTag  does not appear to be a valid tag" ; return }
        if (-not  $Filter)                {$Filter = New-Object -ComObject WIA.ImageProcess }
        if (-not  $Filter.Apply)          {Write-Warning -Message "Could not get the filter object";              return }
        if (      $TypeName )             {$TypeID = $TypeName.value__}
        elseif(   $ExifTypeHash[$ExifID]) {$TypeID =  $ExifTypeHash[$ExifID]}
        if (-not  $TypeID)                {Write-Warning -Message "Could not find a the Exif data type for your tag"; return }

        if ((@(1006,1007) -contains $TypeID) -and (-not $Value) -and ($null -ne $Numerator ) -and $Denominator) {
            $Value =New-Object -ComObject WIA.Rational
            $Value.Denominator = $Denominator
            $Value.Numerator   = $Numerator
        }
        if ((@(1100,1101) -contains $TypeID) -and (-not $Value) -and $String) {$Value = New-Object -ComObject "WIA.Vector"
                                                                               $Value.SetFromString($String)
        }
        if ((1002 -eq $TypeID) -and (-not $Value) -and $String) { $Value    =  $String }

        $Filter.Filters.Add( $Filter.FilterInfos.Item("Exif").FilterId)
        $Filter.Filters.Item($Filter.Filters.Count).Properties.Item("ID")   = "$ExifID"
        $Filter.Filters.Item($Filter.Filters.Count).Properties.Item("Type") = "$TypeID"
        $Filter.Filters.Item($filter.Filters.Count).Properties.Item("Value")=  $Value

        if ($passthru) { return $filter }
    }
}
