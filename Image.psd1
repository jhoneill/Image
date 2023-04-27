
@{
    RequiredModules   = @('GetSQL')
    NestedModules     = "Add-ConversionFilter.ps1",
                        "Add-CropFilter.ps1",
                        "Add-ExifFilter.ps1",
                        "Add-OverlayFilter.ps1",
                        "Add-RotateFlipFilter.ps1",
                        "Add-ScaleFilter.ps1",
                        "ConvertTo-Jpeg.ps1",
                        "ConvertTo-Bitmap.ps1",
                        "Copy-Image.ps1",
                        "Get-Image.ps1",
                        "Helper.ps1",
                        "Read-EXIF.ps1",
                        "Get-IndexedItem.ps1",
                        "New-overlay.ps1",
                        "Save-image.ps1",
                        "Set-ImageFilter.ps1",
                        "Tagging-Support.ps1",
                        "Lightroom.ps1",
                        "ArgumentCompleter.ps1"
    FormatsToProcess  = "Lightroom.Format.ps1xml"

    GUID              = "{e42ab99b-f281-4412-93cd-5f43427d3c8a}"

    Author            = "James Brundage, James O'Neill"

    CompanyName       = "Microsoft Corporation"

    Copyright         = "© Microsoft Corporation 2009. All rights reserved."

    ModuleVersion     = "2.2.0.1"

    PowerShellVersion = "3.0"

    CLRVersion        = "2.0"

    VariablesToExport = "ExifTagNames"

    FunctionsToExport = 'Add-ConversionFilter',
    'Add-CropFilter',
    'Add-EXIFFilter',
    'Add-OverlayFilter',
    'Add-RotateFlipFilter',
    'Add-ScaleFilter',
    'Add-LightRoomCollectionItem',
    'Add-LightRoomItemKeyword',
    'Add-ExifKeyword',
    'Close-LightRoom',
    'Convert-GPStoEXIFFilter',
    'Convert-SuuntoToExifFilter',
    'ConvertFrom-KML',
    'ConvertTo-Bitmap',
    'ConvertTo-DateTime',
    'ConvertTo-GPX',
    'ConvertTo-Jpeg',
    'Copy-GPSImage',
    'Copy-Image',
    'Copy-SuutoDBTOImage',
    'Copy-SuutoDBTOLightRoom',
    'Get-BingPhotos',
    'Get-CSVGPSData',
    'Get-GPSBearing',
    'Get-GPSDistance',
    'Get-GPXData',
    'Get-Image',
    'Get-IndexedItem',
    'Get-LightRoomItem' ,
    'Get-LightRoomCollectionItem',
    'Get-LightRoomCollection',
    'Get-LightRoomKeyword',
    'Get-LightRoomKeywordItem',
    'Get-NearestPoint',
    'Get-NearestSuutoDBPoint',
    'Get-NMEAData',
    'Invoke-FilePrint',
    'Merge-GPSPoints',
    'Move-NonLightRoom',
    'New-Imagefilter',
    'New-Overlay',
    'Out-MapPoint',
    'Read-Exif',
    'Rename-LightRoomItem',
    'Resolve-ImagePlace',
    'Save-Image',
    'Select-SeperateGPSPoints',
    'Set-DefaultPrinter',
    'Set-ImageFilter',
    'Set-LightRoomItemColor',
    'Set-LightRoomItemFlag',
    'Set-LightRoomItemRating',
    'Set-LogOffset',
    'Test-LightRoomItem'
}