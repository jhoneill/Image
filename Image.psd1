@{
    ModuleVersion     =   '2.3.0.0'
    GUID              =   '{e42ab99b-f281-4412-93cd-5f43427d3c8a}'
    Author            =   'James Brundage, James O''Neill'
    Copyright         =   'Portions © Microsoft Corporation 2009. All rights reserved. Portions © James O''Neill 2011-2023'

    PowerShellVersion =   '5.0'
    RequiredModules   = @('GetSQL')
    ScriptsToProcess  = @('Lightroom.ps1')
    FormatsToProcess  = @('Lightroom.Format.ps1xml')
    RootModule        =   'Image.psm1'

    FunctionsToExport = @('Add-ConversionFilter',
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
    )
    VariablesToExport = @('ExifTagNames')

    FileList          = @('Add-ConversionFilter.ps1',
                          'Add-CropFilter.ps1',
                          'Add-ExifFilter.ps1',
                          'Add-OverlayFilter.ps1',
                          'Add-RotateFlipFilter.ps1',
                          'Add-ScaleFilter.ps1',
                          'ArgumentCompleter.ps1',
                          'ConvertTo-Jpeg.ps1',
                          'ConvertTo-Bitmap.ps1',
                          'Copy-Image.ps1',
                          'Get-Image.ps1',
                          'Get-IndexedItem.ps1',
                          'Helper.ps1',
                          'Lightroom.Format.ps1xml',
                          'Lightroom.ps1',
                          'New-overlay.ps1',
                          'Read-EXIF.ps1',
                          'Save-image.ps1',
                          'Set-ImageFilter.ps1',
                          'Tagging-Support.ps1'
    )
}
