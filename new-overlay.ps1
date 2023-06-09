[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | out-null

function New-Overlay {
<#
    .Synopsis
        Creates a new transparent JPG containing text to use as an overlay
    .Description
        Creates a new transparent JPG containing text to use as an overlay
    .Example
        C:\PS>$Overlay = New-overlay -text "© James O'Neill 2009" -size 32 -TypeFace "Arial"  -color red -filename "$Pwd\overLay.jpg"
        C:\PS>$filter = Add-OverlayFilter $overlay -passthru
            
            
        Creates an overlay adds it to the filter list
    .Parameter Text
        The text you want to set in the overlay
    .Parameter Size
        The font size (default 32)
    .Parameter TypeFace
        The font face (Default Arial)
    .Parameter Color
        The font Color (Default grey; type [system.drawing.color]:: and then tab through to see the more obscure color names available)
    .Parameter Filename
        The name for the overlay file (default "Overlay.jpg" in the current folder)
#>
    param ([string]$text="All rights reserved by the copyright owner", 
              [int]$size=32, 
           [string]$TypeFace="Arial", 
           [system.drawing.color]$color=[system.drawing.color]::Gray, 
           [string]$filename="$Pwd\overLay.jpg" 
    )
    $width     = $text.length * $size   # This is a close enough approximation - 
    $height    = $size * 1.5
    $Thumbnail = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $width, $height, [System.Drawing.Imaging.PixelFormat]::Format16bppRgb565
    $graphic   =[system.drawing.graphics]::FromImage($thumbnail)                                                                         
    if ($color -eq [system.drawing.color]::Black) {$b=New-Object -TypeName system.drawing.solidbrush -ArgumentList [system.drawing.color]::white
                                                    $graphic.FillRectangle($b,0,0,$width,$height)                                           
    } 
    $font     = New-Object -TypeName system.drawing.font       -ArgumentList $typeFace,$size    $brush    = New-Object -TypeName system.drawing.solidbrush -ArgumentList $color
    $graphic.DrawString($text,$font ,$brush, 0 , 0 )         
    if ($color -eq [system.drawing.color]::Black) {$Thumbnail.MakeTransparent([system.drawing.color]::White)}
    else                                          {$Thumbnail.MakeTransparent([system.drawing.color]::Black)}                                                                       
    $Thumbnail.Save($filename)
    Get-Image -Path  $filename
}