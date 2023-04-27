#Original code James Brundage / Microsoft 2009 , published in PSIMageTools
function ConvertTo-Jpeg {
    <#
      .Synopsis
        Converts a file to a JPG of the specified quality in the same folder
      .Description
        Converts a file to a JPG of the specified quality in the same folder.
        If the file is already a JPG it will be overwritten at the new quality setting
      .Example
        C:\PS>  Dir -recure -include *.tif | Convert-toJPeg .\myImage.bmp
        Creates creates JPG images of quality 100 for all tif files in the current directory and it's sub directories
      .Example
        C:\PS>  Dir -recure -include *.tif | Convert-toJPeg -quality 75
        Creates JPG images of quality 75 for all tif files in the current directory and it's sub directories
      .Parameter Image
        An image object, a path to an image, or a file object representing an image file. It may be passed via the pipeline.
      .Parameter Quality
        Range 1-100, sets image quality (100 highest), lower quality will use higher rates of compression.
        The default is 100.
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
        $Image,

        [ValidateRange(1,100)]
        [int]$Quality = 100
    )
    process {
        if ((     $Image -is [String]) -or ($Image -is [System.io.FileInfo])) {Get-Image $image | ConvertTo-Jpeg -Quality $Quality ; return}
        if (      $Image.count -gt 1)                                         {          $Image | ConvertTo-Jpeg -Quality $Quality ; return}
        if ( -not $Image.Loadfile  -and  -not   $Image.Fullname)              { return }
        Write-Verbose   -Message ("Processing $($Image.fullName)")
        $noExtension    = $Image.Fullname -replace "\.\w*$",""   # "\.\w*$" means dot followed by any number of alpha chars, followed by end of string - i.e file extension
        $process        = New-Object -ComObject Wia.ImageProcess
        $convertFilter  = $process.FilterInfos.Item("Convert").FilterId
        $process.Filters.Add($convertFilter)
        $process.Filters.Item(1).Properties.Item("Quality")  = $Quality
        $process.Filters.Item(1).Properties.Item("FormatID") = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
        $newImg         = $process.Apply($image.PSObject.BaseObject)
        $newImg.SaveFile("$noExtension.jpg")
    }
}