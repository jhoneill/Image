PS> $Overlay = New-Overlay -text "© James O'Neill 2009" -size 32 -TypeFace "Arial"  -color "red" -filename "$Pwd\overLay.jpg"

PS> $filter = New-Imagefilter |
              Add-EXIFFilter       -passThru -Exiftag  Keywords    -typeid   1101           -string    "Ocean,Bahamas"        |
              Add-EXIFFilter       -passThru -Exiftag  Title       -typeName "vectorofbyte" -string    "fish"                 |
              Add-EXIFFilter       -passThru -Exiftag  Copyright   -typeName "String"       -value     "© James O'Neill 2009" |
              Add-EXIFFilter       -passThru -Exiftag  GPSAltitude -typeName "uRational"    -Numerator 123 -denominator 10    |
              Add-EXIFFilter       -passThru -Exiftag  GPSAltRef   -typeName "Byte"         -value     1                      |
              Add-ScaleFilter      -passThru -height   800         -width    65535    |
              Add-OverlayFilter    -passThru -top      0           -left     0              -image     $Overlay |
              Add-ConversionFilter -passThru -typeName jpg         -quality  70

PS> Get-Image C:\Users\Jamesone\Pictures\IMG_3333.JPG  | Set-ImageFilter -passThru -filter $filter | Save-image -fileName {$_.FullName -replace ".jpg$","-small2.jpg"}

function Copy-Exif{
    Param  ( [Parameter(Mandatory=$true  ,valueFromPipeline=$true )] $DestinationPath,
             [Parameter(Mandatory=$true )] $SourcePath   ,
             [switch]$fixICE)
    process {
        $destinationPath | foreach {               & 'C:\Program Files (x86)\EXIFutils\exifcopy.exe' "/bq" "/o" $SourcePath $_
                                                   & 'C:\Program Files (x86)\EXIFutils\exifedit'    "/b"   "/r" print-im    $_
                                    if ($FixICE) { & 'C:\Program Files (x86)\EXIFutils\exifedit'    "/b"   "/a" "Firm-ver=Microsoft ICE v1.4.4.0" "/r" "orient" "/s" "/t" "a,160" $DestinationPath }
                                    }
    }
}

function Fix-ICeExif{
    Param  ( [Parameter(Mandatory=$true  ,valueFromPipeline=$true )] $DestinationPath)
& 'C:\Program Files (x86)\EXIFutils\exifedit'     "/b"   "/a" "Firm-ver=Microsoft ICE v1.4.4.0" "/r" "orient" "/s"  "/t" "a,160" $DestinationPath
}



$destpath           = [system.environment]::GetFolderPath( [system.environment+specialFolder]::MyPictures )
$keywords           = "Exploring"
$sourcePath         = "E:\DCIM\100PENTX"
$logpath            = "F:\My Documents\My GPS\Track Log\20100404143612.log"
$refDate            = ([datetime]"04/04/2010 16:02:17 +1")
$ReferenceImagePath = "C:\Users\Jamesone\Pictures\Greenham\GC-43272.JPG"
$offset             = $null
$filesToTag         = $null
$NMEAData           = $False
$CSVData            = $True

if ($refDate -and $ReferenceImagePath  ) {
        if  (-not (test-path $ReferenceImagePath) )  {Write-host "Could not find $ReferenceImagePath, exiting "; return }
         $offset = ((Read-Exif $ReferenceImagePath ).DateTaken  - $refDate.ToUniversalTime() ).totalseconds

}
if ($offset -eq $null) {   # Not passed as a param or calculated
    switch  (Select-Item -Caption "Log to camera time offset " -TextChoices "Camera and log are &Sync'd", "&Calculate from a picture of the logger", "&Enter offset manually")
       {   0 { $offset = 0 }
           1 { if (-not $refDate) {$RefDate = ([datetime]( Read-Host ("Please enter the Date & time, in the reference picture, formatted as" + [char]13 + [Char]10 +
                                                                      "Either MM/DD/yyyy HH:MM:SS ±Z or  dd MMMM yyyy HH:mm:ss ±Z")
                                               )).touniversalTime()}
                if (-not $ReferenceImagePath) {$ReferenceImagePath  = Read-Host "Please enter the path to the picture"}
                if ($ReferenceImagePath -and (test-path $ReferenceImagePath) -and $RefDate) {
                   $offset =  ((Read-Exif $ReferenceImagePath ).DateTaken - $refdate).totalSeconds
                }
                else  {Write-host "A valid reference image path and time are needed"; return }
             }
           2 {[long]$offset  = Read-Host ("Please enter the offset in seconds:" + [char]13 + [char]10 +
                                          "negative for logger before camera time, " + [char]13 + [char]10 +
                                          "positive for camera time before logger.")  }
       }
}

if (-not $logpath) {$LogPath = Read-Host "Please enter the path to the log file"}
if ($logpath -and (test-path $LogPath)) {
############ Suunto
   if     ($NMEAData )  { $Points =  Get-NMEAGPSData   -Path $Logpath -offset $offset  }
   elseif ($CsvData  )  { $Points =  Get-EfficaGPSData -Path $Logpath -offset $offset  }
   else   {switch  (Select-Item -Message "What kind of data do you want to add " -TextChoices "GPS data from &Efficasoft", "GPS data in &NMEA format")
                {   0 { $Points =  Get-EfficaGPSData -Path $Logpath -offset $offset  }
                    1 { $Points =  Get-NMEAGPSData   -Path $Logpath -offset $offset  }
                }
   }
}
else {Write-host "Could not find logpath: $logpath exiting "; return }

if ($ReferenceImagePath -and (Test-Path $ReferenceImagePath)) {write-host "Check the following data against the reference picture"
                    get-nearestPoint -DataPoints $points -ColumnName "DateTime" -MatchingTime $pictime | ft * -a | Out-Host
                    $picPath = Split-Path -Path $ReferenceImagePath
}

If (-not $SourcPath) {
    $SourcePath = read-host "Please enter the path to the the pictures - $PicPath"
    If (-not $SourcePath) {$SourcePath = $picPath}
}

If (-not $destpath) {
    $picPath  = [system.environment]::GetFolderPath( [system.environment+specialFolder]::MyPictures )
    $DestPath = read-host "Where would you like the tagged Pictures to be saved - $PicPath"
    If (-not $Destpath) {$DestPath = $picPath}
}

Get-ChildItem -Path (Join-path -Path $SourcePath -ChildPath "*.JPG") | Select-List -Property Name -multiple | Copy-SuutoImage -points $points -DestPath $DestPath -keywords "Ocean;Bahamas"


$VerbosePreference="Continue"
$Points =  Get-EfficaGPSData  -offset 3607 -Path 'f:\My Documents\My GPS\Track Log\20100425115503.log'
dir E:\DCIM\100PENTX\*.jpg | select-list -pro name -mul | %{ Get-Image $_ | Copy-GPSImage -points $Points -DestPath 'C:\users\Jamesone\Pictures\New folder' -keywords "foo" subject "bar"}



$MPApp = New-Object -ComObject "Mappoint.Application"
$MPApp.Visible = $true


$map = $mpapp.ActiveMap
Merge-GPSPoints -points $points | foreach-object { $location=$map.GetLocation($_.AveLat, $_.AveLong)
                           $Pin = $map.AddPushpin($location, ("{0} - {1:00} MPH " -f $_.Minute,$_.Aveknots*1.1508))
                           if     ($_.aveknots -lt 35) {$pin.symbol= 5 }
                           elseif ($_.aveknots -gt 47) {$pin.symbol= 7 }
                           else                      {$pin.symbol= 6 }
                           }


