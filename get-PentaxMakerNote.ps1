function Get-PentaxMakerNoteProperty {
    <#
      .Synopsis
        Returns an single item of Pentax specific maker note information
      .Description
        Returns an single item of Pentax specific maker note information
      .Example
        C:\PS> Get-PentaxMakerNoteProperty -image $image -exifID 5
        Returns the Camera model ID code.
      .Parameter image
        The image from which the data will be read $
      .Parameter ExifID
        The ID of the required data field
    #>
    param ([int]$exifID,   [__ComObject] $image)
    $MakerID ="AOC"
    try   { [byte[]]$wholeMakerNote= $image.Properties.Item("37500").value.binaryData }
    Catch { Write-Warning -Message "Error getting maker note - proably doesn't exist" ; return   }
    0..$MakerId.length | ForEach-Object {if ($wholeMakerNote[$_] -ne [byte]($makerid[$_]) ) {Write-Debug -Message "Wrong maker ID"; Return}}
    $PrevFieldCode = 0
    for ($i = 8 ; $i -lt $WholeMakerNote.count -1 ; $i += 12 ) {
        $FieldCode = 256 * $WholeMakerNote[$i] + $WholeMakerNote[$i + 1]
        If (($FieldCode -lt $PrevFieldCode) -or ($fieldcode -gt $exifID)) {Write-Debug -Message "Field code not found"; return} else {$PrevFieldCode = $FieldCode}
        #If (($WholeMakerNote[$i + 4] -ne 0) -Or ($WholeMakerNote[$i + 5] -ne 0)) {return } # don't recall why
        If ($FieldCode -eq $ExifID) {#& write-host $wholemakernote[$i+3]
            if  ($WholeMakerNote[$i + 3] -eq 3) {$S = [string](256 * $WholeMakerNote[$i + 8] + $WholeMakerNote[$i + 9])
                        If ($WholeMakerNote[$i + 7] -eq 2) {$s = $s + " , " + [String](256 * $WholeMakerNote[$i + 10] + $WholeMakerNote[$i + 11]) }
                        Return $s}
            if (($WholeMakerNote[$i + 3] -eq 4) -and ($WholeMakerNote[$i + 7] -eq 1)) { Return [String](16777216 * $WholeMakerNote[$i + 8] + 65536 * $WholeMakerNote[$i + 9] + 256 * $WholeMakerNote[$i + 10] + $WholeMakerNote[$i + 11])}
            if (( @(1, 6, 7) -contains $WholeMakerNote[$i + 3]) -and ($WholeMakerNote[$i + 7] -le 4))  {($i +8)..($i +7 + $WholeMakerNote[$i + 7] ) | ForEach-Object -Begin {$s=""} -process {$s = $s + [string]$WholeMakerNote[$_]  + " "  } -end {return $s} }
        }
    }
}

function Get-PentaxExif {
    <#
        .Synopsis
            Returns an object containing Pentax specific maker note information
        .Description
            Returns an object containing Pentax specific maker note information
        .Example
            Get-pentaxExif $image
        .Parameter image
            The image from which the data will be read
        .Parameter full
            Pentax data only or all adata

    #>
    param   ([Parameter(ValueFromPipeline=$true,Mandatory=$true)]$image , [Switch]$full)
    process {
        if ($image -is [system.io.fileinfo] ) {$image = $image.FullName }
        if ($image -is [String]             ) {$image = Get-Image $image}
        if ($image.count -gt 1              ) {$image | ForEach-Object {Get-PentaxExif $_ -Full:$full} ; Return}

        $s = Get-PentaxMakerNoteProperty -image $image -exifID 5
        if ($null -eq $s) {
            Write-Warning -Message "No camera ID was found for $($image.fullname); giving up"; return
        }
        else {     $PentaxModel = @{"13" = "Optio 430";"76180"="*ist-D"; "76830"="K10D" ; "77240" = "K7" ; "77430" = "K5"      }[$s]
        if   (-not $PentaxModel) {
                $PentaxModel = "Unknown: $s"
        }}
        Write-Progress -Activity "Getting Pentax Exif Data for" -Status $image.fullname
        $s                      = (Get-PentaxMakerNoteProperty -image $image -exifID 93).Split(" ")
        $d                      = (Get-PentaxMakerNoteProperty -image $image -exifID 6).Split(" ")
        $t                      = (Get-PentaxMakerNoteProperty -image $image -exifID 7).Split(" ")
        [int64]$dno             =               (16777216 * $d[0]) + (65536 * $d[1]) + (256 * $d[2]) +($d[3])
        [int64]$tno             =               (16777216 * $t[0]) + (65536 * $t[1]) + (256 * $t[2]) +($t[3])
        [int64]$sno             = 4294967296 -  (16777216 * $s[0]) - (65536 * $s[1]) - (256 * $s[2]) -($s[3]) - 1
        $PentaxShutterCount     =  $Sno -bXor $dNo -bXor $tNo

        $s= (Get-PentaxMakerNoteProperty -image $image -exifID 92).Split(" ")
        If ([int]$s[1] –band 1)  {    $t =      "SR Enabled, "
            If ([int]$s[3] -band 1)  {$t = $t + " Focal Length: $([int]$S[3]*4)"}
            else                     {$t = $t + " Focal Length: $([int]$S[3]/2)mm"}
            If ([int]$s[0] –band 1)  {$t = $t + " Stabilized." }
            else                     {$t = $t + " Not Stabilized."}
            If ([int]$s[0] –band 64) {$t = $t + " Not Ready."  }
        }
        else                         {$t =     "SR Disabled."    }
        $PentaxShakeReduction     =   $t

        $s = (Get-PentaxMakerNoteProperty -image $image -exifID 51).Split(" ")
        $t = @{0="Program"; 1="Hi-Speed Program"; 2="DOF Program"; 3="MTF Program"; 4="Standard"; 5="Portrait"; 6="Landscape";
            7="Macro"; 8="Sport"; 9="Night Scene"; 10="No Flash"; 11="Soft"; 12="Surf & Snow"; 13="Text"; 14="Sunset";
            15="Kids"; 16="Pet"; 17="Candlelight"; 18="Museum"; 19="Food"; 20="Stage Lighting"; 21="Night Snap";
            30="Self-Portrait"; 31="Illustrations"; 33="Digital Filter"; 37="Museum"; 38="Food"; 40="Green Mode";
            49="Light Pet"; 50="Dark Pet"; 51="Medium Pet"; 53="Underwater"; 54="Candlelight"; 55="Natural Skin Tone";
            56="Synchro Sound Record"; 58="Frame Composite"; 60="Kids"; 61="Blur Reduction"; 255="Digital Filter";
            260="Auto PICT (Standard)"; 261="Auto PICT (Portrait)"; 262="Auto PICT (Landscape)"; 263="Auto PICT (Macro)";
            264="Auto PICT (Sport)"; 512="Program (HyP)"; 513="Hi-speed Program (HyP)"; 514="DOF Program (HyP)";
            515="MTF Program (HyP)"; 534="Shallow DOF (HyP)"; 768="Green Mode"; 1024="Shutter Speed Priority";
            1280="Aperture Priority"; 1536="Program Tv Shift"; 1792="Program Av Shift"; 2048="Manual";
            2304="Bulb"; 2560="Aperture Priority, Off-Auto-Aperture"; 2816 ="Manual, Off-Auto-Aperture";
            3072="Bulb; Off-Auto-Aperture"; 3328 ="Shutter & Aperture Priority AE"; 3840 ="Sensitivity Priority AE";
            4096="Flash X-Sync Speed AE"; 4608="Auto Program (Normal)"; 4609="Auto Program (Hi-speed)";
            4610="Auto Program (DOF)"; 4611="Auto Program (MTF)";
            4630 ="Auto Program (Shallow DOF)"}[ (([int]$s[0] * 256) +[int]$s[1]) ]
        If ($S[2] -eq "0") {
            $PentaxPictureMode = $t + ": 1/2 EV Steps"
        }
        else {
            $PentaxPictureMode = $t + ": 1/3 EV Steps"
        }
        $PentaxBracketing = [boolean][int]((Get-PentaxMakerNoteProperty -image $image -exifID 0x18) -split " , ")[0]
        $s= (Get-PentaxMakerNoteProperty -image $image -exifID 52).Split(" ")
        if ($PentaxBracketing) {
            $t = "Bracket Sequence "
        }
        Elseif ([boolean][int]$s[3]) {
            $t = @{"1"="Multiple Exposure. ";  "16"="HDR. "; "32"="HDR Strong 1. ";  "48"="HDR Strong 2. ";
                    "64"="HDR Strong 3. "     ; "224"="HDR Auto. "}[($S[3])]
        }
        Else {
            $t =             @{"0"="Single-frame. "; "1"="Continuous. "; "2"="Continuous [Hi]. "; "3"="Burst. "}[($S[0])]
        }
        $t =              $t + @{"0"=""; "1"="Self-timer, 12 sec. ";         "2"="Self-timer, 2 sec. "; "16"="Mirror Lock-up. "  }[($S[1])]
        $PentaxDriveMode =$t + @{"0"=""; "1"="Remote control, 3 sec delay. "; "2"="Remote control."    ;  "4"="Remote Continuous Shooting. "}[($S[2])]

        $s               = (Get-PentaxMakerNoteProperty -image $image -exifID 0x71).Split(" ")
        $t               =     @{"0"="Inactive - "; "1"="Active "; "2"="Active (Weak) ";
                                "3"="Active (Strong) ";"4"="Active (Medium) "}[($S[1])]
        $PentaxHighIsoNR = $t+ @{"0"="Off "; "1"="Weakest"; "2"="Weak"; "3"="Strong"; "4"="Medium"; "255"="Auto"; }[($S[0])]
        if ($s[1] -eq 0 -and $S[0] -ne 255) {
                    $PentaxHighIsoNR += " will be enabled if " + @{"48"="ISO>400"; "56"="ISO>800"; "64"="ISO>1600"; "72"="ISO>3200"}[($S[2])]
        }
        $s               = (Get-PentaxMakerNoteProperty -image $image -exifID 0x32).split(" ")
        if ($s[-1] -ne 0) {
                        $PentaxEdit = "Digital Filter"
        }
        Else {           $PentaxEdit =  @{"0" ="None" ; "2"="Cropped";"4"="Parameter Adjust";"6"="Digital Filter";;"16"="Frame Synthethis"}[($s[0])]}
                $PentaxTemperature = (Get-PentaxMakerNoteProperty -image $image -exifID 71)
                $PentaxEffectiveLV = (Get-PentaxMakerNoteProperty -image $image -exifID 0x2d)/1024
            $PentaxProcessingCount = (Get-PentaxMakerNoteProperty -image $image -exifID 0x41)
                    $PentaxAELock = [boolean][int](Get-PentaxMakerNoteProperty -image $image -exifID 0x48)
            $PentaxNoiseReduction = [boolean][int](Get-PentaxMakerNoteProperty -image $image -exifID 0x49)

        $HighLowValues               = @{"0"="Low";      "1"="Normal";    "2"="High"; "3"="Med-Low";"4" = "Med-High";
                                        "5"="Very Low"; "6"="Very High"; "7" = "-4"; "8"="+4"}

                $PentaxSaturation   = $HighLowValues[ (Get-PentaxMakerNoteProperty -image $image -exifID 31) ]
        if (-not $PentaxSaturation)  {
                $PentaxSaturation   = "Unknown"
        }
                $PentaxContrast     = $HighLowValues[(Get-PentaxMakerNoteProperty -image $image -exifID 32)]
        if (-not $PentaxContrast )   {
                $PentaxContrast     = "Unknown"
        }
                $PentaxSharpening   = $HighLowValues[ (Get-PentaxMakerNoteProperty -image $image -exifID 33) ]
        If (-not $PentaxSharpening)  {
                $PentaxSharpening   = "Unknown"
        }
                $PentaxImageTone    = @{"0"="Natural"; "1"="Bright";     "2"="Portrait"; "3"="Landscape";
                                        "4"="Vibrant"; "5"="Monochrome"; "6"="Muted";    "7"="Reversal Film";
                                        "8"="Bleach Bypass"}[ (Get-PentaxMakerNoteProperty -image $image -exifID 79) ]
        if (-not $PentaxImageTone)   {
                $PentaxImageTone    = "Unknown"
        }
            $PentaxWhiteBalance    = @{"0"="Auto";    "1"="Daylight"; "2"="Shade"; "3"="Fluorescent"; "4"="Tungsten";
                                    "5"="Manual";  "6"="Daylight Fluorescent" ; "7"="Day White Fluorescent";
                                    "8"="White Fluorescent";  "9"="Flash"; "10"="Cloudy";
                                    "15"="Color Temperature Enhancement"; "17"="Kelvin"; "65534"="Unknown";
                                "65535"="User-Selected"}[(Get-PentaxMakerNoteProperty -image $image -exifID 25)]
        if (-not $PentaxWhiteBalance){
                $PentaxWhiteBalance = "Unknown"
        }
                $PentaxFocusMode    = @{"0"="Normal";    "1"="Macro"; "2"="Infinity"; "3"="Manual"; "4"="SuperMacro"
                                    "5"="Pan-Focus"; "16"="AF-S"; "17"="AF-C"   ; "18"="AF-A" ;
                                    "32"="Contrast Detect" }[(Get-PentaxMakerNoteProperty -image $image -exifID 13) ]
        if (-not $PentaxFocusMode)   {
                $PentaxFocusMode    = "Unknown"
        }
                $PentaxFocusPoint   = @{"1"="Top Left";     "2"="Top Center";    "3"="Top Right";    "4"="Middle Far-Left";
                                        "5"="Middle Left";  "6"="Middle Center"; "7"="Middle Right"; "8"="Middle Far-Right" ;
                                        "9"="Bottom Left"; "10"="Bottom Center"; "11"="Bottom Right";
                                    "65532"="Face Detection AF" ; "65533"="Auto Tracking AF"; "65534"="Fixed Center";
                                    "65535"="Auto" }[(Get-PentaxMakerNoteProperty -image $image -exifID 14)]
        if (-not $PentaxFocusPoint)  {
                $PentaxFocusPoint   = "Unknown"
        }
                $PentaxMetermode    = @{"0"="Multi-Segment"; "1"="Center-Weighted"; "2"="Spot"}[ (Get-PentaxMakerNoteProperty -image $image -exifID 23) ]
        if (-not $PentaxMetermode)   {
                $PentaxMetermode    = "Unknown"
        }
                $PentaxQuality      = @{"0"="Good"; "1"="Better"; "2"="Best"; "3"="TIFF"; "4"="RAW"; "5"="Premium" }[(Get-PentaxMakerNoteProperty -image $image -exifID 8)]
        if (-not $PentaxQuality )    {
                $PentaxQuality      = "Unknown"
        }
        $i = [INT](Get-PentaxMakerNoteProperty -image $image -exifID 0x14)
                $PentaxIso         = @{3=50;     4=64;      5=80;      6=100;     7=125;    8=160;    9=200;   10=250;   11=320;
                                        12=400;   13=500;    14=640;    15=800;    16=1000;  17=1250;  18=1600;  19=2000;  20=2500;
                                        21=3200;  22=4000;   23=5000;   24=6400;   25=8000;  26=10000; 27=12800; 28=16000; 29=20000;
                                        30=25600; 31=32000;  32=40000;  33=51200;
                                    259=70;   260=100;   261=140;   262=200;   263=280;  264=400;
                                    265=560;  266=800;   267=1100;   68=1600;  269=2200; 270=3200}[$i]
        if (-not $PentaxIso )        {
                $PentaxIso          = $i
        }
        $s =(Get-PentaxMakerNoteProperty -image $image -exifID 63).Split(" ")
        $i = (256 * [int]$S[0]) + [int]$s[1]
                $PentaxLens = @{256="M-Series";                                512="A-Series";
                                812="Sigma 10-20 F4-5.6 EX DC";                814="Sigma 100-300 F4.5-6.7";
                                1039="SMC PENTAX-FA 28-105mm F4-5.6 [IF]";     1048="SMC-PENTAX-FA 77mm F1.8 Limited";
                                1071="SMC PENTAX-FA J 18-35mm F4-5.6 AL";      1284="SMC PENTAX-FA 50mm F1.4";
                                2023="SMC PENTAX-DA 18-250 F3.5-63 ED EL[IF]" }[$i]
        if (-not $PentaxLens) {
                $PentaxLens =  "ID= $i"
        }


        $s = (Get-PentaxMakerNoteProperty -exifID 0x35 -image $image) -split "\s*,\s*"
        $PentaxSensorSize = "" + [float]$s[0]/500  + " x " + [float]$s[1]/500 + "mm"
    <# ? Add ?
    0x0032   ImageEditing

    '0 0' = None
    '0 0 0 0' = None
    '0 0 0 4' = Digital Filter
    '2 0 0 0' = Cropped
    '4 0 0 0' = Digital Filter 4
    '6 0 0 0' = Digital Filter 6
    '16 0 0 0' = Frame Synthesis?
    #>

        $h= @{}
        if ($full) {
                    $e= get-exif $image ;
                    Get-Member -input $e -MemberType noteproperty | ForEach-Object {$h.add($_.name, $e."$($_.name)") }
        }
        else {     $h.add("FullName", $image.FullName) }
        New-Object PSObject -Property (@{PentaxContrast       =  $PentaxContrast
                                        PentaxShutterCount   =  $PentaxShutterCount
                                        PentaxSaturation     =  $PentaxSaturation
                                        PentaxSharpening     =  $PentaxSharpening
                                        PentaxImageTone      =  $PentaxImageTone
                                        PentaxShakeReduction =  $PentaxShakeReduction
                                        PentaxPictureMode    =  $PentaxPictureMode
                                        PentaxDriveMode      =  $PentaxDriveMode
                                        PentaxFocusMode      =  $PentaxFocusMode
                                        PentaxProcessingCount=  $PentaxProcessingCount
                                        PentaxEffectiveLV    =  $PentaxEffectiveLV
                                        PentaxFocusPoint     =  $PentaxFocusPoint
                                        PentaxLens           =  $PentaxLens
                                        PentaxMetermode      =  $PentaxMetermode
                                        PentaxModel          =  $PentaxModel
                                        PentaxQuality        =  $PentaxQuality
                                        PentaxWhiteBalance   =  $PentaxWhiteBalance
                                        PentaxNoiseReduction =  $PentaxNoiseReduction
                                        PentaxHighIsoNR      =  $PentaxHighIsoNR
                                        PentaxIso            =  $PentaxIso
                                        PentaxAELock         =  $PentaxAELock
                                        PentaxSensorSize     =  $PentaxSensorSize
                                        PentaxTemperature    = "$PentaxTemperature°c" } +$h )
    }
    end     {
        Write-Progress -Activity "Getting Pentax Exif Data for" -Completed -Status " "
  }
}