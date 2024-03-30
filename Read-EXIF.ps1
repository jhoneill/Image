#The meaning of some of the EXIF information here is available from many sources, but a lot comes from Phil Harvey's ExifTool pages
# https://sno.phy.queensu.ca/~phil/exiftool/ - some is beter explained there and some doesn't seem to be recorded anywhere else.
# At the very least his work saved me a lot of time, but there are things I would not have even attempted without that resource.
# So thanks (and any Donations you want to make) should go to Phil.

#region Tables for the metadata attribute IDs split by data type. Strictly, not all the data is "EXIF"
$ExifUndefineds        =@(0xA300, 0xA301,0x044D )
$ExifBytes             =@(0x0005)
$ExifStrings           =@(0x0001, 0x0003, 0x0008, 0x0009, 0x000A, 0x000C, 0x000D, 0x000E, 0x000F, 0x0010, 0x0011, 0x0012, 0x001D,
                          0x010E, 0x010F, 0x0110, 0x0131, 0x0132, 0x013B, 0x5041, 0x8298, 0x9003, 0x9004, 0x9290, 0x9291, 0x9292, 0xA430, 0xA431, 0xA433, 0xA434)
$ExifUInts             =@(0x0100, 0x0101, 0x0102, 0x0106, 0x0112, 0x0115, 0x0128, 0x0213, 0x4746, 0x4749, 0x5023, 0x5030, 0x5090, 0x5091, 0x8822, 0x8827, 0x8830, 0x9207, 0x9208, 0x9209,
                          0xA001, 0xA210, 0xA217, 0xA401, 0xA402, 0xA403, 0xA405, 0xA406, 0xA407, 0xA408, 0xA409, 0xA40A, 0xA40C) #8830
$ExifLongs             =@()
$ExifULongs            =@(0x0201, 0x0202, 0xA002, 0xA003, 0x8831, 0x8832)  #8831
$ExifRationals         =@(0x9201, 0x9203, 0x9204)
$ExifuRationals        =@(0x0002, 0x0004, 0x0006, 0x0007, 0x000D, 0x011A, 0x011B, 0x502D, 0x502E, 0x829A, 0x829D, 0x9102, 0x9202, 0x9205, 0x920A, 0xA20E, 0xA20F, 0xA404, 0xA432, 0xA500)
$ExifVectorOfUndefineds=@(0x001b, 0x5042, 0x9000, 0x9101, 0x927C, 0x9286, 0xA000, 0xA300, 0xA302, 0xC4A5)
$ExifVectorOfBytes     =@(0x0000, 0x501B, 0x8773, 0x9C9B, 0x9C9C, 0x9C9D, 0x9C9E, 0x9C9F)
$ExifVectorOfUints     =@(0x5090, 0x5091)
$ExifVectorOfLongs     =@()
$ExifVectorOfULongs    =@()
$ExifVectorOfRationals =@()
$ExifVectorOfURationals=@(0x0002, 0x0004, 0x0007, 0x0214, 0xA432)
$ExifDateStrings       =@(0x0132, 0x9003, 0x9004)                #0x001d
$ExifMultiStrings      =@(0x9C9E)
$ExifnumberStrings     =@(0x9209, 0x9201, 0x9202, 0x0612, 0x0000)
$ExifUnicodeStrings    =@(0x9286, 0x9c9b, 0x9c9c, 0x9c9d, 0x9c9e, 0x9c9f)
$ExifUTF8Strings       =@(0xa000, 0x9286,"Interop:2","2:05","2:40","2:55","2:60","2:62","2:63","2:80","2:90","2:101","2:105","2:110","2:116","2:120")

#endregion

#region Type ID Handling - a hash table and ENUM for the data type IDS, and a giant hash of the all the AttributesIDs and their types IDS.
$ExifTypeIDs           =@{"Undefined"        =1000;         "Byte"=1001; "String"=1002;         "uInt"=1003;         "Long"=1004;          "uLong"=1005;         "Rational"=1006;         "URational"=1007
                          "VectorOfUndefined"=1100; "VectorOfByte"=1101;                "VectorOfUint"=1102; "VectorOfLong"=1103;  "VectorOfULong"=1104; "VectorOfRational"=1105; "VectorOfURational"=1106;}
$ExifTypeHash          =@{} ; foreach ($k in $ExifTypeIDs.Keys ) {(Get-Variable -Name "Exif$k`s").Value | ForEach-Object {$ExifTypeHash[$_] = $ExifTypeIDs[$k]}}

### Create ENUM from TypeIDS
$ExifTypeIDs.Keys  | ForEach-Object -Begin   {$items =""} `
                                    -Process {$items += (",`n {0,20} = {1}" -f $_,$($ExifTypeIDs[$_]) ) } `
                                    -End     {Add-Type -TypeDefinition  "public enum ExifType : int  `n{$($items.Substring(1)) `n}"  }
#endregion

#region Define names For known attribute  IDs
<#Note (a) If a tag ID isn't matched up to a name here it won't be output at the end.
       (b) ORDERED hashtables, to control the output order.
 The following tags are known, but not output
 0x00fe='SubFileType'     ; 0x0103='Compression'         ; 0x0106='PhotometricInterpretation'; 0x0111='StripOffsets'            ; 0x0115='SamplesPerPixel' ;  0x0116='RowsPerStrip';
 0x0117='StripByteCounts' ; 0x011C='PlanarConfig'        ; 0x013E='WhitePoint'               ; 0x013F='PrimaryChromaticities'   ; 0x014a='SubIFDPointer';
 0x0201='ThumbnailOffset' ; 0x0202='ThumbnailLength'     ; 0x0211='YCbCrCoefficients'        ; 0x0213='YCbCrPositioning'        ;
 0x501B='ThumbnailData'   ; 0x5023='ThumbnailCompression'; 0x502D='ThumbnailResolutionX'     ; 0x502E='ThumbnailResolutionY'    ; 0x5030='ThumbnailResolutionUnit'
 0x5090='LuminanceTable'  ; 0x5091='ChrominanceTable'    ; 0x83BB='IPTC-NAAPointer'          ; 0x8649='PhotoShopSettingsPointer';
 0x8769='ExifFDPointer'   ; 0x8773='ICCProfilePointer'   ; 0x8825='GPSFDPointer'             ; 0x9101='ComponentsConfiguration' ; 0x9102='CompressedBitsPerPixel';
 0x9214='SubjectArea'     ; 0x927C='MakerNotePointer'    ;
 0xA005='InterOPFDPointer'; 0xA217='SensingMethod'       ; 0xA301='SceneType'                ; 0xA302='CFAPattern'              ; 0xc4a5='PrintIMPointer';    0xEA1C='EXIFFDPointer'


 REFBlackWhite	0x0214
Comment         0x9C9C
FocalXRes		0xA20E
FocalYRes	    0xA20F
FocalResUnit	0xA210
Gaincontrol	    0xA407


#>
$ExifTagNames      = [ordered]@{
    0x0100='Width'                     ; 0x0101='Height'           ; 0x0112='Orientation'
    0x0102='BitsPerSample'             ;
    0x011A='XResolution'               ; 0x011B='YResolution'      ; 0x0128='ResolutionUnit'
    0x1001='RelatedImageWidth'         ;
    0x1002='RelatedImageHeight'        ;
    0x5041='InteropIndex'              ; 0x5042='InteropVersion'   ;
    0xA001='ColorSpace'                ; 0xA002='PixelXDimension'  ; 0xA003='PixelYDimension'
    0x010F='CameraManufacturer'        ; 0x0110='CameraModel'      ; 0x0131='Software'
    0xa430='OwnerName'                 ; 0xa431='SerialNumber'     ;                                          # <---
    0xA433='LensMake'                  ; 0xA434='LensModel'        ; 0xA432='LensData'
    0x920A='FocalLength'               ; 0xA404='DigitalZoomRatio' ; 0xA405='FocalLengthEquivIn35mmFormat'
    0xA40C='SubjectRange'              ; 0xA420='ImageUniqueID'    ; 0xA435='LensSerialNumber'
    0x829A='Exposuretime'              ; 0x829D='FNumber'          ; 0x9205='MaxApperture'
    0x9201='ShutterSpeedValue'         ; 0x9202='ApertureValue'    ; 0x9203='BrightnessValue'
    0x8827='ISOSpeed'                  ; 0x8830='SensitivityType'  ; 0x8831='StandardOutputSensitivity'
    0x8832='RecommendedExposureIndex'  ;                                                                      # <---
    0x9207='MeteringMode'              ; 0x8822='ExposureProgram'  ; 0x9204='ExposureBiasEV'
    0xA300='FileSource'                ; 0XA401='CustomRender'     ;
    0xA402='ExposureMode'              ; 0xA406='SceneCaptutreMode'; 0xA403='WhiteBalance'
    0xA408='Contrast'                  ; 0xA409='Saturation'       ; 0xA40A='Sharpness'
    0xa500='Gamma'                     ; 0x9208='LightSource'      ; 0x9209='FlashSettings'
    0x9290='SubSecTime'                ; 0x9291='SubSecTaken'      ; 0x9292='SubSecDigitized'
    0x9003='DateTaken'                 ; 0x9004='DateDigitized'    ; 0x0132='DateModified'
    0x9C9D='Author'                    ; 0x013B='Artist'           ; 0x8298='Copyright'
    0x9286='Comment'                   ; 0x9C9B='Title'            ; 0x9C9F='Subject'
    0x010E='ImageDescription'          ; 0x9C9E='Keywords'         ;
    0x4746='Rating'                    ; 0x4749='RatingPercent'   ;
    0xC612='DNGVersion'                ; 0x9000='ExifVersion'      ; 0x9216='TiffStandardID'
    0xA000='FlashPixVersion'           ;
    0x0000='GPSVersion'                ;
    0x0001='GPSLatRef'                 ; 0x0002='GPSLattitude'     ;
    0x0003='GPSLongRef'                ; 0x0004='GPSLongitude'     ;
    0x0005='GPSAltRef'                 ; 0x0006='GPSAltitude'      ;
    0x0007='GPSTimeStamp' 	           ; 0x0008='GPSSatellites'    ;
    0x0009='GPSStatus'                 ; 0x000a='GPSMeasureMode'   ; 0x000b='GPSDofP'   ;
    0x000c='GPSSpeedRef'               ; 0x000d='GPSSpeed'         ;
    0x000e='GPSTrackRef'               ; 0x000f='GPSTrack'         ;
    0x0010='GPSImgDirectionRef'        ; 0x0011='GPSImgDirection'  ;
    0x0012='GPSMapDatum'               ;
    0x001b='GPSProcessingMethod'       ;
    0x001d='GPSDateStamp'              ; 9999='GPSSummary'
   }
#We know these maker note tags but don't output them       0x0050 ='_ColorTemperature' ;    0x0207 ='_LensInfo' #Can't decode properly ;
#     0x0224 ='_EVStepInfo';    0x0010 = '_FocusPositon';  0x022D = '_WhitebalanceLevels';  0x022A = '_FilterInfo';
#     0x0228 = '_FaceSize';     0x0228 = '_FacePos';       0x007a = '_ISOAutoParameters';   0x0092 = 'IntervalShooting';
#     0x0095 = 'SkinToneCorrection';                       0x0096 = 'ClarityControl';       0x004d ='_FlashExposureComp'

$PentaxTagNames    = [ordered]@{
    0x0005='PentaxModel'               ; 0x0215='PentaxBodyInfo'   ;  0x0229='PentaxSerialNo'     ;
    0x0035='PentaxSensorSize'          ; 0x2001='PentaxFirmwareDate'                              ;
    0x0027='PentaxDSPFirmwareVersion'  ;
    0x0028='PentaxCPUFirmwareVersion'  ;
    0x005d='PentaxShutterCount'        ; 0x0029='PentaxFrameNumber';
    0x3006='PentaxHomeTown'            ; 0x006B='PentaxTimeZoneInfo'; 0x0216='PentaxBatteryInfo'  ;  # as yet can't properly decode
    0x0047='PentaxTemperature'         ;
    0x0037='PentaxColorSpace'          ; 0x0008='PentaxQuality'    ;  0X0080='PentaxAspectRatio'  ;
    0x003f='PentaxLens'                ; 0x001d='PentaxFocalLength';
    0x000d='PentaxFocusMode'           ; 0x000e='PentaxFocusPoint' ;  0x0072='PentaxAFAdjustment' ;
    0x000c='PentaxFlashMode'           ; 0x0034='PentaxDriveMode'  ;
    0x0017='PentaxMetermode'           ; 0x002d='PentaxEffectiveLV'; 0x0048='PentaxAELockEnabled' ;
    0x0012='PentaxExposureTime'        ; 0x0013='PentaxFNumber'    ; 0x0014='PentaxIso'           ;
    0x0016='PentaxExposureComp'        ;
    0x0033='PentaxPictureMode'         ; 0x005c='PentaxSR'         ; #
    0x022b='PentaxLevelInfo'           ; #as yet can't properly decode
    0x0243='PentaxPixelShiftInfo'      ; 0x0245='PentaxAFInfo'      ; 0x0085 = 'PentaxHDRInfo'    ;
    0x0019='PentaxWB'                  ;
    0x0049='PentaxNREnabled'           ; 0x0071='PentaxHighIsoNRMode'                             ;
    0x0062='PentaxRawDevProcessing'    ; 0x0041='PentaxProcessingCount'                           ;
    0x004f='PentaxImageTone'           ;
    0x0073='PentaxMonoFilter'          ; 0x0074='PentaxMonoToning'                                ;
    0x007f='PentaxBleachBypassToning'  ; 0x007b='PentaxCrossProcess'                              ;
    0x001f='PentaxSaturation'          ; 0x0020='PentaxContrast'   ; 0x0021='PentaxSharpening'    ;
    0x0067='PentaxHue'                 ;
    0x0069='PentaxDRExpansion'         ;
    0x006c='PetnaxHighLowKeyAdjust'    ;
    0x006d='PentaxContrastHighlight'   ;
    0x006e='PentaxContrastshaddow'     ;
    0x006f='PentaxContrastHighLightShaddowAdjust';
    0x0070='PentaxFineSharpness'       ;
    0x0079='PentaxShadowCorrect'       ;
    0x007d='PentaxLensCorrection'      ;
    0x3002='PentaxQuality'             ; # 0x0008 and 0x3002 are mutually exclusive
    0x3003='PentaxFocusMode'           ; # 0x000d and 0x3003 are mutually exclusive
    0x3014='PentaxIso'                 ; # 0x0014 and 0x3014 are mutually excusive
 }

$AppleTagNames     = [ordered]@{
    0x0003='AppleRunTime'              ;
    0x0008='AppleAccelerationVector'   ;
    0x000a='AppleHDRImageType'         ;
    0x000b='AppleBurstUUID'            ;
    0x0015='AppleImageGUID'
}

$IPTCTagNames      = [ordered]@{
    '1:90'  = 'IPTCCodedCharacterSet'
    '2:00'  = 'IPTCRecordVersion'
    '2:05'  = 'IPTCObjectName'
    '2:25'  = 'IPTCKeywords'
    '2:40'  = 'IPTCSpecialInstructions'
    '2:80'  = 'IPTCByline'
    '2:55'  = 'IPTCDateCreated'
    '2:60'  = 'IPTCTimeCreated'
    '2:62'  = 'IPTCDigitalCreationDate'
    '2:63'  = 'IPTCDigitalCreationTime'
    '2:90'  = 'IPTCCity'
    '2:101' = 'IPTCCountryPrimaryLocationName'
    '2:105' = 'IPTCHeadline'
    '2:110' = 'IPTCCredit'
    '2:116' = 'IPTCCopyrightNotice'
    '2:120' = 'IPTCCaptionAbstract'
}

$InterOpTagNames   = [ordered]@{        0x0001='InteropID'        ; 0x0002='InteropVersion'}
#endregion

#region Enumerations for EXIF attributes e.g. Attribute 0x0128  is 'ResolutionUnit', and is 1 for "none" 2 for "inch" 3 for "cm".
<# Strings known but not processed currently
@{
0x5041=@{ 'R03'= "R03 - DCF option file (Adobe RGB)"; 'R98'="R98 - DCF basic file (sRGB)"; 'THM' = "THM - DCF thumbnail file"}
0x0009=@{   'A'= "Measurement Active";                  'V'="Measurement Void"}
0x000A=@{	'2'= "2-Dimensional Measurement";           '3'="3-Dimensional Measurement"}
0x000C=@{	'K'= "km/h" ;                               'M'="mph" ; 'N' = "knots"}
0x000E=@{   'M'= "Magnetic North";                      'T'="True North"}
0x0010=@{   'M'= "Magnetic North";                      'T'="True North"}
}
#>
#int lookups
$ExifLookUps       =@{
    0x0128=@{     1="No Unit"          ;      2="Inch"             ;      3="Centimetre"          }
    0x8822=@{     1="Manual"           ;      2="Program: Normal"  ;      3="Aperture Priority"   ;      4="Shutter Priority";
                  5="Program: Creative";      6="Program: Action"  ;      7="Portrait Mode"       ;      8="Landscape Mode"
    }
    0x9207=@{     1="Av"               ;      2="Centre"           ;      3="Spot"                ;      4="Multi-Spot"             ;     5="Multi-Segment"         ;     6="Partial"           }
    0x0112=@{     1="0"                ;      3="180"              ;      6="270"                 ;      8="90"                     }
    0xA001=@{     1="sRGB"             ;      2="Adobe RGB"        ;  65533="Wide Gamut RGB"      ;  65534="ICC Profile"            ; 65535="Uncalibrated"}
    0xA300=@{     1="Film scanner"     ;      2="Print scanner"    ;      3="Digital still camera"}
    0xA401=@{     0="Normal"           ;      1="Custom"           ;      3="HDR"                 ;      4="No HDR Effect";
                  6="Panorama"         ;      8="Portrait"         ;      9="No Portrait Effect"  }
    0xa402=@{     0="Auto"             ;      1="Manual"           ;      2="Auto Bracket"        }
    0xa403=@{     0="Auto"             ;      1="Manual"           }
    0xa406=@{     0="Standard"         ;      1="Landscape"        ;      2="Portrait"            ;      3="NightScene"             }
    0xa408=@{     0="Normal"           ;      1="Soft"             ;      2="Hard"                }
    0xa409=@{     0="Normal"           ;      1="Low"              ;      2="High"                }
    0xa40a=@{     0="Normal"           ;      1="Soft"             ;      2="Hard"                }
    0xA40C=@{     1="Macro"            ;      2="Close"            ;      3="Distant"             }
    0x9208=@{     0="Auto"             ;      1="Daylight"         ;      2="Fluorescent"         ;      3="Tungsten"               ;     4="Flash"                 ;     9="Fine Weather"     ;
                 10="Cloudy Weather"   ;     11="Shade"            ;     12="Daylight Fluorescent";     13="Day White Fluorescent"  ;    14="Cool White Fluorescent";    15="White Fluorescent";
                 17="Standard Light A" ;     18="Standard Light B" ;     19="Standard Light C"    ;     20="D55"                    ;    21="D65"                   ;    22="D75"              ;
                 23="D50"              ;     24="ISO Studio Tungsten"
    }
    0x8830=@{     0='Unknown'          ;      1='Standard Sensitivity';   2='Recommended EI'      ;      3='ISO Speed'              ;     4='Standard Sensitivity & Recommended EI';
                  5='Standard Sensitivity & ISO Speed'             ;      6='Recommended EI & ISO Speed' ;
                  7='Standard Sensitivity , Recommended EI & ISO Speed'}
}
$PentaxLookUps     =@{
    0x0008 = @{   0="Good"             ;      1="Better"           ;      2="Best"                ;
                  3="TIFF"             ;      4="RAW"              ;      5="Premium"             ;      7="RAW (Pixel-shift enabled)"
    }  #quality
    0x000d=@{     0="Normal"           ;      1="Macro"            ;      2="Infinity"            ;      3="Manual"                 ;
                  4="Super Macro"      ;      5="Pan Focus"        ;     16="AF-S"                ;     17="AF-C"                   ;
                 18="AF-A"             ;     32="Contrast-detect"  ;     33="Tracking Contrast-detect"                              ;
                272="AF-S (Release-priority)"                      ;    273="AF-C (Release-priority)"                               ;
                274="AF-A (Release-priority)"                      ;    288="Face Detect"
    }  #Focus mode
    0x0017=@{     0="Multi-Segment"    ;      1="Center-Weighted"  ;      2="Spot"}  #Metering mode
    0x0019=@{     0="Auto"             ;      1="Daylight"         ;      2="Shade"               ;      3="Fluorescent"            ;     4="Tungsten"          ;
                  5="Manual"           ;      6="Daylight Fluorescent"                            ;      7="Day White Fluorescent"  ;     8="White Fluorescent" ;     9="Flash"            ;
                 10="Cloudy"           ;     15="Color Temperature Enhancement"                   ;     17="Kelvin"                 ; 65534="Unknown"           ; 65535="User-Selected"
    } #White Balance
    0x0037=@{     0="sRGB"             ;      1="Adobe RGB"       }       #Colour space
    0x004f=@{     0="Natural"          ;      1="Bright"          ;       2="Portrait"            ;      3="Landscape"              ;     4="Vibrant"           ;
                  5="Monochrome"       ;      6="Muted"           ;       7="Reversal Film"       ;      8="Bleach Bypass"          ;     9="Radiant"
    } #Image Tone
    0x0062=@{     1="K10D,K200D,K2000,K-m"                        ;       3="K20D"                ;      4="K-7"                    ;     5="K-x"               ;
                  6="645D"             ;  7="K-r"                 ;       8="K-5,K-5 II,K-5 II s" ;      9="Q"                      ;    10="K-01,K-30"         ;
                 11="Q10"              ; 12="MX-1"                ;      13="K-3,K-3 II"          ;     14="645Z"                   ;    15="K-S1,K-S2"         ;
                 16="K-1"              ; 17="K-70"
    } #Processing version
    0x0067=@{     0=-2                 ;  1="Normal"              ;       2=2                     ;      3=-1                       ;     4=1                   ;
                  5=-3                 ;  6=3                     ;       7=-4                    ;      8=4                        ; 65535="None"
    } #Hue
    0x0073=@{     1="Green"            ;  2="Yellow"              ;       3="Orange"              ;      4="Red"                    ;     5="Magenta"           ;
                  6="Blue "            ;  7="Cyan"                ;       8="Infrared"                                              ; 65535="None"
    } #Monochrome Filter Effect
    0x0074=@{     0=-4                 ;  1=-3                    ;       2=-2                    ;      3=-1                       ;     4=0                   ;
                  5=1                  ;  6=2                     ;       7=3                     ;      8=4                        ; 65535="None"
    } #Monochrome Toning level
    0x007f=@{     1="Green"            ;  2="Yellow"              ;       3="Orange"              ;      4="Red"                    ;     5="Magenta"           ;
                  6="Purple"           ;  7="Blue"                ;       8="Cyan"                                                  ; 65535="Off"
    }# BleachBypassToning
    0x3002=@{     1="Economy"          ;  2="Normal"              ;       3="Fine"                }
    0x3003=@{     0="Manual"           ;  1="Focus Lock"          ;       2="Macro"               ;      3="Single-Area Auto Focus"                             ;     5="Infinity"         ;  6="Multi-Area Auto Focus"                               ;      8="Super Macro"
    }
    0x0014=@{
         3=50    ;   4=64   ;   5=80   ;    6=100   ;    7=125    ;   8=160  ;   9=200  ;  10=250  ;  11=320  ;  12=400   ; 13=500   ; 14=640   ;
        15=800   ;  16=1000 ;  17=1250 ;   18=1600  ;   19=2000   ;  20=2500 ;  21=3200 ;  22=4000 ;  23=5000 ;  24=6400  ; 25=8000  ; 26=10000 ;
        27=12800 ;  28=16000;  29=20000;   30=25600 ;   31=32000  ;  32=40000;  33=51200;  34=64000;  35=80000;  36=102400; 37=128000; 38=160000;
        39=204800;  50=50   ; 100=100  ;  200=200   ;  258=50     ; 259=70   ; 260=100  ; 261=140  ; 262=200  ; 263=280   ; 264=400  ; 265=560  ;
       266=800   ; 267=1100 ; 268=1600 ;  269=2200  ;  270=3200   ; 271=4500 ; 272=6400 ; 273=9000 ; 274=12800; 275=18000 ; 276=25600; 277=36000;
       278=51200 ; 400=400  ; 800=800  ; 1600=1600  ; 3200=3200
    } #ISO
}
$PentaxFocus11Pts  =@{
                    1="Top Left"       ;     2="Top Center"       ;     3="Top Right"             ;     4="Middle Far-Left"         ;     5="Middle Left"       ;
                    6="Middle Center"  ;     7="Middle Right"     ;     8="Middle Far-Right"      ;     9="Bottom Left"             ;    10="Bottom Center"     ;
                   11="Bottom Right"   ; 65532="Face Detection AF"; 65533="Auto Tracking AF"      ; 65534="Fixed Center"            ; 65535="Auto"
}
$PentaxK1FocusPts  =@{
       0='None'                        ;   1='Top-left'           ;      2='Top Near-left'        ;      3='Top'                    ;     4='Top Near-right'    ;
       5='Top-right'                   ;   6='Upper Far-left'     ;      7='Upper-left'           ;      8='Upper Near-left'        ;     9='Upper-middle'      ;
      10='Upper Near-right'            ;  11='Upper-right'        ;     12='Upper Far-right'      ;     13='Far Far Left'           ;    14='Far Left'          ;
      15='Left'                        ;  16='Near-left'          ;     17='Center'               ;     18='Near-right'             ;    19='Right'             ;
      20='Far Right'                   ;  21='Far Far Right'      ;     22='Lower Far-left'       ;     23='Lower-left'             ;    24='Lower Near-left'   ;
      25='Lower-middle'                ;  26='Lower Near-right'   ;     27='Lower-right'          ;     28='Lower Far-right'        ;    29='Bottom-left'       ;
      30='Bottom Near-left'            ;  31='Bottom'             ;     32='Bottom Near-right'    ;     33='Bottom-right'           ;   263='Zone Select Upper-left';
     264='Zone Select Upper Near-left' ; 265='Zone Select Upper Middle'                           ;
     266='Zone Select Upper Near-right'; 267='Zone Select Upper-right'                            ;
     270='Zone Select Far Left'        ; 271='Zone Select Left'   ;    272='Zone Select Near-left';
     273='Zone Select Center'          ; 274='Zone Select Near-right'                             ;
     275='Zone Select Right'           ; 276='Zone Select Far Right'                              ;
     279='Zone Select Lower-left'      ; 280='Zone Select Lower Near-left'                        ;
     281='Zone Select Lower-middle'    ; 282='Zone Select Lower Near-right'                       ;
     283='Zone Select Lower-right'     ;
   65531='AF Select'                   ; 65532='Face Detect AF'   ;  65533='Automatic Tracking AF';
   65534='Fixed Center'                ; 65535='Auto'
}
$PentaxAspectRatio =@{
       0="4:3"                         ;     1="3:2"              ;      2="16:9"                 ;      3="1:1"
}
$PentaxFlash0      =@{
  0x0000="Auto, Did not fire"                                     ; 0x0001="Off, Did not fire"    ;
  0x0002="On, Did not fire"                                       ;
  0x0003="Auto (Red-eye reduction), Did not fire"                 ; 0x0005="On (Wireless [Master]) Did not fire"                    ;
  0x0100="Auto, Fired"                 ; 0x0102="On, Fired"       ; 0x0103="Auto (Red-eye reduction), Fired"                        ;
  0x0104="On (Red-eye reduction)"                                 ; 0x0105="On (Wireless [Master])"                                 ;
  0x0106="On (Wireless [Control])"     ; 0x0108="On, Soft"        ; 0x0109="On, Slow-sync"                                          ;
  0x010a="On (Red-eye reduction), Slow-sync"                      ; 0x010b="On, Trailing-curtain Sync"
}
$PentaxFlash1      =@{
  0x0000="n/a - Off-Auto-Aperture"                                ; 0x003f="Internal"              ; 0x0100="External, Auto"        ;
  0x023f="External, Flash Problem"                                ; 0x0300="External, Manual"      ; 0x0304="External, P-TTL Auto"  ;
  0x0305="External, Contrast-control Sync"                        ; 0x0306="External, High-speed Sync"                              ;
  0x030c="External, Wireless"                                     ; 0x030d="External, Wireless, High-speed Sync"
}
$PentaxLens        =@{
      1039="smc PENTAX-FA 28-105mm F4-5.6 [IF]"                   ;    814="Sigma 100-300 F4.5-6.7"                                 ;  1071="smc PENTAX-FA J 18-35mm F4-5.6 AL"            ;
      1048="smc PENTAX-FA 77mm F1.8 Limited"                      ;    812="Sigma or Tamron Lens"                                   ;  2023="smc PENTAX-DA 18-250mm F3.5-6.3 ED AL [IF]"   ;
       785="smc PENTAX-FA 85mm F2.8 SOFT"                         ;   1284="smc PENTAX-FA 50mm F1.4"                                ;  2010="smc PENTAX-DA 18-55mm F3.5-5.6 AL WR"         ;
      1027="smc PENTAX-FA 43mm F1.9 Limited"                      ;   2111="HD PENTAX-D FA 15-30mm F2.8 ED SDM WR"                  ;  2114="HD Pentax-D FA 85mm F1.4 ED SDM AW"           ;
      2115="HD PENTAX-D FA 21mm F2.4 ED Limited DC WR"            ;   1069="TAMRON 28-300mm F3.5-6.3 Ultra zoom XR"                 ;   578="rmc TOKINA 28mm F2.8"                         ;
      1043="Tamron SP AF 90mm F2.8"                               ;    512="A Series Lens"                                          ;   256="K or M Lens"                                  ;
}
$PentaxModel       =@{
        13="Optio 430"                 ;  76180="*ist-D"          ;  76830="K10D"                  ;  77240="K-7"                   ;  77430="K-5"              ; 77680="K-5 II"           ;
     77681="K-5 II s"                  ;  77760="K-3"             ;  77980="K-3 II"                ;  77970="K-1"                   ;  77750="K-50"             ; 78370="K-70"
     78400="K-1 II"
}
$PentaxHiLo        =@{
         0="Low"                       ;      1="Normal"          ;      2="High"                  ;      3="Med-Low"               ;     4="Med-High"          ;
         5="Very Low"                  ;      6="Very High"       ;      7="-4"                    ;      8="+4"
}
$PentaxPicMode     =@{
         0="Program"                   ;      1="Hi-Speed Program";      2="DOF Program"           ;      3="MTF Program"           ;     4="Standard"          ;     5="Portrait"         ;
         6="Landscape"                 ;      7="Macro"           ;      8="Sport"                 ;      9="Night Scene"           ;    10="No Flash"          ;    11="Soft"             ;
        12="Surf & Snow"               ;     13="Text"            ;     14="Sunset"                ;     15="Kids"                  ;    16="Pet"               ;    17="Candlelight"      ;
        18="Museum"                    ;     19="Food"            ;     20="Stage Lighting"        ;     21="Night Snap"            ;    23='Blue Sky'          ;    24='Sunset'           ;
        26='Night Scene HDR'           ;     27='HDR'             ;     28="Quick Macro"           ;     29='Forest'                ;    30="Self-Portrait"     ;
        31="Illustrations"             ;     33="Digital Filter"  ;     37="Museum"                ;     38="Food"                  ;    40="Green Mode"        ;    49="Light Pet"        ;
        50="Dark Pet"                  ;     51="Medium Pet"      ;     53="Underwater"            ;     54="Candlelight"           ;    55="Natural Skin Tone" ;
        56="Synchro Sound Record"      ;     58="Frame Composite" ;     60="Kids"                  ;     61="Blur Reduction"        ;
       255="Digital Filter"            ;    260="Auto PICT (Standard)"                             ;    261="Auto PICT (Portrait)"  ;
       262="Auto PICT (Landscape)"     ;    263="Auto PICT (Macro)";    264="Auto PICT (Sport)"    ;    512="Program (HyP)"         ;
       513="Hi-speed Program (HyP)"    ;    514="DOF Program (HyP)"                                ;    515="MTF Program (HyP)"     ;
       534="Shallow DOF (HyP)"         ;    768="Green Mode"                                       ;   1024="Shutter Speed Priority";
       1280="Aperture Priority"        ;   1536="Program Tv Shift"                                 ;   1792="Program Av Shift"      ;
       2048="Manual"                   ;   2304="Bulb"                                             ;   2560="Aperture Priority, Off-Auto-Aperture"              ;
       2816="Manual, Off-Auto-Aperture";   3072="Bulb; Off-Auto-Aperture"                          ;   3328="Shutter & Aperture Priority AE"                    ;
       3840="Sensitivity Priority AE"  ;   4096="Flash X-Sync Speed AE"                            ;   4608="Auto Program (Normal)" ;
       4609="Auto Program (Hi-speed)"  ;   4610="Auto Program (DOF)"                               ;   4611="Auto Program (MTF)"    ;  4630="Auto Program (Shallow DOF)"                   ;
       4864="Astrotracer"              ;   5142="Blur Control"
}
$PentaxCities      =@{
          0='Pago Pago'                ;      1='Honolulu'        ;      2='Anchorage'             ;      3='Vancouver'             ;     4='San Francisco'     ;
          5='Los Angeles'              ;      6='Calgary'         ;      7='Denver'                ;      8='Mexico City'           ;     9='Chicago'           ;
         10='Miami'                    ;     11='Toronto'         ;     12='New York'              ;     13='Santiago'              ;    14='Caracus'           ;
         15='Halifax'                  ;     16='Buenos Aires'    ;     17='Sao Paulo'             ;     18='Rio de Janeiro'        ;    19='Madrid'            ;
         20='London'                   ;     21='Paris'           ;     22='Milan'                 ;     23='Rome'                  ;    24='Berlin'            ;
         25='Johannesburg'             ;     26='Istanbul'        ;     27='Cairo'                 ;     28='Jerusalem'             ;    29='Moscow'            ;
         30='Jeddah'                   ;     31='Tehran'          ;     32='Dubai'                 ;     33='Karachi'               ;    34='Kabul'             ;
         35='Male'                     ;     36='Delhi'           ;     37='Colombo'               ;     38='Kathmandu'             ;    39='Dacca'             ;
         40='Yangon'                   ;     41='Bangkok'         ;     42='Kuala Lumpur'          ;     43='Vientiane'             ;    44='Singapore'         ;
         45='Phnom Penh'               ;     46='Ho Chi Minh'     ;     47='Jakarta'               ;     48='Hong Kong'             ;    49='Perth'             ;
         50='Beijing'                  ;     51='Shanghai'        ;     52='Manila'                ;     53='Taipei'                ;    54='Seoul'             ;
         55='Adelaide'                 ;     56='Tokyo'           ;     57='Guam'                  ;     58='Sydney'                ;    59='Noumea'            ;
         60='Wellington'               ;     61='Auckland'        ;     62='Lima'                  ;     63='Dakar'                 ;    64='Algiers'           ;
         65='Helsinki'                 ;     66='Athens'          ;     67='Nairobi'               ;     68='Amsterdam'             ;    69='Stockholm'         ;
         70='Lisbon'                   ;     71='Copenhagen'      ;     72='Warsaw'                ;     73='Prague'                ;    74='Budapest'
}
$PentaxXProcess    =@{
          0="Off"                      ;      1="Random"          ;      2="Preset 1"              ;      3="Preset 2"              ;
          4="Preset 3"                 ;     33="Favorite 1"      ;     34="Favorite 2"            ;     35="Favorite 3"
}
#endregion

#region Create EXIF type for completion
$Global:ExifTagValues = [Ordered]@{} ; foreach ($k in $ExifTagNames.keys) {$ExifTagValues[$ExifTagNames.$k] = $K}
$CodeFrag = @"
public struct Exif
{    public string          Fullname ;
     public string[]        JPEGSegements;
     public string          CreatorTool;
     public string          People;
     public object[]        History;
     public object[]        ToneCurve;
     public object[]        DescAttrib;

"@
foreach ($k in $ExifTagNames.keys  )  {
     if     ($ExifDateStrings       -contains $k)   {$CodeFrag += "  public System.DateTime " + $ExifTagNames.$k +" ;`r`n" }
     elseif ($ExifMultiStrings      -contains $k)   {$CodeFrag += "  public string[]        " + $ExifTagNames.$k +" ;`r`n" }
     elseif ($ExifnumberStrings     -contains $k -or $ExifUnicodeStrings -contains $K -or $ExifUTF8Strings -contains $K -or $ExifLookUps.ContainsKey($K) -or $ExifStrings -contains $K )
                                                    {$CodeFrag += "  public string          " + $ExifTagNames.$k +" ;`r`n" }
     elseif ($ExifBytes             -contains $k -or $ExifUInts -contains $k -or $ExifLongs -contains $K -or $ExifULongs -contains $K)
                                                    {$CodeFrag += "  public int             " + $ExifTagNames.$k +" ;`r`n" }
     elseif ($ExifRationals         -contains $k -or $ExifuRationals -contains $K)
                                                    {$CodeFrag += "  public float           " + $ExifTagNames.$k +" ;`r`n" }
     elseif ($ExifVectorOfBytes     -contains $k -or $ExifVectorOfUints -contains $k -or $ExifVectorOfLongs -contains $K -or $ExifVectorOfULongs -contains $K)
                                                    {$CodeFrag += "  public int[]           " + $ExifTagNames[$k] +" ;`r`n" }
     elseif ($ExifVectorOfRationals -contains $k -or $ExifVectorOfURationals -contains $K)
                                                    {$CodeFrag += "  public float[]         " + $ExifTagNames.$k +" ;`r`n" }
}
$PentaxTagNames.Values | Where-Object {$CodeFrag -notmatch $_} | ForEach-Object  {if ($CodeFrag -notmatch $_)
                                                    {$CodeFrag += "  public string          $_;`r`n" }}
$IPTCTagNames.Values   | ForEach-Object             {$CodeFrag += "  public string          $_;`r`n" }
Add-Type -TypeDefinition ($CodeFrag += "}")
#endregion

function Find-IPTCPosition {
    <#
        .Synopsis Only intended to be called from another function.
    #>
    param ($segments , [System.IO.FileStream]$Stream, [Byte[]]$Array)
    $PhShSeg = ($Segments -match "Photoshop" | Select-Object -First 1)
    if ($PhShSeg) {# PhotoShop seegment maker will be in the form "ffed Photoshop 3.0 9876"
        #we'll need the offset 9876 and the length of the "photoshop 3.0" part
        $PSLabelLen      = ($PhShSeg -replace "^.*(Photoshop.*)\s\d+$",'$1').Length
        $Stream.Position = 0+($PhShSeg -replace "^.*?\s(\d+)$",'$1')
        $ArrayPos        = $Stream.Position % 64kb
        $posdiff         = $Stream.Position - $ArrayPos
        #Read the data at the start of the segment and figure out how long the segment is.
        $BytesRead       = $Stream.Read($Array,$ArrayPos,20)
        $SegLength       = $Array[$ArrayPos+2] * 256 + $Array[$ArrayPos+3]
        $SegEnd          = $ArrayPos + $SegLength
        $Stream.Position = 0+($PhShSeg -replace "^.*?\s(\d+)$",'$1')
        $ArrayPos        = $Stream.Position % 64kb
        $BytesRead       = $Stream.Read($Array,$ArrayPos,($SegLength+2) )
        $ArrayPos        += ($PSLabelLen+5)
        while ($ArrayPos -lt $Segend -and [char[]]$Array[($ArrayPos)..($ArrayPos+3)] -join ""   -eq '8BIM') {
            if ([string]$Array[($ArrayPos+4)..($ArrayPos+5)] -eq "4 4") { $IPTCPos  = $posdiff + $ArrayPos + 12 ; $ArrayPos = $Segend+1  }
            else {      $ArrayPos += ($Array[$ArrayPos+10] * 256 + $Array[$ArrayPos+11] +12 + $(if ($Array[$ArrayPos+11] % 2) {1}  ) )}
        }
        return $IPTCPos
    }
}

function Read-IPTCData     {
    <#
        .Synopsis Only intended to be called from another function.
    #>
    param ($IPTCPos , [System.IO.FileStream]$Stream, [Byte[]]$Array, [System.Collections.Specialized.OrderedDictionary]$PropHash, [switch]$raw )
    $UTF8Encoding  = New-object -TypeName System.Text.UTF8Encoding
    Write-Verbose -Message ( "IPTC-NAA data found at offset {0:x4} ({0})" -f $IPTCPos)
    $IPTCResults      = @{}
    $Stream.Position  = $IPTCPos
    $ArrayPos         = $Stream.Position % 64KB
    $BytesRead        = $Stream.Read($Array,$ArrayPos,5)
    while ($Array[$ArrayPos] -eq 28)   #Oh Joy another way to write tags.
    {   $IPTCTagID    = "" + $Array[$ArrayPos+1] + ":" + $Array[$ArrayPos+2].ToString("00")
        $IPTCLength   =     256 * $Array[$ArrayPos+3] + $Array[$ArrayPos+4]
        $ArrayPos    += 5
        $BytesRead    = $Stream.Read($Array,$ArrayPos,$IPTCLength)
        $IPTCBytes    = $Array[$ArrayPos..($ArrayPos + $IPTCLength -1)]
        $IPTCItem     = New-Object -TypeName psobject -Property @{TagID=$IPTCTagID; TagName=$IPTCTagNames[$IPTCTagID];   length=$IPTCLength; Value=$IPTCBytes ; meaning=$null }
        if ($IPTCResults.ContainsKey($IPTCTagID))  {$IPTCResults[$IPTCTagID] = ,$IPTCItem + $IPTCResults[$IPTCTagID] }
        else                                       {$IPTCResults[$IPTCTagID] =  $IPTCItem }
        $ArrayPos    += $IPTCLength
        $BytesRead    = $Stream.Read($Array,$ArrayPos,5)
    }
    if ($IPTCResults["2:025"]) { $PropHash["IPTCKeywords"]=($IPTCResults["2:025"] | ForEach-Object { $_.meaning = $UTF8Encoding.GetString($_.Value) ; $_.meaning })}
    $ExifUTF8Strings | ForEach-Object {if ($IPTCResults[$_].Value) {$IPTCResults[$_].meaning = $PropHash[$IPTCTagNames[$_]] = $UTF8Encoding.getString($IPTCResults[$_].Value)}}
    if ($Raw) {$IPTCResults.Values}
}

function Read-XAPData      {
    <#
        .Synopsis Only intended to be called from another function.
   #>
    param ($XMPPos , $XMPLength , [System.IO.FileStream]$Stream, [Byte[]]$Array, [System.Collections.Specialized.OrderedDictionary]$PropHash, [switch]$raw, $DumpAsXML )
    Write-Verbose -Message ("XAP Segment found at offset {0:x4} ({0})" -f $XMPPos)
    $Stream.Position     = $XMPPos + 2
    $ArrayPos            = $Stream.Position   % 64kb
    $BytesRead           = $Stream.Read($Array,$ArrayPos,($XMPLength) )
    $XMLtext             = $UTF8Encoding.GetString($Array[($ArrayPos)..($ArrayPos + $XMPLength -1)]).trim([char]0)
    $XMLtext             = $XMLtext   -replace "(?s)^\s*<\?.*?\?>\s*<","<"  -replace "(?s)>\s+<\?.*?\?\>\s*$",">"
    $Global:xapXML       = [xml]$XMLtext
    if ($XMLtext -match "CreatorTool") {
        $PropHash["CreatorTool"]  = $xapXML.xmpmeta.rdf.Description | ForEach-Object {$_.CreatorTool }       | Where-Object {$_}
    }
    if ($XMLtext -match "usercomment" -and -not $PropHash["Comment"]) {
        $PropHash["Comment"]      = $xapXML.xmpmeta.rdf.Description.usercomment.alt.li.'#text' | Where-Object {$_}
    }
    if ($XMLtext -match "persondisplayname") {
        $PropHash["People"]       = ($xapXML.xmpmeta.rdf.Description | ForEach-Object{$_.regioninfo.description.regions.bag.li.description.persondisplayname."#text"} |
                                                                                                             Where-Object {$_ -match "\w+"}) -join "; "
    }
    if ($XMLtext -match "RegionList") {
        $PropHash["Regions"]      = ($xapXML.xmpmeta.RDF.Description.Regions.RegionList.seq.li.type -join '; ')
    }
    if ($XAPDetails) {
        Write-Verbose -Message $XMLtext
        $Prophash["History"]    =  $xapXML.xmpmeta.RDF.Description | ForEach-Object {$_.History.Seq.li }   | Select-Object -Property when,action,softwareagent
        $Prophash["ToneCurve"]  =  $xapXML.xmpmeta.RDF.Description | ForEach-Object {$_.ToneCurve.seq.li } | Where-Object {$_}
        $DescAttrib             =  @()
        $DescAttrib            +=  $xapXML.xmpmeta.RDF.Description | ForEach-Object {$_.attributes}        | Where-Object -Property prefix -ne "xmlns" |
                                                                     Select-Object -Property   @{N="Attribute";e={$_.localname}},@{N="Value";e={$_."#text"}}
        $DescAttrib            +=  $xapXML.xmpmeta.RDF.Description.ChildNodes | Where-Object {$_."#text" -is [string]} | Select-Object -Property @{n="Attribute"; e={$_.name -replace "^.+:",""}}, @{n="Value"; e={$_."#Text"}}
        $DescAttrib            +=  New-Object -TypeName pscustomobject -Property  @{"Attribute"="Gradients";        "Value"=($xapXML.xmpmeta.RDF.Description.GradientBasedCorrections.seq.li.Count)}
        $DescAttrib            +=  New-Object -TypeName pscustomobject -Property  @{"Attribute"="CircularGradients";"Value"=($xapXML.xmpmeta.RDF.Description.CircularGradientBasedCorrections.seq.li.Count)}
        $DescAttrib            +=  New-Object -TypeName pscustomobject -Property  @{"Attribute"="Retouches";        "Value"=($xapXML.xmpmeta.RDF.Description.RetouchInfo.seq.li.Count)}
        $ExtXMP                 =  $DescAttrib  | Where-Object {$_.attribute -eq "HasExtendedXMP"}
        if ($ExtXMP) {
            Write-Verbose -Message ("XAP Segment has 'ExtendedXMP' flag" )
            $XMLtext = ""
            $Segresults | Where-Object {$_.tagname -eq "http://ns.adobe.com/xmp/extension/"} | ForEach-Object {
                $Stream.Position  = $_.offset
                Write-Verbose -Message ("XMP Extension found at offset{0:x4} ({0})" -f $_.offset)
                $ArrayPos         = $_.offset % 64kb
                $BytesRead        = $Stream.Read($Array,$ArrayPos ,($_.length +2))
                $XMLtext         += $UTF8Encoding.GetString($Array[($ArrayPos+5+$_.TagName.length + $ExtXMP.Value.length+8)..($ArrayPos + $_.Length+1)])
            }
            $XMLtext              = $XMLtext   -replace "(?s)^\s*<\?.*?\?>\s*<","<"  -replace "(?s)>\s+<\?.*?\?\>\s*$",">"
            Write-Verbose -Message  $XMLtext
            $DescAttrib          += (([xml]$XMLtext).xmpmeta.rdf.Description | ForEach-Object {$_.attributes}        | Where-Object -Property  prefix -ne "xmlns" |
                                                                         Select-Object -Property @{N="Attribute";e={$_.localname}},@{N="Value";e={$_."#text"}} )
        }
        $Prophash["DescAttrib"] = $DescAttrib
    }
}

function Read-ExifFD       {
    <#
        .Synopsis Only intended to be called from another function.
    #>
    param ([Parameter(Mandatory=$true)]
           [System.IO.Stream ]$Stream ,
           [Parameter(Mandatory=$true)]
           $ImgDirStart ,
           [boolean]$LittleEndian,
           [Parameter(Mandatory=$true)]
           $TagNames,
           [Parameter(Mandatory=$true)]
           $Dataoffset,
           [switch]$makerNoteOffset
    )
    $Results                        = @{}
    $UTF7Encoding                   = New-Object -TypeName System.Text.UTF7Encoding
    $BytesPerExifRow                = 12    # Each Exif row is is 12 bytes
    #The first two byte word in an Exif Image directory indicates how many rows of Exif data are in that directory; read that and the rows themselves.
    $Stream.Position                = $ImgDirStart + $Dataoffset
    $ArrayPos                       = $Stream.Position % 64KB
    $BytesRead                      = $Stream.Read($Array,$ArrayPos ,2)
    if ($LittleEndian)  { $RowCount =      $Array[$ArrayPos] +(256*$Array[$ArrayPos + 1])}
    else                { $RowCount = (256*$Array[$ArrayPos])+     $Array[$ArrayPos + 1] }
    if ($RowCount -gt  2000)  {Write-Warning -Message "$Path Seems to have a data block with $rowcount rows, skipping that."
                               return  $Results }
    $TableStart                     = $ArrayPos    + 2
    $TableLength                    = $RowCount * $BytesPerExifRow
    $BytesRead                      = $Stream.Read($Array,$TableStart,$TableLength)

    # Now parse the 12 byte rows: the first 2 are the ID, the next 2 are the datatype, the next 4 are EITHER data or an offset to it.
    for ($RowNo = 0; $RowNo -lt $RowCount; $RowNo ++) {
         $RowStart = $TableStart + ($RowNo * $BytesPerExifRow)
         if ($LittleEndian) {$ExifTagID                 = $Array[$RowStart    ] + (256 * $Array[$RowStart + 1])
                             $ExifTypeID                = $Array[$RowStart + 2] + (256 * $Array[$RowStart + 3])
                             $ExifDataLength            = $Array[$RowStart + 4] + (256 * $Array[$RowStart + 5]) #+ (65536 * $Array[$RowStart + 6]) + (16777216 * $Array[$RowStart + 7])
         }
         else               {$ExifTagID                 = $Array[$RowStart + 1] + (256 * $Array[$RowStart    ])
                             $ExifTypeID                = $Array[$RowStart + 3] + (256 * $Array[$RowStart + 2])
                             $ExifDataLength            = $Array[$RowStart + 7] + (256 * $Array[$RowStart + 6]) #+ (65536 * $Array[$RowStart + 5]) + (16777216 * $Array[$RowStart + 4])
         }
         $ExifMeaning = $null
         write-verbose -Message ("TagID {0:x4}" -f $exifTagID)

         switch ($ExifTypeID) {  #Different processing for each data type
                 1 {$ExifTypeName = "Byte"
                     if ($ExifDataLength -gt 4) {
                        if ($LittleEndian) {$ExifOffset = $Array[$RowStart + 8] + (256      * $Array[$RowStart + 9]) + (65536 * $Array[$RowStart +10]) + (16777216 * $Array[$RowStart +11]) }
                        else               {$ExifOffset = $Array[$RowStart +11] + (256      * $Array[$RowStart +10]) + (65536 * $Array[$RowStart + 9]) + (16777216 * $Array[$RowStart + 8]) }
                        $Stream.Position                = ($Dataoffset + $ExifOffset)
                        $ArrayPos                       = $Stream.Position % 64KB
                        $BytesRead                      = $Stream.Read($Array,$ArrayPos,$ExifDataLength)
                        $ExifValue                      = $Array[$ArrayPos..($ArrayPos + $ExifDataLength -1)]
                      }
                      else  {               $ExifValue  = $Array[($RowStart + 8)..($RowStart + 7 + $ExifDataLength)]
                                            $ExifOffset = $null
                      }
                 }
                 2 {$ExifTypeName                     = "String[$ExifDataLength]"
                      if ($ExifDataLength -gt 4) {
                        if ($LittleEndian) {$ExifOffset = $Array[$RowStart + 8] + (256      * $Array[$RowStart + 9]) + (65536 * $Array[$RowStart +10]) + (16777216 * $Array[$RowStart +11]) }
                        else               {$ExifOffset = $Array[$RowStart +11] + (256      * $Array[$RowStart +10]) + (65536 * $Array[$RowStart + 9]) + (16777216 * $Array[$RowStart + 8]) }
                        if ($makerNoteOffset)    {
                                  $Stream.Position      = $exifoffset + $imgdirstart -10  + $Dataoffset
                           }
                        else {    $Stream.Position      = $ExifOffset + $Dataoffset }
                        $ArrayPos                       = $Stream.Position % 64KB
                        $BytesRead                      = $Stream.Read($Array,$ArrayPos,$ExifDataLength)
                        $ExifValue                      = $UTF7Encoding.GetString($Array[$ArrayPos..($ArrayPos + $ExifDataLength -1)]).Trim([char]0) # ([char[]]$Array[$ArrayPos..($ArrayPos + $ExifDataLength -1)] -join "").trim([char]255).Trim([char]0)
                      }
                      else  {               $ExifValue  = $UTF7Encoding.GetString($Array[($RowStart + 8)..($RowStart + 7 + $ExifDataLength)]).Trim([char]0) #  ([char[]]$Array[($RowStart + 8)..($RowStart + 7 + $ExifDataLength)] -join "").Trim([char]0)
                                            $ExifOffset = $null
                      }
                 }
                 3 {$ExifTypeName   = "Word"
                      if ($LittleEndian)  {$s1          = $Array[$RowStart + 8] + (256 * $Array[$RowStart + 9]) ; $s2 =    ($Array[$RowStart +10]) + (256      * $Array[$RowStart +11])}
                      else                {$s1          = $Array[$RowStart + 9] + (256 * $Array[$RowStart + 8]) ; $s2 =    ($Array[$RowStart +11]) + (256      * $Array[$RowStart +10])}
                      if     ($ExifDataLength -gt 200) { Write-verbose -Message "Skipping because it is an array of $ExifDataLength"}
                      elseif ($ExifDataLength -gt 2)   {
                        $ExifValue                      = @()
                        if ($LittleEndian) {$ExifOffset = $s1 + (65536 * $s2)}
                        else               {$ExifOffset = $s2 + (65536 * $s1)}
                        if ($makerNoteOffset)    {
                                  $Stream.Position      = $exifoffset + $imgdirstart -10  + $Dataoffset
                           }
                        else {    $Stream.Position      = $ExifOffset + $Dataoffset }
                        $ArrayPos                       = $Stream.Position % 64KB
                        $BytesRead                      = $Stream.Read($Array,$ArrayPos,2 * $ExifDataLength)
                        if ($LittleEndian) {for ($i=0 ; $i -lt $exifdatalength ; $i ++ ) {
                                            $exifValue += $Array[$Arraypos + $i *2    ] +256 * $Array[$Arraypos + $i *2 + 1] }}
                         Else               {for ($i=0 ; $i -lt $exifdatalength ; $i ++ ) {
                                            $exifValue += $Array[$Arraypos + $i *2 +1 ] +256 * $Array[$Arraypos + $i *2]     }}
                      }
                      elseif ($ExifDataLength -eq 2) {
                        $ExifValue                      = @($s1, $s2)
                        $ExifOffset                     = $null
                      }
                      elseif ($ExifDataLength -eq 1) {
                        $ExifValue                      = $s1
                        $ExifOffset                     = $null
                     }
                 }
                 4 {$ExifTypeName   = "DWord"
                      if ($LittleEndian)  {$long        = $Array[$RowStart + 8] + (256      * $Array[$RowStart + 9]) + (65536 * $Array[$RowStart +10]) + (16777216 * $Array[$RowStart +11]) }
                      else                {$long        = $Array[$RowStart +11] + (256      * $Array[$RowStart +10]) + (65536 * $Array[$RowStart + 9]) + (16777216 * $Array[$RowStart + 8]) }
                      if ($ExifDataLength -gt 1) {
                                           $ExifOffset  = $long ;  $ExifValue = @()
                      for ($ItemNo = 0; $ItemNo -lt $ExifDataLength; $itemNo ++) {   ###############  for maker note this looks like it should be $imgdirstart + 2 + exifoffset
                           if ($makerNoteOffset)    {
                                  $Stream.Position      = ($ItemNo * 4) + $exifoffset + $imgdirstart -10  + $Dataoffset
                           }
                           else { $Stream.Position      = ($ItemNo * 4) + $ExifOffset + $Dataoffset }
                           $ArrayPos                    = $Stream.Position % 64KB
                           $BytesRead                   = $Stream.Read($Array,$ArrayPos,4)
                           if ($LittleEndian)  {$long   = $Array[$ArrayPos    ] + (256      * $Array[$ArrayPos + 1]) + (65536 * $Array[$ArrayPos + 2]) + (16777216 * $Array[$ArrayPos + 3]) }
                           else                {$long   = $Array[$ArrayPos + 3] + (256      * $Array[$ArrayPos + 2]) + (65536 * $Array[$ArrayPos + 1]) + (16777216 * $Array[$ArrayPos    ]) }

                           $ExifValue                  += $long
                           }
                      }
                      else                {$ExifOffset  = $null ;  $ExifValue     = $long }
                 }
                 5 {$ExifTypeName   = "Rational"
                      if ($LittleEndian)  {$ExifOffset  = $Array[$RowStart + 8] + (256      * $Array[$RowStart + 9]) + (65536 * $Array[$RowStart +10]) + (16777216 * $Array[$RowStart +11]) }
                      else                {$ExifOffset  = $Array[$RowStart +11] + (256      * $Array[$RowStart +10]) + (65536 * $Array[$RowStart + 9]) + (16777216 * $Array[$RowStart + 8]) }
                      $ExifValue = @()
                      for ($ItemNo = 0; $ItemNo -lt $ExifDataLength; $itemNo ++) {
                           $Stream.Position             = ($Dataoffset + $ExifOffset) + $ItemNo * 8;
                           $ArrayPos                    = $Stream.Position % 64KB
                           $BytesRead                   = $Stream.Read($Array,$ArrayPos,8)
                           if ($LittleEndian)  {$long1  = $Array[$ArrayPos    ] + (256      * $Array[$ArrayPos + 1]) + (65536 * $Array[$ArrayPos + 2]) + (16777216 * $Array[$ArrayPos + 3])
                                                $long2  = $Array[$ArrayPos + 4] + (256      * $Array[$ArrayPos + 5]) + (65536 * $Array[$ArrayPos + 6]) + (16777216 * $Array[$ArrayPos + 7]) }
                           else                {$long1  = $Array[$ArrayPos + 3] + (256      * $Array[$ArrayPos + 2]) + (65536 * $Array[$ArrayPos + 1]) + (16777216 * $Array[$ArrayPos    ])
                                                $long2  = $Array[$ArrayPos + 7] + (256      * $Array[$ArrayPos + 6]) + (65536 * $Array[$ArrayPos + 5]) + (16777216 * $Array[$ArrayPos + 4]) }
                           if  ($long1 -eq 0 -and $long2 -eq 0)
                                {$ExifValue            += 0                 }
                           else {$ExifValue            += [math]::Round(($long1 / $long2),6) }
                      }
                      if ($ExifValue.Length -eq 1) {
                        $ExifValue                      = $ExifValue[0]
                        #if ($long1 -eq 1) {$ExifMeaning = "1/$long2"}
                      }
                 }
                 6 {$ExifTypeName = "SignedByte"   #not applying twos compliment to process signs!
                     if ($ExifDataLength -gt 4) {
                        if ($LittleEndian) {$ExifOffset = $Array[$RowStart + 8] + (256      * $Array[$RowStart + 9]) + (65536 * $Array[$RowStart +10]) + (16777216 * $Array[$RowStart +11]) }
                        else               {$ExifOffset = $Array[$RowStart +11] + (256      * $Array[$RowStart +10]) + (65536 * $Array[$RowStart + 9]) + (16777216 * $Array[$RowStart + 8]) }
                        $Stream.Position                = ($Dataoffset + $ExifOffset)
                        $ArrayPos                       = $Stream.Position % 64KB
                        $BytesRead                      = $Stream.Read($Array,$ArrayPos,$ExifDataLength)
                        $ExifValue                      = @()
                        foreach        (  $b    in        $Array[$ArrayPos..($ArrayPos  + $ExifDataLength -1)] ) {
                                                if ($b -gt 0x7f) {$b = -(0xff -bxor ($b - 1))}
                                                $ExifValue += $b
                        }
                      }
                      else  {               $ExifValue  = @()
                                            foreach ($b in $Array[($RowStart + 8)..($RowStart + 7 + $ExifDataLength)]) {
                                                 if ($b -gt 0x7f) {$b = -(0xff -bxor ($b - 1))}
                                                 $ExifValue += $b
                                            }
                                            $ExifOffset = $null
                      }
                      if ($ExifValue.Length -eq 1) {$ExifValue = $ExifValue[0]  }
                  }
                 7 {$ExifTypeName = "Undefined"
                     if ($ExifDataLength -gt 4) {
                        if ($LittleEndian) {$ExifOffset = $Array[$RowStart + 8] + (256      * $Array[$RowStart + 9]) + (65536 * $Array[$RowStart +10]) + (16777216 * $Array[$RowStart +11]) }
                        else               {$ExifOffset = $Array[$RowStart +11] + (256      * $Array[$RowStart +10]) + (65536 * $Array[$RowStart + 9]) + (16777216 * $Array[$RowStart + 8]) }
                        if ($makerNoteOffset)    {
                                  $Stream.Position      = $exifoffset + $imgdirstart -10  + $Dataoffset
                           }
                        else {    $Stream.Position      = $ExifOffset + $Dataoffset }
                        $ArrayPos                       = $Stream.Position % 64KB
                        $BytesRead                      = $Stream.Read($Array,$ArrayPos,$ExifDataLength)
                        $ExifValue                      = $Array[$ArrayPos..($ArrayPos + $ExifDataLength -1)]
                      }
                      else  {               $ExifValue  = $Array[($RowStart + 8)..($RowStart + 7 + $ExifDataLength)]
                                            $ExifOffset = $null
                      }
                      if ($ExifValue.Length -eq 1) {$ExifValue = $ExifValue[0]  }
                 }
                 8 {$ExifTypeName   = "Signed Short"
                      if ($LittleEndian)  {$s1          = $Array[$RowStart + 8] + (256 * $Array[$RowStart + 9]) ; $s2 =    ($Array[$RowStart +10]) + (256      * $Array[$RowStart +11])}
                      else                {$s1          = $Array[$RowStart + 9] + (256 * $Array[$RowStart + 8]) ; $s2 =    ($Array[$RowStart +11]) + (256      * $Array[$RowStart +10])}
                      if ($ExifDataLength -gt 2) {
                        if ($LittleEndian) {$ExifOffset = $s1 + (65536 * $s2)}
                        else               {$ExifOffset = $s2 + (65536 * $s1)}
    ############ Not reading the array !!!
                        $ExifValue                      = $null
                      }
                      if ($s1 -gt 0x7fff) {$s1 =  -(0xffff -bxor ($s1-1)) }
                      if ($s2 -gt 0x7fff) {$s2 =  -(0xffff -bxor ($s2-1)) }
                      if ($ExifDataLength -eq 2) {
                        $ExifValue                      = @($s1, $s2)
                        $ExifOffset                     = $null
                      }
                      elseif ($ExifDataLength -eq 1) {
                        $ExifValue                      = $s1
                        $ExifOffset                     = $null
                      }
                 }
                 9 {$ExifTypeName   = "Signed Long"
                      if ($LittleEndian)  {$Long        = $Array[$RowStart + 8 ] + (256 * $Array[$RowStart + 9 ]) + (65536 * $Array[$RowStart +10]) + (16777216 * $Array[$RowStart +11])}
                      else                {$Long        = $Array[$RowStart + 11] + (256 * $Array[$RowStart + 10]) + (65535 * $Array[$RowStart +9 ]) + (16777216 * $Array[$RowStart +8])}
                      if ($ExifDataLength -gt 1) {
                        if ($LittleEndian) {$ExifOffset = $long}
                        else               {$ExifOffset = $long}
    ############ Not reading the array !!!
                        $ExifValue                      = $null
                      }
                      if ($long -gt 0x7fffffff)  {$long =  -(0xffffffff -bxor ($long-1)) }
                      $ExifValue                        = $Long
                      $ExifOffset                       = $null

                 }
                10 {$ExifTypeName   = "SignedRational"
                      if ($LittleEndian)  {$ExifOffset          = $Array[$RowStart + 8] + (256      * $Array[$RowStart + 9]) + (65536 * $Array[$RowStart +10]) + (16777216 * $Array[$RowStart +11]) }
                      else                {$ExifOffset          = $Array[$RowStart +11] + (256      * $Array[$RowStart +10]) + (65536 * $Array[$RowStart + 9]) + (16777216 * $Array[$RowStart + 8]) }
                      $ExifValue =@()
                      for ($ItemNo = 0; $ItemNo -lt $ExifDataLength; $itemNo ++) {
                           $Stream.Position                     = ($Dataoffset + $ExifOffset) + $ItemNo * 8;
                           $ArrayPos                            = $Stream.Position % 64KB
                           $BytesRead                           = $Stream.Read($Array,$ArrayPos,8)
                           #negative numbers use 2's compliment so first make sure the number is an unsigned int 32 ....
                           if ($LittleEndian)  {[uint32]$long1  = $Array[$ArrayPos    ] + (256      * $Array[$ArrayPos + 1]) + (65536 * $Array[$ArrayPos + 2]) + (16777216 * $Array[$ArrayPos + 3])
                                                [uint32]$long2  = $Array[$ArrayPos + 4] + (256      * $Array[$ArrayPos + 5]) + (65536 * $Array[$ArrayPos + 6]) + (16777216 * $Array[$ArrayPos + 7]) }
                           else                {[uint32]$long1  = $Array[$ArrayPos + 3] + (256      * $Array[$ArrayPos + 2]) + (65536 * $Array[$ArrayPos + 1]) + (16777216 * $Array[$ArrayPos    ])
                                                [uint32]$long2  = $Array[$ArrayPos + 7] + (256      * $Array[$ArrayPos + 6]) + (65536 * $Array[$ArrayPos + 5]) + (16777216 * $Array[$ArrayPos + 4]) }
                           #If numerator is bigger than maxint/2   Subtract 1, flip all bits, put a Minus sign in front.
                           if ($long1 -gt 0x7fffffff)
                                                   {[int]$long1 = -(0xffffffff -bxor ($long1 -1) )}
                           if  ($long1 -eq 0 -or $long2 -eq 0) # avoid division by zero error, return 0 for both 0 (or either)
                                {$ExifValue                    += 0                 }
                           else {$ExifValue                    += [math]::Round(($long1 / $long2),6) }

                      }
                      if ($ExifValue.Length -eq 1) {$ExifValue = $ExifValue[0]
                                                  #  if ($long1 -eq 1) {
                                                  #  $ExifMeaning = "$long1/$long2"}
                      }
                  }
         }
         $Results[$ExifTagID] = New-Object -TypeName psobject -Property @{TagID=$ExifTagID.ToString("x4"); TagName=$TagNames[[object]$ExifTagID];  TypeID=$ExifTypeID; length=$ExifDataLength; offset=$ExifOffset;    TypeName=$ExifTypeName; Value=$ExifValue ; meaning=$ExifMeaning}
    }
    $Results
}

function Read-EXIF         {
    <#
      .Synopsis
       Gets Metadata from JPEG, PSD, TIFF, DNG and compatible RAW image files
      .Description
       Read-EXIF extracts the meta data from Image files. Not all of this data is "EXIF data"
       In the case of JPEG files there can be multiple data segments; the function will read some information from
       non-EXIF segments; and can dump these segments as HEX or XML for further investigation.
       It understands some Pentax maker note data.
      .Example
        Read-Exif img64525.jpg
        Returns  the exif data for the specified file
      .Example
        Read-Exif *.jpg | where {$_.JPEGSegements -match "ICC_PROFILE"} | ft -a Fullname,Software
        Finds files which have an embedded colour profile and outputs a table of file name and software which produced it.
      .Example
        Read-Exif *.dng,*.tif | Group-Object -NoElement -Property FocalLength
        Outputs a list of focal lengths used, with a count of each. [Note it takes about 0.3 seconds per file]
      .Example
        Read-Exif $file -DumpAsXML  "http://ns.adobe.com/xap/1.0/"
        Dumps the XAP data segment as XML; allowing, for example, the position of a facial recognition rectangle to be discovered
      .Example
        Read-EXIF IMG75771.DNG -Raw | Out-GridView
        Displays all the discovered tags with their raw data and name (if known) and interpreted data (where interpretation is understood)
    #>
    [cmdletbinding()]
    [OutputType([Exif])]
    param   (
        #Path To the file
        [parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true )]
        [alias("FullPath","FullName")]
        [string]$Path ,
        #Dump raw data with tag ID (with interpretted name),length Value (with interpretted meaning)
        [switch]$Raw,
        #Dump a matching EXIF segment as XML  (the Value is treated as a regular expression)
        [String]$DumpAsXML,
        #Dump a Matchig EXIF segment as HEX (the Value is treated as a regular expression)
        [String]$DumpAsHex,
        #If an adobe XAP Segment is present, add information from it to the output
        [switch]$XAPDetails
    )
    begin   {
        [byte[]]$Array            = ,0 * 128KB
        $UTF8Encoding             = New-object -TypeName System.Text.UTF8Encoding
        $UnicodeEncoding          = New-object -TypeName System.Text.UnicodeEncoding
    }
    process {
        $Path                      = Resolve-Path -Path $Path -ErrorAction SilentlyContinue
        if (-not $Path) {
            Write-Warning "'$($PSBoundParameters.path)' did not resolve to any files."
            return
        }
        if ($Path.count -gt 1)   {
            $PSBoundParameters.Remove("Path") | Out-Null
            $Path | Read-Exif @PSBoundParameters
            return
        }
        Write-Verbose -Message "Reading EXIF for $Path"
        #Open a stream and read the first few bytes.
        $Dataoffset               = 0
        $XMPPos     =  $XMPLength = 0
        try {
            $File                 = Get-Item $Path
            $Stream               = $file.OpenRead()
            $PropHash             = [ordered]@{"Fullname" = $File.FullName; "Path"          = $File.FullName
                                               "Length"   = $File.Length  ; "LastWriteTime" = $File.LastWriteTime
            }
        }
        catch {
            Write-Warning "Error trying to read '$Path'."
            return
        }
        $BytesRead                = $Stream.Read($Array,0 ,50)
        if ($BytesRead -lt 50)    { Write-Warning -Message "$Path seems impossibly small - could only read $BytesRead bytes - skipping"; return}
    #region find blocks in JPEGs or PhotoShop files: we should emerge from here knowing where the EXIF is (or exit having dumped a JPG segment as hex or XML)
        if     ([string]$Array[0..1]  -eq "255 216") {
             #If the file is a JPG bytes 0&1 will be the marker 0xFFD8. Bytes 2 & 3should be the marker for an App block.
             #Bytes 4&5 hold the size of the block, then we hope to find the the signature Exif. We may need to look at more than one block to find it
             #Inside the block we should find the same stuff as the start of a TIFF - which is found below
             Write-Verbose -Message ("Found JPEG Marker (0xFFD8) in $Path")
             $SegStart            = $ArrayPos = 2 #JPG marker in bytes 0 &1, Data segment starts at byte 2
             $Segments            = @()
             $SegResults          = @()
             $SegID               = (($Array[$ArrayPos] * 256) + $Array[($ArrayPos+1)]).ToString("x4")
             do {  #find and save offset to segment from start of file, length ID, name  , then jump forward by length and if we have a valid segnment ID there, repeat the process
                 $SegLength       = (($Array[$ArrayPos+2] * 256) + $Array[($ArrayPos+3)])
                 $SegName         = ""
                 $i               = $ArrayPos + 4
                 while ($Array[$i] -gt 1 -and $i -lt ($ArrayPos + 70) -and $i -lt ($ArrayPos + 2 + $SegLength))
                     {$SegName   += [char]$Array[$i]; $i++}
                 Write-Verbose -Message ("   Marker: 0x$SegID, @ Offset: 0x{0:x4} ({0}), Identifier: $SegName, Length: 0x{1:x4} ({1})" -f $SegStart,$SegLength  )
                 if ($SegName -eq "http://ns.adobe.com/xap/1.0/") {
                    $XMPPos       = $SegStart  + 31
                    $XMPLength    = $SegLength - 31
                 }
                 $Segments       += "$SegID $SegName $($SegStart)"
                 $SegResults     += New-Object -TypeName psobject -Property @{TagID=$SegID; TagName=$SegName;  TypeID=$null;  Length=$SegLength; Offset=$SegStart;
                                                                             TypeName="JPEG Segment";           Value=$null; Meaning=$null}
                 $SegStart        = $SegStart + 2 + $SegLength
                 $Stream.Position = $SegStart
                 $ArrayPos        = $SegStart % 64kb
                 $BytesRead       = $Stream.Read($Array,$ArrayPos ,100)
                 $SegID           = (($Array[$ArrayPos] * 256) + $Array[($ArrayPos+1)]).ToString("x4")
             }
             until  ($Segid -notmatch "ff[fe]\w")
             if     ($Raw)       {$SegResults}
             elseif ($DumpAsXML) {
                 $SegResults | Where-Object {$_.tagname -match $DumpAsXML} | ForEach-Object {
                     $Stream.Position   = $_.Offset
                     $ArrayPos          = $_.Offset % 64kb
                     $BytesRead         = $Stream.Read($Array,$ArrayPos ,($_.length +2))
                     $XMLtext           = $UTF8Encoding.GetString($Array[($ArrayPos+5+$_.TagName.length)..($ArrayPos + $_.Length+1)])
                     $XMLtext           = $XMLtext.trim([char]0)   -replace "(?s)^\s*<\?.*?\?>\s*<","<"  -replace "(?s)>\s+<\?.*?\?\>\s*$",">"
                     $Sw                = New-Object -TypeName System.IO.StringWriter
                     $Writer            = New-Object -TypeName System.Xml.XmlTextWriter -ArgumentList $Sw -Property @{Formatting = [System.xml.formatting]::Indented }
                     $global:ExifXML    = $null  #So if the $XMLText is not valid XML we get nothing.
                     $global:ExifXML    = [xml]$XMLtext
                     foreach ($Doc in [xml]$exifxml) {
                         $Doc.WriteContentTo($Writer)
                         $Sw.ToString()
                    }
                }
             return
             }
             elseif ($DumpAsHex) {
                $SegResults | Where-Object {$_.tagname -match $DumpAsHex} | ForEach-Object {
                     $Stream.Position    = ($_.offset -($_.offset %32) )
                     $ArrayPos           = $Stream.Position % 64kb
                     $Width              = 32
                     $CountSoFar         = $Stream.Position
                     $BytesRead          = $Stream.Read($Array,$ArrayPos ,$Width)
                     while ($CountSoFar -le ($_.offset + $_.length + $Width) ) {
                      $CountSoFar.ToString("x4") + ":" + ($CountSoFar + $Width -1 ).ToString("x4") + "  " +
                         ((($Array[ $ArrayPos            ..($ArrayPos + ($Width/2)-1)] | ForEach-Object {$_.tostring("x2")}) -join " ") + "-" +
                          (($Array[($ArrayPos+($Width/2))..($ArrayPos +  $Width   -1)] | ForEach-Object {$_.tostring("x2")}) -join " ") ).padright($Width * 3 + 4," ")   +
                          (($Array[ $ArrayPos            ..($ArrayPos +  $Width   -1)] | Foreach-Object { if ($_ -ge 32 -and $_ -le 127 ) { [char]$_ } else { "." }} )  -join  '' )
                      $CountSoFar       += $Width
                      $Arraypos         += $Width
                      $BytesRead         = $Stream.Read($Array,$ArrayPos ,$Width)
                  } }
             return
             }
             #Add the list of Segments found to the propery hash. If none is named EXIF, bail out, but try to show any XAP data first.
             $PropHash["JPEGSegements"] = ($Segresults.tagname -join ", ")
             if   (-not( $Segments -match "ffe1.*exif" ))  {
                Write-Warning -Message ("{0} Does not appear to contain any EXIF data." -f $Path)
                 if ($XMPPos -and $XMPLength)   {
                     Read-XAPData   -Stream $Stream -Array $Array -PropHash $PropHash -raw:$raw -XMPPos $XMPPos -XMPLength $XMPLength
                     if (-not $Raw) {New-Object -TypeName pscustomobject -Property $PropHash}
                 }
                 return
             }

             $Stream.Position          = $Dataoffset = ($SegResults | Where-Object {$_.tagID -eq "ffe1" -and $_.tagname -eq "Exif"} |
                                             Microsoft.PowerShell.Utility\Select-Object -First 1 -ExpandProperty offset) + 10 #10 bytes is 2 for segment ID. 2 for segment length, 4 for name 'EXIF' + 2 for 00-00 marking End of header. Next two bytes indicate big endian or little endian
             $BytesRead                = $Stream.Read($Array,0 ,16)
        }
        elseif ([string]$Array[0..11] -eq "56 66 80 83 0 1 0 0 0 0 0 0") {
             Write-Verbose -Message ("Found Photoshop Marker (0x8BPS01000000) in $Path")
             $psColorDataLength   = (16777216 * $Array[26]) + (65536 * $Array[27]) + (256 * $Array[28]) + $Array[29]
             $SegStart            = $ArrayPos = $Stream.Position = 30 + $psColorDataLength
             $BytesRead           = $Stream.Read($Array,$ArrayPos,20)
             $SegLength           = (16777216 * $Array[$ArrayPos]) + (65536 * $Array[$ArrayPos+1]) + (256 * $Array[$ArrayPos+2]) + $Array[$ArrayPos+3]
             $ArrayPos            = $Stream.Position = 30 + $psColorDataLength + 4
             $SegEnd              = $ArrayPos + $SegLength
             $BytesRead           = $Stream.Read($Array,$ArrayPos,($SegLength) )
             while ($ArrayPos -lt $Segend -and [char[]]$Array[($ArrayPos)..($ArrayPos+3)] -join ""   -eq '8BIM') {
                    if ([string]$Array[($ArrayPos+4)..($ArrayPos+5)] -eq "4 4" ) { $IPTCPos    =  $ArrayPos + 12   }
                    if ([string]$Array[($ArrayPos+4)..($ArrayPos+5)] -eq "4 34") { $ExifPos    =  $ArrayPos + 12   }
                    if ([string]$Array[($ArrayPos+4)..($ArrayPos+5)] -eq "4 36") { $XMPPos     =  $ArrayPos + 10
                                                                                   $XMPLength  = ($Array[$XMPPos] * 256 + $Array[$XMPPos+1] ) }
                    $ArrayPos += ($Array[$ArrayPos+10] * 256 + $Array[$ArrayPos+11] +12 + $(if ($Array[$ArrayPos+11] % 2) {1}  ) )
             }
             $Stream.Position     = $ExifPos
             $BytesRead           = $Stream.Read($Array,0 ,16)
             $Dataoffset          = $ExifPos
        }
    #endregion
    <#  We either have moved up a JPEG or PSD file to what we think is a block of EXIF data
        Or if the file is a TIF, or TIF derrived (DNG, PEF etc) the file STARTS with the EXIF data block and we should have skipped over stuff above.
        The block goes like this: the first 2 bytes  are either "MM" - Motorola, big endian byte order,
                                                             or "II" - Intel, Little endian byte order
        The bytes 2 and 3 form a word using this endian-ness and the Value is always 42.
        The bytes 4,5,6,7 form a dword using this endian-ness which points to the start of 1st Image Directory
       (this usally comes straight after II/MM [42] [Pointer] i.e position 0008)
    #>

    #region Get data from EXIF, GPS and Interop sections
        if     ($Array[0] -eq $Array[1]  -and $Array[1] -eq 77) {
                    if ( (256*$Array[2]) +    $Array[3] -ne 42) {Write-Warning -Message "Data block signature not found, this may not be a valid file"}
                    else                                        {Write-Verbose -Message "Found Big Endian data block"}
                    $LittleEndian  = $false
                    $ImgDirStart   = (16777216 * $Array[4]) + (65536 * $Array[5]) + (256 * $Array[6]) + $Array[7]
        }
        elseif ($Array[0] -eq $Array[1]  -and $Array[1] -eq 73) {
                    if ( (256*$Array[3]) +    $Array[2] -ne 42) {Write-Warning -Message "Data block signature not found, This may not be a valid file"}
                    else {write-Verbose -Message "Found Little Endian data block"}
                    $LittleEndian  = $True  ;
                    $ImgDirStart   = (16777216 * $Array[7]) + (65536 * $Array[6]) + (256 * $Array[5]) + $Array[4]
        }
        else     {Write-Warning -Message "Could not find Endian tag in $Path, can't process"; Return }
        Write-Verbose -Message ("Reading data from offset {0:x4} ({0})" -f $ImgDirStart)
        $Results                   = Read-ExifFD -Stream $Stream -ImgDirStart $ImgDirStart -LittleEndian $LittleEndian -TagNames $ExifTagNames -Dataoffset $Dataoffset
        [void]$Results.Remove(0xEA1C)
        if ($Results[0x00fe].Value -eq 1 -and $Results[0x014a]) { #subfile type = Reduced resolution image & other image subfiles exist
            foreach ($subFile in $Results[0x014a].Value) {         #Go through the subfiles, attributes from any that aren't tagged reduced res take precedence
                $res2              = Read-ExifFD -Stream $Stream -ImgDirStart $Subfile               -LittleEndian $littleEndian -TagNames $ExifTagNames -dataoffset $Dataoffset
                if ($res2[0x00fe]  -ne 1) {$res2.keys | ForEach-Object {$Results[$_] = $res2[$_]}}
            }
        }
        #if the 1st image Directory has pointers to Exif,GPS and/or interop data process it
        if ( $Results[0x8769] ) { #ExifFD
                Write-Verbose -Message ("EXIF data continues at offset {0:x4} ({0})" -f $Results[0x8769].Value)
                $ExifResults       = Read-ExifFD -Stream $Stream -ImgDirStart $Results[0x8769].Value -LittleEndian $LittleEndian -TagNames $ExifTagNames -Dataoffset $Dataoffset
                $exifresults.Keys  | ForEach-Object {$Results[$_] = $ExifResults[$_] }
        }
        if ( $Results[0x8825] ) {#GpsFD
                Write-Verbose -Message ( "GPS data at offset {0:x4} ({0})" -f $Results[0x8825].Value)
                $GPSResults       = Read-ExifFD -Stream $Stream -ImgDirStart $Results[0x8825].Value -LittleEndian $LittleEndian -TagNames $ExifTagNames -Dataoffset $Dataoffset
                $GPSResults.Keys  | ForEach-Object {$Results[$_] = $GPSResults[$_]   }
        }
        if ( $Results[0xA005] ) {#InterOpID
                Write-Verbose -Message ("Interop data at offset {0:x4} ({0})" -f $Results[0xA005].Value)
                $InterOPResults = Read-ExifFD -Stream $Stream -ImgDirStart $Results[0xA005].Value -LittleEndian $LittleEndian -TagNames $InterOpTagNames -Dataoffset $Dataoffset
                $InterOPResults.Keys | ForEach-Object {$Results["InterOp:$_"] = $InterOPResults[$_]   }
        }
    #endregion
    #region Process returned data something easier to read
        #Translate enumeration fields to their text
        $ExifLookUps.Keys   | ForEach-Object { if ($Results[$_]) {$Results[$_].Meaning= $ExifLookUps[$_][[int]$Results[$_].Value]}}
        #Munge various fields to their meanings stating with Dates as dates
        $ExifDateStrings    | ForEach-Object { if ($Results[$_].Value -match "^\d{4}:\d\d:\d\d\s\d\d:\d\d:\d\d$" -and $Results[$_].Value -ne "0000:00:00 00:00:00") {
                $Results[$_].Meaning =([DateTime]::ParseExact($Results[$_].Value ,"yyyy:MM:dd HH:mm:ss",[System.Globalization.CultureInfo]::InvariantCulture) )
        } }
        $ExifUTF8Strings    | ForEach-Object { if ($Results[$_] -and $Results[$_].TypeID -notin @(2,4)) {
                $Results[$_].meaning =   $UTF8Encoding.GetString($Results[$_].Value).trim([char]0)
            if ($Results[$_].meaning -match "^UNICODE" -and $Results[$_].Value[8]) {
                $Results[$_].meaning = $UnicodeEncoding.GetString($Results[$_].Value[8..($Results[$_].Value.count-1)]).trim([char]0)}
        elseif ($Results[$_].meaning -match "^UNICODE") {
                $Results[$_].meaning      = $UnicodeEncoding.GetString($Results[$_].Value[9..($Results[$_].Value.count-2)]).trim([char]0)} }
        }
        $ExifUnicodeStrings | ForEach-Object { if ($Results[$_] -and $Results[$_].TypeID -notin @(2,4)) {
                $Results[$_].meaning      = $UnicodeEncoding.GetString($Results[$_].Value).trim([char]0)}
        }
        if ($Results[0x9c9e] -and $Results[0x9c9e].meaning -like "*;*") { $Results[0x9c9e].meaning =  $Results[0x9c9e].meaning -split "\s*;\s*"}
        if ($Results[0x9209]) {
            $myvar =[int]$Results[0x9209].Value
            $flash=""
            If (    $myvar –band 1)           {$Flash += "Flash fired"}
            If (    $myvar -bAnd 4)           {$Flash += ", return"
               If  ($myvar -bAnd 2)           {$flash += " detected"}
                else                          {$Flash += "not detected"}
            }
            If     ($myvar -bAnd 8) {
               If  (($myvar-bAnd 16) -eq 16)  {$Flash =  "Flash Auto, " + $flash}
              else                            {$flash =  "Flash on, "   + $flash} }
            ElseIf ($myvar -bAnd 16)          {$flash =  "Flash off " }
            If     ($myvar -bAnd 32)          {$flash =  "No Flash function"    }
            If     ($myvar -bAnd 64)          {$flash += ", Red Eye reduction"  }
            $Results[0x9209].Meaning = $flash
        }
        if ($Results[0x9000] -and $Results[0x9000].typeID -eq 7)   { $Results[0x9000].meaning =   $UTF8Encoding.GetString($Results[0x9000].Value)}
        if ($Results[0x9201] -and $Results[0x9201].Value  -ge 1)   { $Results[0x9201].meaning = "{0:N3}Tv (1/{1:N0} Sec)" -f $Results[0x9201].Value,   [math]::Pow(2,$Results[0x9201].Value)}
        if ($Results[0x9201] -and $Results[0x9201].Value  -lt 1)   { $Results[0x9201].meaning = "{0:N3}Tv ({1:N1} Sec)"   -f $Results[0x9201].Value,(1/[math]::Pow(2,$Results[0x9201].Value))}
        if ($Results[0x9202] )                                     { $Results[0x9202].meaning = "{0:N3}Av (f/{1:N1})"     -f $Results[0x9202].Value,[math]::Sqrt([math]::Pow(2,$Results[0x9202].Value)) }
        if ($Results[0x0000])    {$Results[0x0000].meaning =         $Results[0x0000].Value   -join "" }
        if ($Results[0xc612] )   {$Results[0xc612].meaning  =        $Results[0xc612].Value   -join "" }
        if (-not($Results[0x9286]) -and ($Results[0x9c9c])  )      { $Results[0x9c9c].Tagname = "Comment" ; $Results[0x9286] =$Results[0x9c9c]}

        #IF NOT 35MM EQUIV BUT FOCAL LENGTH (920A) AND PIXELS (A002 (x) , A003(Y) AND PIXELS PER INCH, a20e  (X) a20f (Y)  A210 Units (2=inch,3=cm,4=mm) CALCULATE 35MM EQUIV
         if (-not $Results[0xA405].Value -and $Results[0x920a].Value -and $Results[0xa002].Value -and $Results[0xa003].Value -and`
                  $Results[0xa210].Value -and $Results[0xa20f].Value -and $Results[0xa20e].Value      )  {
             $WidthMM         = $Results[0xa002].Value / $Results[0xa20e].Value * @{2=25.4;3=10;4=1}[$Results[0xa210].Value]
             $heightMM        = $Results[0xa003].Value / $Results[0xa20f].Value * @{2=25.4;3=10;4=1}[$Results[0xa210].Value]
             $Fl35            = [math]::round($Results[0x920a].Value * 43.266 / [math]::Sqrt(($heightmm * $heightmm) + ($Widthmm * $Widthmm))  ,0)
             $Results[0xA405] = New-Object -TypeName psobject -Property @{TagID="A405"; TagName="FocalLengthEquivIn35mmFormat";  TypeID=3; length=1; offset=0;    TypeName="Word"; Value=$Fl35 }
        }
        #Turn the GPS position fields into a single string
        if ($Results[1]           -and $Results[2] -and $Results[3] -and $Results[4] ) {
            $gps                   =   ""
            $l                     =   $Results[2].Value
            if     ($l[2]) {$gps  +=  "$($l[0])°$($l[1])'$($l[2])"""}
            elseif ($l[1]) {$gps  +=  "$($l[0])°$($l[1])'" }
            else           {$gps  +=  "$($l[0])°" }
            $gps                  +=   $Results[1].Value + ", "
            $l                     =   $Results[4].Value
            if     ($l[2]) {$gps  +=  "$($l[0])°$($l[1])'$($l[2])"""}
            elseif ($l[1]) {$gps  +=  "$($l[0])°$($l[1])'" }
            else           {$gps  +=  "$($l[0])°" }
            $gps                  +=   $Results[3].Value
            if ($Results[5]       -and $Results[6] ) {
                if ($Results[5].Value[0] -EQ 0) {$gps += ", $($Results[6].Value)M above Sea Level"}
                if ($Results[5].Value[0] -EQ 1) {$gps += ", $($Results[6].Value)M below Sea Level"}
            }

            if ($Results[0x10]       -and $Results[0x11] ) {
                if ($Results[0x10].Value[0] -EQ "M") {$gps += ", Bearing $($Results[0x11].Value)° Magnetic"}
                if ($Results[0x10].Value[0] -EQ "T") {$gps += ", Bearing $($Results[0x11].Value)° True"}
            }

            if ($Results[0x12]) { $gps += ", $($Results[0x12].Value)"}


            $Results[9999] =  New-Object -TypeName psobject -Property @{TagID="9999"; TagName="GPSsummary";  TypeID=2; length=$gps.Length; offset=0;    TypeName="String"; Value=$gps ; meaning=$gps}

        }
    #endregion
        if ($Raw){$Results.Values }
        #Add properties to the hash table - tagnames collections are ordered, and so is the property hash, so output should have the order we want
        $Results[($ExifTagNames.keys + ( $InterOpTagNames.Keys | ForEach-Object {"Interop:$_"})  )] |
              Where-Object {$_.tagname} | ForEach-Object {if ($null -ne $_.meaning)       {$PropHash[$_.tagname] = $_.meaning}
                                                          else                            {$PropHash[$_.tagname] = $_.Value} }
        #Exif has some function duplication. If a common variation is absent and the other one is present, fill in the missing one.
        if ($PropHash["PixelXDimension"]  -and -not $PropHash["Width"]   )  {$PropHash["Width"]    = $PropHash["pixelxdimension"]}
        if ($PropHash["PixelYDimension"]  -and -not $PropHash["Height"]  )  {$PropHash["Height"]   = $PropHash["pixelydimension"]}
        if ($PropHash["CreatorTool"]      -and -not $PropHash["Software"])  {$PropHash["Software"] = $PropHash["CreatorTool"]}
        if ($PropHash["Artist"]           -and -not $PropHash["Author"]  )  {$PropHash["Author"]   = $PropHash["Artist"]}
        if ($PropHash["Author"]           -and -not $PropHash["Artist"]  )  {$PropHash["Artist"]   = $PropHash["Author"]}
        if ($PropHash["ImageDescription"] -and -not $PropHash["Subject"] )  {$PropHash["Subject"]  = $PropHash["ImageDescription"]}

    #region Process Maker note EXIF data. (Almost all Pentax :-)
        $makerNotePreamble = ""
        if     ($Results[0x927C] ) {$makerNotePreamble = [char[]]$Results[0x927C].Value[0..19] -join "" ; $MakerNoteOffset = $Results[0x927C].offset }
        elseif ($Results[0xc634] ) {$makerNotePreamble = [char[]]$Results[0xc634].Value[0..19] -join "" ; $MakerNoteOffset = $Results[0xc634].offset }

        if     ($makerNotePreamble -match "AOC" -or $makerNotePreamble -match "Pentax") {
            $makernoteresults = @{}
            if     ($makerNotePreamble.IndexOf("II") -gt 0) {
                Write-Verbose -Message ("Pentax maker note block @ {0} with little endian tag " -f $MakerNoteOffset.tostring("X4") )
                $makernoteresults = Read-ExifFD -Stream $Stream -ImgDirStart $( $MakerNoteOffset + $makerNotePreamble.IndexOf("II") + 2 ) -LittleEndian $true  -TagNames $PentaxTagNames -dataoffset $Dataoffset -makerNoteOffset
            }
            elseif ($makerNotePreamble.IndexOf("MM") -gt 0) {
                Write-Verbose -Message ("Pentax maker note block @ {0} with big endian tag " -f $MakerNoteOffset.tostring("X4") )
                $makernoteresults = Read-ExifFD -Stream $Stream -ImgDirStart $( $MakerNoteOffset + $makerNotePreamble.IndexOf("MM") + 2 ) -LittleEndian $false -TagNames $PentaxTagNames -dataoffset $Dataoffset -makerNoteOffset
            }
            else {
                Write-Verbose -Message ("Pentax maker note block @ {0} without endian tag. Making Aassumptions " -f $MakerNoteOffset.tostring("X4") )
                if ($LittleEndian) {$makernoteresults = Read-ExifFD -Stream $Stream -ImgDirStart $( $MakerNoteOffset + 6 + $Dataoffset) -LittleEndian $false -TagNames $PentaxTagNames -dataoffset 0 }
                else {$makernoteresults = Read-ExifFD -Stream $Stream -ImgDirStart $( $MakerNoteOffset + 6  ) -LittleEndian $false -TagNames $PentaxTagNames -dataoffset $Dataoffset }
            }

            if ($makernoteresults[0x0002].Value.count -eq 2) {$makernoteresults[0x0002].meaning = $makernoteresults[0x0002].Value -join " x "                    }
            if ($makernoteresults[0x0013].value -gt 0) {$makernoteresults[0x0013].meaning = "f/" + ($makernoteresults[0x0013].Value / 10)                   }
            if ($makernoteresults[0x0016]) {$makernoteresults[0x0016].meaning = $makernoteresults[0x0016].Value[0] / 10 - 5                      }
            if ($makernoteresults[0x001d].value -gt 0) {$makernoteresults[0x001d].meaning = "{0:#.#}mm" -f ($makernoteresults[0x001d].Value / 100)          }
            if ($makernoteresults[0x0027]) {$makernoteresults[0x0027].meaning = ($makernoteresults[0x0027].Value |ForEach-Object {255 - $_} ) -join "."}
            if ($makernoteresults[0x0028]) {$makernoteresults[0x0028].meaning = ($makernoteresults[0x0028].Value |ForEach-Object {255 - $_} ) -join "."}
            if ($makernoteresults[0x002d]) {$makernoteresults[0x002d].meaning = $makernoteresults[0x002d].Value / 1024                          }
            if ($makernoteresults[0x0035]) {$makernoteresults[0x0035].meaning = "{0:N2} x {1:N2} mm" -f ($makernoteresults[0x35].Value[0] / 500), ( $makernoteresults[0x0035].Value[1] / 500)}
            if ($makernoteresults[0x0041]) {$makernoteresults[0x0041].meaning = $makernoteresults[0x0041].Value + 0                            }
            if ($makernoteresults[0x0047]) {$makernoteresults[0x0047].meaning = $makernoteresults[0x0047].Value[0].ToString() + "°C"           }  #19
            if ($makernoteresults[0x0048]) {$makernoteresults[0x0048].meaning = [boolean]$makernoteresults[0x48].Value                         }
            if ($makernoteresults[0x0049]) {$makernoteresults[0x0049].meaning = [boolean]$makernoteresults[0x49].Value                         }
            if ($makernoteresults[0x007b]) {$makernoteresults[0x007b].meaning = $PentaxXProcess[[int]$makernoteresults[0x007b].Value[0]]       }
            if ($makernoteresults[0x0005]) {
                $makernoteresults[0x0005].meaning = $PentaxModel[$makernoteresults[0x0005].Value]
                if (-not  $makernoteresults[0x0005].meaning ) {    $makernoteresults[0x0005].meaning = "ID = " + $makernoteresults[0x0005].Value }
            } # Model ID
            if ($makernoteresults[0x0012]) {
                if ($makernoteresults[0x0012].Value -gt 50000) {$makernoteresults[0x0012].meaning = "{0:N1}sec " -f ($makernoteresults[0x0012].Value / 100000 ) }
                else {$makernoteresults[0x0012].meaning = "1/{0:#}sec " -f (100000 / $makernoteresults[0x0012].Value )}
            } # Exposure time
            if ($makernoteresults[0x000E]) {
                if (  $makernoteresults[0x0005].value -in @(77970,78400)) {
                    #K1
                    $t = $PentaxK1FocusPts[$makernoteresults[0x000E].value[0]]
                    if ($makernoteresults[0x000E].value[1] -eq 1 ) { $t += "Expanded area (S)" }
                    if ($makernoteresults[0x000E].value[1] -eq 3 ) { $t += "Expanded area (M)" }
                    if ($makernoteresults[0x000E].value[1] -eq 5 ) { $t += "Expanded area (L)" }
                    $makernoteresults[0x000E].meaning = $t
                }
                else {
                    $makernoteresults[0x000E].meaning = $PentaxFocus11Pts[$makernoteresults[0x000E].value[0]]
                }
            } # AF Point
            if ($makernoteresults[0x0245]) {
                [int[]]$s = $makernoteresults[0x0245].Value
                $numberOfPoints = $s[2]
                $focusPointBitmap = ($s[4..12] | ForEach-Object {[convert]::ToString($_, 2).padleft(8, "0")} ) -join ""
                $pointsSelected = @()
                $pointsInFocus = @()
                0..($numberOfPoints - 1) | ForEach-Object {
                    $point = $focusPointBitmap.substring((2 * $_),2)
                    if ($point -ne "00"  ) {$pointsSelected += ($_ + 1 )}
                    if ($point -like "1*") {$pointsInFocus += ($_ + 1 )}
                }
                if     ( $numberOfPoints -eq  $pointsSelected.count) {$t = "Focus points 1-$numberOfPoints, all selected"}
                elseif ( $numberOfPoints -and $pointsSelected      ) {$t = "Focus points 1-$numberOfPoints, selected: " + ($pointsSelected -join ', ')   }
                else   {$t = "" }
                if     ($t -and $pointsInFocus) {$t = $t + ", in focus: " + ($pointsInFocus -join ', ') + "."}
                elseif ($t)                     {$t = $t + ", none in focus."}
                if     ($t) {$makernoteresults[0x0245].meaning = $t}
            } # AF Point info
            if ($makernoteresults[0x0033]) {
                [int[]]$s = $makernoteresults[0x0033].Value
                $t = $PentaxPicMode[ (([int]$s[0] * 256) + [int]$s[1]) ]
                If ($S[2] -eq "0") {$makernoteresults[0x0033].meaning = $t + ": 1/2 EV Steps"}
                else {$makernoteresults[0x0033].meaning = $t + ": 1/3 EV Steps"}
            } # Picture mode
            if ($makernoteresults[0x0034]) {
                [int[]]$s = $makernoteresults[0x0034].Value
                if ($makernoteresults[0x0018].Value[0]) {$t = "Bracket Sequence "    }
                Elseif ($s[3])                          {$t = @{1 = "Multiple Exposure. "; 16 = "HDR. "       ; 32 = "HDR Strong 1. "   ; 48 = "HDR Strong 2. "; 64 = "HDR Strong 3. "; 224 = "HDR Auto. "}[($S[3])] }
                Else                                    {$t = @{0 = "Single-frame. "     ;  1 = "Continuous. ";  2 = "Continuous [Hi]. ";  3 = "Burst. "}[($S[0])]
                }
                $t = $t + @{0 = ""; 1 = "Self-timer, 12 sec. "; 2 = "Self-timer, 2 sec. "; 16 = "Mirror Lock-up. "  }[$S[1]]
                $t = $t + @{0 = ""; 1 = "Remote control, 3 sec delay. "; 2 = "Remote control."    ; 4 = "Remote Continuous Shooting. "}[$S[2]]
                $makernoteresults[0x0034].meaning = $t -replace "\.\s*$", ""
            } # Drive mode
            if ($makernoteresults[0x003f]) {
                $i = (256 * $makernoteresults[0x003f].Value[0] + $makernoteresults[0x003f].Value[1])
                if ($i) {
                    if ($PentaxLens.ContainsKey($i)) {    $makernoteresults[0x003f].meaning = $PentaxLens[$i]}
                    else                             {    $makernoteresults[0x003f].meaning = "ID= $i" }
                }
            } # Lens
            if ($makernoteresults[0x005c]) {
                [int[]]$s = $makernoteresults[0x005c].Value
                $t = "SR " + @{ 0 = "Off"         ; 1 = "On"     ;  4 = "Off (AA simulation off)"   ;  5 = "On but Disabled" ; 6 = "On (Video)" ;
                                7 = "On (AA simulation off)"     ; 12 = "Off (AA simulation type 1)"; 15 = "On (AA simulation type 1)" ;
                               20 = "Off (AA simulation type 2)" ; 23 = "On (AA simulation type 2)"                }[$s[1]]
                If     ($s[3] -band 1)  {   $t = $t + " Focal Length: $($S[3]*4)mm"}
                elseif ($s[3])          {   $t = $t + " Focal Length: $($S[3]/2)mm"}
                If     ($s[0] –band 1)  {   $t = $t + " Stabilized."    }
                else                    {   $t = $t + " Not Stabilized."}
                If     ($s[0] –band 64) {   $t = $t + " Not Ready."     }
                $makernoteresults[0x005c].meaning = $t
            } # Shake reduction
            if ($makernoteresults[0x005d]) {
                $s = $makernoteresults[0x005d].Value
                $d = $makernoteresults[0x0006].Value
                $t = $makernoteresults[0x0007].Value
                [int64]$dno = (16777216 * $d[0]) + (65536 * $d[1]) + (256 * $d[2]) + ($d[3])
                [int64]$tno = (16777216 * $t[0]) + (65536 * $t[1]) + (256 * $t[2]) + ($t[3])
                [int64]$sno = 4294967296 - (16777216 * $s[0]) - (65536 * $s[1]) - (256 * $s[2]) - ($s[3]) - 1
                $makernoteresults[0x005d].meaning = ($Sno -bXor $dNo -bXor $tNo)
            } # Shutter count
            if ($makernoteresults[0x0069]) {
                $drExpansion = [int[]]$makernoteresults[0x0069].value
                if ($drExpansion[0] -eq 0) { $t = "Off"}
                elseif ($drExpansion[1] -eq 1) { $t = "Enabled"}
                elseif ($drExpansion[1] -eq 2) { $t = "Auto"}
                else {$t = "None"}
                $makernoteresults[0x0069].meaning = $t
            } # DR Expansion
            if ($makernoteresults[0x006B]) {
                $timeinfo = $makernoteresults[0x006B].value
                if ($timeinfo.count -eq 4) {
                    $hometown = $PentaxCities[[int]$timeinfo[2]]
                    $destination = $PentaxCities[[int]$timeinfo[3]]
                    if ($timeinfo[0] -band 1) {$CurrentTz = $destination} else {$CurrentTz = $hometown}
                    if ($timeinfo[0] -band 2) {$hometown += " (DST)" }
                    if ($timeinfo[0] -band 4) {$destination += " (DST)" }
                    $makernoteresults[0x006B].meaning = "Location=$currentTz; Hometown=$hometown; Destination=$destination"
                }
            } # Time zone info
            if ($makernoteresults[0x006f]) {
                if ($makernoteresults[0x006f].value[0] -eq 0) { $makernoteresults[0x006f].meaning = "Off"}
                else { $makernoteresults[0x006f].meaning = "On" }
            } # Contrast Highlight/Shadow adjust
            if ($makernoteresults[0x0070]) {
                if     ($makernoteresults[0x0070].value[0] -eq 0) { $makernoteresults[0x0070].meaning = "Off"}
                elseif ($makernoteresults[0x0070].value[1] -eq 0) { $makernoteresults[0x0070].meaning = "Normal"}
                elseif ($makernoteresults[0x0070].value[1] -eq 2) { $makernoteresults[0x0070].meaning = "Extra-Fine"}
                else {  $makernoteresults[0x0070].meaning = "Unknown"}
            } # Fine Sharpness
            if ($makernoteresults[0x0071]) {
                $s = $makernoteresults[0x0071].Value
                $t = @{0 = "Inactive - "; 1 = "Active "; 2 = "Active (Weak) "; 3 = "Active (Strong) "; 4 = "Active (Medium) "              }[$S[1]] +
                     @{0 = "Off "       ; 1 = "Weakest"; 2 = "Weak";           3 = "Strong";           4 = "Medium";          255 = "Auto" }[$S[0]]
                if ($s[1] -eq 0 -and $S[0] -ne 255 -and $s[2])  {
                    $t += " will be enabled if " + @{48 = "ISO>400"; 56 = "ISO>800"; 64 = "ISO>1600"; 72 = "ISO>3200"}[$S[2]]
                }
                $makernoteresults[0x0071].meaning = $t
            } # High ISO NR
            if ($makernoteresults[0x0079]) {
                $ShadowCorr = [int[]]$makernoteresults[0x0079].value
                if ($ShadowCorr[0] -eq 0) {$t = "Off" }
                elseif ($ShadowCorr[0] -eq 2) {$t = "Auto" }
                elseif ($ShadowCorr[0] -eq 1) {
                    if ($ShadowCorr[1] -eq 1) {$t = "Weak" }
                    if ($ShadowCorr[1] -eq 2) {$t = "Normal"}
                    if ($ShadowCorr[1] -eq 3) {$t = "Strong"}
                }
                else {$t = "Unknown"}
                $makernoteresults[0x0079].meaning = $t
            } # Shadow Correction
            if ($makernoteresults[0x007d]) {
                $LensCorr = $makernoteresults[0x007d].value
                $t = @()
                if ($LensCorr[0]  ) {$t += "Distortion"}
                if ($LensCorr[1]  ) {$t += "Chromatic aberration"}
                if ($LensCorr[2]  ) {$t += "Peripheral illumination"}
                if ($LensCorr[3]  ) {$t += "Diffraction"}
                if ($t.Count -eq 0) {$makernoteresults[0x007d].meaning = "None"  }
                else {$makernoteresults[0x007d].meaning = $t -join ", " }
            } # Lens Correction
            if ($makernoteresults[0x0080]) {
                $makernoteresults[0x0080].meaning = $PentaxAspectRatio[[int]$makernoteresults[0x80].Value[0]]
            } # Aspect ratio
            if ($makernoteresults[0x0085]) {
                $hdr = [int[]]$makernoteresults[0x0085].value
                if ($hdr[0] -eq 0) {$t = "HDR Off"}
                else {
                    $t = @{1 = "HDR Auto"; 2 = "HDR 1"; 3 = "HDR 2"; 4 = "HDR 3"; 5 = "HDR Advanced"}[$hdr[0]]
                    if ($hdr[1] -eq 1) {$t = "$t , Auto Align" }
                    $t += @{0 = ""; 4 = "1 EV"; 8 = "2 EV"; 12 = "3 EV"}[$hdr[2]]
                }
                $makernoteresults[0x0085].meaning = $t
            } # HDR
            if ($makernoteresults[0x0215]) {
                $m = $makernoteresults[0x0215]
                if ($m.value[1] -gt 20000000 -and $m.value[1] -lt 21000000) {
                    #$t  = $PentaxModel[$m.value[0]] +", "
                    $t = "Manufactured {0}-{1}-{2}" -f [int]( $m.value[1] / 10000), [int](($m.value[1] % 10000) / 100).tostring("00"), ($m.value[1] % 100).tostring("00")
                    $t += ", Production code {0}.{1}" -f $m.value[2], $m.value[3]
                    $t += ", Internal Serial no = {0}" -f $m.value[4]
                    $makernoteresults[0x0215].meaning = $t
                }
                else {$makernoteresults[0x0215].meaning = "Format not recognized" }
            } # Camera info
            if ($makernoteresults[0x0216]) {
                $S = $makernoteresults[0x0216].value
                if     (($s[0] -band 0x0f) -eq 2) {
                    $t = "Body Battery"
                    $level = $s[1] -band 0xf0
                    switch ($level) {
                        0x10 {$t += ": empty."}
                        0x20 {$t += ": almost empty."}
                        0x30 {$t += ": half full."}
                        0x40 {$t += ": almost full."}
                        0x50 {$t += ": full."}
                    }
                }
                elseif (($s[0] -band 0x0f) -eq 3) {
                    $t = "Grip Battery"
                    $level = $s[1] -band 0x0f
                    switch ($level) {
                        1 {$t += ": empty."}
                        2 {$t += ": almost empty."}
                        3 {$t += ": half full."}
                        4 {$t += ": almost ful.l"}
                        5 {$t += ": full."}
                    }
                }
                elseif ($s[0] -band 0x0f -eq 4) {$t = "External Power"}
                else                            {$t = "Unknown Power"}
                $makernoteresults[0x0216].Meaning =$t
            } # Battery info
            if ($makernoteresults[0x022b]) {
                $levelInfo = $makernoteresults[0x022b].Value
                $t = switch ($levelInfo[0] -band 0x0f) {
                    1 {"Horizonal"}
                    2 {"Rotate 180"}
                    3 {"Rotate  90 CW"}
                    4 {"Rotate 270 CW"}
                    9 {"Horizonal; Off Level"}
                    10 {"Rotate 180; Off Level"}
                    11 {"Rotate  90 CW; Off Level"}
                    12 {"Rotate 270 CW; Off Level"}
                    13 {"Upwards"}
                    14 {"Downwards"}
                    Default {"Unknown"}
                }
                $t += switch ($levelInfo[0] -band 0xf0) {
                    0x20 {" with Composition Adjust"}
                    0xa0 {" with Composition Adjust & Horizion Correction"}
                    0xc0 {" with Horizion Correction"}
                }
                if ($levelinfo[2] -gt 0x7f) { $t += "; Pitch " + ((0xff -bxor ($levelinfo[2] - 1)) / 2).ToString("0.0° Up") }
                elseif ($levelinfo[2] -ne 0   ) { $t += "; Pitch " + (             $levelinfo[2] / 2).ToString("0.0° Down") }
                if ($levelinfo[1] -gt 0x7f) { $t += "; Roll " + ((0xff -bxor ($levelinfo[1] - 1) ) / 2).ToString("0.0° Clockwise")  }
                elseif ($levelinfo[1] -ne 0   ) { $t += "; Roll " + (             $levelinfo[1] / 2).ToString("0.0° AntiClockwise")}

                #value 5,6,7 are adjust steps up, right, clockwise step = 1/16 mm or 1/8°
                $makernoteresults[0x022b].Meaning = $t
            } # Level Info
            if ($makernoteresults[0x0243]) {
                if ($makernoteresults[0x0243].value[0] -eq 0) {
                    $makernoteresults[0x0243].meaning = "Pixel shift off"
                }
                else { $makernoteresults[0x0243].meaning = "Pixel shift on" }
            } # Pixel Shift info
            if ($makernoteresults[0x2001].Value)               {
                Write-Verbose -Message ("$Path " + $makernoteresults[0x2001].Value);
                $FwDateString = $makernoteresults[0x2001].Value -replace "\x00", ""
                if ($FwDateString -match "\d{10}") {
                    $makernoteresults[0x2001].meaning = [DateTime]::ParseExact(($FwDateString), "yyMMddHHmm", [System.Globalization.CultureInfo]::InvariantCulture)
                }
            } # Firmware Date
            if ($makernoteresults[0x000c].Value.length -eq 2 ) {
                $makernoteresults[0x000c].meaning = $PentaxFlash1[$makernoteresults[0x000c].Value[1]] + ", " + $PentaxFlash0[$makernoteresults[0x000c].Value[0]]
            } # Flash mode
            if ($makernoteresults[0x0229].value -match "\D" )  { $makernoteresults[0x0229].meaning = "Format not recognized" }  #Serial no (string, all digits)

            0x1f, 0x20, 0x21      | ForEach-Object {if ($makernoteresults[$_] ) {
                    $makernoteresults[$_].meaning = $PentaxHiLo[$makernoteresults[$_].Value[0]]
                    if (-not $makernoteresults[$_].meaning) {$makernoteresults[$_].meaning = "Unknown" }
                }
            }  #Saturation contrast and sharpness
            0x6c, 0x6d, 0x6e      | ForEach-Object {
                if ($makernoteresults[$_] ) {$makernoteresults[$_].meaning = $makernoteresults[$_].value[0]  }
            }  #High/Low key adjust, Contrast Highlight, Contrast shadow
            $PentaxLookUps.Keys   | ForEach-Object {if ($makernoteresults[$_] ) {
                    $makernoteresults[$_].meaning = $PentaxLookUps[$_][$makernoteresults[$_].Value]
                    if (-not $makernoteresults[$_].meaning) {$makernoteresults[$_].meaning = "Unknown" }
                }
            }  #Simple lookups

            if ($Raw) {$makernoteresults.Values}
            $makernoteresults[$PentaxTagNames.keys] | Where-Object {$_.tagname} | ForEach-Object {
                if ($null -ne $_.meaning ) {$PropHash[$_.tagname] = $_.meaning}
                else {$PropHash[$_.tagname] = $_.Value}
            }
            if (-not ($Results[0x8827].Value) -and $makernoteresults[0x3014].Value)   { $PropHash["ISOSpeed"]  = $makernoteresults[0x3014].Value }
            if (-not ($Results[0xA434].Value) -and $makernoteresults[0x003f].meaning) { $PropHash["LensModel"] = $makernoteresults[0x003f].meaning }
        }
        elseif ($makerNotePreamble -and ($Results[0x010F].Value -match "Canon") ) {
            #Apolgies to any Canon user this just makes sure I have the lens details for my canon compact.
            Write-Verbose -Message "Maker ID matches Cannon, Found maker note, attempting to parse it."
            $makernotes = Read-ExifFD -Stream $Stream -ImgDirStart $makernoteoffset -LittleEndian $true -tagnames @{} -dataoffset $Dataoffset
            if (-not ($Results[0xA434].Value) -and $makernotes[0x0095].Value) {$PropHash["LensModel"] = $makernotes[0x0095].Value }
            elseif (-not ($Results[0xA434].Value) -and $makernotes[1].Value[25]) {
                $PropHash["LensModel"] = "Unknown {0} to {1}mm lens" -f ($makernotes[1].Value[24] / $makernotes[1].Value[25]), ($makernotes[1].Value[23] / $makernotes[1].Value[25])
            }
        }
        elseif ($makerNotePreamble -match "Apple") {
            $makernotes = Read-ExifFD -Stream $Stream -ImgDirStart $( $MakerNoteOffset + $makerNotePreamble.IndexOf("MM") + 2 ) -LittleEndian $false -TagNames $AppleTagNames -dataoffset $Dataoffset -makerNoteOffset
            #Not much in the Apple maker note is documented.
            #   0x0003 is gives a time in since boot (see exif tool Run Time Value / Run Time Scale    = Runtime seconds
            #   0x0008 is 3 rationals which give the acceleration vector
            #   0x0015 is a GUID which is unique for each picture.
            # However Apples's offsets appear to calculated differently (they're off by 4 compared with canon and Pentax) so I haven't bothered.
            if ($makernotes[10].value -eq 3) {$PropHash["AppleHDRType"] = 'HDR Image'      }
            elseif ($makernotes[10].value -eq 4) {$PropHash["AppleHDRType"] = 'Original Image' }
        }
    #endregion
    #region Process IPTC & XAP Data & Apps which write a comment as a section name
        #IPTC-NAA data is textual - it either has an EXIF pointer or shows up in the photoshop JPEG segment, under tag 0404
        if ( $Results[0x83bb] )        { $IPTCPos = ($Results[0x83bb].offset + $Dataoffset) }
        else                           { $IPTCPos =  Find-IPTCPosition -segments $Segments -Stream $stream -Array $Array             }
        if ($IPTCPos)                  { Read-IPTCData  -Stream $Stream -Array $Array -PropHash $PropHash -raw:$raw -IPTCPos $IPTCPos}
        if ($XMPPos -and $XMPLength)   { Read-XAPData   -Stream $Stream -Array $Array -PropHash $PropHash -raw:$raw -XMPPos  $XMPPos -XMPLength $XMPLength }
        if (-not $prophash["Comment"]) {$PropHash["Comment"] = ($segresults | Where-Object -Property tagid -Match fff) | ForEach-Object {$_.tagname.trim()} }
    #endregion
        $Stream.Close()
        #Everything we will return is in $PropHash so make it into a PS Custom Object
        if (-not $Raw) {New-Object -TypeName pscustomobject -Property $PropHash}
    }
}
