function ConvertTo-DateTime {
    <#
      .Synopsis
        Takes a date and time as text, parses it and returns a [DateTime]
      .Description
        Takes a date and time as text, parses it and returns a [DateTime]
      .Example
        C:\PS> ConvertTo-DateTime "2010-01-31 12:23:34"  "yyyy-MM-dd HH:mm:ss"
        Returns 31 January 2010 12:23:34
      .Parameter date
        A text string containing the date
      .Parameter format
        A text string containing the formatting information
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Date,
        [Parameter(Mandatory=$true)]
         [string]$Format
    )
    [DateTime]::ParseExact($Date,$Format,[System.Globalization.CultureInfo]::InvariantCulture)
}

function Set-LogOffset      {
    <#
      .Synopsis
        Sets a global variable $PictureOffset from a picture.
      .Description
        Calculates the offset between the time  displayed by a data logger
        and the time taken recorded by the camera when the logger was photographed
      .Example
        C:\PS> Set-LogOffset
        Will prompt for the path to the picture and the date and time in it
      .Example
        C:\PS> Set-LogOffset -image "D:\DCIM\100PENTX\IMG43210.jpg" -time "12:23:34"
        Will calculate the offset from image , given the date in the picture

    #>
    [CmdletBinding()]
    param (
        #The reference image (object, or path to one. Alias Fullname or path. Can come from the pipeline
        [Parameter(mandatory=$true, HelpMessage="Please enter the path to the picture",ValueFromPipelineByPropertyName=$true)][Alias('FullName','Path')]
        $Image ,
        #The date on the photographed logging devices formatted as HH:mm:ss"
        [Parameter(mandatory=$true, HelpMessage="Please enter the time on the loggin device the reference picture, formatted as HH:mm:ss")]
        [string]$TimeInPicture
    )
    [datetime]$picTime = (Read-Exif -path $Image -verbose:$false).DateTaken
    $Global:LogOffset = -( $picTime.TimeOfDay.Subtract($TimeInPicture).TotalSeconds)
    Write-Verbose -Message "Camera time=$picTime  , logger time=$TimeInPicture. Add $Global:LogOffset seconds to picture time to give log time"
}

function Get-NearestPoint   {
    <#
      .Synopsis
        From a set of timestamped data points, returns the one nearest to a given time
       .Description
        From a set of timestamped data points, returns the one nearest to a given time
      .Example
        C:\PS> $point = get-nearestPoint -DataPoints $points -ColumnName "DateTime" -MatchingTime $dt
        Returns the point in $pointswhere the "DateTime" column is nearest to $dt
      .Parameter DataPoints
        An array containing the data points
      .Parameter ColumnName
        The name of the column in the points array that holds the dateTime to match against
      .Parameter MatchingTime
        The time of the item being sought
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $DataPoints ,
        [Parameter(Mandatory=$true)]
        $ColumnName ,
        [Parameter(Mandatory=$true)]
        $MatchingTime
    )
    Write-Verbose -Message "Checking $($DataPoints.count) points for one where $ColumnName is closest to $MatchingTime"
    $variance  = [math]::Abs(($DataPoints[0].$ColumnName - $MatchingTime).totalseconds)
    $i         = 1
    do {
        # write-progress -Activity "looking" -Status "looking" -CurrentOperation $i
        $v = [math]::Abs(($DataPoints[$i].$ColumnName - $MatchingTime).totalseconds)
        if   ($v -le $variance)      {$i ++ ; $variance = $v }
    } while (($v -eq $variance) -and ($i -lt $datapoints.count))
    Write-Verbose -Message "Point $I matched with variance of $variance seconds"
    $datapoints[($i -1)]
    # write-progress -Activity "looking" -Status "looking" -Completed
}

function Get-CSVGPSData     {
    <#
      .Synopsis
        Gets GPS Data from a CSV file
      .Description
        Gets GPS Data from a CSV file
      .Example
        C:\PS> $points = Get-CSVGPSData .\20100420161012.log -offset $offset
        Reads the GPS data from the Comma seperated log file,
        applying the offset in $offset - storing the result in $points.
      .Parameter Path
        The path to the file
      .Parameter Offset
        The offset to apply to the logged data
  #>
    param  (
        [Parameter(Mandatory=$true )]
        [Alias("Filename","FullName")]$Path ,
        $offset = 0
    )
    ## Check your date format different GPS devices return different numbers of decimals for the .f (fraction part).
    $Dateformat = "yyyy-MM-dd HH:mm:ss"
    Import-Csv -Path $path -Header "Date","Lat","Lon","altitude","bearing","MetersPerSec","H_acc","V_acc","blank","Network"|
        Select-Object  -property  @{Name="DateTime"; Expression = {(   ConvertTo-DateTime $_.Date $DateFormat ).addSeconds($offset) }}  ,
                                    @{Name="MPH";      Expression = {    [system.math]::Round((2.237  * $_.MetersPerSec),1) }} , @{Name="KPH"; Expression = { [system.math]::Round((3.6 * $_.Knots),1) }} ,
                                    knots, bearing ,
                                    @{Name="LatDMS";   Expression = {  @([math]::truncate([math]::Abs( [double]$_.lat                               )   )  ,
                                                                        [math]::truncate([math]::Abs( [double]$_.lat      - [math]::truncate( [double]$_.lat) )*60)  ,
                                                                        [math]::round(   [math]::Abs(([double]$_.lat *60) - [math]::truncate(([double]$_.lat *60)) )*60 ,2) )  }}   ,
                            lat , @{Name="NS";       Expression = {if ([double]$_.lat -gt 0) {"N"} else {"S"}  }} ,
                                    @{Name="LonDMS";   Expression = {  @([math]::truncate([math]::Abs( [double]$_.lon                               )   )  ,
                                                                        [math]::truncate([math]::Abs( [double]$_.lon      - [math]::truncate( [double]$_.lon) )*60)  ,
                                                                        [math]::round(   [math]::Abs(([double]$_.lon *60) - [math]::truncate(([double]$_.lon *60)) )*60 ,2) )  }}  ,
                            lon , @{Name="EW";        Expression = {if ([double]$_.lon -gt 0) {"E"} else {"W"}  }} ,
                                    @{Name="AltM";      Expression = {    [math]::round(  ([double]$_.Altitude       ), 1) }} ,
                                    @{Name="AltFT";     Expression = {    [math]::round(  ([double]$_.Altitude * 3.28), 1) }}  |
            Sort-object -property datetime
}

function Get-GPXData        {
    <#
      .Synopsis
        Gets GPS Data from a GPX format XML file
      .Description
        Gets GPS Data from a GPX format XML file
      .Example
        C:\PS> $points = Get-GPXData .\20100420161012.GPX -offset $offset
        Reads the GPS data from the GPX XML log file,
        applying the offset in $offset - storing the result in $points.
      .Parameter Path
        The path to the file
      .Parameter Offset
        The offset to apply to the logged data
    #>
    param  (
        [Parameter(Mandatory=$true)]
        [Alias("FileName","FullName")]$Path ,
        $offset = 0
    )
    ([xml](Get-Content-path $path)).gpx.trk.trkseg.trkpt |
      Select-Object -property  @{Name="DateTime"; Expression = { (  [dateTime]$_.time).toUniversalTime().addSeconds($offset) }}  ,
                               @{Name="MPH";      Expression = {    [math]::Round((0.621  * $_.Speed),1) }} ,
                               @{Name="KPH";      Expression = {  $_.speed }} ,
                               @{Name="knots";    Expression = {    [math]::Round(($_.Speed / 1.852 ),1) }},
                               @{Name="Bearing";  Expression = {  $_.course}}  ,
                               @{Name="LatDMS";   Expression = {  @([math]::truncate([math]::Abs( [double]$_.lat                               )   )  ,
                                                                    [math]::truncate([math]::Abs( [double]$_.lat      - [math]::truncate( [double]$_.lat) )*60)  ,
                                                                    [math]::round(   [math]::Abs(([double]$_.lat *60) - [math]::truncate(([double]$_.lat *60)) )*60 ,2) )  }}   ,
                               lat , @{Name="NS"; Expression = {if ([double]$_.lat -gt 0) {"N"} else {"S"}  }} ,
                               @{Name="LonDMS";   Expression = {  @([math]::truncate([math]::Abs( [double]$_.lon                               )   )  ,
                                                                    [math]::truncate([math]::Abs( [double]$_.lon      - [math]::truncate( [double]$_.lon) )*60)  ,
                                                                    [math]::round(   [math]::Abs(([double]$_.lon *60) - [math]::truncate(([double]$_.lon *60)) )*60 ,2) )  }}  ,
                               lon , @{Name="EW"; Expression = {if ([double]$_.lon -gt 0) {"E"} else {"W"}  }} ,
                               @{Name="AltM";     Expression = {    [math]::round(  ([double]$_.ele       ), 1) }} ,
                               @{Name="AltFT";    Expression = {    [math]::round(  ([double]$_.ele * 3.28), 1) }}  |
                Sort-object -property datetime
}

function Get-NMEAData       {
    <#
      .Synopsis
        Gets GPS Data from a text file of NMEA sentences
      .Description
        Gets GPS Data from a text file of NMEA sentences
      .Example
        C:\PS> $points = Get-NMEAData .\20100420161012.LOG -offset $offset
        Reads the GPS data from the NMEA file,
        applying the offset in $offset - storing the result in $points.
      .Parameter Path
        The path to the file
      .Parameter Offset
        The offset to apply to the logged data
    #>
    param  ([
        Parameter(Mandatory=$true)]
        [Alias("FileName","FullName")]
        $Path ,
        $Offset = 0 ,
        [Switch]$NoAltitude
)
    ## Check your date format different GPS devices return different numbers of decimals for the .f (fraction part).
    $Dateformat = "ddMMyyHHmmss.f"
    $TimeFormat = "HHmmss.f"
    if (-not $NoAltitude) {$altPoints = (Import-Csv -path $Path -Header "type","time","lat","ns","lon","ew","quality","sattelites","HDofP","Altitude","Units","age","ref" |
       Where-Object {$_.type -eq '$GPGGA'} )  |
          Select-Object -property "lat","ns","lon","ew","altitude",
                                   @{Name="DateTime"; Expression = {(ConvertTo-DateTime $_.time $TimeFormat).timeofday }} | Sort-Object -Property datetime
    }
    Import-Csv -Path $Path -Header "Type","Time","status","lat","NS","lon","EW","Knots","bearing","Date","blank","checksum" |
       Where-Object {$_.type -eq '$GPRMC' -and $_.Time } |
          Select-Object -property  @{Name="DateTime"; Expression = { (ConvertTo-DateTime ($_.Date+$_.Time) $DateFormat).addSeconds($Offset) }}  ,
                                           knots,
                                   @{Name="MPH";      Expression = {  [math]::Round((1.15  * $_.Knots),1) }} ,
                                   @{Name="KPH";      Expression = {  [math]::Round((1.852 * $_.Knots),1) }} ,
                                           bearing,   lat, NS,
                                   @{Name="LatDMS";   Expression = {@([math]::truncate($_.lat / 100) ,
                                                                      [math]::truncate($_.lat % 100) ,
                                                                      [math]::round((60 * ($_.lat - [math]::truncate($_.lat ))),2))}},
                                           lon, EW,
                                   @{Name="LonDMS";   Expression = {@([math]::truncate($_.lon / 100) ,
                                                                      [math]::truncate($_.lon % 100) ,
                                                                      [math]::round((60 * ($_.lon - [math]::truncate($_.lon ))),2))}},
                                   @{Name="AltM";     Expression = {if ($alt) {(Get-NearestPoint -datapoints $altPoints -columnname "DateTime" `
                                                                                -matchingtime (ConvertTo-DateTime $_.time $TimeFormat).timeofday).altitude }}}
}

function ConvertFrom-KML    {
    [cmdletbinding()]
    param(
        #File to convert
        $Path
    )

    process {
        foreach ($p in $Path) {
            $kml= ([xml](Get-Content $p)).kml
            if (-not $kml.Document) {Write-Warning "Failed to read KML from $path" ; return }
            foreach ($Placemark in  $kml.Document.Folder.where({$_.name -eq "Points"}).placemark ) {
                $lonLatAlt = $Placemark.point.coordinates.split(",")
                if ($kml.Document.Folder.where({$_.name -eq "Points"}).placemark.description."#cdata-section"[1] -match
                   "(\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}) UTC\<br\>Longitude:(\d+)&deg;(\d+)'([\d.]+)""(\w)\<br\>Latitude:(\d+)&deg;(\d+)'([\d.]+)""(\w)") {
                    New-Object -TypeName psobject -Property ([ordered]@{
                        Time = [dateTime]$Matches[1]
                        NS   = $Matches[9] ;  LatDMS = $Matches[6,7,8] ; Lat  = [math]::Abs($lonLatAlt[1]) ;
                        EW   = $Matches[5] ;  LonDMS = $Matches[2,3,4] ; Lon  = [math]::Abs($lonLatAlt[0]) ;
                        AltM = $lonLatAlt[2]})
                }
            }
        }
    }
}

function Convert-GPStoEXIFFilter {
    <#
      .Synopsis
        Builds a collection of EXIF WIA filters to set GPS data from a point
      .Description
        Builds a collection of EXIF WIA filters to set GPS data from a point
      .Example
        C:\PS> $filter = Convert-GPStoEXIFFilter 51,36,7 "N" 1,33,54 "W"
        Creates a new filter chain and adds the Exif Filters to add the GPS data to images
      .Parameter LATDMS
        The lattitude as an array of 3 numbers for Degrees, Minutes and Seconds
      .Parameter NS
        N for lattitude north of the equator, S for lattitude South of the equator
      .Parameter LONDMS
        The longitude as an array of 3 numbers for Degrees, Minutes and Seconds
      .Parameter EW
        E for longitude East of Greenwich W for longitude West of GreenWich
      .Parameter AltM
        Altitude in Meters above mean Sea level
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]$LatDMS,
        [Parameter(Mandatory=$true)]$NS,
        [Parameter(Mandatory=$true)]$LONDMS,
        [Parameter(Mandatory=$true)]$EW,
        $AltM
    )
    process {
        $filter = New-Imagefilter
        if (-not $filter.Apply)  {Write-Warning -Message "Couldn't get filter"; return }
        $ExifVervalue = New-Object -ComObject "WIA.Vector"
        $ExifVervalue.Add([byte]2)
        $ExifVervalue.Add([byte]2)
        $ExifVervalue.Add([byte]0)
        $ExifVervalue.Add([byte]0)

        $LongDMSValue = New-Object -ComObject "WIA.Vector"
        $LatDMSValue  = New-Object -ComObject "WIA.Vector"
        $LonDMS | ForEach-Object {$v = New-Object -ComObject wia.rational ; $v.numerator = [int32]($_ * 1000000) ; $v.denominator = 1000000 ; $longDmsValue.add($v) }
        $LatDMS | ForEach-Object {$v = New-Object -ComObject wia.rational ; $v.numerator = [int32]($_ * 1000000) ; $v.denominator = 1000000 ; $latDmsValue.add($v) }


        Add-EXIFFilter -Filter $filter -ExifTag GPSVersion   -TypeName VectorOfByte      -Value $ExifVervalue
        Add-EXIFFilter -Filter $filter -ExifTag GPSLatRef    -TypeName String            -Value $NS
        Add-EXIFFilter -Filter $filter -ExifTag GPSLongRef   -TypeName String            -Value $EW
        Add-EXIFFilter -Filter $filter -ExifTag GPSLongitude -TypeName VectorOfURational -Value $longDmsValue
        Add-EXIFFilter -Filter $filter -Exiftag GPSLatRef    -TypeName VectorOfURational -Value $LatDMSValue
        Write-Verbose -Message "Created EXIF filter for Lattitude and Longitude"
        if ($AltM -or $AltM -eq 0)  {
            if ($altM -ge 0) { Add-EXIFFilter -Filter $filter -ExifTag GPSAltRef -TypeName Byte -Value ([byte]0) }
            else             { Add-EXIFFilter -Filter $filter -ExifTag GPSAltRef -typename Byte -value ([byte]1) }
            Add-EXIFFilter -Filter $filter -ExifTag GPSAltitude -TypeName URational "$ExifUnsignedRational" -Numerator ([uint32]([math]::Abs($AltM)  * 100) ) -denominator 100
            Write-Verbose -Message "Added EXIF filter for altitude"
           }
           return $filter
    }
}

function Copy-GPSImage {
    <#
      .Synopsis
        Copies an image, applying EXIF data from GPS data points
      .Description
        Copies an image, applying EXIF data from GPS data points
      .Example
        C:\PS>  Dir E:\dcim –inc IMG*.jpg –rec | Copy-GpsImage -Points $Points -Keywords "Oxfordshire" -rotate -DestPath "$env:userprofile\pictures\oxford" -replace  "IMG","OX-"
        Copies IMG files from folders under E:\DCIM to the user's picture\Oxford folder, replacing IMG in the file name with OX-.
        The Keywords field is set to Oxfordshire, pictures are GeoTagged with the data in $points and rotated.
      .Parameter Image
        A WIA image object, a path to an image, or a file object representing an image file. It may be passed via the pipeline.
      .Parameter Points
        An array of GPS data points
      .Parameter Destination
        The FOLDER to which the file should be saved.
      .Parameter Keywords
        If specified, sets the keywords Exif field.
      .Parameter Title
        If specified, sets the Title Exif field..
      .Parameter Replace
        If specified, this contains two values seperated by a comma specifying a replacement in the file name
      .Parameter Rotate
        If this switch is specified, the image will be auto-rotated based on its orientation filed
      .Parameter NoClobber
        Unless this switch is specified, a pre-existing image WILL be over-written
      .Parameter ReturnInfo
        If this switch  is specified the GPS point associated with each image is returned.
        Note that the point in the points collection is updated as a side effect, so this can be combined with Out-Null
        The points returned - or isolated from the collection with a a where command can be used
        to plot picture locations on a map
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        [Parameter(ValueFromPipeline=$true, Mandatory=$true)][Alias("Path","FullName")]$image ,
        [Parameter(Mandatory=$true)]$points,
        [Parameter(Mandatory=$true)]$Destination ,
        $Keywords ,
        $Title,
        $Replace,
        [switch]$Rotate,
        [switch]$NoClobber,
        [switch]$ReturnInfo,
        [Parameter(DontShow=$true)]
        $psc = $pscmdlet
    )
    process {
      #  if ($psc -eq $null)  {$psc = $pscmdlet} ; if (-not $PSBoundParameters.psc) {$PSBoundParameters.add("psc",$psc)}
        if ($image -is [system.io.fileinfo] ) {$image = $image.FullName }
        if ($image -is [String]             ) {$image = Get-Image $image}
        if ($Image.count -gt 1              ) {[Void]$PSBoundParameters.Remove("Image")
                                               $Image | ForEach-object {Copy-GPSImage -image $_ @PSBoundParameters}
                                               return
        }
        if ($image -is [__comObject])  {
            $dt =(Read-Exif -path $image).DateTaken
            if ($dt) {# write-verbose $dt
                      $point  = Get-NearestPoint  -DataPoints $points      -ColumnName "DateTime" -MatchingTime $dt
                      $filter = Convert-GPStoEXIFFilter -LatDMS $point.Latdms -NS $point.NS -LONDMS $point.londms -EW $point.ew -AltM $point.altM
                      [Void]$PSBoundParameters.Remove("Points")
                      $path = Copy-Image -filter $filter @PSBoundParameters
                      If ($ReturnInfo) {if ($point.paths ) {if ($point.paths -notcontains $path ) {$point.paths +=  $path ; $point}}
                                        else               {$point | Add-Member -MemberType Noteproperty -Name "paths" -Value @($path) -PassThru -force}
                      }
            }
        }
    }
}

Add-Type -TypeDefinition  @"
public struct suuntoItem
{
   public string          Author       ;
   public int             BottomTemp   ;
   public string          Copyright    ;
   public string          Description  ;
   public int             DiveNumber   ;
   public System.DateTime EndTime      ;
   public string          FileID       ;
   public string          GPSLattitude ;
   public string          GPSLongitude ;
   public string          LocationName ;
   public string          Path         ;
   public double          Pressure     ;
   public double          SampleTime   ;
   public double          SegmentDepth ;
   public string          SiteName     ;
   public System.DateTime StartTime    ;
}
"@
#I built this with : Get-LightRoomItem -include IMG64286 | Get-NearestSuutoDBPoint | gm -MemberType Property | foreach -begin {"public struct suuntoItem";"{"} -Process {$_.definition -replace "(\w+)\s+(\w+)\s.*$","   public `$1  `$2  ;"} -end {"}"} |clip

function Get-NearestSuutoDBPoint{
    <#
      .Synopsis
        Finds the nearest point in the suunto database
      .Description
        Returns data from the database which matches either a Date & time (when) or the DateTaken EXIF field in an image
        If  Set-LogOffset  has been run the error between the date/time being sought is allowed for
        So if the camera time isn't synced to dive computer time, photograph the computer
        and run set-logoffset <filename> <time on the computer>
        To help functions further down the pipe line if an image is passed the returned data contains extra properties
      .Example
            Get-NearestSuutoDBPoint .\DIVE56886.jpg

            dir d*.jpg | Get-NearestSuutoDBPoint  | group -Property Divenumber  | %{ $_.group | measure -max -Minimum -Property fileID | Add-Member -MemberType NoteProperty -Name "DiveID" -value $_.name -PassThru }  | ft -AutoSize diveID,Count,Minimum,maximum
            DiveID Count Minimum Maximum
            ------ ----- ------- -------
            212       12   56886   56938
            213       10   56948   56978
            214       23   56981   57041
    #>
    [cmdletBinding()]
    [OutputType([suuntoItem])]
    param (
        #Date to find the database point to
        [Parameter(ParameterSetName='Date',Mandatory=$true)]
        $When,
        #Image to get the database point for
        [Parameter(ParameterSetName='Image',ValueFromPipelineByPropertyName=$true, Mandatory=$true,Position=0)]
        [alias ("Path","FullName")]
        $image,
        #Database connection string if there isn't and ODBC source named "Suunto"
        $Connection="DSN=suunto"
    )
    process {
        $path = $title = $copyright = ""
        $Keywords = @()
        if ($image) {if ($image -is [System.Data.DataRow]) {
                        $path       = $image.Path
                        $fileID     = $image.id_local
                        $when       = [datetime]$image.captureTime
                        $copyright  = $image.copyright -replace "'","''"
                    }
                    else {
                        # Since we need to get exif data here, we'll save some of the values which will be useful later
                        $Exif       = (Read-Exif -path $image -Verbose:$false)
                        $path       = $Exif.path
                        $fileId     = $path -replace "^.*?(\d{3,6}).*\.jpg","`$1"
                        $When       = $Exif.DateTaken
                        $Keywords   = $Exif.Keywords
                        $title      = $Exif.Title     -replace "'","''" -replace "(.+)",'$1, '
                        $copyright  = $Exif.Copyright -replace "'","''"
                        $author     = $Exif.Author    -replace "'","''"
                        if (-not $author) {
                            $author  = $Exif.Artist    -replace "'","''"}
                    }
        }
        if ($null -eq $global:LogOffset) {Write-Warning -Message "No offset has been set, assuming times are sync'd" ; $global:LogOffset = 0}
        if ($when -is [datetime])        {$when = $when.AddSeconds($global:LogOffset).ToString("MM/dd/yyyy HH:mm:ss") }
        elseif ($when -isnot [string] )  {Write-Warning -Message "Could not establish the date taken, giving up on image $path" ;  return}
        #Select TrackPoints joined to Dives - where the photo time is after the start of the dive and before the end of it.
        #Embed Path and any existing Author & Copyright notice in the Data
        #Merge Existing title, depth, location site, and temperature into Description
        Write-Verbose -Message "Searching for dive info logged at $when"
        $result = Get-SQL -connection $Connection -Session Suunto -verbose:$false -quiet  -SQL "
SELECT     Items.StartTime, TrackPoints.SampleTime ,  Items.EndTime ,
           Items.i_custom11                        AS DiveNumber    ,
           Items.t_custom3                         AS LocationName  ,
           Items.t_custom4                         AS SiteName      ,
           Items.t_custom15                        AS GPSLattitude  ,
           Items.t_custom16                        AS GPSLongitude  ,
           round(trackpoints.Altitude,1)           AS SegmentDepth  ,
           Items.i_custom14                        AS BottomTemp    ,
           round(TrackPoints.d_custom3/1000 , 0)   AS Pressure      ,
           '$copyright'                            AS Copyright     ,
           '$Author'                               As Author        ,
           '$path'                                 As Path          ,
           '$fileID'                               As FileID        ,
           '$title' &
           SegmentDepth & 'M, '& siteName   & ', ' &
           LocationName & '. ' & BottomTemp & '°C' AS Description
FROM       TrackPoints
INNER JOIN Items ON TrackPoints.LogID = Items.ItemID
WHERE (   (items.type      = 33)
       and(Items.starttime < #$when#)
       and(Items.endtime   > #$when#))
ORDER BY   Items.i_custom11, TrackPoints.SampleTime
"       | Add-Member        -Name     KeyWords -MemberType NoteProperty   -PassThru -Value $Keywords   |
           Add-Member       -Name     LAT      -MemberType ScriptProperty -PassThru -Value {$this.GpsLattitude -split "\s*,\s*"}  |
            Add-Member      -Name     LON      -MemberType ScriptProperty -PassThru -Value {$this.GPSLongitude -split "\s*,\s*"}  |
             Add-Member     -Name     GpsLat   -MemberType ScriptProperty -PassThru -Value {[math]::Round([int]$this.lat[0]+$this.lat[1]/60 + $this.lat[2]/3600,5) * $(if ($this.lat[3] -eq "s") {-1} else {1} )}  |
              Add-Member    -Name     GpsLon   -MemberType ScriptProperty -PassThru -Value {[math]::Round([int]$this.lon[0]+$this.lon[1]/60 + $this.lon[2]/3600,5) * $(if ($this.lon[3] -eq "w") {-1} else {1} )}  |
               Add-Member   -Name     OffSet   -MemberType ScriptProperty -PassThru -Value {[math]::abs($this.StartTime.AddSeconds($this.SampleTime).subtract([datetime]$when).totalseconds)}|
                 Sort-Object -Property Offset |
                  Select-Object -First 1
   if (-not $result)  {Write-Warning -Message "No Depth records found for a dive at $when" }
   else               {Write-Verbose -Message "Found match $($result.sampletime) seconds into dive number $($result.DiveNumber). $path"
                       return $result }
  }
 }

function Convert-SuuntoToExifFilter {
    [CmdletBinding()]
    <#
      .Synopsis
        Builds a collection of WIA filters to set EXIF data in pictures using Scuba data
      .Description
        This is not expected to be used as on its own - it is called by other commands
        which want to put scuba data into a picture. They get a data point which corresponds
        to the moment in the dive and call this function to get a collection of EXIF WIA filters;
        They then apply the filters to the picture and save it.
        In addition to the data which comes back from the dive log, this function can add the
        Author, Copyright and Keyword tags to the this of filters.
      .Example
        $filter = Get-NearestSuutoDBPoint .\DIVE56886.jpg  | Convert-SuuntotoExifFilter -Verbose -Keywords "oman","ocean" -Author "James O'Neill"
        Creates a new filter chain to set the dive description, depth, gps, author, copyright and keyword tags.
    #>
    param(
        #A Suunto data point
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]$Point,
        #Keywords to add with the data
        [String[]]$Keywords,
        #Text to add at the start of the title
        [String]$Title,
        #Text to set as the Author Name if it is not set already
        [String]$Author
    )
    process {
        if (($Point.LAT.Count -eq 4 ) -and ($Point.Lon.Count -eq 4 )) {
                 $filter = Convert-GPStoEXIFFilter -LatDMS ([single[]]($Point.LAT[0..2])) `
                                                       -NS (           $point.LAT[3] -replace "\s","") `
                                                   -LONDMS ([single[]]($Point.LON[0..2])) `
                                                       -EW (           $Point.LON[3] -replace "\s","")
        }
        else {   $filter = New-Imagefilter }
        if (-not $filter.Apply)  { return }
        if ($Author -and -not $Point.Author) {
            Add-EXIFFilter      -Filter $filter -ExifTag Author -TypeName VectorOfByte -String $Author
            Write-Verbose -Message "   Author $Author"
            if (-not $Point.Copyright) {
                 $CopyrightMsg =   ([char]169 + " " + $Author + " " + $point.EndTime.Year + ", All Rights Reserved.")
                 Add-EXIFFilter -Filter $filter -Exiftag Copyright -TypeName String -Value  $CopyrightMsg
                 Write-Verbose -Message "   Copyright $CopyrightMsg"
            }
        }
        $altitudeValue   =  New-Object -ComObject WIA.Rational -Property @{"Numerator" = ([uint32]([double]$point.SegmentDepth * 100)) ; "Denominator" = 100}
        Add-EXIFFilter          -Filter $filter -ExifTag GPSAltitude -TypeName URational    -Value $altitudeValue
        Add-EXIFFilter          -Filter $filter -ExifTag GPSAltRef   -TypeName Byte         -Value ([byte]1)
        if     ($title)        { $Point.Description = $title          + ": " + $Point.Description }
        Add-EXIFFilter          -Filter $filter -ExifTag Subject     -TypeName VectorOfByte -String $Point.Description
        Write-Verbose -Message "   Depth: $($Point.SegmentDepth). Description: $($Point.Description)"
        if ($Keywords) {
            $Point.keywords | ForEach-Object {if ($Keywords -notcontains $_) {$Keywords += $_} }
            Add-EXIFFilter -Filter $filter -ExifTag Keywords          -TypeName VectorOfByte -string ($Keywords -join ",")
            Write-Verbose -Message ("   Keywords: " + ($Keywords -join ",") )
        }
        return $filter
    }
}

function Copy-SuutoDBTOImage {
    <#
      .Synopsis
        Applies EXIF data from Scuba diving log data points
      .Description
        Copies an image, applying EXIF data from Scuba diving log data points
      .Example
            $DiveFolder =  "C:\Users\Jamesone\Pictures\Ecuador\Dive\"
            Set-PictureOffset -image "$DiveFolder\LightRoom\CRW_5337.jpg" -RefDate "09 APRIL 2012 14:37:18 "
            dir "$DiveFolder\LightRoom\*.jpg" | %{Copy-SuutoImage -image $_ -points $points -Destination  $DiveFolder\Finished" -keywords "Ocean;Galapagos" -replace "crw_|IMG_","Dive" -title $titlehash[$_.FullName -replace ".*\\(.*?\.jpg)",'$1']  }
            Sets the time offset using the CRW_5337.jpg where the time on computer reads "14:37:18 on 9th April 2012";
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
         #A WIA image object, a path to an image, or a file object representing an image file. It may be passed via the pipeline.
        [Parameter(ValueFromPipeline=$true, Mandatory=$true)][Alias("Path","FullName")]$Image ,
        #The FOLDER to which the file should be saved.
        [Parameter(Mandatory=$true)]$Destination ,
        #If specified, sets the keywords Exif field.
        $Keywords ,
        #If specified, adds text to the start of the Title Exif field. N.b  If Windows explorer sets a different field as the title and uses that if it is present.
        $Title,
        #Text to set as the Author Name if it is not set already
        [String]$Author,
        #If specified, this contains two values seperated by a comma specifying a replacement in the file name
        [validateCount(2,2)][string[]]$Replace,
        #Unless this switch is specified, a pre-existing image WILL be over-written
        [switch]$NoClobber,
        [Parameter(DontShow=$true)]
        $psc = $pscmdlet
    )
    process {
        #Ensure that if called recusively the "yes to all" and "No to All" work by the same PScmdlet everywhere
       # moved into default for $psc if ($psc -eq $null)  {$psc = $pscmdlet} ; if (-not $PSBoundParameters.psc) {$PSBoundParameters.add("psc",$psc)}
        #Whatever we were passed in $image make sure we end up with a single image object (even if we call ourself recursively) .
        if ($Image -is [system.io.fileinfo] ) {$Image = $image.FullName }
        if ($Image -is [String]             ) {$Image =(Resolve-Path -Path $Image -errorAction "SilentlyContinue") | ForEach-Object {$_.path }}
        if ($Image.count -gt 1              ) {[Void]$PSBoundParameters.Remove("Image")
                                               $Image | ForEach-Object {Copy-SuutoDBTOImage -Image $_ @PSBoundParameters}
                                               return
        }
        if ($Image -is [String]             ) {$Image = Get-Image -Path $Image}
        if ($Image -is [__comObject])  {
              #Get the data point for this image, and turn it into a set of exif filters
              $filter = Get-NearestSuutoDBPoint -Image $Image  |  Convert-SuuntotoExifFilter -Keywords $Keywords -Title $Title -Author $Author
              if ($Filter.Filters.Count) {
                    $savePath = Join-Path -Path (Resolve-Path -Path $Destination) -ChildPath ((Split-Path -path $Image.FullName -Leaf) -Replace $Replace)
                    Write-Verbose -Message "Applying $($Filter.Filters.Count) filters to update the EXIF data"
                    foreach ($f in $Filter) { $Image = $f.Apply($Image.PSObject.BaseObject) }
                    #Note if $Replace is empty, -replace $null returns the unmodified string
                    #Check we can save, and if we can, save there.
                    if (Test-Path -Path $SavePath) {
                         if     ($noclobber) {Write-Warning -Message "$SavePath exists and WILL NOT be overwritten"; if ($passthru) {$image} ; Return }
                         elseIF ($psc.ShouldProcess($SavePath,"Delete file")) {Remove-Item  -Path $SavePath -Force -Confirm:$false }
                         else   {Return}
                    }
                    if ((Test-Path -Path $SavePath -IsValid) -and ($pscmdlet.shouldProcess($SavePath,"Write image")))  {
                         $image.SaveFile($SavePath)
                     }
             }
             else {Write-Warning -Message "Couldn't get EXIF Filters to update $($image.FullName)"}
        }
   }
}

function Copy-SuutoDBTOLightRoom {
    param (
        [parameter(mandatory=$true)]
        $Path         ,
        $Connection   = "DSN=LR"
    )
    process {
   Get-LightRoomItem -Include $Path | ForEach-Object {
     $point     = Get-NearestSuutoDBPoint $_
     if ($point.Description -and  $_.caption -is [system.dbnull]) {
           Get-SQL -Confirm:$false -Connection $connection -Session LR -Table AgLibraryIPTC           -Where image -EQ $_.id_local -Set caption                                     -Values $point.Description
     }
     else {Write-Warning -Message "$($_.Path) has caption already set or no new caption provided, not updating"}

     if ($point.GpsLat -and $point.GpsLon -and -not $_.hasGPS) {
           Get-SQL -Confirm:$false -Connection $connection -Session LR -table AgHarvestedExifMetadata -Where image -EQ $_.id_local -Set gpsLatitude,gpsLongitude,hasGPS,gpsSequence -Values $point.GpsLat,$point.GpsLon,1,2.0
     }
     else {Write-Warning -Message "$($_.Path)has GPS Already, or no new gps provided not updating"}

     $xmp       = $_.xmp
     $xmlxmp    = [xml]$xmp
     $updateXMP = $false
     if ($point.Description -and -not $xmlxmp.xmpmeta.RDF.Description.title) {
       $TitleXML   = @"

   <dc:title>
    <rdf:Alt>
     <rdf:li xml:lang="x-default">$($point.Description)</rdf:li>
    </rdf:Alt>
   </dc:title>
"@
        $xmp       = $xmp -replace "(?s)(?<=>)(?=\s*</rdf:description>)",$TitleXML
        $updateXMP = $true
     }
     if ($xmp -notmatch "exif:GPSLatitude=" -and $xmp -notmatch "exif:GPSLongitude=" -and $point.LAT.Length -eq 4 -and $point.LON.Length -eq 4) {
        $gpslatDMS = "{0},{1:##.000}{2}" -f $point.LAT[0],([int]$point.LAT[1] + [int]$point.LAT[2]/60),$point.LAT[3]
        $gpsLonDMS = "{0},{1:##.000}{2}" -f $point.LON[0],([int]$point.LON[1] + [int]$point.LON[2]/60),$point.LON[3]
        $gpsxml    = @"

   exif:GPSVersionID="2.2.0.0"
   exif:GPSLatitude="$gpslatDMS"
   exif:GPSLongitude="$gpsLonDMS"
"@
        $xmp       = $xmp -replace "(?s)(?<=exif:ExifVersion=`"\d+`")(?=\s+exif:)",$gpsxml
        $updateXMP = $true
      }
      if ($updateXMP) {
              Get-SQL -Confirm:$false -Connection $connection -Session LR -Table Adobe_AdditionalMetadata -Set xmp -Values ($xmp -replace "'","''") -Where image -EQ $_.id_local
      }
      else {Write-Warning -Message "No Need to update additional Meta data fields for $($_.Path)"}
   }
 }
}

function Merge-GPSPoints {
    <#
      .Synopsis
        Merges a set of GPS points, producing 1 point per minute (or longer)
      .Description
        Averages the points logged each minute, or longer interval specifed by -interval
      .Example
        C:\PS>  merge-gpsPoints -points $points
        Returns the points inc $points combined to 1 average point per minute
      .Parameter Points
        An array of GPS data points
      .Parameter Interval
        The interval in minutes over which points should be averaged (default 1)
    #>
    param    ([Parameter(ValueFromPipeline=$true, Mandatory=$true)]$points, $interval=1)
    begin   {$PointsToMerge = @()     }
    process {$PointsToMerge += $points}
    end     {
         $pointsToMerge | Select-Object -Property knots, @{Name="Minute"       ; Expression={$_.dateTime.date.addhours($_.datetime.hour).addMinutes($interval*[math]::Truncate($_.dateTime.minute/$interval)).tostring("dd MMMM yyyy HH:mm")  }} ,
                                                         @{name="formattedLon" ; expression={$lon = $_.londms[0] + ($_.londms[1] /60) + ($_.londms[2] /3600)  ;  if ($_.EW -eq "W") {$lon *= -1};  $lon}} ,
                                                         @{name="formattedLat" ; expression={$lat = $_.latdms[0] + ($_.latdms[1] /60) + ($_.latdms[2] /3600)  ;  if ($_.NS -eq "S") {$Lat *= -1};  $lat}} |
                               Group-Object -Property minute  |
                                 Select-Object -Property @{name="DateTime"   ;expression={[DateTime]$_.name}},@{name="aveKnots"; expression={( $_.GROUP | MEASURE-OBJECT -Property knots -Average).AVERAGE}},
                                                         @{name="lat"        ;expression={( $_.GROUP | MEASURE-OBJECT -Property formattedLat -Average).AVERAGE}},
                                                         @{name="lon"        ;expression={( $_.GROUP | MEASURE-OBJECT -Property formattedLon -Average).AVERAGE}}
    }
}

function Get-GPSDistance {
    param ($point1 ,$point2 , $Units="KM")

    $conv =[system.math]::pi  /180
    $lat1 = $conv * $point1.lat
    $lon1 = $conv * $point1.lon

    $lat2 = $conv * $point2.lat
    $lon2 = $conv * $point2.lon

    $distanceInKm = [Math]::Acos([math]::Sin($lat1)*[math]::Sin($lat2) +  [math]::Cos($lat1)*[math]::Cos($lat2)*[math]::Cos($Lon2 - $lon1) ) *6371

    switch ($units) {
        "KM"     {$distanceInKm}
        "Miles"  {$distanceInKm * 0.6213 }
        "NM"     {$distanceInKm * 0.54   }
        "Meters" {$distanceInKm * 1000   }
    }
}

function Get-GPSBearing {
    param ($point1 ,$point2 )

    $conv =[system.math]::pi  /180
    $lat1 = $conv * $point1.lat
    $lon1 = $conv * $point1.lon

    $lat2 = $conv * $point2.lat
    $lon2 = $conv * $point2.lon

    #   tc1=mod(atan2(sin(lon2-lon1)*cos(lat2), cos(lat1)*sin(lat2)-sin(lat1)*cos(lat2)*cos(lon2-lon1)),2*pi)

    [math]::Round(((([math]::Atan2( ([math]::Sin($lon2- $lon1) * [math]::Cos($lat2))   ,
                                    [math]::Cos($Lat1)        * [math]::Sin($Lat2) - [math]::Sin($Lat1)*[math]::Cos($lat2)*[math]::Cos($Lon2 - $lon1)   ) + 2*[math]::pi ) % (2*[math]::pi )) / $conv),0)

}

function Select-SeperateGPSPoints {
    param    ([Parameter(ValueFromPipeline=$true, Mandatory=$true)]$points , $KM=0.04)
    begin   {$m= @() }
    process {$m+= $points}
    end     {0..($m.Count-1) | ForEach-Object -Begin {$previous=$m[0]}`
                                -Process {$d=(get-gpsDistance  $m[$_] $previous -Units "KM")
                                          if  ( $d -gt $KM) {Add-Member -force -InputObject $Previous -MemberType Noteproperty -Name "aveKnots" -Value ($d/(($m[$_].dateTime - $previous.DateTime ).totalhours) )
                                                             Add-Member -force -InputObject $previous -MemberType Noteproperty -Name "Bearing" -Value ([single](get-gpsBearing  $Previous $m[$_] ))
                                                              $previous
                                                              $Previous = $m[$_] }
                                         }`
                                -end  {$m[-1]  }
}}

function ConvertTo-GPX{
    <#
      .Synopsis
        Converts a set of GPS points to a GPX file to be imported by other programs
      .Description
        Converts a set of GPS points to a GPX file to be imported by other programs
      .Example
        C:\PS> merge-gpsPoints -points $points | convertto-GPX | Set-Content 2010-04-04.gpx -Encoding utf8
        Takes the result of merging the Points in $points and writes it as a GPX file.
        Note that the file will not read properly if output as unicode so -encoding UTF8 is required
      .Example
        C:\PS> $points | where {$_.paths} ) | convertto-GPX -name {Split-Path $_.paths[0] -Leaf} | out-file -Encoding utf8 -FilePath temp.gpx
        Takes the points which have a path set after using Copy-GPS with the -ReturnInfo switch,
        and outputs a file with the short file name of the the picture as the label for the data point.
        Note that the file will not read properly if output as unicode so -encoding UTF8 is required.
      .Parameter Points
        An array of GPS data points - may be passed via the the pipeline
      .Parameter Name
        A label for the data point written as a code block. The default is the dateTime specified as {$_.dateTime}
    #>
    param   (
          [Parameter(ValueFromPipeline=$true,Mandatory=$true)]$points,
          [ScriptBlock]$Name={$_.DateTime}
        )
    begin   {
                '<?xml version="1.0" encoding="UTF-8"?><gpx xmlns="http://www.topografix.com/GPX/1/1" ' +
                ' creator="PowerShell" version="1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '+
                ' xsi:schemaLocation="http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd"> '+
                ' <metadata><name>PowerShellExport</name> </metadata>'
    }
    process { $points | ForEach-Object {
                '<wpt lat="{0}" lon="{1}"><name>{2}</name> </wpt>' -f $_.lat,$_.lon,( Invoke-Command -ScriptBlock $Name)
            } }
    end     {   '</gpx>'}
}

function Out-MapPoint {
    <#
      .Synopsis
        Creates a set of PushPins in Map point from a set of GPS points
      .Description
        Creates a set of PushPins in Map point from a set of GPS points
      .Example
       C:\ps> merge-gpsPoints $points | Out-MapPoint -name  {"{0} - {1:00} MPH " -f $_.DateTime,$_.Aveknots*1.1508 } -symbol {if ($_.aveknots -lt 10) {5}  elseif ($_.aveknots -gt 40) {7} else {6}}
        Takes the result of merging the Points in $points and creates a Mappoint map
        The label is based on the the date and speed converted to MPH for example "14:00 - 30 MPH", and the symbol is colour coded based on speed
      .Example
        C:\ps> (merge-gpsPoints $points | select lat,lon,paths,datetime  ) +
                ($points | where {$_.paths} | select lat,lon,paths,datetime) | Out-MapPoint -symbol {if ($_.paths) {79} else {1}}
        Merges the GPS points, and combines them with those which have a path to a picture
        set by using copy-gps with the -return info switch.
        The "walking" points are give a red push pin and the photo sites are given a camera symbol.
        The label defaults to the date and time.
      .Example
            C:\ps> Get-NMEAData F:\copilot\gpstracks\Jul2110.gps | Merge-GPSPoints  | Select-SeperateGPSPoints -KM 0.03
            | Out-MapPoint -linkPoints -symbol {Switch ([math]::truncate((([single]$_.bearing)+22.5)/45)) { 0 {128} ; 1 {136} ; 2 {131} ; 3 {138} ;
                4 {129} ; 5 {139} ; 6{130} ; 7 {137 }; 8 {128} }}  -linecol {if     ($_.dateTime -lt [datetime]"07/21/2010 12:36:00") {16711680}
                                                                            elseif ($_.dateTime -lt [datetime]"07/21/2010 13:06:00") {65280} else {255} }
            reads , merges and selects distinct GPS points. then sets pins on the map based on the bearing and draws lines coloured for different stages of the journey
      .Parameter Points
        An array of GPS data points - may be passed via the the pipeline
      .Parameter Name
        The Pushpin Label, written as a code block. The default is the dateTime specified as {$_.dateTime}
      .Parameter Symbol
        The Pushpin type ID, written as a code block - a pin type maybe specified as  {79}
      .Parameter LinkPoints
        If Specified, lines will be drawn between the points
      .Parameter LineCol
        line colour as a written as a code block {1} = black, 255=Red, 65280=Green, 16711680=Blue, 65535=Yellow, 16711935=Magenta, 16776960=cyan
    #>
    param  (
        [Parameter(valueFromPipeLine=$true)]$points ,
        [ScriptBlock]$name={$_.DateTime} ,
        [ScriptBlock]$symbol,
        [switch]$linkPoints,
        [ScriptBlock]$LineCol={1}
    )
    begin   {
        if (-not $Global:mpapp) {$Global:MPApp = New-Object -ComObject "Mappoint.Application"}
        $Global:MPApp.Visible = $true
        $map                  = $Global:mpapp.ActiveMap
        $Script:prevLoc       = $null
    }
    process {
        $points | foreach-object {
        $location        = $map.GetLocation($_.Lat, $_.lon)
        $nameText        = Invoke-Command -ScriptBlock  $name
        if ($symbol -is [scriptblock])    {$symbolID = Invoke-Command -ScriptBlock $Symbol}
        $Pin            = $map.AddPushpin($location, $nameText)
        if ($symbol)     {$pin.symbol = $symbolID}
        if ($linkPoints) {
            if ($Script:prevLoc) {
                $line = $map.shapes.addLine($Script:prevLoc,$location)
                $line.zorder(5)
                $line.line.forecolor = Invoke-command -ScriptBlock $LineCol
                $line.line.weight=2
                $line.line.EndArrowhead = $true
            }
            $Script:prevLoc = $location
        }
    }
  }
}

function Resolve-ImagePlace {
    <#
      .Synopsis
        Queries the GeoNames Web service to translate EXIF Lat/Long information to a place name
      .Description
        Queries the GeoNames Web service to translate EXIF Lat/Long information to a place name
      .Example
        C:\ps> resolve-ImagePlace ".\IMG_1234.jpg"
        Returns the place information for the image
      .Parameter Image
        The image object or file to test - may be passed via the pipeline
    #>
    param (
        [Parameter(ValueFromPipeline=$true, Mandatory=$true)]
        [Alias("Path","FullName")]$image
    )
    if ($image -is [system.io.fileinfo] ) {$image = $image.FullName }
    if ($image -is [String]             ) {$image =(Resolve-Path -Path $image -errorAction "SilentlyContinue") | ForEach-Object {$_.path }}
    if ($Image.count -gt 1              ) {$Image | ForEach-object {Resolve-ImagePlace -image }
                                            return
    }

    $ExifData = Read-Exif -path $image
    $l=$ExifData.GPSLattitude
    if ($l.count -eq 3) {$lat  =  ($l[0]) + ($l[1] /60 ) + ($l[2] / 3600)}
    if ($l.count -eq 2) {$lat  =  ($l[0]) + ($l[1])/60 }
    if ($ExifData.GPSLatRef   -eq "S") {$lat= $lat * -1}
    $lat = [math]::Round($lat  ,  5)
    $l=$ExifData.GPSLongitude
    if ($l.count -eq 3) {$lon  =  ($l[0]) + ($l[1] / 60) + ($l[2] / 3600) }
    if ($l.count -eq 2) {$lon  =  ($l[0]) + ($l[1] / 60) }
    if ($exifdata.GPSLongRef  -eq "W") {$lon= $lon * -1}
    $lon = [math]::Round($lon,5)
    Write-Verbose -Message ("Lat {0}, Lon {1}" -f $lat,$lon)
    $url = "http://ws.geonames.org/extendedFindNearby?lat=$lat&lng=$lon"
    Write-Debug -Message $url
    if ($null -eq  $Script:WebClient) {$Script:WebClient=New-Object -TypeName System.Net.WebClient  }
    $x = (([xml]($Script:WebClient.DownloadString($url)))  )
    $x.geonames.geoname[-1] | ForEach-Object {write-verbose -Message ("Lat {0}, Lon {1}, Name {2}, ID {3}" -f $_.lat,$_.lng,$_.name,$_.geoNameID) }
    $x.geonames.geoname     | ForEach-Object -Begin {$n=""} -Process {$n = $_.name +", " +$n} -End {$n -replace '\,\s$',""}
}
