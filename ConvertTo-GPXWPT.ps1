
function ConvertTo-GPXWpt {
    <#
        .Synopsis
        Converts a set of GPS points to a GPX file to be imported by other programs
    #>
    param   (
        #Points to plot, look for lat, lon, alt, time, Message, & 'CUSTOM.distance [m]'
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]$points,
        $Path ,
        #Minimum gap to plot a point
        $Latitude       = 'lat',
        $Longitude      = 'lon',
        $Elevation      = 'alt',
        $Time           = {[datetime]::FromOADate($d[0].time).tostring("yyyy-MM-ddTHH:mm:ss")},
        $Description    = {(($_.message -split "\(")[0] -split "\.")[0] -replace '[\d\W]*$','' -replace '^[\d\W]*',''},
        $Exclude        = '^In flight',
        $Distance       ,#= 'CUST.distance',
        $MinSpace       = 10
    )
    begin   {
        $xml       = "<?xml version=`"1.0`" encoding=`"UTF-8`"?>`r`n" +
                        "<gpx version=`"1.1`" creator=`"PowerShell Script for DJI`" xmlns=`"http://www.topografix.com/GPX/1/1`">`r`n"
        $wpt       = "  <wpt lat=""{0}"" lon=""{1}""><ele>{2}</ele><time>{3}</time><name>Point{4}</name><desc>{5}</desc></wpt>`r`n"
        $trk       = "  <trk>`r`n"+
                        "    <trkseg>`r`n"
        $trkpt     = "      <trkpt lat=""{0}"" lon=""{1}""><ele>{2}</ele><time>{3}</time></trkpt>`r`n"
        $endXML    = "    </trkseg>`r`n"+
                        "  </trk>`r`n" +
                        "</gpx>"
        $n         = 0
        $prevDist     = - $minspace
        $prevLat   = 0
        $prevLon   = 0
        $prevDesc  = ''
    }
    process {
        $points | ForEach-Object {
            if     (-not ($_.$Latitude -and $_.$Longitude)) {continue}  #skip lines without lat/lon
            if     ($Description -is [scriptblock])                     {$Desc      =  Invoke-Command $Description }
            elseif ($Description -is [string] -and $Description -ne '') {$Desc      =  $_.$Desc}
            else   {$Desc = $null}

            if     ($Time        -is [scriptblock])                     {$timeStamp =  Invoke-Command $Time }
            elseif ($Time        -is [string] -and $Time -ne '')        {$timeStamp =  $_.$Time}
            else   {$TimeStamp = $null}

            if     ($Distance    -is [scriptblock])                     {$Dist      = ((Invoke-Command $Distance) -as [double]) - $prevDist  }
            elseif ($Distance    -is [string] -and $Distance -ne '')    {$Dist      =  $_.$Distance - $prevDist }
            else   {
                #Quick pythagoras. North south = Diff in lat * Polar M per degree; East west = Diff in Long * Cos(Lat) * Equitoral M Per Degree
                #Not perfect but OK to say "Have we gone far enough to plot this point, or go on to the next one"
                $ew = ($_.$Longitude - $PrevLon) * 111320 * [math]::Cos(($_.$Longitude + $prevLat) * [math]::pi /360 )
                $ns = ($_.$Latitude  - $PrevLat) * 111133
                $Dist  = [math]::Sqrt($ew*$ew + $ns*$ns )
            }
            #If we have travelled far enough or we have label to plot ...
            if ($Dist -gt $MinSpace -or ($Desc -and $Desc -notmatch $Exclude -and $Desc -ne $prevDesc)) {
                # increment point counter and draw track line
                $n    ++
                $trk  +=    $trkpt -f $_.$Latitude, $_.$Longitude, $_.$Elevation, $timeStamp
                #If  description exists and isn't excluded (including duplicates if we've moved) plot it as a way point
                if ($Desc -and $Dev -notmatch $exclude) {
                    $xml  += $wpt   -f $_.$Latitude, $_.$Longitude, $_.$Elevation, $timeStamp, $n, $Desc
                    $prevDesc = $Desc
                }
                $prevDist, $prevLat, $PrevLon  = $Dist, $_.$Latitude,  $_.$Longitude
            }
        }
    }
    end     {
        if ($Path) {Set-Content -Value ($xml  +  $trk + $endXML) -Encoding utf8 -Path $Path  }
        else      {                    $xml  +  $trk + $endXML }
    }
}
