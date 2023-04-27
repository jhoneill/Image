#requires -module ImportExcel
<#
  .Synopsis
    Exports log files from a DJI drone to Excel and GPX files
  .Description
    Calls the publically available deccoder to convert a txt file saved by the remote control app to a .csv
    Reads the .csv and selects and renames columns. Exports to the data to Excel with formatting,
    and calculaton macros and/or to .GPX files which can be imported into a mapping product. The Excel
    files are .xlsM files and need to run with macros enabled to work correctly. The file is created
    from scratch and the Macro code can be reviewed below before trusting Excel to enable macros.
    Limitations. Decoding is done by other software available at :
    https://phantompilots.com/threads/tool-win-offline-txt-flightrecord-to-csv-converter.70428/
    Please read about that software there.
    Some of its output columns are blank on some models, and some have  names which are not readily
    understood, or data values which do not make sense, and this script removes the parts which are less
    useful.
    Copyright 2020 James O'Neill. The rights authors rights set out by the Copyrights Designs
    and Patents act 1988 have been asserted by him, THIS SCRIPT IS DISTRIBUTED UNDER THE MIT LICENCE
    which you can read at https://opensource.org/licenses/MIT
    In particular it is provided WITHOUT WARRANTY OF ANY KIND. As such it MUST NOT BE USED AS THE
    BASIS FOR SAFETY DECISIONS ABOUT FUTURE FLIGHTS but only to assist review of past flights, to
    that end you are encouraged to change the selection, ordering and format of fields to suit your needs.
#>
[cmdletbinding(DefaultParameterSetName='NoGPX')]
param   (
    #Path to Txt file from the controlling phone, or pre-processed CSV file.
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [Alias('DataPath','TxtPath')]
    $Path,

    #Excel file path. Defaults to DJIFlights in Documents folder. Note it contains macros so must end xlsM not xlsX.
    $XLPath         = (Join-Path -Path ([environment]::GetFolderPath([System.Environment+SpecialFolder]::MyDocuments)) -ChildPath 'DJIFlights.xlsM'),

    #Convert date in file name to a worksheet name : default would is 1JAN 16_00. (: is not allowed in a sheet name).
    #For fixed text use a string containing "" marks: for example '"day2" HH_mm'
    $WsDateFormat   = "dMMM H_mm",

    #If specified, writes GPX data to a named file (should end .GPX).
    [Parameter(ParameterSetName='GPXName',Mandatory=$true)]
    $GPXPath,

    #If specified, creates a GPX path based on the Excel folder and source file name. If neither GPXpath or AutoGPX is specified no GPX will be output.
    [Parameter(ParameterSetName='AutoGPX',Mandatory=$true)]
    [switch]$AutoGpxName,

    #If specified, only outputs lines recorded on whole seconds or with messages.
    [switch]$OneSecondSampling,

    #If specified, opens the .xlsM file in Excel after processing.
    [switch]$Show,

    #If specified, doesn't show the summary flight and drone information after processing.
    [switch]$NoSummaryInfo,

    #If specified, skips Excel output (and only outputs .GPX). Note that the preprocessor can also generate GPX files (and export embedded JPGs)
    [Parameter(ParameterSetName='GPXName')]
    [Parameter(ParameterSetName='AutoGPX')]
    [switch]$GPXOnly,

    #Location of the conversion tool  #EDIT THE DEFAULT FOR YOUR VERSION AND LOCATION
    $TxtLogToolPath = '~\Downloads\TXTlogToCSVtoolMM.exe'
)

begin   {
    # define one local function, some constants, and variables for the process block
    function ConvertTo-GPXWpt {
        <#
          .Synopsis
            Converts a set of GPS points to a GPX file to be imported by other programs
        #>
        param   (
            #Points to plot, look for lat, lon, alt, time, Message, & 'CUSTOM.distance [m]'
            [Parameter(ValueFromPipeline=$true,Mandatory=$true)]$points,
            #Minimum gap to plot a point
            $minSpace  = 10
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
            $prevD     = - $minspace
            $flyStart  = $null
        }
        process {
            $points | ForEach-Object {
                if ($null -eq $flyStart) {$flyStart = $_.flytime}
                if ($_.message -and $_.message -notmatch '^In flight') {
                    $n ++
                    $time  = $_.date.AddSeconds($_.flytime - $flystart).tostring("yyyy-MM-ddTHH:mm:ss")
                    $desc  = (($_.message -split "\(")[0] -split "\.")[0] -replace '[\d\W]*$','' -replace '^[\d\W]*',''
                    $xml  += $wpt   -f $_.lat, $_.lon, $_.alt, $time, $n, $desc
                    $trk  += $trkpt -f $_.lat, $_.lon, $_.alt, $time
                    $prevD = $_.'CUST.distance'
                }
                elseif ($_.'CUST.distance' - $prevD -gt $MinSpace) {
                    $n ++
                    $time  = $_.date.AddSeconds($_.flytime).tostring("yyyy-MM-ddTHH:mm:ss")
                    $desc  = 'Moved {0:0.0}' -f ($_.'CUST.distance' - $prevD)
                    $trk  += $trkpt -f $_.lat, $_.lon, $_.alt, $time
                    $prevD = $_.'CUST.distance'
                }
            }
        }
        end     {   $xml  +  $trk + $endXML }
    }

    #region Define functions and formulas to use them UPDATE if the column order changes  - need to be an XLSM to use the functions.
    $VBACode        = @"
Public Function Haversine(Lat1 As Double, Lon1 As Double, Lat2 As Double, Lon2 As Double)
Dim rad_lat1 As Double, rad_lat2 As Double, rad_deltaLat As Double, rad_deltaLon As Double
Dim m_Radius As Double, a As Double, rad_c As Double

m_Radius = 6371000 'of the earth, in Meters

rad_lat1 = Excel.WorksheetFunction.Radians(Lat1)
rad_lat2 = Excel.WorksheetFunction.Radians(Lat2)
rad_deltaLat = rad_lat2 - rad_lat1
rad_deltaLon = Excel.WorksheetFunction.Radians(Lon2 - Lon1)

a = (Sin(rad_deltaLat / 2) ^ 2) + (Cos(rad_lat1) * Cos(rad_lat2) * (Sin(rad_deltaLon / 2) ^ 2))

rad_c = 2 * Excel.WorksheetFunction.Atan2(Sqr(1 - a), Sqr(a))

Haversine = (m_Radius * rad_c)
End Function

Public Function Pythagoras(Lat1 As Double, Lon1 As Double, Lat2 As Double, Lon2 As Double)

Dim m_poles As Double, m_equator As Double, m_here As Double, HSqrd As Double
m_poles = 111133   'circumference round the poles
m_equator = 111320 'circumference round the equator
m_here = m_equator * Cos(Excel.WorksheetFunction.Radians((Lat2 + Lat1) / 2))

HSqrd = ((m_poles * (Lat2 - Lat1)) ^ 2) + ((m_here * (Lon2 - Lon1)) ^ 2)

Pythagoras = Sqr(HSqrd)
End Function

Public Function Bearing(Lat1 As Double, Lon1 As Double, Lat2 As Double, Lon2 As Double)

Dim rad_lat1 As Double, rad_lat2 As Double, rad_lon1 As Double, rad_lon2 As Double
Dim rad_angle As Double, p1 As Double, p2 As Double
rad_lat1 = Excel.WorksheetFunction.Radians(Lat1)
rad_lat2 = Excel.WorksheetFunction.Radians(Lat2)
rad_lon1 = Excel.WorksheetFunction.Radians(Lon1)
rad_lon2 = Excel.WorksheetFunction.Radians(Lon2)

p1 = Cos(rad_lat1) * Sin(rad_lat2) - Sin(rad_lat1) * Cos(rad_lat2) * Cos(rad_lon2 - rad_lon1)
p2 = Sin(rad_lon2 - rad_lon1) * Cos(rad_lat2)
rad_angle = Excel.WorksheetFunction.Atan2(p1, p2)

Bearing = (Excel.WorksheetFunction.Degrees(rad_angle))
End Function

"@

    #For long distance - aircraft to homepoint - we use the haversine function just defined above. Don't bother if the homepoint is blank
    $HaverSine      = '=IF(OR(AY{0}="",AZ{0}=""),"",Haversine(D{0},E{0},AY{0},AZ{0}))'
    #For short distances pythagoras is accurate and quicker, skip if previous is blank or points are the same
    $pythagoras     = '=IF(AND(D{0}<>"",E{0}<>"",D{0}<>D{1},E{0}<>E{1}), Pythagoras(D{0},E{0},D{1},E{1}) ,0)'
    #Angle between two points
    $bearingFormula = '=IF(AND(D{0}<>"",E{0}<>"",D{0}<>D{1},E{0}<>E{1}), MOD(Bearing(D{0},E{0},D{1},E{1})+360,360), "")'
    #endregion
    #    name= Name for excel Column  ; exp= expression. 'SimpleFieldName' or {powershell Script block}            calc=Excel range calculation;       fmt=@{params for Set-Format}  cnd=@{params for add-ConditionalFormat}
    $Properties     = @(
        @{name='Date'                 ; exp= {$_.'Details.Timestamp' -as [datetime]}         ;                                                         fmt=@{Num= 'Short Date'; hid=$true}}, # flight start: ensure it is sent as a datetime, format as date later
        @{name='Time'                 ; exp= {'=A{0}+((C{0}-$C$2)/86400)' -f ($rowCount+1)}  ;                                                         fmt=@{Num='HH:mm:ss.0'     }}, # start time + net flytime, format as time to .1 sec
        @{name='flytime'              ; exp= {$_.'OSD.flyTime [s]'}                ;                                                                   fmt=@{Num='0.0'             }; # flytime to .1 sec. May not start at zero if not powered off between flights
                                        cnd= $(if ($OneSecondSampling){ @{RuleType='Expression'; ConditionValue='c3-c2 <> 1'; BackgroundColor='Yellow'}} else {$null})             }, # on OneSecondSampling highlight, if gap is not 1 sec = a message or signal drop
        @{name='Lat'                  ; exp= 'OSD.latitude'                        ;                                                                   fmt=@{Num='0.000000'       }}, # GPS Given to 6 places of decimals
        @{name='Lon'                  ; exp= 'OSD.longitude'                       ;                                                                   fmt=@{Num='0.000000'       }},
        #properties with [] don't work for e=propertyname
        @{name='Alt'                  ; exp= {$_.'OSD.altitude [m]'}               ;                                                 calc='Max';       fmt=@{Num='0.0'            }}, # GPS ALT
        @{name='Height'               ; exp= {$_.'OSD.height [m]'}                 ;                                                 calc='Max';       fmt=@{Num='0.0'            }}, # Measured distance above/below home. .
        @{name='Track'                ; exp= {if ($rowCount -gt 1)  {$bearingFormula        -f  $rowCount,    ($rowCount+1) }  };                      fmt=@{Num='0'              }}, # GPS Direction from last point
        @{name='ΔDistance'            ; exp= {if ($rowCount -gt 1)  {$pythagoras            -f  $rowCount,    ($rowCount+1) }  };    calc='Sum';       fmt=@{Num='0.0;-0.0;'      }}, # GPS Distance from last point; Short distance use pythagoras format to hide 0
        @{name='SpeedOTG'             ; exp= {if ($rowCount -gt 1 -and $OneSecondSampling) {
                                                      "=If(C{1}<>C{0},I{1}/(C{1}-C{0}),0)"  -f  $rowCount,    ($rowCount+1)}
                                    elseif ($rowCount -ge 10) {($pythagoras+"/(C{1}-C{0})") -f ($rowCount-8), ($rowCount+1)}
                                    else {$null}                               };                                                    calc='Max';       fmt=@{Num='0.0;-0.0;'       }; cnd=@{DataBarColor='LightGreen' }}, # GPS Speed OTG; format to hide 0 indicate speed with green databar
        @{name='DistanceFromHome'     ; exp= {$global:rowCount ++ ;  $HaverSine             -f  $rowCount };                         calc='Max';       fmt=@{Num='0.0'             }; cnd=@{DataBarColor='Blue'       }}, # Long dist use haversine show blue databar
        @{name='Message'              ; exp= {if ($_.'APP_MESS.message'){$_.'APP_MESS.message' -replace '^[\d\W]*(.*)[\d\W]*$','$1'}
                                              elseif ($_.'CUSTOM.isVideo' -ne $prevVideo -and $_.'CUSTOM.isVideo' -eq 'Recording') {'Video Start'}
                                              elseif ($_.'CUSTOM.isVideo' -ne $prevVideo -and $_.'CUSTOM.isVideo' -eq '')          {'Video End'}
                                              elseif ($_.'CUSTOM.isPhoto') {'Photo'}
                                              else {$null}                        };                                                                   fmt=@{Width=40             }},  #don't send empty string, try to send camera info; max column width = 40
        @{name='Battery'              ; exp= 'SMART_BATTERY.battery'              }, #percentage as 0-100 int.
        @{name='GPSNum'               ; exp= 'OSD.gpsNum'                          ;                                                                   cnd=@{DataBarColor='LightGray'}  }, #Gray databar for sat strength
        @{name='XSpeed'               ; exp= {$_.'OSD.xSpeed [m/s]'               };                                                                   fmt=@{Num='0.0;-0.0;'      }}, # format to hide 0
        @{name='YSpeed'               ; exp= {$_.'OSD.ySpeed [m/s]'               };                                                                   fmt=@{Num='0.0;-0.0;'      }},
        @{name='ZSpeed'               ; exp= {$_.'OSD.zSpeed [m/s]'               };                                                                   fmt=@{Num='0.0;-0.0;'      }},
        @{name='hSpeed'               ; exp= {$_.'CUSTOM.hSpeed [m/s]'            };                                                                   fmt=@{Num='0.0'; Hid=$true }},
        @{name='CALC.hSpeed'          ; exp= {$_.'CALC.hSpeed [m/s]'              };                                                                   fmt=@{Num='0.0'; Hid=$true }}, #don't know how these are derived
        @{name='CUST.distance'        ; exp= {$_.'CUSTOM.distance [m]'            };                                                                   fmt=@{Num='0.0'; Hid=$true }},
        @{name='CALC.distance'        ; exp= {$_.'CALC.distance [m]'              };                                                                   fmt=@{Num='0.0'; Hid=$true }},
        @{name='CALC.travelled'       ; exp= {$_.'CALC.travelled [m]'             };                                                                   fmt=@{Num='0.0'; Hid=$true }},
        @{name='Pitch'                ; exp= 'OSD.pitch'                           ;                                                                   fmt=@{Num='0.0'            }},
        @{name='Roll'                 ; exp= 'OSD.roll'                            ;                                                                   fmt=@{Num='0.0'            }},
        @{name='Yaw'                  ; exp= 'OSD.yaw'                             ;                                                                   fmt=@{Num='0.0'            }},
        @{name='GimbalYaw'            ; exp= 'GIMBAL.yawAngle'                     ;                                                                   fmt=@{Num='0.0'            }},
        @{name='Aileron'              ; exp= {$_.'RC.aileron'  / 10000            };                                                                   fmt=@{Num='0%;-0%;'        }},#scale from -100000 to +10000 to percentage don't show 0
        @{name='Elevator'             ; exp= {$_.'RC.elevator' / 10000            };                                                                   fmt=@{Num='0%;-0%;'        }},
        @{name='Throttle'             ; exp= {$_.'RC.throttle' / 10000            };                                                                   fmt=@{Num='0%;-0%;'        }},
        @{name='Rudder'               ; exp= {$_.'RC.rudder'   / 10000            };                                                                   fmt=@{Num='0%;-0%;'        }},
        @{name='Gimbal'               ; exp= {$_.'RC.gimbal'   / 10000            };                                                                   fmt=@{Num='0%;-0%;'        }},
        @{name='RCGoHome'             ; exp= 'RC.goHome'                           ;                                                                   fmt=@{Num='0;-0;'          }},
        @{name='RCMode'               ; exp= 'RC.mode'                            },
        @{name='modeChannel'          ; exp= 'OSD.modeChannel'                    },
        @{name='MotorUp'              ; exp= {if ($_.'OSD.isMotorUp' -eq 'False') {$false} else {$null} };                                             fmt=@{Foreground='Red'     }}, #only send if motor is off, and make red
        @{name='GroundOrSky'          ; exp= 'OSD.groundOrSky'                     ;                                                                   cnd=@{RuleType='ContainsText';ConditionValue='Sky';    Foreground='White'} }, #Hide when flying
        @{name='GoHomeStatus'         ; exp= 'OSD.goHomeStatus'                    ;                                                                   cnd=@{RuleType='ContainsText';ConditionValue='Standby';Foreground='White'} }, #Only show when active,
        @{name='FlightAction'         ; exp= 'OSD.flightAction'                    ;                                                                   cnd=@{RuleType='ContainsText';ConditionValue='None';   Foreground='White'} }, #Hide "None"
        @{name='FlyCCommand'          ; exp= 'OSD.flycCommand'                     ;                                                                   cnd=@{RuleType='ContainsText';ConditionValue='AutoFly';Foreground='White'} }, #Hide "Autofly"
        @{name='FlyCState'            ; exp= 'OSD.flycState'                      },
        @{name='IsPhoto'              ; exp= {if ($_.'CUSTOM.isPhoto'){$_.'CUSTOM.isPhoto'} else {$null} };                                            fmt=@{Foreground='Green'   }}, #don't send empty string show green when something to say always empty?
        @{name='EnabledPhoto'         ; exp= 'CAMERA_INFO.enabledPhoto'            ;                                                                   fmt=@{Num='0;-0;'          }}, # 0 or 1 format to hide 0
        @{name='CameraIsStoring'      ; exp= 'CAMERA_INFO.isStoring'               ;                                                                   fmt=@{Num='0;-0;'          }},
        @{name='CameraIsTimePhotoing' ; exp= 'CAMERA_INFO.isTimePhotoing'          ;                                                                   fmt=@{Num='0;-0;'          }},
        @{name='IsVideo'              ; exp= {
                                          $Global:PrevVideo = $_.'CUSTOM.isVideo'   #Store the video state so we can see if it changes on next record and update message if it does
                                         if ($PrevVideo){$PrevVideo} else {$null} };                                                                   fmt=@{Foreground='Green'   }}, #Don't send empty string, show green when something to say
        @{name='CurrentVideoTime'     ; exp= 'CAMERA_INFO.videoRecordTime'        },   #blank if not recording then 1,2,3... for current vid.time
        @{name='SDCardMBFree'         ; exp= 'CAMERA_INFO.sdCardFreeSize'          ;                                                                   fmt=@{Num='#,###'          }}, #if > 1GB free want thousand sep
        @{name='ShotsRemaining'       ; exp= 'CAMERA_INFO.remainedShots'           ;                                                                   fmt=@{Num='#,###'          }}, #May have more than 1000 shots / seconds
        @{name='RemainingTime'        ; exp= 'CAMERA_INFO.remainedTime'            ;                                                                   fmt=@{Num='#,###'          }},
        @{name='IsHomeRecorded'       ; exp= {$_.'HOME.isHomeRecord' -eq 'True'}   ;                                                                   fmt=@{Foreground = 'lightgreen'}; cnd=@{RuleType='Equal'; ConditionValue=$False; Foreground='Red'}},
        @{name='HomeLat'              ; exp= 'HOME.latitude'                       ;                                                                   fmt=@{Num='0.000000'       }}, # GPS Given to 6 places of decimals
        @{name='HomeLon'              ; exp= 'HOME.longitude'                      ;                                                                   fmt=@{Num='0.000000'       }},
        @{name='HomeHeight'           ; exp= {$_.'HOME.height [m]'                };                                                                   fmt=@{Num='0.0'            }}, # pressure altitude with sea level @ 1013.25mb
        @{name='FailSafeAction'       ; exp= 'MC_PARAM.failSafeAction'             ;                                                                   fmt=@{Foreground='Red' };
                                        cnd= @{RuleType='ContainsText';ConditionValue='GoHome';Foreground='White'}                                                                 }, #Hide normal state. Show others in red
        @{name='DistanceLimitReached' ; exp= {if ($_.'HOME.isReachedLimitDistance' -eq 'True')  {$true } else {$null}};                                fmt=@{Foreground = 'Red'   }},
        @{name='HeightLimitReached'   ; exp= {if ($_.'HOME.isReachedLimitHeight'   -eq 'True')  {$true } else {$null}};                                fmt=@{Foreground = 'Red'   }},
        @{name='HOME.isBigGale'       ; exp= {if ($_.'HOME.isBigGale'              -eq 'True')  {$true } else {$null}};                                fmt=@{Foreground = 'Red'   }},
        @{name='HOME.isBigGaleWarning'; exp= {if ($_.'HOME.isBigGaleWarning'       -eq 'True')  {$true } else {$null}};                                fmt=@{Foreground = 'Red'   }},
        @{name='VoltageWarning'       ; exp= 'OSD.voltageWarning'                  ;                                                                   fmt=@{Num='[red]0;[red]-0;'}}, #show number in red. Hide 0
        @{name='Batt_UsefulSecs'      ; exp= {$_.'SMART_BATTERY.usefulTime [s]'   };                                                                   fmt=@{Num='#,###'          }},#Seconds use 1000 sepeator (maybe > 17mins)
        @{name='Batt_GoHomeSecs'      ; exp= {$_.'SMART_BATTERY.goHomeTime [s]'  }},
        @{name='Batt_LandSecs'        ; exp= {$_.'SMART_BATTERY.landTime [s]'    }},
        @{name='Batt_GoHomePercent'   ; exp= 'SMART_BATTERY.goHomeBattery'        },
        @{name='Batt_LandPercent'     ; exp= 'SMART_BATTERY.landBattery'          },
        @{name='Batt_volts'           ; exp= { $_.'SMART_BATTERY.voltage [V]'}     ;                                                                   fmt=@{Num='0.000'          }}, #Voltage given to 3 places,
        @{name='Batt_Amps'            ; exp= { $_.'CENTER_BATTERY.current [A]'}    ;                                                 calc='Max';       fmt=@{Num='0.0'            }}, #Current to one  place
        @{name='Batt_Temp'            ; exp= {($_.'CENTER_BATTERY.temperature [C]' -32)*5/9 -As [int] }},   #label says C think it is F
        @{name='Batt_SafeFlyRadius'   ; exp= 'SMART_BATTERY.safeFlyRadius'         ;                                                                   fmt=@{Num='#,###'          }}, # use thousand seperator values from mini don't make much sense
        @{name='Batt_Consume'         ; exp= 'SMART_BATTERY.volumeConsume'         ;                                                                   fmt=@{Num='#,###'          }}, # don't know what this is but >1000
        @{name='GPSLevel'             ; exp= 'OSD.gpsLevel'    },
        @{name='IsGPSused'            ; exp= {if ($_.'OSD.isGPSused'   -eq 'False') {$false} else {$null} };                                           fmt=@{Foreground='Red'     }},  #Only Show GPS off in red.
        @{name='IsVisionUsed'         ; exp= {if ($_.'OSD.isVisionUsed'-eq 'True' ) {$true } else {$null} };                                           fmt=@{Foreground='Green'   }},  #Only show vision used on in green
        @{name='IsSWaveWork'          ; exp= 'OSD.isSwaveWork' },
        @{name='SwaveHeight'          ; exp= {$_.'OSD.sWaveHeight [m]'} },
        @{name='VPSenabled'           ; exp= {$_.'MC_PARAM.VPSenabled' -eq "True"}},
        @{name='VideoSeconds'         ; exp= {'=IF(AND(AT{0}<>0,AT{1}=0),AT{0},"")'         -f  $rowCount, ($rowCount+1)};           calc='Sum'}
    )

    #region properties we'll use for selecting, or to output summary info
    $Selectprops    = @()
    foreach ($p in $properties) {$selectprops += @{n=$p.name; e=$p.exp}  }

    $FlightProps    = @( 'Filename',
        @{n='Lines';          e='DETAILS.recordLineCount'},
        @{n='Start';          e='DETAILS.timestamp'},
        @{n='Lat';            e='DETAILS.latitude'},
        @{n='Lon';            e='DETAILS.longitude'},
        @{n='Alt';            e={$_.'DETAILS.takeOffAltitude [m]'}},
        @{n='Photos';         e='DETAILS.photoNum'},
        @{n='VideoTime';      e={$_.'DETAILS.videoTime [s]'}},
        @{n='TotalTime';      e={$_.'DETAILS.totalTime [s]'}},
        @{n='TotalDistance';  e={$_.'DETAILS.totalDistance [m]'}},
        @{n='Max_Height';     e={$_.'DETAILS.maxHeight [m]'}},
        @{n='Max_HSpeed';     e={$_.'DETAILS.maxHorizontalSpeed [m/s]'}},
        @{n='Max_VSpeed';     e={$_.'DETAILS.maxVerticalSpeed [m/s]'}}
    )
    $DroneProps     = @(
        @{n='Drone Type';     e='DETAILS.droneType'},
        @{n='Actived';        e={$_.'DETAILS.activeTimestamp' -as [datetime]}},
        @{n='Aircraft Sn';    e='DETAILS.aircraftSnBytes'},
        @{n='Battery Sn';     e='DETAILS.batterySn'},
        @{n='Camera Sn';      e='DETAILS.cameraSn'},
        @{n='RC Sn';          e='DETAILS.rcSn'},
        @{n='App Version';    e='DETAILS.appVersion'}
    )
    #endregion

    $TempCsvPath    = Join-Path ([System.IO.Path]::GetTempPath()) "temp_DJI.csv"
    $firstRows      = @()
    $sheetCount     = 0
}
process {
    $file           = $Path | Get-Item
    foreach ($f in $file) { #normally there is only one $f in $file but $path might be an array or wild card
        $rawData = $selection = $null
        #region decide worksheet and GPX file name -Use the flight date stamp (if there is one) otherwise use sheet1, sheet2 etc
        $sheetCount ++
        if ($f.BaseName -match  "(\d{4}-\d{2}-\d{2}).*(\d{2}-\d{2}-\d{2})" ) {
            $date = [datetime]::ParseExact($matches[1]+$matches[2],"yyyy-MM-ddHH-mm-ss",[System.Globalization.CultureInfo]::InvariantCulture)
            $worksheetName =  $date.ToString($WsDateFormat)
            if ($AutoGpxName) {$GPXPath = Join-path (Split-Path -path $XLPath -Parent) "$worksheetName.gpx"}
        }
        else {
            $worksheetName = "Sheet$SheetCount"
            if ($AutoGpxName) {$GPXPath = Join-path (Split-Path -path $XLPath -Parent) "$($f.BaseName).gpx"}
        }
        #endregion
        #region call the converter and make sure we don't continue until it's finished, read data back in
        if ($f.Extension -eq ".csv") {$rawData = Import-Csv $f}
        else {
            Write-Progress -Activity 'Processing DJI Data' -CurrentOperation "Converting $f.name to CSV format"
            Start-Process  -Wait -FilePath $TxtLogToolPath  -ArgumentList $f.FullName, $tempCsvPath
            Write-Progress -Activity 'Processing DJI Data' -CurrentOperation "Reading back CSV file"
            $rawData = Import-Csv -Path $tempCsvPath
        }
        if (-not $rawData) {Write-error -Message "Could not import the data from $($f.FullName) "; return}
        #endregion
        #region save data for per-flight summary, process the raw data, output the GPX file; exit if -GPSOnly specified
        $firstRows += Add-Member -InputObject $rawData[0] -NotePropertyName Filename -NotePropertyValue $f.Name -PassThru

        Write-Progress -Activity 'Processing DJI Data' -CurrentOperation "Processing CSV file"
        $global:rowCount = 1
        if ( -not $OneSecondSampling){ #read the lot.
            $selection   =  $rawData | Select-Object -Property $selectprops
        }
        else { # Only keep rows for whole seconds or messages
            $selection   = $rawData | Where-Object {$_.'OSD.flyTime [s]' -eq [int]$_.'OSD.flyTime [s]' -or $_.'APP_MESS.message' -or $_.'CUSTOM.isPhoto' } |
            Select-Object -Property $selectprops
        }
        if ($GPXPath) {
            Write-Progress -Activity 'Processing DJI Data' -CurrentOperation "Saving GPX data to $GPXPath"
            $gpx = ConvertTo-GPXWpt -points $Selection
            Set-Content -Path $GPXPath -Value $gpx -Encoding utf8
        }
        if ($GPXOnly) {
            Remove-Item -Path $tempCsvPath -ErrorAction SilentlyContinue
            continue
        }
        #endregion
        #region send data and, if not there already, the VBA for Haversine etc to the Excel file
        Write-Progress -Activity 'Processing DJI Data' -CurrentOperation "Creating Spreadsheet" -Status "Creating $worksheetname in $XLPath with $($selection.count) rows"
        if (-not $xl) {
            try {$xl = Open-ExcelPackage -Path  $XLPath -Create -ErrorAction stop}
            catch {throw "Failed to open '$XLPath'." }
        }
        $xl = $selection | Export-Excel -ExcelPackage $xl -PassThru -WorksheetName $worksheetName -ClearSheet -FreezeTopRow -AutoFilter
        if (-not $xl.Workbook.VbaProject) {
            $xl.Workbook.CreateVBAProject()
            $null = $xl.Workbook.VbaProject.Modules.AddModule("Module1")
            $xl.Workbook.VbaProject.Modules["module1"].code = $VBACode
        }
        #Will need the worksheet and last row/col for formatting.
        $ws         = $xl.$worksheetName
        $lastRow    = $ws.Dimension.End.Row
        $lastCol    = $ws.Dimension.End.Column
        #endregion
        #region apply formats & conditional formats defined with properties & autosize,
        Write-Progress -Activity 'Processing DJI Data' -CurrentOperation "Applying formatting"
        $col             = 1
        $reapply         = @{} #Autosize will break hidden, fixed width, so note those.
        foreach ($p in $properties) {
            $p['Column'] = $col
            if ($p.fmt) {
                $format  = $p.fmt
                Set-Format -Address $ws.Column($col) @format
                if ($format['hid'] -or $format['width']) {$reapply[$col] = $format}
            }
            if  ($p.cnd) {
                $cFormat = $p.cnd
                $addr    = [OfficeOpenXml.ExcelAddress]::GetAddress(2, $col, $lastRow, $col)
                Add-ConditionalFormatting -WorkSheet $ws -Address $addr @cFormat
            }
            if ($p.calc) {
                $ws.Cells[($lastRow+1), $col].formula = $p.calc + '(' + [OfficeOpenXml.ExcelAddress]::GetAddress(2, $col, $lastRow, $col) +')'
            }
            $col ++
        }
        #Autosize with formatting applied
        if   ($lastRow -le 1000) {
                  $AutosizeRange = [OfficeOpenXml.ExcelAddress]::GetAddress(1, 1, $lastRow, $LastCol)
        }
        else {    $AutosizeRange = [OfficeOpenXml.ExcelAddress]::GetAddress(1, 1, 1000,     $LastCol)}
        $ws.Cells[$AutosizeRange].AutoFitColumns()
        #restore the formatting autosize just broke
        foreach ( $k   in   $reapply.keys ) {
            $format       = $reapply[$k]
            Set-Format -Address $ws.Column($K) @format
        }
        #Where columns have had color set, that will change the header, so set row 1 back to black
        Set-Format -Address $ws.Row(1) -ForegroundColor Black
        #lastrow + 1 had totals, max, min, average, whatever format to suit
        Set-Format -Address $ws.Row($lastrow + 1) -Bold -BorderTop Double
        #endregion
        #region battery is a special case, find the column and customize the conditional format to fixed numbers.
        $BatteryCol       = $properties.where({$_.name -eq "Battery"}).column
        $addr             = [OfficeOpenXml.ExcelAddress]::GetAddress(2, $BatteryCol, $lastRow, $BatteryCol)
        $cf               = Add-ConditionalFormatting -WorkSheet $ws -Address $addr -FiveIconsSet Rating -PassThru  #Battery
        $cf.Icon1.Type    =  $cf.Icon2.Type = $cf.Icon3.Type = $cf.Icon4.Type = $cf.Icon5.Type = "Num"
        #endregion
        #clean up temp file
        Remove-Item -Path $tempCsvPath -ErrorAction SilentlyContinue
    }
}
end     {
    #Save Excel file and output a summary of information about this drone
    if (-not $GPXOnly -and $xl) {
        Write-Progress -Activity 'Processing DJI Data' -CurrentOperation "Saving"
        Close-ExcelPackage $xl -Show:$show
    }
    Write-Progress -Activity 'Processing DJI Data' -CurrentOperation "Completed"
    if ($firstRows -and -not $NoSummaryInFo) {
        Write-Output "Drone data for '$($rawData[0].'DETAILS.aircraftName')' "
        $firstRows[-1] | Select-object -property $DroneProps
        $firstRows     | Select-object -property $FlightProps | Format-Table | Out-String #Better for capture / redirection
    }
}
