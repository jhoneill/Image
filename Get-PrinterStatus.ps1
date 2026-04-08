function Get-PrinterStatus {
    param ($PrinterIP  = "192.168.0.181")
    enum SNMPPrinterStatus {
        Other       = 1
        Unknown     = 2
        Idle        = 3
        Printing    = 4
        Warmup      = 5
    }
    enum SNMPDoorStaus     {
        other          = 1
        DoorOpen       = 3
        DoorClosed     = 4
        InterlockOpen  = 5
        interlockClosed= 6
    }
    enum SNMPPrintUnits    {
    TenThousandthsOfInches = 3
    Micrometers            = 4
    Characters             = 5
    Lines                  = 6
    Impressions            = 7
    Sheets                 = 8
    DotRow                 = 9
    Hours                  = 11
    Feet                   = 16
    Meters                 = 17
    }

    #I have used the SNMP COM object to save installing the SNMP module.
    $snmp               = New-Object -ComObject olePrn.OleSNMP
    $snmp.open($PrinterIP, 'public', 2, 1000)
    $printerUptime      = [timespan]::FromSeconds( $snmp.get(
                                        ".1.3.6.1.2.1.1.3.0") / 100)    # Epson does not recognise ".1.3.6.1.2.1.25.3.2.1.5"
                                    #   ".1.3.6.1.2.1.1.1.0"            # Description
                                    #   ".1.3.6.1.2.1.1.5.0"            # Hostname
    $printerFirmware    = $snmp.Get(    ".1.3.6.1.2.1.2.2.1.2.1")       # EPSON     $snmp.Get(".1.3.6.1.4.1.1248.1.2.2.2.1.1.2.1.4")
    $printerModelName   = $snmp.Get(    ".1.3.6.1.2.1.25.3.2.1.3.1")    # EPSON     $snmp.get(".1.3.6.1.4.1.1248.1.1.3.1.3.8.0")  or  1.3.6.1.2.1.43.5.1.1.16.1
    $printerStatus      = $snmp.Get(    ".1.3.6.1.2.1.43.5.1.1.3.1")  -as [SNMPPrinterStatus] # 1.3.6.1.2.1.25.3.5.1.1.1"
    $printerSerialNo    = $snmp.Get(    ".1.3.6.1.2.1.43.5.1.1.17.1")   # EPSON     $snmp.get(".1.3.6.1.4.1.1248.1.2.2.1.1.1.5.1")
    $coverNames         = $snmp.GetTree(".1.3.6.1.2.1.43.6.1.1.2.1")
    $coverStatuses      = $snmp.GetTree(".1.3.6.1.2.1.43.6.1.1.3.1")
    $covers             = for($i=0; $i -lt $coverNames.length/2; $i++){
        [pscustomobject][ordered]@{Name = $coverNames[1,$i] ; Status = ($coverStatuses[1,$i] -as [SNMPDoorStaus])}
    }
    if ($covers.Status -like "*open*") {Write-host -ForegroundColor Yellow "Cover open"}
    #$Outputdestination = $snmp.Get(    ".1.3.6.1.2.1.43.9.2.1.7.1.1")
    $printCountUnit     = $snmp.Get(    ".1.3.6.1.2.1.43.10.2.1.3.1.1") -as [SNMPPrintUnits]
    $printCountTotal    = $snmp.Get(    ".1.3.6.1.2.1.43.10.2.1.4.1.1") # EPSON  doesn't work ".1.3.6.1.4.1.1248.1.2.2.1.1.1.6"
    #$colorantsCount    = $snmp.Get(    ".1.3.6.1.2.1.43.10.2.1.6.1.1")
    $consumableNames    = $snmp.GetTree(".1.3.6.1.2.1.43.11.1.1.6.1")   # EPSON $snmp.GetTree('.1.3.6.1.4.1.1248.1.2.2.28.1.1.5.1')
    $consumableLevels   = $snmp.GetTree(".1.3.6.1.2.1.43.11.1.1.9.1")   # EPSON $snmp.GetTree('.1.3.6.1.4.1.1248.1.2.2.28.1.1.2.1')
    $consumableMaximums = $snmp.GetTree(".1.3.6.1.2.1.43.11.1.1.8.1")
    $snmp.Close()

    Write-host "$printerModelName, $printerFirmware, Serial no: $printerSerialNo. $printCountTotal $PrintCountUnit printed, Up-time: $printerUptime, Status: $printerStatus"

    for ($i=0; $i -lt $consumableNames.length/2; $i++) {[pscustomobject][ordered]@{
            Name    = $consumableNames[1,$i];
            Level   = $consumableLevels[1,$i]
            Max     = $consumableMaximums[1,$i]
    }}
}
<#
$MIB_INFO                           = [ordered]@{
#      "Print inputs"  names/status    = "1.3.6.1.2.1.43.8.2.1.13.1" / "1.3.6.1.2.1.43.8.2.1.11.1"
        "Output Tray Status"            = "1.3.6.1.2.1.43.9.2.1.10"          # prtOutputStatus
        "Error Status"                 = "1.3.6.1.2.1.43.16.5.1.2"          # prtConsoleDisplayBufferText
        "MAC Addr"                      = "1.3.6.1.4.1.1248.1.1.3.1.1.5.0"
        "Model short"                   = "1.3.6.1.4.1.1248.1.1.3.1.3.8.0"
        "IP Address"                    = "1.3.6.1.4.1.1248.1.1.3.1.4.19.1.3.1"
        "IPP_URL"                       = "1.3.6.1.4.1.1248.1.1.3.1.4.46.1.2.1"
        "WiFi"                          = "1.3.6.1.4.1.1248.1.1.3.1.29.2.1.9.0"
        "Driver"                        = "1.3.6.1.4.1.1248.1.1.3.1.29.3.1.27.0"
        "Epson device id"               = "1.3.6.1.4.1.1248.1.2.2.1.1.1.1.1"
        "Epson Printer Name"            = "1.3.6.1.4.1.1248.1.2.2.1.1.1.2.1"
        "Epson Personal Name"           = "1.3.6.1.4.1.1248.1.2.2.1.1.1.3.1"
        "LPR_URL"                       = "1.3.6.1.4.1.2699.1.2.1.3.1.1.4.1.1"
    }
#>

<#
$page           = Invoke-WebRequest "http://$PrinterIP/PRESENTATION/ADVANCED/INFO_PRTINFO/TOP" -SkipCertificateCheck
$tankRegex      = [regex]::new("<div class='tank'>.*\n.*height='(\d+)'.*\n</div>\s*<div.*?>(.*?)<","IgnoreCase")
$tankMatches    = $tankRegex.Matches($page.Content)
$h              = [ordered]@{}
$tankMatches    | ForEach-Object {
    if (-not ($label = $_.Groups[2].value)) {$label = "Waste"}
    $h[$Label] = 2* $_.Groups[1].value
}
$h
#>
