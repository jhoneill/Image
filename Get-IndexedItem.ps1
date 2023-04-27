
#Alias definitions take the form  AliasName = "Full.Cannonical.Name" ;
#Any defined here will be accepted as input field names in -filter and -OrderBy parameters
#and will be added to output objects as AliasProperties.
 $PropertyAliases   = @{Width         = "System.Image.HorizontalSize" ;       Height = "System.Image.VerticalSize";  Name   = "System.FileName" ;
                        Extension     = "System.FileExtension" ;        CreationTime = "System.DateCreated"       ;  Length = "System.Size" ;
                        LastWriteTime = "System.DateModified" ;              Keyword = "System.Keywords"          ;  Tag    = "System.Keywords"
                        CameraMaker   = "System.Photo.Cameramanufacturer" ; Software = "System.ApplicationName"}

 $FieldTypes = "System","Photo","Image","Music","Media","RecordedTv","Search","Audio"
#For each of the field types listed above, define a prefix & a list of fields, formatted as "Bare_fieldName1|Bare_fieldName2|Bare_fieldName3"
#Anything which appears in FieldTypes must have a prefix and fields definition.
#Any definitions which don't appear in fields types will be ignored
#See https://docs.microsoft.com/en-gb/windows/win32/properties/props for property info.

#https://docs.microsoft.com/en-gb/windows/desktop/properties/document-bumper
 $SystemPrefix     = "System."            ;     $SystemFields = "ItemName|ItemUrl|FileExtension|FileName|FileAttributes|FileOwner|ItemType|ItemTypeText|KindText|Kind|MIMEType|Size|DateModified|DateAccessed|DateImported|DateAcquired|DateCreated|Author|Company|Copyright|Subject|Title|Keywords|Comment|SoftwareUsed|Rating|RatingText|ApplicationName|ItemPathDisplay"
 $PhotoPrefix      = "System.Photo."      ;      $PhotoFields = "fNumber|ExposureTime|FocalLength|IsoSpeed|PeopleNames|DateTaken|Cameramodel|Cameramanufacturer|orientation"#https://docs.microsoft.com/en-gb/windows/desktop/properties/photo-bumper
 $ImagePrefix      = "System.Image."      ;      $ImageFields = "Dimensions|HorizontalSize|VerticalSize" #https://docs.microsoft.com/en-gb/windows/desktop/properties/image-bumper
 $MusicPrefix      = "System.Music."      ;      $MusicFields = "AlbumArtist|AlbumID|AlbumTitle|Artist|BeatsPerMinute|Composer|Conductor|DisplayArtist|Genre|PartOfSet|TrackNumber" #https://docs.microsoft.com/en-gb/windows/desktop/properties/music-bumper
 $AudioPrefix      = "System.Audio."      ;      $AudioFields = "ChannelCount|EncodingBitrate|PeakValue|SampleRate|SampleSize"
 $MediaPrefix      = "System.Media."      ;      $MediaFields = "Duration|Year"
 $RecordedTVPrefix = "System.RecordedTV." ; $RecordedTVFields = "ChannelNumber|EpisodeName|OriginalBroadcastDate|ProgramDescription|RecordingTime|StationName"
 $SearchPrefix     = "System.Search."     ;     $SearchFields = "AutoSummary|HitCount|Rank|Store" #https://docs.microsoft.com/en-gb/windows/desktop/properties/search-bumper

 $SelectFields     = $FieldTypes  | ForEach-Object { (Get-Variable -Name "$($_)Fields").value -split "\|" }
 $IndexFields      = $PropertyAliases.keys  + $SelectFields | Sort-Object
 $IndexFields      | ForEach-Object -Begin {$CodeFrag =  "public struct IndexedItem`r`n{"} `
                              -Process {$CodeFrag += "    public string $_;`r`n" } -end { Add-Type -TypeDefinition ($CodeFrag + "`r`n}") }

function Get-IndexedItem {
 <#
    .SYNOPSIS
        Gets files which have been indexed by Windows desktop search
    .Description
        Searches the Windows index on the local computer or a remote file serving computer
        Looking for file properties or free text searching over contents
    .PARAMETER Filter
        Alias INCLUDE
        A single string containing a WHERE condition, or multiple conditions linked with AND
        or Multiple strings each with a single Condition, which will be joined together.
        The function tries to add Prefixes and single quotes if they are omitted
        If no =, >,< , Like or Contains is specified the terms will be used in a FreeText contains search
        Syntax Information for CONTAINS and FREETEXT can be found at
        http://msdn.microsoft.com/en-us/library/dd626247(v=office.11).aspx
    .PARAMETER OrderBy
        Alias SORT
        Either a single string containing one or more Order BY conditions,
        or multiple string each with a single condition which will be joined together
    .PARAMETER Path
        A single string containing a path which should be searched.
        This may be a UNC path to a share on a remote computer
    .PARAMETER First
        Alias TOP
        A single integer representing the number of items to be returned.
    .PARAMETER Value
        Alias GROUP
        A single string containing a Field name.
        If specified the search will return the Values in this field, instead of objects
        for the items found by the query terms.
    .PARAMETER Recurse
        If Path is specified only a single folder is searched Unless -Recurse is specified
        If path is not specified the whole index is searched, and recurse is ignored.
    .PARAMETER List
        Instead of querying the index produces a list of known field names, with short names and aliases
        which may be used instead.
    .PARAMETER NoFiles
        Normally if files are found, the command returns a file object with additional properties,
        which can be piped into commands which accept files. This switch prevents the file being fetched
        improving performance when the file object is not needed.
    .EXAMPLE
        Get-IndexedItem -Filter "Contains(*,'Stingray')", "kind = 'picture'", "keywords='portfolio'" -path ~ -recurse
        Finds picture files anywhere on the current users profile, which have 'Portfolio' as a keyword tag,
        and 'stringray' in any indexed property.
    .EXAMPLE
        Get-IndexedItem Stingray, kind=picture, keyword=portfolio -recurse ~ | copy -destination e:\
        Finds the same pictures as the previous example but uses Keyword as an alias for KeywordS, and
        leaves the ' marks round Portfolio and Contains() round stingray to be automatically inserted;
        then copies the found files to drive E:
    .EXAMPLE
        Get-IndexedItem -filter stingray -path OneIndex16:// -recurse
        Finds OneNote items containing "Stingray"
        (note, nothing will be found without -recurse and the number after Index is office version specific)
    .EXAMPLE
        Get-IndexedItem -filter stingray -path ([system.environment]::GetFolderPath( [system.environment+specialFolder]::MyPictures )) -recurse
        Looks for pictures with stingray in any indexed property, limiting the scope of the search
        to the current user's 'My Pictures' folder and its subfolders.
    .EXAMPLE
        Get-IndexedItem -Filter "system.kind = 'recordedTV' " -order "System.RecordedTV.RecordingTime" -path "\\atom-engine\users" -recurse | format-list path,title,episodeName,programDescription
        Finds recorded TV files on a remote server named 'Atom-Engine' which are accessible via a share named 'users'.
        Field name prefixes are specified explicitly instead of letting the function add them.
        Results are displayed as a list using a subset of the available fields specific to recorded TV.
    .EXAMPLE
        Get-IndexedItem -Value "kind" -path \\atom-engine\users  -recurse
        Lists the kinds of files available on the on the 'users' share of a remote server named 'Atom-Engine'
    .EXAMPLE
        Get-IndexedItem -Value "title" -filter "kind=recordedtv" -path \\atom-engine\users  -recurse
        Lists the titles of RecordedTv files available on the on the 'users' share of a remote server named 'Atom-Engine'
    .EXAMPLE
        Start (Get-IndexedItem -path "\\atom-engine\users" -recurse -Filter "title= 'Formula 1' " -order "System.RecordedTV.RecordingTime DESC" -top 1 )
        Finds files entitled "Formula 1" on the 'users' share of a remote server named 'Atom-Engine'
        Selects the most recent one by TV recording date, and opens it on the local computer.
        Note: start does not support piped input.
    .EXAMPLE
        Get-IndexedItem -Filter "System.Kind = 'Music' AND AlbumArtist like '%'  " -path $null -NoFiles | Group-Object -NoElement -Property "AlbumArtist" | sort -Descending -property count
        Gets all music files with an Album Artist set, using a single combined where condition and a mixture
        of implicit and explicit field prefixes.  Setting path to Null searches the whole computer.
        The result is grouped by Artist and sorted to give popular artist first
    .EXAMPLE
        Get-IndexedItem -Filter "Kind=music","DateModified>'2012-05-31'" -NoFiles | Select-Object -ExpandProperty name
        Gets Music files which have been modified since a given date, and shows just their names.
        Note the date format; and note that the date is actually a date time, so DataModified= will only match files saved at midnight.
    .EXAMPLE
        Get-IndexedItem "itemtype='.mp3'","AlbumArtist like '%'","RatingText <> '1 star'" -NoFiles -orderby encodingBitrate,size -path $null | ft -a AlbumArtist,
            Title, @{n="size"; e={($_.size/1MB).tostring("n2")+"MB" }},@{n="duration";e={$_.duration.totalseconds.tostring("n0")+"sec"}},
            @{n="Byes/Sec";e={($_.size/128/$_.duration.totalSeconds).tostring("n0")+"Kb/s"}},@{n="Encoding";e={($_.EncodingBitrate/1000).tostring("n0")+"Kb/s"}},
            @{n="Sample Rate";e={($_.sampleRate/1000).tostring("n1")+"KHz"}}
        Shows MP3 files with Artist and Track name, showing Size, duration, actual and encoding bits per second and sample rate
    .EXAMPLE
        Get-IndexedItem -path c:\ -recurse  -Filter cameramaker=pentax* -Property focallength | group focallength -no | sort -property @{e={[double]$_.name}}
        Gets all the items which have a the camera maker set to pentax, anywhere on the C: driv
        but ONLY get thier focallength property, and return a sorted count of how many of each focal length there are.
 #>
 #$t=(Get-IndexedItem -Value "title" -filter "kind=recordedtv" -path \\atom-engine\users  -recurse | Select-List -Property title).title
 #start (Get-IndexedItem -filter "kind=recordedtv","title='$t'" -path \\atom-engine\users  -recurse | Select-List -Property ORIGINALBROADCASTDATE,PROGRAMDESCRIPTION)
 [CmdletBinding(DefaultParameterSetName='Filter')]
 [OutputType([IndexedItem],[system.io.fileinfo])]
 Param ([parameter(ParameterSetName="Filter",       Mandatory=$true,Position=0 )]
        [Alias("Include")][String[]]$Filter ,
        [parameter(Position=1)]
        [String]$Path = $pwd,
        [parameter(ParameterSetName="WhereEQ",      Mandatory=$true,Position=0)]
        [parameter(ParameterSetName="WhereNE",      Mandatory=$true,Position=0)]
        [parameter(ParameterSetName="WhereGT",      Mandatory=$true,Position=0)]
        [parameter(ParameterSetName="WhereLT",      Mandatory=$true,Position=0)]
        [parameter(ParameterSetName="WhereLike",    Mandatory=$true,Position=0)]
        [parameter(ParameterSetName="WhereContains",Mandatory=$true,Position=0)]
        [String]$Where,

        [parameter(ParameterSetName="WhereEQ",      Mandatory=$true)]
        [String]$EQ,

        [parameter(ParameterSetName="WhereNE",      Mandatory=$true)]
        [String]$NE,

        [parameter(ParameterSetName="WhereGT",      Mandatory=$true)]
        [String]$GT,

        [parameter(ParameterSetName="WhereLT",      Mandatory=$true)]
        [String]$LT,

        [parameter(ParameterSetName="WhereLike",    Mandatory=$true)]
        [String]$Like,

        [parameter(ParameterSetName="WhereContains",Mandatory=$true)]
        [String]$Contains,

        [parameter(ParameterSetName="Filter")]
        [parameter(ParameterSetName="WhereEQ")]
        [parameter(ParameterSetName="WhereNE")]
        [parameter(ParameterSetName="WhereGT")]
        [parameter(ParameterSetName="WhereLT")]
        [parameter(ParameterSetName="WhereLike")]
        [parameter(ParameterSetName="WhereContains")]
        [Alias("Sort")][String[]]$OrderBy = "ITEMURL",

        [parameter(ParameterSetName="Filter")]
        [parameter(ParameterSetName="WhereEQ")]
        [parameter(ParameterSetName="WhereNE")]
        [parameter(ParameterSetName="WhereGT")]
        [parameter(ParameterSetName="WhereLT")]
        [parameter(ParameterSetName="WhereLike")]
        [parameter(ParameterSetName="WhereContains")]
        [Alias("Top")][int]$First,

        [parameter(ParameterSetName="ValueList",Mandatory=$true)]
        [Alias("Group")][String]$Value,

        [parameter(ParameterSetName="Filter")]
        [parameter(ParameterSetName="WhereEQ")]
        [parameter(ParameterSetName="WhereNE")]
        [parameter(ParameterSetName="WhereGT")]
        [parameter(ParameterSetName="WhereLT")]
        [parameter(ParameterSetName="WhereLike")]
        [parameter(ParameterSetName="WhereContains")]
        [Alias("Select")][String[]]$Property,
        [Switch]$Recurse,

        [parameter(ParameterSetName="PropertyList",Mandatory=$true)]
        [Switch]$List,
        [parameter(ParameterSetName="Filter")]
        [parameter(ParameterSetName="WhereEQ")]
        [parameter(ParameterSetName="WhereNE")]
        [parameter(ParameterSetName="WhereGT")]
        [parameter(ParameterSetName="WhereLT")]
        [parameter(ParameterSetName="WhereLike")]
        [parameter(ParameterSetName="WhereContains")]
        [Switch]$NoFiles,
        [parameter(ParameterSetName="Filter")]
        [parameter(ParameterSetName="WhereEQ")]
        [parameter(ParameterSetName="WhereNE")]
        [parameter(ParameterSetName="WhereGT")]
        [parameter(ParameterSetName="WhereLT")]
        [parameter(ParameterSetName="WhereLike")]
        [parameter(ParameterSetName="WhereContains")]
        [Switch]$Bare,
        [String]$OutputVariable
        )

        if ($List)  {  #Output a list of the fields and aliases we currently support.
        $( foreach ($type in $FieldTypes) {
                (Get-Variable -name "$($type)Fields").value -split "\|" | select-object @{n="FullName" ;e={(Get-Variable -Name "$($type)prefix").value+$_}},
                                                                                        @{n="ShortName";e={$_}}
            }
        ) + ($PropertyAliases.keys | Select-Object  @{name="FullName" ;expression={$PropertyAliases[$_]}},
                                                    @{name="ShortName";expression={$_}}
        ) | Sort-Object -Property @{e={$_.FullName -split "\.\w+$"}},"FullName"
        return
        }

        #Make a giant SELECT clause from the field lists; replace "|" with ", " - field prefixes will be inserted later.
        #There is an extra comma to ensure the last field name is recognized and gets a prefix. This is tidied up later
        if ($EQ        -and $Where)  { $Filter +=           "$where =    $EQ"          }
        if ($NE        -and $Where)  { $Filter +=           "$where <>   $NE"          }
        if ($GT        -and $Where)  { $Filter +=           "$where >    $GT"          }
        if ($LT        -and $Where)  { $Filter +=           "$where <    $LT"          }
        if ($Like      -and $Where)  { $Filter +=           "$where LIKE $Like"        }
        if ($Contains  -and $Where)  { $Filter += "Contains( $where ,  '$Contains'  )" }
        if ($First)    {$SQL =  "SELECT TOP $First "}
        else           {$SQL =  "SELECT "}
        if ($Property) {$SQL += ($Property     -join ", ") + ", "}
        else           {$SQL += ($SelectFields -join ", ") + ", "}

        #IF a UNC name was specified as the path, build the FROM ... WHERE clause to include the computer name.
        if ($Path -match "\\\\([^\\]+)\\.") {
            $SQL += " FROM $($Matches[1]).SYSTEMINDEX WHERE "
        }
        else {$SQL += " FROM SYSTEMINDEX WHERE "}

        #If a WHERE condidtion was provided via -Filter, add it now
        if ($Filter) { #Convert * to % , unless preceded by open bracket
                $Filter = $Filter -replace "(?<!\(\s*)\*","%"
                #Insert quotes where needed any condition specified as "keywords=stingray" is turned into "Keywords = 'stingray' "
                $Filter = $Filter -replace "\s*(=|<|>|like)\s*([^\''\d][^\d\s\'']*)$"  , ' $1 ''$2'' '
                # Convert "= 'wildcard'" to "LIKE 'wildcard'"
                $Filter = $Filter -replace "\s*=\s*(?='.*%.*'\s*$)" ," LIKE " # was "\s*=\s*(?='.+%'\s*$)" ," LIKE "
                #If a no predicate was specified, use the term in a contains search over all fields.
                $Filter = ($Filter | ForEach-Object {
                                if ($_ -match "'|=|<|>|like|contains|freetext|is") {$_}
                                else {"Contains(*,'$_')"}
                })
                #if $filter is an array of single conditions join them together with AND
                $SQL += $Filter -join " AND "
        }

        #If a path was given add SCOPE or DIRECTORY to WHERE depending on whether -recurse was specified.
        if ($Path)   {
            if ($Path -notmatch "\w{4}:") {$Path = "file:" + (resolve-path -path $path).providerPath}  # Path has to be in the form "file:C:/users"
            $Path  = $Path -replace "\\","/"
            if ($SQL -notmatch "WHERE\s$") {$SQL += " AND " }                     #If the SQL statement doesn't end with "WHERE", add "AND"
            if ($Recurse)                  {$SQL += " SCOPE     = '$path' " }     #INDEX uses SCOPE <folder> for recursive search,
            else                           {$SQL += " DIRECTORY = '$path' " }     # and DIRECTORY <folder> for non-recursive
        }

        if ($Value)  {
            if ($SQL -notmatch "WHERE\s$") {$SQL += " AND " } #If the SQL statement doesn't end with "WHERE", add "AND"
                                            $SQL += " $Value is not null"
                                            $SQL =  $SQL -replace "^SELECT.*?FROM","SELECT $Value, FROM"
        }

        #If the SQL statement Still ends with "WHERE" we'd return everything in the index. Bail out instead
        if ($SQL -match "WHERE\s*$")  { Write-Warning -Message "You need to specify either a path , or a filter." ; return}

        #Add any order-by condition(s). Note there is an extra trailing comma to ensure field names are recognised when prefixes are inserted .
        if ($Value)        {$SQL  =  "GROUP ON $Value, OVER ( $SQL )"}
        elseif ($OrderBy)  {$SQL += " ORDER BY " + ($OrderBy   -join " , " ) + ","}

        # For each entry in the PROPERTYALIASES Hash table look for the KEY part being used as a field name
        # and replace it with the associated value. The operation becomes
        # $SQL  -replace "(?<=\s)CreationTime(?=\s*(=|\>|\<|,|Like))","System.DateCreated"
        # This translates to "Look for 'CreationTime' preceeded by a space and followed by ( optionally ) some spaces, and then
        # any of '=', '>' , '<', ',' or 'Like' (Looking for these prevents matching if the word is a search term, rather than a field name)
        # If you find it, replace it with "System.DateCreated"

        $PropertyAliases.Keys | ForEach-Object { $SQL= $SQL -replace "(?<=\s)$($_)(?=\s*(=|>|<|,|Is|Like))",$PropertyAliases.$_}

        # Now a similar process for all the field prefixes: this time the regular expression becomes for example,
        # $SQL -replace "(?<!\s)(?=(Dimensions|HorizontalSize|VerticalSize))","System.Image."
        # This translates to: "Look for a place which is preceeded by space and  followed by 'Dimensions' or 'HorizontalSize'
        # just select the place (unlike aliases, don't select the fieldname here) and put the prefix at that point.
        foreach ($type in $FieldTypes) {
            $Fields = (Get-Variable -Name "$($type)Fields").value
            $Prefix = (Get-Variable -Name "$($type)Prefix").value
            $SQL = $SQL -replace "(?<=\s)(?=($Fields)\s*(=|>|<|,|Is|Like))" , $Prefix
        }

        # Some commas were  put in just to ensure all the field names were found but need to be removed or the SQL won't run
        $SQL = $SQL -replace "\s*,\s*FROM\s+" , " FROM "
        $SQL = $SQL -replace "\s*,\s*OVER\s+" , " OVER "
        $SQL = $SQL -replace "\s*,\s*$"       , ""

        #Finally we get to run the query: result comes back in a dataSet with 1 or more Datatables. Process each dataRow in the first (only) table
        Write-Debug -Message $SQL
        $adapter = New-Object -TypeName system.data.oledb.oleDBDataadapter -argumentlist $sql, "Provider=Search.CollatorDSO;Extended Properties=’Application=Windows’;"
        $ds      = New-Object -TypeName system.data.dataset
        if ($adapter.Fill($ds)) {
            if ($OutputVariable) {Set-Variable -Scope 1 -Name $OutputVariable -Value $ds.Tables[0] -Visibility Public}
            else {
                foreach ($row in $ds.Tables[0])  {
                    #If the dataRow refers to a file output a file obj with extra properties, otherwise output a PSobject
                    if     ($Value) {$row | Select-Object -Property @{name=$Value; expression={$_.($ds.Tables[0].columns[0].columnname)}}}
                    elseif ($Bare)  {$row}
                    else {
                        if (($row."System.ItemUrl" -match "^file:") -and (-not $NoFiles)) {
                            $obj = (Get-Item -force -LiteralPath (($row."System.ItemUrl" -replace "^file:","") -replace "\/","\"))
                            if (-not $obj) {$obj = New-Object -TypeName psobject }
                        }
                        else {
                            if ($row."System.ItemUrl") {
                                $obj = New-Object -TypeName psobject -Property @{Path = $row."System.ItemUrl"}
                                Add-Member -Force -InputObject $obj -Name "ToString"  -MemberType "scriptmethod" -Value {$this.path}
                            }
                            else {$obj = New-Object -TypeName psobject }
                        }
                        if ($obj) {
                            #Add all the the non-null dbColumns removing the prefix from the property name.
                                    foreach ($prop in (Get-Member -InputObject $row -MemberType property | Where-Object {$row."$($_.name)" -isnot [system.dbnull] })) {
                            Add-member -ErrorAction "SilentlyContinue" -InputObject $obj -MemberType NoteProperty  -Name (($prop.name -split "\." )[-1]) -Value  $row."$($prop.name)"
                        }
                            #Add aliases
                                foreach ($prop in ($PropertyAliases.Keys | Where-Object {  ($row."$($propertyAliases.$_)" -isnot [System.DBNull] ) -and
                                                                                   ($row."$($propertyAliases.$_)" -ne $null )})) {
                            Add-Member -ErrorAction "SilentlyContinue" -InputObject $obj -MemberType AliasProperty -Name $prop -Value ($propertyAliases.$prop  -split "\." )[-1]
                        }
                            #Overwrite duration as a timespan not as 100ns ticks
                            If ($obj.duration) { $obj.duration =([timespan]::FromMilliseconds($obj.Duration / 10000) )}
                            $obj
                        }
                    }
          }
            }
        }

}

