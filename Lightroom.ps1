#requires -modules getsql
if (-not $env:lrPath ) {
    $env:lrPath  = Join-Path -Path ([environment]::GetFolderPath([System.Environment+SpecialFolder]::MyPictures)) -ChildPath "Lightroom\Catalog-2-v12.lrcat"
}

if ($PSVersionTable.PSVersion.Major -gt 5 ) {$refAssy = "System.Xml.ReaderWriter"} else {$refAssy = "System.Xml" }
Add-Type -ReferencedAssemblies $refAssy -TypeDefinition @"
public struct LightRoomItem
{
     public double ApertureValue ;
     public string BaseName ;
     public string bitDepth ;
     public string CameraModel ;
     public string Caption ;
     public string colorChannels ;
     public string ColorLabels ;
     public string Copyright ;
     public string DateDay ;
     public string DateMonth ;
     public string DateTaken ;
     public string DateYear ;
     public string Directory ;
     public double ExposureTime ;
     public string Extension ;
     public string FileFormat ;
     public int    FlashFired ;
     public double fNumber ;
     public string FocalLength ;
     public string FullName ;
     public string GPSLatitude ;
     public string GPSLongitude ;
     public int    GrayScale ;
     public int    HasGPS ;
     public double Height ;
     public string ID_global ;
     public int    ID_local ;
     public double ISOSpeed ;
     public int    IsRawFile ;
     public string Keywords ;
     public string LensModel ;
     public string Orientation ;
     public string Path ;
     public double Rating;
     public double ShutterSpeedValue ;
     public double Width ;
     public System.Xml.XmlElement XAPDescription ;
     public string XMP ;
}
"@

function Get-LightRoomItem           {
<#
  .Synopsis
    Returns files or Folders known to lightroom
  .EXAMPLE
    Get-LightRoomItem -ListFolders -include $pwd
    Lists folders below the current one, in the LightRoom Library
  .EXAMPLE
    Get-LightRoomItem  -include "dive"
    Lists files in LightRoom Library where the path contains
    "dive" in the folder or filename
  .EXAMPLE
    Get-LightRoomItem  | Group-Object -no -Property "lensModel"  | sort count | ft -a count,name
    Produces a summary of lightroom items by lens used.
  .EXAMPLE
    Get-LightRoomItem -Where Aperture -LE f/2.8 | measure
    Counts the number of images where the apperture is below f/2.8
  .EXAMPLE
    Get-LightRoomItem -Where LensModel -Like 'smc PENTAX-DA 18-55mm F3.5-5.6 AL WR' | group -Property Directory -No | ft -AutoSize
    Produces a break down how many pictures in each folder were shot with a specific lens
  .EXAMPLE
    Get-LightRoomItem -Where ShutterSpeed -ge 1/750 | group ExposureTime -No | sort -desc @{e={[double]$_.name}} | ft Count,@{n="Name";e={"1/"+[math]::round((1/$_.name),0)}} -AutoSize
    Returns summary of the number of shots using a shutter speed of 1/750th or faster
    Note that the speed is converted from a time in seconds to a a logarithmic "shutter value" where higher numbers equal faster speeds
    So "greater" is "faster"
  .EXAMPLE
    Get-LightRoomItem -Include IM100740.DNG  -KeyWords | ft path,caption,keywords
    Gets a named file, regardless of the folder it is in, and adds the files keywords
    Keywords are not normally displayed, so in this case the full path, Caption and Keywords are shown as a table
#>
  [CmdletBinding(DefaultParameterSetName="Default")]
  [OutputType([LightRoomItem])]
  param (
     # Files to include
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [Alias("FullName")]
       [string]$Include  = "C:\" ,
     # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR"
       [alias('Path')]
       $Connection = $env:lrPath ,
     # Field to select on
       [parameter(ParameterSetName="WhereGT",  Mandatory=$true)]
       [parameter(ParameterSetName="WhereGE",  Mandatory=$true)]
       [parameter(ParameterSetName="WhereEQ",  Mandatory=$true)]
       [parameter(ParameterSetName="WhereNE",  Mandatory=$true)]
       [parameter(ParameterSetName="WhereLE",  Mandatory=$true)]
       [parameter(ParameterSetName="WhereLT",  Mandatory=$true)]
       [parameter(ParameterSetName="WhereLike",Mandatory=$true)]
       [parameter(ParameterSetName="WhereNoLi",Mandatory=$true)]
       [ValidateSet("Aperture", "CameraModel", "Caption", "ColorLabels", "Copyright", "DateTaken", "Extension", "Fileformat",  "FocalLength", "HasGPS", "keyword", "IsoSpeed", "IsRawFile",  "LensModel", "Pick", "Rating", "ShutterSpeed")]
       $Where,
       [parameter(ParameterSetName="WhereGT",  Mandatory=$true)]
       [string]$GT,
       [parameter(ParameterSetName="WhereGE",  Mandatory=$true)]
       [string]$GE,
       [parameter(ParameterSetName="WhereEQ",  Mandatory=$true)]
       [string]$EQ,
       [parameter(ParameterSetName="WhereNE",  Mandatory=$true)]
       [string]$NE,
       [parameter(ParameterSetName="WhereLE",  Mandatory=$true)]
       [string]$LE,
       [parameter(ParameterSetName="WhereLT",  Mandatory=$true)]
       [string]$LT,
       [parameter(ParameterSetName="WhereLike",Mandatory=$true)]
       [string]$Like,
       [parameter(ParameterSetName="WhereNoLi",Mandatory=$true)]
       [string]$NotLike,
     # Return values in specific fields
       [parameter(ParameterSetName="Values",Mandatory=$true)]
       [ValidateSet("Aperture", "CameraModel", "ColorLabels", "Copyright", "Extension", "Fileformat",  "FocalLength", "HasGPS", "Keyword", "IsoSpeed", "IsRawFile",  "LensModel", "ShutterSpeed")]
       $Values,
     # switch to list known folders instead of files
       [parameter(ParameterSetName="Folders",Mandatory=$true)]
       [switch]$ListFolders,
     # switch to list only exported files
       [parameter(ParameterSetName="Default"  )]
       [parameter(ParameterSetName="WhereGT"  )]
       [parameter(ParameterSetName="WhereGE"  )]
       [parameter(ParameterSetName="WhereEQ"  )]
       [parameter(ParameterSetName="WhereNE"  )]
       [parameter(ParameterSetName="WhereLE"  )]
       [parameter(ParameterSetName="WhereLT"  )]
       [parameter(ParameterSetName="WhereLike")]
       [parameter(ParameterSetName="WhereNoLi")]
       [switch]$Exports,
     # switch to list only printed files - note if specified with exports, -prints is ignored.
       [parameter(ParameterSetName="Default"  )]
       [parameter(ParameterSetName="WhereGT"  )]
       [parameter(ParameterSetName="WhereGE"  )]
       [parameter(ParameterSetName="WhereEQ"  )]
       [parameter(ParameterSetName="WhereNE"  )]
       [parameter(ParameterSetName="WhereLE"  )]
       [parameter(ParameterSetName="WhereLT"  )]
       [parameter(ParameterSetName="WhereLike")]
       [parameter(ParameterSetName="WhereNoLi")]
       [switch]$Prints,
    # switch to Add Keywords
       [parameter(ParameterSetName="Default"  )]
       [parameter(ParameterSetName="WhereGT"  )]
       [parameter(ParameterSetName="WhereGE"  )]
       [parameter(ParameterSetName="WhereEQ"  )]
       [parameter(ParameterSetName="WhereNE"  )]
       [parameter(ParameterSetName="WhereLE"  )]
       [parameter(ParameterSetName="WhereLT"  )]
       [parameter(ParameterSetName="WhereLike")]
       [parameter(ParameterSetName="WhereNoLi")]
       [Alias("Tags")]
       [switch]$KeyWords
    )
  begin {
    $lite = Test-Path $Connection
    $null = Get-SQL -Session LR -Connection $Connection -lite:$lite
  }
  process {
    foreach ($i in $Include)   {
        if      ($ListFolders) {
            $rows = Get-SQL -Session LR  -Quiet -SQL  @"
                SELECT   RootFolder.absolutePath || Folder.pathFromRoot as Path  , totals.ItemCount
                FROM     AgLibraryFolder          Folder
                JOIN     AgLibraryRootFolder      RootFolder ON RootFolder.id_local = Folder.rootFolder
                JOIN     (select count (id_global) ItemCount ,folder from AgLibraryFile group by folder) AS totals ON Totals.folder = Folder.id_local
                WHERE    RootFolder.absolutePath || Folder.pathFromRoot like '%$($i -replace "\\","/")%'
                ORDER BY Path
"@
            $rows | ForEach-Object {$_.Path = $_.Path -replace "/","\"}
            return  $rows
        }
        elseif  ($Values)      {
            switch ($Values) {
                "CameraModel"  {(Get-SQL -Session LR -Quiet -Distinct -Select value          -OrderBy  value           -Table AgInternedExifCameraModel).value}
                "Extension"    {(Get-SQL -Session LR -Quiet -Distinct -Select extension      -OrderBy  extension       -Table AgLibraryFile        ).extension}
                "Fileformat"   {(Get-SQL -Session LR -Quiet -Distinct -Select fileformat     -OrderBy  fileformat      -Table Adobe_images        ).fileformat}
                "LensModel"    {(Get-SQL -Session LR -Quiet -Distinct -Select value          -OrderBy  value           -Table AgInternedExifLens       ).value}
                "Keyword"      {(Get-SQL -Session LR -Quiet -Distinct -Select name           -OrderBy  lc_name         -Table AgLibraryKeyword          ).name}
                "ColorLabels"  {(Get-SQL -Session LR -Quiet -Distinct -Select ColorLabels    -OrderBy  ColorLabels     -Table Adobe_images `
                                                                                -Where    ColorLabe ls    -ne "''").colorlabels}
                "Copyright"    {(Get-SQL -Session LR -Quiet -Distinct -Select copyright      -OrderBy  copyright       -Table AgLibraryIPTC `
                                                                                -Where    copyright        -ne "''").copyright}
                "FocalLength"  {(Get-SQL -Session LR -Quiet -Distinct -Select FocalLength    -OrderBy  FocalLength     -Table AgharvestedExifMetadata  `
                                                                                -Where    FocalLength     "is not null").FocalLength}
                "IsoSpeed"     {(Get-SQL -Session LR -Quiet -Distinct -Select ISOSpeedRating -OrderBy  ISOSpeedRating  -Table AgharvestedExifMetadata  `
                                                        -Where  ISOSpeedRating "is not null" )  |  ForEach-Object {[int]$_.ISOSpeedRating}}
                "ShutterSpeed" {(Get-SQL -Session LR -Quiet -Distinct -Select ShutterSpeed   -OrderBy  ShutterSpeed    -Table AgharvestedExifMetadata  `
                                                        -Where  ShutterSpeed   "is not null" ) |
                                        ForEach-Object {[math]::Round((1/[math]::Pow(2,$_.ShutterSpeed)),6) } |
                                                ForEach-Object{if($_ -gt 0.25) {[math]::Round($_,2)} else {"1/" + [math]::Round((1/$_),0)} }
                            }
                "Aperture"     {(Get-SQL -Session LR -Quiet -Distinct -Select Aperture       -OrderBy  Aperture        -Table AgharvestedExifMetadata `
                                                                                -Where   Aperture       "is not null" ) |
                                        ForEach-Object {"f/" + [math]::round([math]::Sqrt([math]::Pow(2,$_.Aperture)),1) }   }
                "HasGPS"       {@($true,$false)}
                "IsRawFile"    {@($true,$false)}
                "Rating"       {@(0,1,2,3,4,5)}
                "Pick"         {@(-1,0,1)}
            }
            return
        }
        else                   {
            # admetadata.xmp removed  see https://stackoverflow.com/questions/62825586/lightroom-sqlite-database-binary-xmp-format
            $SQL = @"
                SELECT rootFolder.absolutePath || folder.pathFromRoot || rootfile.baseName || '.' || rootfile.extension       AS fullName,
                      rootFolder.absolutePath  || folder.pathFromRoot     AS directory,
                        image.fileFormat         , image.id_global      , image.id_local            , Camera.Value            AS cameraModel,
                        rootfile.extension       , image.orientation    , Image.fileWidth  AS width , image.fileHeight        AS height ,
                        metadata.dateDay         , metadata.dateMonth   , metadata.dateYear         , Image.captureTime       AS dateTaken,
                        metadata.hasGPS          , metadata.GPSLatitude , metadata.GPSLongitude     , metadata.Aperture       AS apertureValue,
                        metadata.focalLength     , metadata.flashFired  , rootfile.baseName         , metadata.ShutterSpeed   AS shutterSpeedValue,
                        IPTC.copyright           , IPTC.caption         , settings.grayscale        ,
                        image.colorLabels        , image.rating         , image.pick                , metadata.ISOSpeedRating AS ISOSpeed,
                        admetadata.israwFile     , image.bitdepth       , image.colorChannels       , LensRef.value           AS lensModel
                FROM        AgLibraryIPTC                  IPTC
                JOIN        Adobe_images                  image  ON      image.id_local =       IPTC.image
                JOIN        AgLibraryFile              rootFile  ON   rootfile.id_local =      image.rootFile
                JOIN        AgLibraryFolder              folder  ON     folder.id_local =   rootfile.folder
                JOIN        AgLibraryRootFolder      rootFolder  ON rootFolder.id_local =     folder.rootFolder
                JOIN        AgharvestedExifMetadata    metadata  ON      image.id_local =   metadata.image
                JOIN        Adobe_AdditionalMetadata adMetaData  ON      image.id_local = admetadata.image
                JOIN        Adobe_imageDevelopSettings settings  ON      image.id_local =   settings.image
                LEFT JOIN   AgInternedExifLens          LensRef  ON    LensRef.id_Local =   metadata.lensRef
                LEFT JOIN   AgInternedExifCameraModel    Camera  ON     Camera.id_local =  metadata.cameraModelRef
"@
        if     ($Exports -or $Prints) {$SQL = ($SQL -replace "SELECT ", "SELECT DISTINCT ") + @"

                JOIN        Adobe_libraryImageDevelopHistoryStep history on  image.id_local = history.image
                WHERE       history.name LIKE '$(if ($Exports) {'Export'} elseif ($Prints) {'Print'})%'
                  AND
"@
                if ($Exports -and $Prints)          {Write-Warning "If -Exports and -Prints are both specified, -Prints is ignored."}
        }
        elseif ($Where -eq "Keyword" -and -not $EQ) {Write-Warning "Keywords only work with -EQ"}
        elseif ($Where -eq "Keyword") {
                        $SQL = $SQL + @"

                JOIN        AgLibraryKeywordImage       keyword  ON      image.id_local =    keyword.image
                JOIN        AgLibraryKeyword            tags     ON       tags.id_local =    keyword.tag
                WHERE       tags.lc_name = '$($EQ.toLower())'
                  AND
"@      }
        else          { $SQL = $SQL +  @"

                WHERE
"@      }
             $SQL = $SQL + "       RootFolder.absolutePath || Folder.pathFromRoot || RootFile.baseName || '.' || RootFile.extension  like '%$($i -replace "\\","/" -replace '^\.\\|^\./','')%' " +
                           "`r`n                ORDER BY    FullName"
        }
        if      ($Where -and $where -ne 'Keyword')       {
            $Condition = ($GT + $GE +  $EQ + $NE + $LE + $LT + $Like + $NotLike )
            switch ($Where) {
                "FocalLength"  { $Condition = [convert]::ToDouble($Condition)    }
                "Pick"         { $Condition = [convert]::ToInt16( $Condition)    }
                "Rating"       { $Condition = [convert]::ToDouble($Condition)    }
                "HasGPS"       { if ([convert]::ToBoolean($eq)) { $Condition = 1 }
                                else                           { $Condition = 0 }
                            }
                "IsRawFile"    { if ([convert]::ToBoolean($eq)) { $Condition = 1 }
                                else                           { $Condition = 0 }
                            } #TimeValue = log2  1/ExposureTimeInSeconds ; stored to 6 digits
                "ShutterSpeed" { if ($Condition -match "^1/" ) {$Condition = [math]::Round([math]::Log(  ([convert]::ToDouble(($Condition -replace "^1/",""))),2),6)}
                                else                          {$Condition = [math]::Round([math]::Log((1/[convert]::ToDouble( $Condition                   )),2),6)}
                            } #ApertureValue = Log2 ( fnumber squared)  ; stored to 6 digits
                "Aperture"     { $Condition = [math]::round([math]::log([math]::Pow(($Condition -replace "f/",""),2),2),6) }
                "DateTaken"    { $Condition = [convert]::ToDateTime($Condition).ToString("yyyy-MM-dd") }
            }
            if   ($Condition -is [string]) {
                  $Condition = "'$($Condition.ToUpper())'"
                  $w     = "upper($where)"
            }
            else {$w   = $where}
            $SQL = $SQL -replace "(?<=Where)\s*",$( switch ($PSCmdlet.ParameterSetName) {
                "WhereGT"   {  "       $w >        {0} AND " -f  $Condition }
                "WhereGE"   {  "       $w >=       {0} AND " -f  $Condition }
                "WhereEQ"   {  "       $w  =       {0} AND " -f  $Condition }
                "WhereNE"   {  "       $w <>       {0} AND " -f  $Condition }
                "WhereLE"   {  "       $w <=       {0} AND " -f  $Condition }
                "WhereLT"   {  "       $w <        {0} AND " -f  $Condition }
                "WhereLike" { ("       $w LIKE     {0} AND " -f ($Condition -replace "\*","%") ) }
                "WhereNoLi" { ("       $w NOT LIKE {0} AND " -f ($Condition -replace "\*","%") ) }
            })
        }

        $rows = @() + (Get-SQL -Session LR -Quiet -SQL $SQL)
        Write-Verbose -Message "$SQL `r`nReturned $($rows.count) rows"
     #  removed Add-Member -PassThru -MemberType ScriptProperty -name "XAPDescription" -Value {([xml]($this.xmp)).xmpmeta.RDF.Description } |
        $rows | Add-Member -PassThru -MemberType ScriptProperty -name "Path"           -Value {$this.fullname -replace "/","\"}   |
                Add-Member -PassThru -MemberType ScriptProperty -name "FNumber"        -Value {[math]::Round([math]::Sqrt([math]::Pow(2,$this.ApertureValue)),1)} |
                Add-Member -PassThru -MemberType ScriptProperty -name "ExposureTime"   -Value {[math]::Round((1/[math]::Pow(2,$this.ShutterSpeedValue)),6)} |
            ForEach-Object {
                $_.pstypenames.add("LightRoomItem") ;
                if (-not $KeyWords) {$_}
                else {
                    $k =  (Get-SQL -session lr -Quiet -SQL (@"
                            SELECT  AgLibraryKeyword.Name
                            FROM    AgLibraryKeywordImage
                            JOIN    AgLibraryKeyword on AgLibraryKeyword.id_local = AgLibraryKeywordImage.Tag
                            WHERE   image  =  $($_.id_local)
"@                        )).name -join "; "
                    Add-Member -InputObject $_ -PassThru -NotePropertyName "Keywords" -NotePropertyValue $K
                }
            }
    }
  }
}

function Get-LightRoomCollectionItem {
  <#
  .Synopsis
    Returns Collection items known to LightRoom
  .EXAMPLE
    Get-LightRoomCollectionItem | out-gridview
    Displays all items in all collections
  .EXAMPLE
    Get-LightRoomCollectionItem -include music
    Displays items in the music collection
  .EXAMPLE
    Get-LightRoomCollectionItem -include music | copy -Destination e:\raw\music
    Copies the original files in the music collection
#>
  [CmdletBinding()]
  [OutputType([System.Data.DataRow])]
  param (
     # Collections to include
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [String]$Include  = "",
     # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR" or "Driver={SQLite3 ODBC Driver};Database=<<path>>"
       [alias('Path')]
       $Connection = $env:lrPath
  )
  begin {
    $lite = Test-Path $Connection
    $null = Get-SQL -Session LR -Connection $Connection -lite:$lite
  }
  process { foreach ($i in $Include) {
    $SQL = @"
      SELECT   Collection.name AS CollectionName ,  image.id_global,       image.id_local,  RootFile.baseName,
               RootFolder.absolutePath || Folder.pathFromRoot || RootFile.baseName || '.' || RootFile.extension AS FullName
      FROM     AgLibraryCollection      Collection
      JOIN     AgLibraryCollectionimage cimage     ON collection.id_local = cimage.Collection
      JOIN     Adobe_images             Image      ON      Image.id_local = cimage.image
      JOIN     AgLibraryFile            RootFile   ON   Rootfile.id_local = image.rootFile
      JOIN     AgLibraryFolder          Folder     ON     folder.id_local = RootFile.folder
      JOIN     AgLibraryRootFolder      RootFolder ON RootFolder.id_local = Folder.rootFolder
      WHERE    Collection.name     LIKE '$i%'
      ORDER BY CollectionName, FullName
"@
    $rows = @() + (Get-SQL -Session LR -Quiet -SQL $SQL)
    Write-Verbose -Message "$SQL `r`nReturned $($rows.count) rows"
    $rows | Add-Member -MemberType ScriptProperty -Name "Path" -Value {$this.FullName -replace "/","\"} -PassThru
  }}
}

function Get-LightRoomCollection     {
<#
  .Synopsis
    Returns Lightroom Collections
  .EXAMPLE
    Get-LightRoomCollection
    Lists all collections
  .EXAMPLE
    Get-LightRoomCollection -include unsaved
    Lists collections with names that begin "unsaved"
#>
  [CmdletBinding()]
  [OutputType([System.Data.DataRow])]
  param(
     # Collections to include
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [String]$Include  = "" ,
     # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR" or "Driver={SQLite3 ODBC Driver};Database=<<path>>"
       [alias('Path')]
       $Connection = $env:lrPath
  )
  begin {
    $lite = Test-Path $Connection
    $null = Get-SQL -Session LR -Connection $Connection -lite:$lite
  }
  process {
    foreach ($i in $Include) {
        $SQL = @"
            SELECT   Collection.name AS CollectionName
            FROM     AgLibraryCollection      Collection
            WHERE    Collection.name     LIKE '$i%'
            ORDER BY CollectionName
"@
        if (-not $Connection -and $Path) {$Connection="Driver={SQLite3 ODBC Driver};Database=$Path"}
        $rows = @() + (Get-SQL -Session LR -Quiet -SQL $SQL)
        Write-Verbose -Message "$SQL `r`nReturned $($rows.count) rows"
        $rows
  }}
}

# AgLibraryCollectionContent defines search / sort settings
# AgLibraryImport / AgLibraryImportimage
function Add-LightRoomCollectionItem  {
   <#
  .Synopsis
    Adds existing Lightroom items to a Collection
  .EXAMPLE
     dir raw | Where-Object {Test-Path ".\$($_.basename)*ig*.jpg" } | Get-LightRoomItem | Add-LightRoomCollectionItem -Collection "IG Uploads"
      Looks in the RAW subdirectory for files which have been saved in the current folder
      gets the Lightroom items corresponding to the these files,
      and inserts them into the IG uploads collection
    #>
    [CmdletBinding()]
    param(
     # Files to include
       [Parameter(ValueFromPipeline=$true,mandatory=$true)]
       [System.Data.DataRow[]]$InputObject,
       [String]$Collection ,
     # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR" or "Driver={SQLite3 ODBC Driver};Database=<<path>>"
       [alias('Path')]
       $Connection = $env:lrPath
    )
    begin {
        $lite  = Test-Path $Connection
        $null  = Get-SQL -Session LR -Connection $Connection -lite:$lite

        $col   = Get-SQL -Session LR -Table "AgLibraryCollection" -Select id_local,Name -Where "name" -EQ $collection -Quiet
        $nextid = 2 + ( Get-SQL -Session LR -sql 'select max(id_local) as maxid from AgLibraryCollectionImage' -Quiet ).maxid
        if (-not $col.id_local) {throw "Could not find ID for collection $Collection"; return}
        $existingItems =  @{}
        Get-LightRoomCollectionItem -Include $Collection | ForEach-Object {$existingItems[$_.FullName] = $true}
    }
    process {
        if ($existingItems[$InputObject.fullName]) {Write-Verbose "$Collection contained    $($InputObject.fullName)"}

        else {
            Get-Sql -Session LR -SQL @"
                Insert into AgLibraryCollectionimage (id_local, collection, image, pick )
                Values ($nextid , $($col.id_local), $($InputObject.id_local),0 )
"@          -Quiet
            $nextid += 2
            $existingItems[$InputObject.fullName] = $true
            Write-Verbose "$Collection now contains $($InputObject.fullName)"
        }
    }
}

function Convert-LRSerialToDate       {
    param ( $value)

    [datetime]::new(2001,1,1,0.0,0, [System.DateTimeKind]::Utc).AddSeconds($value).ToLocalTime()
}

# AgLibraryKeyword   JOINs via AgLibraryKeywordImage to Adobe_images
function Get-LightRoomKeyword        {
<#
  .Synopsis
    Returns Lightroom Keywords
  .EXAMPLE
    Get-LightRoomKeyword
    Lists all keywords
  .EXAMPLE
    Get-LightRoomKeyword oxfordshire*
    Lists keywords with Oxfordshire, Oxfordshire/Abingdon, Oxfordshire/Banbury, Oxfordshire/Oxford
#>
    [CmdletBinding()]
    [OutputType([System.Data.DataRow])]
    param(
       [Parameter(ValueFromPipeline=$true)]
       $Keyword = '*',
       # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR" or "Driver={SQLite3 ODBC Driver};Database=<<path>>"
       [alias('Path')]
       $Connection = $env:lrPath
    )
    begin {
        $lite = Test-Path $Connection
        $null = Get-SQL -Session LR -Connection $Connection -lite:$lite
    }
    process {    foreach ($k in $Keyword) {
        $K = $k.toLower() -replace '\*','%'
        $SQL = @"
            SELECT     name, AgLibraryKeyword.id_local, lastApplied, Count(AgLibraryKeywordImage.id_local) As Images
            FROM       AgLibraryKeyword
            LEFT JOIN  AgLibraryKeywordImage On AgLibraryKeywordImage.tag = AgLibraryKeyword.id_local
            WHERE      lc_name     LIKE '$k'
            GROUP BY   AgLibraryKeyword.id_local
"@
        if (-not $Connection -and $Path) {$Connection="Driver={SQLite3 ODBC Driver};Database=$Path"}
        $rows = @() + (Get-SQL -Session LR -Quiet -SQL $SQL)
        Write-Verbose -Message "$SQL `r`nReturned $($rows.count) rows"
        $rows | Select-Object -Property @{n='Name';e='name'}, id_local, Images , @{n='LastApplied';e={Convert-LRSerialToDate $_.lastApplied}} |
          ForEach-Object {$_.pstypenames.add("LightRoomKeyword") ; $_ }  | Sort-Object -Property Name
  }}

}

function Get-LightRoomKeywordItem    {
<#
  .Synopsis
    Returns Lightroom Items that match a keyword
  .EXAMPLE
    Get-LightRoomKeywordItem Ferrari
    Lists Lightroom items with the keyword "Ferrari"
  .EXAMPLE
    Get-LightRoomKeywordItem Ferrari,Mecedes
    List lightroom items with either or both of the keywords (any with both will be listed twice)
  .EXAMPLE
   Get-LightRoomKeyword | where images -gt 0 | where lastapplied -lt (get-date -year 2018) | Get-LightRoomKeywordItem
   Finds keyword tags which are in use in the catalog, but have not been used since this date in 2018,
   and displays the pictures using those tags.
#>
    [CmdletBinding()]
    [OutputType([System.Data.DataRow])]
    param(
       [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
       $Keyword,
       # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR" or "Driver={SQLite3 ODBC Driver};Database=<<path>>"
       [alias('Path')]
       $Connection = $env:lrPath
    )
    begin {
        $lite = Test-Path $Connection
        $null = Get-SQL -Session LR -Connection $Connection -lite:$lite
    }
    process { foreach ($k in $Keyword) {
        if ($k -is [string]) {
                Get-LightRoomKeyword $K | Get-LightRoomKeywordItem
        }
        elseif ($k.id_Local) {$k = $k.id_local}
        if ($k -is [valuetype]) {
            # admetadata.xmp removed  see https://stackoverflow.com/questions/62825586/lightroom-sqlite-database-binary-xmp-format
            $SQL = @"
                SELECT rootFolder.absolutePath || folder.pathFromRoot || rootfile.baseName || '.' || rootfile.extension       AS fullName,
                        rotFolder.absolutePath  || folder.pathFromRoot     AS directory,
                        image.fileFormat         , image.id_global      , image.id_local            , Camera.Value            AS cameraModel,
                        rootfile.extension       , image.orientation    , Image.fileWidth  AS width , image.fileHeight        AS height ,
                        metadata.dateDay         , metadata.dateMonth   , metadata.dateYear         , Image.captureTime       AS dateTaken,
                        metadata.hasGPS          , metadata.GPSLatitude , metadata.GPSLongitude     , metadata.Aperture       AS apertureValue,
                        metadata.focalLength     , metadata.flashFired  , rootfile.baseName         , metadata.ShutterSpeed   AS shutterSpeedValue,
                        IPTC.copyright           , IPTC.caption         , settings.grayscale        ,
                        image.colorLabels        , image.rating         , image.pick                , metadata.ISOSpeedRating AS ISOSpeed,
                        admetadata.israwFile     , Image.bitdepth       , image.colorChannels       , LensRef.value           AS lensModel,
                        tags.name AS keyword
                FROM        AgLibraryIPTC                  IPTC
                JOIN        Adobe_images                  image  ON      image.id_local =       IPTC.image
                JOIN        AgLibraryFile              rootFile  ON   rootfile.id_local =      image.rootFile
                JOIN        AgLibraryFolder              folder  ON     folder.id_local =   rootfile.folder
                JOIN        AgLibraryRootFolder      rootFolder  ON rootFolder.id_local =     folder.rootFolder
                JOIN        AgharvestedExifMetadata    metadata  ON      image.id_local =   metadata.image
                JOIN        Adobe_AdditionalMetadata adMetaData  ON      image.id_local = admetadata.image
                JOIN        Adobe_imageDevelopSettings settings  ON      image.id_local =   settings.image
                JOIN        AgLibraryKeywordImage       keyword  ON      image.id_local =    keyword.image
                JOIN        AgLibraryKeyword            tags     ON       tags.id_local =    keyword.tag
                LEFT JOIN   AgInternedExifLens          LensRef  ON    LensRef.id_Local =   metadata.lensRef
                LEFT JOIN   AgInternedExifCameraModel    Camera  ON     Camera.id_local =  metadata.cameraModelRef
                WHERE       keyword.tag = $K
"@
            $rows = @() + (Get-SQL -Session LR -Quiet -SQL $SQL)
            Write-Verbose -Message "$SQL `r`nReturned $($rows.count) rows"
            # REMOVED Add-Member -PassThru -MemberType ScriptProperty -name "XAPDescription" -Value {([xml]($this.xmp)).xmpmeta.RDF.Description } |
            $rows | Add-Member -PassThru -MemberType ScriptProperty -name "Path"           -Value {$this.fullname -replace "/","\"}   |
                    Add-Member -PassThru -MemberType ScriptProperty -name "FNumber"        -Value {[math]::Round([math]::Sqrt([math]::Pow(2,$this.ApertureValue)),1)} |
                    Add-Member -PassThru -MemberType ScriptProperty -name "ExposureTime"   -Value {[math]::Round((1/[math]::Pow(2,$this.ShutterSpeedValue)),6)} |
                ForEach-Object {$_.pstypenames.add("LightRoomItem") ; $_ }
        }
  }}
}

function Add-LightRoomItemKeyword    {
    <#
      .Synopsis
        Adds a keyword to existing Lightroom items
      .EXAMPLE
        Add-LightRoomItemKeyword -Include IM104489.DNG -Keyword "Black and White"
        Adds the keyword "Black and white" to the the matching file(s) - in this case one file name is given.
      .EXAMPLE
        Get-LightRoomItem -Include IM10447%.dng | Where-Object GrayScale -eq 1 |  Add-LightRoomItemKeyword -Keyword "Black and White" -Verbose
        Gets images IM104470 to IM104479 and filters them to gray-scale images, then adds the keyword  "Black and White" to those selected.
    #>
    [CmdletBinding(DefaultParameterSetName='Include')]
    param(
     # Files to include
       [Parameter(ParameterSetName='Include',Mandatory=$true,Position=0)]
       $Include,
    #Items passed from Get-Lightroomitem
       [Parameter(ParameterSetName='IO',ValueFromPipeline=$true,Mandatory=$true)]
       [System.Data.DataRow[]]$InputObject,
       [Parameter(Mandatory=$true,Position=1)]
       [string]$Keyword,
     # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR" or "Driver={SQLite3 ODBC Driver};Database=<<path>>"
       [alias('Path')]
       $Connection = $env:lrPath
    )
    begin {
        $lite  = Test-Path $Connection
        $null  = Get-SQL -Session LR -Connection $Connection -lite:$lite
        $kw    = Get-SQL -Session LR -Table "AgLibraryKeyword" -Select id_local,name -Where "name" -EQ $Keyword -Quiet
        $nextid = 2 + ( Get-SQL -Session LR -sql 'select max(id_local) as maxid from AgLibraryKeywordImage' -Quiet ).maxid
        if (-not $kw.id_local) {throw "Could not find ID for keyword $keyword"; return}
    }
    process {
        if ($Include -and -not $InputObject) {
            Get-LightRoomItem -Include $Include -verbose:$false | Add-LightRoomItemKeyword -Keyword $Keyword
            return
        }
        $existingKeyword = Get-Sql -Session LR -Quiet -SQL ("SELECT * FROM AgLibraryKeywordImage " +
                                                           " WHERE tag = $($kw.id_local) and image = $($InputObject.id_local) ")
        if ( $existingKeyword ) {Write-Verbose "Keyword '$Keyword' is already present on file $($inputobject.Path)"}
        else {
            Get-Sql -Session LR -SQL @"
                Insert into AgLibraryKeywordImage (id_local, tag, image )
                Values ($nextid , $($kw.id_local), $($InputObject.id_local) )
"@          -Quiet
            $nextid += 2
            Write-Verbose "Keyword '$Keyword' added to $($inputobject.Path)"
        }
    }
}

function Rename-LightRoomItem        {
  <#
  .Synopsis
    Renames items in LightRoom database and on disk
  .EXAMPLE
    Rename-LightRoomItem -include $pwd\IMG*.DNG -replace "IMG","DIVE" -Verbose
    Renames files with a DNG extension from IMGxxxxx.DNG to DIVExxxxx.DNG
#>
  [CmdletBinding(SupportsShouldProcess=$True)]
  param (
     # Files to include
       [Parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true)]
       [string]$Include,
     # Replace what with what in the name e.g IMG_,DIVE
       [Parameter(Mandatory=$true)]
       [string[]]$Replace,
     # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR" or "Driver={SQLite3 ODBC Driver};Database=<<path>>"
       [alias('Path')]
       $Connection = $env:lrPath
  )
  begin {
    $lite = Test-Path $Connection
    $null = Get-SQL -Session LR -Connection $Connection -lite:$lite
  }
  process {
    foreach ($i in $Include) {
        $SQL = @"
        SELECT rootFolder.absolutePath || folder.pathFromRoot || rootfile.baseName || '.'  || rootfile.extension AS fullName,
        rootfile.id_local, baseName, extension	, folder, idx_filename, importHash, lc_idx_filename, lc_idx_filenameExtension, originalFilename
        FROM   AgLibraryFile              rootFile
        JOIN   AgLibraryFolder            folder     ON     folder.id_local = rootfile.folder
        JOIN   AgLibraryRootFolder        rootFolder ON rootFolder.id_local =   folder.rootFolder
        WHERE  RootFolder.absolutePath || Folder.pathFromRoot || RootFile.baseName || '.' || RootFile.extension  like '%$($i -replace "\\","/" -replace "\*","%" )%'
        ORDER BY FullName
"@
        $rows = @() + (Get-SQL -Session LR -Quiet -SQL $SQL)
        Write-Verbose -Message "$SQL `r`nReturned $($rows.count) rows"
        $rows  | ForEach-Object {
                $SQL      = "UPDATE AgLibraryFile SET basename='{0}', idx_filename='{1}', importHash='{2}', lc_idx_filename='{3}', originalFilename='{4}' WHERE id_local={5} " -f ($_.basename -replace $Replace) ,($_.idx_filename -replace $Replace) , ($_.importHash -replace $Replace) , ($_.lc_idx_filename -replace $Replace).ToLower()  , ($_.originalFilename -replace $Replace), ($_.id_local)
                $SQL      = $SQL -replace ",importHash=\s*,"," ,"
                Write-Verbose -Message $SQL
                if ($psCmdlet.ShouldProcess($_.originalFilename , "Update")) {
                    Get-SQL -Confirm:$false -Session LR -sql $sql -Quiet
                    Rename-Item -Path ($_.fullname -replace "/","\") -NewName ($_.idx_filename -replace $Replace)
                }
        }
    }
  }
}

function Set-LightRoomItemColor      {
<#
  .Synopsis
    Sets the ColorLabels attribute on Lightroom images.
  .EXAMPLE
    Get-LightRoomItem -include $pwd -Exports | Set-LightRoomItemColor -colour "Red" -Verbose
    Puts all exported files from the current folder into the "Red" group.
  .EXAMPLE
    Get-LightRoomCollectionItem -include music  | Set-LightRoomItemColor -colour "Green" -Verbose
    Puts all files in the music collection current folder into the "Green" group.
#>
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
     # Files to include
       [Parameter(ValueFromPipeline=$true,mandatory=$true)]
       [System.Data.DataRow[]]$InputObject,
     # Path to the LightRoom catalog file
       [Parameter(mandatory=$true)][ValidateSet("Red", "Yellow", "Green", "Blue", "Purple")]
       [String]$Colour ,
     # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR" or "Driver={SQLite3 ODBC Driver};Database=<<path>>"
       [alias('Path')]
       $Connection = $env:lrPath
  )
  begin {
    $lite = Test-Path $Connection
    $null = Get-SQL -Session LR -Connection $Connection -lite:$lite
  }
  process {
    $InputObject | ForEach-Object {
        If ($psCmdlet.ShouldProcess($_.Basename , "Update to $Colour")) {
            Get-SQL -Session LR -Confirm:$false -Table Adobe_images -Where id_local -eq $_.id_local -Set colorLabels -Values $Colour
        }
    }
  }
}

function Set-LightRoomItemFlag       {
<#
  .Synopsis
    Sets the Pick (flag) attribute on Lightroom images: -1 = Rejected, 1 = flagged, 0 = Unflagged
  .EXAMPLE
    Get-LightRoomItem -include $pwd -Exports | Set-LightRoomItemFlag -Flag 1  -Verbose
    Flags all exported files from the current folder
#>
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
     # Files to include
       [Parameter(ValueFromPipeline=$true,mandatory=$true)]
       [System.Data.DataRow[]]$InputObject,
     # Path to the LightRoom catalog file
       [Parameter(mandatory=$true)][ValidateRange(-1,1)]
       [Int]$Flag ,
     # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR" or "Driver={SQLite3 ODBC Driver};Database=<<path>>"
       [alias('Path')]
       $Connection = $env:lrPath
  )
  begin {
    $lite = Test-Path $Connection
    $null = Get-SQL -Session LR -Connection $Connection -lite:$lite
  }
  process {
    $InputObject | ForEach-Object {
        If ($psCmdlet.ShouldProcess($_.Basename , "Update 'pick' flag to $Flag")) {
            Get-SQL -Session LR -Confirm:$false -Table Adobe_images -Where id_local -eq $_.id_local -Set pick -Values $flag
        }
    }
  }
}

function Set-LightRoomItemRating     {
<#
  .Synopsis
    Sets the Rating (Stars) attribute on Lightroom images from 0 to 5
  .EXAMPLE
    Get-LightRoomItem -include $pwd -Exports | Set-LightRoomItemRating -Stars 5  -Verbose
    Gives 5 starts to exported files from the current folder
#>
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
     # Files to include
       [Parameter(ValueFromPipeline=$true,mandatory=$true)]
       [System.Data.DataRow[]]$InputObject,
     # Path to the LightRoom catalog file
       [Parameter(mandatory=$true)][ValidateRange(0,5)]
       [Int]$Stars ,
     # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR" or "Driver={SQLite3 ODBC Driver};Database=<<path>>"
       [alias('Path')]
       $Connection = $env:lrPath
  )
  begin {
    $lite = Test-Path $Connection
    $null = Get-SQL -Session LR -Connection $Connection -lite:$lite
  }
  process {
    $InputObject | ForEach-Object {
        If ($psCmdlet.ShouldProcess($_.Basename , "Update rating to $Stars")) {
            Get-SQL -Session LR -Confirm:$false -Table Adobe_images -Where id_local -eq $_.id_local -Set Rating -Values $Stars
        }
    }
  }
}

function Test-LightRoomItem          {
  <#
.Synopsis
       Returns files or Folders known to LightRoom
.EXAMPLE
      Test-LightRoomItem .\IMG_4704.DNG
      Returns true if the IMG_4704.DNG in the current folder is in the lightroom database
.EXAMPLE
      Test-LightRoomItem *.dng,*.jpg -Passthru -not | move -Destination ..\Scrap -WhatIf
      Gets JPG & DNG files in the current folder and moves those which are not in lightroom to a "Scrap" folder.
      -Whatif allows the files to be confirmed before being moved.
#>
  [CmdletBinding()]
  [OutputType([Boolean])]
  Param(
      # Path to check
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias("FullName")]
        $Path ,
      # Pass matching files on to the next step
        [switch]$Passthru,
        # Invert the selection return true / file if it is not in the lightroom DB
        [switch]$Not,
      # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR" or "Driver={SQLite3 ODBC Driver};Database=<<path>>"
        $Connection = $env:lrPath
  )
  begin {
    $lite = Test-Path $Connection
    $null = Get-SQL -Session LR -Connection $Connection -lite:$lite
  }
  process {
    (Resolve-Path -Path $Path) | ForEach-Object {
        $p        = $_ -replace "'","''"
        $FileName = (Split-Path -leaf -Path $P )
        # In SQL Lite , = is case sensitive. Use LIKE for case insensitive
        $result   = [boolean](Get-SQL -Session LR  -Quiet -SQL (@"
          SELECT  folderPath,   AgLibraryFile.baseName,   AgLibraryFile.extension
          FROM                  AgLibraryFile
          JOIN (  SELECT        AgLibraryFolder.id_local, AgLibraryRootFolder.absolutePath || AgLibraryFolder.pathFromRoot AS folderPath
                  FROM          AgLibraryFolder
                  JOIN          AgLibraryRootFolder
                             ON AgLibraryRootFolder.id_local = AgLibraryFolder.rootFolder
                ) fullfolder ON          fullfolder.id_local = AgLibraryFile.folder
          WHERE   baseName    like '{0}'
          AND     extension   like '{1}'
          AND     folderPath  like '{2}/'
"@ -f ($FileName -split "\.")[0],($FileName -split "\.")[1],   ((Split-Path -Parent -Path $P ) -replace "\\","/")))
        if     ($not)          {$result = -not $result}
        if     ($Passthru -and  $result) {$_}
        elseif (-not $Passthru){$result}
  }}
}

function Move-NonLightRoom           {
  <#
.Synopsis
       Moves files or folders not found in lightroom's database
.EXAMPLE
      Move-NonLightRoom -destination ..\scrap
      Moves files ending .JPG or .DNG in the current directory but not in LR into the scrap directory
.EXAMPLE
      Move-NonLightRoom raw\*.dng -destination .\scrap -whatif
      But limits the files checked to .DNGs and shows what would b moved between between RAW and SCRAP
#>
[Cmdletbinding(SupportsShouldProcess=$true)]
Param (
    #Files to check against LightRoom
    $Path = @("*.dng","*.jpg"),
    #Directory where they should be sent
    [Parameter(Mandatory=$true)]$Destination
)
 Test-LightRoomItem -Path $path -Not   -Passthru |
   Move-Item -Destination $Destination -PassThru |
     Measure-Object  -sum -Property  length      |
       Format-Table       -Property  Count,@{n="Bytes";e={$_.sum.tostring("N0")}}
}

function Close-LightRoom             {
    Get-Sql -Session LR -Close -WarningAction SilentlyContinue
}