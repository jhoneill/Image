param (
    $Path =$pwd,
        # Path to the LightRoom catalog file  or  Database connection string  e.g. "DSN=LR"
       $Connection = $env:lrPath
)
$lite = Test-Path $Connection
$null = Get-SQL -Session LR -Connection $Connection -lite:$lite


$sql = @"
SELECT
    IPTC.*,
    rootFolder.absolutePath || folder.pathFromRoot || rootfile.baseName || '.' || rootfile.extension      AS fullName,
    Image.captureTime       AS dateTaken
FROM       AgLibraryIPTC              IPTC
    JOIN   Adobe_images              image  ON      image.id_local =     IPTC.image
    JOIN   AgLibraryFile          rootFile  ON   rootfile.id_local =    image.rootFile
    JOIN   AgLibraryFolder          folder  ON     folder.id_local = rootfile.folder
    JOIN   AgLibraryRootFolder  rootFolder  ON rootFolder.id_local =   folder.rootFolder
WHERE  (RootFolder.absolutePath || Folder.pathFromRoot) Like '$($Path -replace "\\","/")%'
"@

Get-SQL  -Session LR -SQL $SQL  | ForEach-Object {
    $c  = Get-NearestSuutoDBPoint -when ([datetime]$_.dateTaken)
    If ($c.Description) {
        #Get-SQL -Session LR -Table AgLibraryIPTC -Set Caption,CopyRight -Values ($c.Description -replace "'","''"), ($_.CopyRight -replace "2016", "2019" -replace "'","''") -Where id_local -eq $_.id_Local -Confirm:$false
        Get-SQL -Session LR -Table AgLibraryIPTC -Set Caption -Values ($c.Description -replace "'","''") -Where id_local -eq $_.id_Local -Confirm:$false
        Write-Verbose "$($_.fullName)       $($c.Description)"
    }
}