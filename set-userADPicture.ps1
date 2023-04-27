
function Set-UserADPicture {
  Param (
    $LDAPFilter  = "(SAMAccountName=$env:USERNAME)",
    [Parameter(Mandatory=$true)]$Path
  )
    try   { [byte[]]$fileData = Get-Content -Path $Path -Encoding Byte -ReadCount 0  -ErrorAction Stop}
    catch { Write-Warning -Message "Could not read file from $Path" ; return }
    $user = [adsi](([adsisearcher]"(&(objectClass=user)$LDAPFilter)").FindAll().path -replace "^GC://","LDAP://")
    $user.Properties["thumbnailPhoto"].Clear()
    $result = $user.Properties["thumbnailPhoto"].Add($fileData)
    if ($result -eq "0") {Write-Verbose -Message "Photo successfully uploaded.   $LDAPFilter"}
    else                 {Write-Warning -Message "Photo was not uploaded."}
    $user.CommitChanges()
}
