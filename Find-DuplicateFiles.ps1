$Names  = @{}
$hashes = @{}
dir -Recurse -File | get-filehash | ForEach-Object {
  if   ($hashes[$_.hash]) {
        "$($_.Path)  =   $($hashes[$_.Hash])"
  }
  else {$hashes[$_.hash] = $_.path}
}