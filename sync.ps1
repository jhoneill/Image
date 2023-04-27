Param (
    [Parameter(Mandatory=$true,position=0,ValueFromPipeline=$true)]
    [string]$
)


Set-Location "D:\Pictures\Models\$dir"
$TotalFiles = (Get-ChildItem .\raw\).count
$count = (Get-ChildItem .\raw\ | Where-Object { -not (Test-Path (Join-path "C:\Not-Indexed\$dir\Raw\" $_.name))}  |  Measure-Object).count
if (-not $count) {Write-host "Nothing to move from $dir\raw to scrap"
}
else {
    write-host "Moving $count of $TotalFiles items from $dir\raw to $dir\scrap"
    Get-ChildItem .\raw\ | Where-Object { -not (Test-Path (Join-path "C:\Not-Indexed\$dir\Raw\" $_.name))}  |  move-item -Destination .\Scrap\ -Confirm
}
Get-ChildItem .\scrap\ | ForEach-Object { Get-Item (Join-path "C:\Not-Indexed\$dir\Scrap\" $_.name)  -ErrorAction SilentlyContinue }  | Remove-Item -Verbose
Get-ChildItem "C:\not-indexed\$dir\Raw\" -File | Where-Object {-not (Test-Path (Join-Path $pwd\raw $_.name))} | Copy-Item -Destination $pwd\raw -Verbose
Get-ChildItem "C:\not-indexed\$dir\"     -File | Where-Object {-not (Test-Path (Join-Path $pwd     $_.name))} | Copy-Item -Destination $pwd      -Verbose