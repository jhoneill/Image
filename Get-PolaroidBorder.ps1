function Get-PolaroidBorder {
param ([Parameter(Mandatory=$true)][int]$edge)
    "Top        {0:N0}" -f ($edge * 0.08)  
    "Bottom     {0:N0}" -f ($edge * 0.28)
    "left/Right {0:N0}" -f ($edge * 0.06)
}