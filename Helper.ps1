
function ConvertTo-Enum {
<#
    .SYNOPSIS
        Converts a hashtable to an enum.

    .PARAMETER Table
        Specifies the hashtable to convert.

    .PARAMETER Name
        Specifies the name of the int-based enum to create.

    .EXAMPLE
        Creates an enum based on a hashtable with the names and values of the three lowest-valued U.S. coins.

        PS > @{"Penny" = 1;"Nickle" = 5;"Dime" = 10} | ConvertTo-Enum -Name SmallChange
        PS > [SmallChange]::1
        Penny
        PS > [SmallChange]10
        Dime

#>
    [CmdletBinding()]
    [outputType([string])]
    param(
        [parameter(Mandatory = $true, ValueFromPipeLine = $true)]
        [Hashtable]
        [ValidateNotNullOrEmpty()]
        $Table,

        [parameter(Mandatory = $true)]
        [string]
        [ValidateNotNullOrEmpty()]
        $Name,

        [Switch]$CodeOnly
    )

    foreach ($key in $Table.Keys) {$items += (",`n {0,20} = {1}" -f $key,$($Table[$key]) ) }
    $code = "public enum $Name : int  `n{$($items.Substring(1)) `n}"
    if ($codeOnly) {$code }
    else           {Add-Type -TypeDefinition  $code}
}

function Select-Item   {
# .ExternalHelp  Maml-Helper.XML
   [CmdletBinding()]
   param ([parameter(ParameterSetName="p1",Position=0)][String[]]$TextChoices,
          [Parameter(ParameterSetName="p2",Position=0)][hashTable]$HashChoices,
          [String]$Caption="Please make a selection",
          [String]$Message="Choices are presented below",
          [int]$default=0,
          [Switch]$returnKey
    )
    $choicedesc = New-Object -TypeName System.Collections.ObjectModel.Collection[System.Management.Automation.Host.ChoiceDescription]
    switch ($PsCmdlet.ParameterSetName) {
        "p1" {$TextChoices | ForEach-Object       { $choicedesc.Add((New-Object -TypeName "System.Management.Automation.Host.ChoiceDescription" -ArgumentList $_                      )) }  }
        "p2" {foreach ($key in $HashChoices.Keys) { $choicedesc.Add((New-Object -TypeName "System.Management.Automation.Host.ChoiceDescription" -ArgumentList $key,$HashChoices[$key] )) }  }
    }
    If ($returnkey) { $choicedesc[$Host.ui.PromptForChoice($caption, $message, $choicedesc, $default)].label }
    else            {             $Host.ui.PromptForChoice($caption, $message, $choicedesc, $default) }

}

function Select-List  {
#  .ExternalHelp  Maml-Helper.XML
    param   ([Parameter(Mandatory=$true  ,valueFromPipeline=$true )][Alias("items")]$InputObject,
             [Parameter(Mandatory=$true)]$Property, [Switch]$multiple)

    begin   { $i= @()  }
    process { $i += $inputobject  }
    end     { if ($i.count -eq 1) {$i[0]} elseif ($i.count -gt 1) {
                  $Script:counter=-1
                  $Property=@(@{Label="ID"; Expression={ ($Script:Counter++) }}) + $Property
                  Format-Table -InputObject $i -AutoSize -Property $Property | out-host
                  if ($multiple) { $response = Read-Host -Prompt "Which one(s) ?" }
                  else           { $response = Read-Host -Prompt "Which one ?"    }
                  if ($response -gt "") {
                       if ($multiple) { $response.Split(",") | ForEach-Object -Begin {$r = @()} -process {if ($_ -match "^\d+$") {$r += $_} elseif ($_ -match "^\d+\.\.\d+$") {$r += (Invoke-Command -ScriptBlock ([scriptblock]::Create( $_)))}} -end {$I[$r] }}
                       else           { $I[$response] }
                  }
              }
            }
}

function Select-EnumType {
#  .ExternalHelp  Maml-Helper.XML
    param ([type]$EType , [int]$default)
    $NamesAndValues= $etype.getfields() |
        foreach-object -begin {$list=@()} `
                       -process { if (-not $_.isspecialName) {$list += $_.getValue($null)}} `
                       -end {$list  | sort-object -property value__ } |
            select-object -property  @{name="Name"; expression={$_.tostring() -replace "_"," "}},value__
    $value = (Select-List -Input $namesAndValues -property Name).value__
    if ($null -eq $value ) {$default} else {$value}
}

function Out-Tree {
#  .ExternalHelp  Maml-Helper.XML
    param   (
        [Parameter(Mandatory=$true,ValueFromPipeLine = $true)][alias("items")]$inputObject,
        [Parameter(Mandatory=$true)]$startAt,
        [string]$path="Path", [string]$parent="Parent",[string]$label="Label", [int]$indent=0
             )
     begin   { $items = @() }
     process { $items += $inputobject}
     end     {
         $children = $items | where-object {( "" + $_.$parent) -eq $startAt.$path.ToString()}  | Sort-Object -Property $label
         if ($null -ne $children) {(("| " * $indent) -replace "\s$","-") + "+$($startAt.$label)"
                                   $children | ForEach-Object {Out-Tree -inputObject $items -startAt $_ -path $path -parent $parent -label $label -indent ($indent+1)}
         }
         else                     {("| " * ($indent-1)) + "|--$($startAt.$label)" }
     }
}

function Select-Tree {
    #  .ExternalHelp  Maml-Helper.XML
    param   (
           [Parameter(Mandatory=$true,ValueFromPipeLine = $true)][alias("items")]$inputObject,
           [Parameter(Mandatory=$true)]$startAt,
           [string]$path="Path", [string]$parent="Parent", [string]$label="Label", $indent=0,
           [Switch]$multiple
     )
    begin   { $items  = @()          }
    process { $items += $inputobject }
    end     {
        if ($Indent -eq 0)  {$script:treeCounter = -1 ;  $script:treeList=@()  }
        $script:treeCounter++
        $script:treeList=$script:treeList + @($startAt)
        $children = $items | where-object {$_.$parent -eq $startat.$path.ToString()}
        if   ($null -eq $children) { "{0,-4} {1}|--{2} " -f  $script:Treecounter, ("| " * ($indent-1)) , $startAt.$Label | out-Host }
        else                       { "{0,-4} {1}+{2} "   -f  $script:treeCounter, ("| " * ($indent)) ,   $startAt.$label | Out-Host
                                     $children | Sort-Object $label | ForEach-Object {
                                           Select-Tree -inputObject $items -StartAt $_ -Path $path -parent $parent -label $label -indent ($indent+1)
                                      }
        }
        if ($Indent -eq 0) {if ($multiple) { $script:treeList[ [int[]](Read-Host -Prompt "Which one(s) ?").Split(",")] }
                            else           { $script:treeList[        (Read-Host -Prompt "Which one ?")] }
                           }
  }
}

function Test-Admin {
#  .ExternalHelp  Maml-Helper.XML
    [CmdletBinding()]
    param()

    $currentUser = New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList $([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    Write-Verbose -Message "isUserAdmin? $isAdmin"
    $isAdmin
}

function Convert-DiskIDtoDrive {
#  .ExternalHelp  Maml-Helper.XML
    param ([parameter(ParameterSetName="p1",Position=0, ValueFromPipeLine = $true)][ValidateScript({ $_ -ge 0 })][int]$diskIndex,
           [Parameter(ParameterSetName="p2",Position=0 , ValueFromPipeline=$true)] [System.Management.ManagementObject]$Inputobject
    )
    process{
        switch ($PsCmdlet.ParameterSetName) {
            "p1"  { Get-WmiObject -query "select * from win32_diskpartition where diskindex = $DiskIndex" | ForEach-Object{$_.getRelated("win32_Logicaldisk")} | ForEach-Object {$_.deviceID} }
            "p2"  { get-wmiobject  -computername $inputObject.__SERVER -namespace "root\cimv2"  -Query "associators of {$($inputObject.__path)} where resultclass=Win32_DiskPartition" |
                    ForEach-Object{$_.getRelated("win32_Logicaldisk")} | ForEach-Object {$_.deviceID}
                  }

        }
    }
}

function Get-FirstAvailableDriveLetter {
#  .ExternalHelp  Maml-Helper.XML
    $UsedLetters = Get-WmiObject -Query "SELECT * FROM Win32_LogicalDisk"  | ForEach-object {$_.deviceId.substring(0,1)}
    [char]$l="A"
    do {
        if   ($usedLetters -notcontains $L) {return $l}
        else {  [char]$l=([byte][char]$l +1 ) }
       }
    while ($l -le 'Z')
    Write-Warning -Message "No free drive letters found"
}

