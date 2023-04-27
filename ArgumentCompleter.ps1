function IndexColumnCompletion      {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)

        $parameters = (Get-IndexedItem -List).shortname
        $parameters |  Where-Object { $_ -like "$wordToComplete*" } | Sort-Object |ForEach-Object {
            New-Object System.Management.Automation.CompletionResult "$_", "$_", ([System.Management.Automation.CompletionResultType]::ParameterValue) , $_
        }

}

function IndexColumnValueCompletion {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $ColumnName = $fakeBoundParameter['Where']
    [Void]$fakeBoundParameter.Remove("Where")
    if ($ColumnName) { (Get-IndexedItem -Value $ColumnName @fakeBoundParameter).$ColumnName  |
                            Where-Object { $_ -like "$wordToComplete*" } | Sort-Object | ForEach-Object {
                    New-Object System.Management.Automation.CompletionResult "$_", "$_", ([System.Management.Automation.CompletionResultType]::ParameterValue) , $_
                }
    }
}

function ExifIDCompletion              {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    Get-Variable -Name ExifID*  | Where-Object { "`$$($_.name)" -like "*$wordToComplete*" } |
        ForEach-Object {
                $completionText= "`$$($_.Name)"
                $toolTip       = ("0x" + $_.value.ToString("x4"))
                New-Object System.Management.Automation.CompletionResult $completionText, $completionText, ([System.Management.Automation.CompletionResultType]::ParameterValue) , $toolTip
        }
}

function ExifTagCompletion             {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $ExiftagValues.keys  | Where-Object {$_ -like "*$wordToComplete*" } |
            ForEach-Object {New-Object System.Management.Automation.CompletionResult "'$_'", "$_", ([System.Management.Automation.CompletionResultType]::ParameterValue) , $_}

}

function LightRoomPropertyCompletion   {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    if ($fakeBoundParameter['Where'] -in @("Aperture"   , "CameraModel", "ColorLabels", "Copyright", "Extension", "Fileformat",
                                           "FocalLength", "HasGPS"     , "IsRawFile"  , "IsoSpeed" , "LensModel" , "ShutterSpeed")) {
        Get-LightRoomItem -Values $fakeBoundParameter['Where'] |  Where-Object { $_ -like "$wordToComplete*" } | Sort-Object |
            ForEach-Object {New-Object System.Management.Automation.CompletionResult "'$_'", $_, ([System.Management.Automation.CompletionResultType]::ParameterValue) , $_}
    }
}

function LightRoomCollectionCompletion {
 param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $parameters =  (Get-LightRoomCollection).CollectionName
    $parameters  |  Where-Object { $_ -like "$wordToComplete*" } | Sort-Object |
            ForEach-Object {New-Object System.Management.Automation.CompletionResult "'$_'", $_, ([System.Management.Automation.CompletionResultType]::ParameterValue) , $_ }
}

#In PowerShell 3 and 4 Register-ArgumentCompleter is part of TabExpansion ++. From V5 it is part of Powershell.core
if (Get-Command -ErrorAction SilentlyContinue -name Register-ArgumentCompleter) {
 Register-ArgumentCompleter -CommandName 'Get-IndexedItem'             -ParameterName 'Where'      -ScriptBlock $Function:IndexColumnCompletion
 Register-ArgumentCompleter -CommandName 'Get-IndexedItem'             -ParameterName 'Property'   -ScriptBlock $Function:IndexColumnCompletion
 Register-ArgumentCompleter -CommandName 'Get-IndexedItem'             -ParameterName 'Orderby'    -ScriptBlock $Function:IndexColumnCompletion
 Register-ArgumentCompleter -CommandName 'Get-IndexedItem'             -ParameterName 'Value'      -ScriptBlock $Function:IndexColumnCompletion
 Register-ArgumentCompleter -CommandName 'Get-IndexedItem'             -ParameterName 'EQ'         -ScriptBlock $Function:IndexColumnValueCompletion
 Register-ArgumentCompleter -CommandName 'Add-EXIFFilter'              -ParameterName 'EXIFTag'    -ScriptBlock $Function:ExifTagCompletion
 Register-ArgumentCompleter -CommandName 'Get-LightRoomItem'           -ParameterName 'EQ'         -ScriptBlock $Function:LightRoomPropertyCompletion
 Register-ArgumentCompleter -CommandName 'Get-LightRoomItem'           -ParameterName 'NE'         -ScriptBlock $Function:LightRoomPropertyCompletion
 Register-ArgumentCompleter -CommandName 'Get-LightRoomItem'           -ParameterName 'LT'         -ScriptBlock $Function:LightRoomPropertyCompletion
 Register-ArgumentCompleter -CommandName 'Get-LightRoomItem'           -ParameterName 'LE'         -ScriptBlock $Function:LightRoomPropertyCompletion
 Register-ArgumentCompleter -CommandName 'Get-LightRoomItem'           -ParameterName 'GT'         -ScriptBlock $Function:LightRoomPropertyCompletion
 Register-ArgumentCompleter -CommandName 'Get-LightRoomItem'           -ParameterName 'GE'         -ScriptBlock $Function:LightRoomPropertyCompletion
 Register-ArgumentCompleter -CommandName 'Get-LightRoomCollectionItem' -ParameterName 'Include'    -ScriptBlock $Function:LightRoomCollectionCompletion
 Register-ArgumentCompleter -CommandName 'Add-LightRoomCollectionItem' -ParameterName 'Collection' -ScriptBlock $Function:LightRoomCollectionCompletion
 }

