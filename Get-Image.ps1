#requires -version 2.0
#Original code James Brundage / Microsoft 2009 , published in PSIMageTools
function Get-Image {
<#
    .Synopsis
        Returns an image object for a file
    .Description
        Uses the Windows Image Acquisition COM object to get image data
    .Example
        Get-ChildItem $env:UserProfile\Pictures -Recurse | Get-Image
    .Parameter file
        The file to get an image from
#>
    param(
        [Parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true)]
        [Alias('FullName',"FileName")]
        [ValidateScript({Test-path -Path $_ })][string]$Path
    )

    process {
        foreach ($file in (Resolve-Path -Path $path) ) {
            $image  = New-Object -ComObject Wia.ImageFile
            try   {
                Write-Verbose -Message "Loading file $($file.Path)"
                $null = $image.LoadFile($file.path)
                $image |
                    Add-Member -MemberType NoteProperty -Name FullName               -Value $File.Path -PassThru |
                    Add-Member -MemberType NoteProperty -Name OriginalName           -Value $File.Path -PassThru |
                    Add-Member -MemberType ScriptMethod -Name Resize                 -Value {
                        param($width, $height, [switch]$DoNotPreserveAspectRatio)
                        $image  = New-Object -ComObject Wia.ImageFile
                        $image.LoadFile(  $this.FullName)
                        $filter = Add-ScaleFilter @psBoundParameters -passThru -image $image
                        $image  = $image | Set-ImageFilter -filter $filter -passThru
                        Remove-Item -Path $this.Fullname
                        $image.SaveFile(  $this.FullName)
                    } -PassThru |
                    Add-Member -MemberType ScriptMethod -Name Crop                   -Value {
                        param([Double]$left, [Double]$top, [Double]$right, [Double]$bottom)
                        $image  = New-Object -ComObject Wia.ImageFile
                        $image.LoadFile(  $this.FullName)
                        $filter = Add-CropFilter @psBoundParameters -passThru -image $image
                        $image  = $image | Set-ImageFilter -filter $filter -passThru
                        Remove-Item -Path $this.Fullname
                        $image.SaveFile(  $this.FullName)
                    } -PassThru |
                    Add-Member -MemberType ScriptMethod -Name FlipVertical           -Value {
                        $image  = New-Object -ComObject Wia.ImageFile
                        $image.LoadFile(  $this.FullName)
                        $filter = Add-RotateFlipFilter -flipVertical -passThru
                        $image  = $image | Set-ImageFilter -filter $filter -passThru
                        Remove-Item -path $this.Fullname
                        $image.SaveFile(  $this.FullName)
                    } -PassThru |
                    Add-Member -MemberType ScriptMethod -Name FlipHorizontal         -Value {
                        $image = New-Object -ComObject Wia.ImageFile
                        $image.LoadFile(  $this.FullName)
                        $filter = Add-RotateFlipFilter -flipHorizontal -passThru
                        $image = $image | Set-ImageFilter -filter $filter -passThru
                        Remove-Item -Path $this.Fullname
                        $image.SaveFile(  $this.FullName)
                    } -PassThru |
                    Add-Member -MemberType ScriptMethod -Name RotateClockwise        -Value {
                        $image = New-Object -ComObject Wia.ImageFile
                        $image.LoadFile(  $this.FullName)
                        $filter = Add-RotateFlipFilter -angle 90 -passThru
                        $image = $image | Set-ImageFilter -filter $filter -passThru
                        Remove-Item -path $this.Fullname
                        $image.SaveFile(  $this.FullName)
                    } -PassThru |
                    Add-Member -MemberType ScriptMethod -Name RotateCounterClockwise -Value {
                        $image = New-Object -ComObject Wia.ImageFile
                        $image.LoadFile(  $this.FullName)
                        $filter = Add-RotateFlipFilter -angle 270 -passThru
                        $image = $image | Set-ImageFilter -filter $filter -passThru
                        Remove-Item -path $this.Fullname
                        $image.SaveFile(  $this.FullName)
                    } -PassThru

            }
            catch { Write-Verbose -Message $_ }
        }
    }
}