[CmdletBinding()]
Param 
(
    [ValidateNotNullOrEmpty()]
    [string]
    $SysinternalsPath = 'C:\SysinternalsSuite', #Specify a Path where you like to install and maintain your Sysinternals Installation
    [ValidatePattern('\.zip$')] #Validate that the Source-Path ends with .zip
    [string]
    $SourcePath = 'https://download.sysinternals.com/files/SysinternalsSuite.zip', #You only need to change this, if you like to Install / Update from a local Path
    [switch]
    $Force # Force closing all running Sysinternals Tools
)

$PSDefaultParameterValues = @{'*:ErrorAction' = 'Stop'} #Ensure that all Errors from Cmdlets are "catchable" => Terminating Errors
#To Ensure you can install the Tools to a Path that requires elevated rights for Installation
if (-not ([System.Security.Principal.WindowsPrincipal][System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))
{
    if ($PSVersionTable['PSEdition'] -eq 'Core')
    {
        $executableName = 'pwsh.exe'
    }
    else
    {
        $executableName = 'powershell.exe'
    }
    $arguments = '-File "{0}"' -f $MyInvocation.MyCommand.Definition
    $powerShellPath = [System.IO.Path]::Combine($PSHome, $executableName)
    Start-Process $powerShellPath -Verb runAs -ArgumentList $arguments #Start Process elevated
    return #Do no proceed from here
}
try 
{
    $SysinternalsZipPath = [System.IO.Path]::Combine($env:TEMP, 'sysinternals.zip') #Temp-Path for the downloaded Zip-File
    Write-Host 'Start Installing / Updating Sysinternal Tools'
    if ($SourcePath -match '^https?') #When downloading from the Internet or local Webserver use BITS
    {
        Start-BitsTransfer -Source $SourcePath -Destination $SysinternalsZipPath 
    }
    else #Otherwise copy the zip-File via File-Copy
    {
        Copy-Item -Path $SourcePath -Destination $SysinternalsZipPath -Force
    }

    if (-not (Test-Path $SysinternalsPath))
    {
        $null = New-Item -ItemType Directory -Path $SysinternalsPath
    }

    $runningProcs = Get-ChildItem -Path $SysinternalsPath -File -Filter '*.exe'   | Select-Object -ExpandProperty Name | ForEach-Object { Get-Process -Name ([System.IO.Path]::GetFileNameWithoutExtension($PSItem)) -ErrorAction Ignore }
    if ($runningProcs)
    {
        if($Force)
        {
            Stop-Process -Name ($runningProcs.Name) -Force 
        }
        else 
        {
            $runningProcs | Select-Object -Property Name | Out-GridView -Title 'Please close the following processes' -OutputMode Multiple | ForEach-Object { Stop-Process -Name $PSItem.Name -Force }
        }
    }

    Expand-Archive -Path $SysinternalsZipPath -DestinationPath $SysinternalsPath -Force
    Write-Host 'Finished'
}
catch
{
    Write-Host "$PSItem" #Display any Error-Messages
}
finally
{
    Remove-Item -Path $SysinternalsZipPath -Force
    Read-Host
}
