<#
Disclaimer

This sample script is not supported under any Microsoft standard support program or service. 
The sample script is provided AS IS without warranty of any kind. Microsoft further disclaims 
all implied warranties including, without limitation, any implied warranties of merchantability 
or of fitness for a particular purpose. The entire risk arising out of the use or performance 
of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, 
or anyone else involved in the creation, production, or delivery of the scripts be liable for 
any damages whatsoever (including, without limitation, damages for loss of business profits, 
business interruption, loss of business information, or other pecuniary loss) arising out of the 
use of or inability to use the sample scripts or documentation, even if Microsoft has been advised 
of the possibility of such damages
#>

<#
.SYNOPSIS
  Installs or Updates the Sysinternals Tools on your System
.DESCRIPTION
  Installs or Updates the Sysinternals Tools on your System from the 
  official Sysinternals Repository or optional from a given Fileshare, Folder or Webserver
.LINK
   https://github.com/ChrisNiebuhrMSFT/Install-Sysinternals
.EXAMPLE
  .\Install-Sysinternals.ps1 # Installs/Updates the current Sysinternals Version under C:\SysinternalSuite
.EXAMPLE
  .\Install-Sysinternals.ps1 -SourcePath \\Myfileshare\Sysinternals.zip # Installs/Updates Sysinternal Suite from a Custom Fileshare
.EXAMPLE
  .\Install-Sysinternals.ps1 -Force # Install/Updates Sysinternal Suite and closes any running Sysinternals Process which is actually running on the System without confirmation
  #When the -Force Parameter is not specified and any running Sysinternal Process in running while updating the user will be propmpted to close the running Processes
.NOTES
  Runs with Windows PowerShell 5.x and PowerShell Core
  Author: Microsoft - Chris Niebuhr
  Date:   06/28/2022
#>

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
