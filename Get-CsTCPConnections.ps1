
<#PSScriptInfo

.VERSION 1.1

.GUID 33820af4-2938-4e51-b881-4618e8ae7bf6

.AUTHOR David Paulino

.COMPANYNAME UC Lobby

.COPYRIGHT

.TAGS Lync LyncServer SkypeForBusiness SfBServer Network TCP

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
  Version 1.0: 2016/08/19 - Initial release.
  Version 1.1: 2023/10/07 - Updated to publish in PowerShell Gallery.

.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 Returns TCP Established Connections Performance Monitor Counter from Lync/Skype for Business. 

#> 

[CmdletBinding()]
param(
[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)]
	[string] $ServerFqdn,
[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)]
    [string] $PoolFqdn,
[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)]
	[switch] $AllPools
)

# Store all the start up variables so you can clean up when the script finishes.
if ($startupvariables) { try {Remove-Variable -Name startupvariables  -Scope Global -ErrorAction SilentlyContinue } catch { } }
New-Variable -force -name startupVariables -value ( Get-Variable | ForEach-Object { $_.Name } ) 

$errpref = $ErrorActionPreference #save actual preference
$ErrorActionPreference = "Stop"

function Clean-Memory {
Get-Variable |
 Where-Object { $startupVariables -notcontains $_.Name } |
 ForEach-Object {
  try { Remove-Variable -Name "$($_.Name)" -Force -Scope "global" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue}
  catch { }
 }
}

$startTime=Get-Date;
$FEList = New-Object System.Collections.ArrayList
$sPoolFqdn = "NA"

function getNetConnectionsServer($server){
    Write-Host "Fetching data from:" $server -ForegroundColor Cyan
    $counter = "\\" + $server + "\TCPv4\Connections Established" 
    try {
        $TCPestablished = (Get-Counter -Counter $counter | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue)
    } catch {
        Write-Host "Error while fetching data from:" $server -ForegroundColor Red
    }
    
    $NetCounters = New-Object PSObject -Property @{            
                    Pool   = $sPoolFqdn
                    FQDN       = $server
                    TCPEstablished = $TCPestablished 
                  }
    [void]$FElist.Add($NetCounters)
    
}

function getNetConnectionsPool($Pool){
    Write-Host "Processing Front End Pool:" $Pool -ForegroundColor Green
    $feServers = Get-CsComputer -Pool $Pool | Sort-Object identity
    foreach($feServer in $feServers){
        getNetConnectionsServer($feServer.Fqdn)
        
    }
}
    if($AllPools)
    {
        $fePools = (Get-CsService -Registrar) | Sort-Object Version
        foreach ($fepool in $fePools){
            $sPoolFqdn = $fepool.PoolFqdn
            $feServers = Get-CsComputer -Pool $fepool.PoolFqdn | Sort-Object identity
            getNetConnectionsPool($fepool.PoolFqdn)
        }
        $FElist | Select Pool, FQDN, TCPEstablished | ft -AutoSize
    } else {
        if($PoolFqdn){
            $sPoolFqdn = $PoolFqdn
            getNetConnectionsPool($PoolFqdn)
        

        } else {
            if(!$ServerFqdn) {
                $ServerFqdn = [System.Net.Dns]::GetHostByName((hostname)).HostName
            }
            
            getNetConnectionsServer($ServerFqdn)
        }
        $FElist | Select FQDN, TCPEstablished | ft -AutoSize
    }
$endTime = Get-Date
$totalTime= [math]::round(($endTime - $startTime).TotalSeconds,2)
Write-Host "Execution time:" $totalTime "seconds" -ForegroundColor Cyan

#Cleanup the variables used by the script
Clean-Memory   
