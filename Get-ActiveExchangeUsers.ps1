<#

.SYNOPSIS

Created by: https://ingogegenwarth.wordpress.com/
Version:    42 ("What do you get if you multiply six by nine?")
Changed:    08.11.2018

.DESCRIPTION

The script is enumerating all Exchange 2010 CAS and Exchange 2013/2016 servers in the current or given AD site and queries specific performance counters

.LINK

http://mikepfeiffer.net/2011/04/determine-the-number-of-active-users-on-exchange-2010-client-access-servers-with-powershell/
https://technet.microsoft.com/library/hh849685.aspx
https://www.granikos.eu/en/justcantgetenough/PostId/197/inline-css-with-convertto-html
https://ingogegenwarth.wordpress.com/2016/05/09/get-activeexchangeusers-2-0/

.PARAMETER ADSite

here you can define in which ADSite is searched for Exchange server. If omitted current AD site will be used.

.PARAMETER Summary

if used the script will sum up the active user count across all servers per protocol

.PARAMETER UseASPDOTNET

switch to use IIS performance counters (ASP.NET Apps v4.0.30319\Requests Executing) for gathering current requests per protocol

.PARAMETER HTTPProxyAVGLatency

the script will collect for each protocol the performance counter "\MSExchange HttpProxy(protocoll)\Average ClientAccess Server Processing Latency"

.PARAMETER HTTPProxyOutstandingRequests

the script will collect for each protocol the performance counter "\MSExchange HttpProxy(protocoll)\Outstanding Proxy Requests"

.PARAMETER HTTPProxyRequestsPerSec

the script will collect for each protocol the performance counter "\MSExchange HttpProxy(protocoll)\\Proxy Requests/Sec"

.PARAMETER E2EAVGLatency

the script will collect for main protocols counters like "\MSExchangeIS Client Type(*)\RPC Average Latency",\MSExchange RpcClientAccess\RPC Averaged Latency","\MSExchange MapiHttp Emsmdb\Averaged Latency"

.PARAMETER TimeInGC

collects and compute the average of the following GC performance counters "\.NET CLR Memory(w3w*)\% Time in GC","\.NET CLR Memory(w3*)\Process ID","\W3SVC_W3WP(*)\Active Requests"

.PARAMETER SpecifiedServers

filtering for specific servers, which were found in given AD site

.PARAMETER MaxSamples

as the script uses the CmdLet Get-Counter you can define the number of MaxSamples. Default is 1

.PARAMETER SendMail

switch to send an e-mail with a CSV attached

.PARAMETER From

define the sender address

.PARAMETER Recipients

define the recipients

.PARAMETER SmtpServer

which SmtpServer to use

.PARAMETER IISMemoryUsage

collects the following performance counters "\Process(w3wp*)\working set - private". Note: This will return only the workes with no hint, which worker process it is. Use UseCIM switch for details.

.PARAMETER UseCIM

Uses CIM for gathering IIS memory usage for application pools, which returns friendly name of application pools.

.PARAMETER CimTimeoutSec

Timeout for CIM connection. Default is 30 seconds.

.EXAMPLE

Get users from current AD site
.\Get-ActiveExchangeUsers.ps1

Get users from given AD site
.\Get-ActiveExchangeUsers.ps1 -ADSite HQ-Site

Get summary

.\Get-ActiveExchangeUsers.ps1 -Summary

Get number of outstanding proxy requests for 60 samples

.\Get-ActiveExchangeUsers.ps1 -HTTPProxyOutstandingRequests -MaxSamples 60

Get number of average processing time of proxy requests

.\Get-ActiveExchangeUsers.ps1 -HTTPProxyAVGLatency

Get number of proxy requests per second

.\Get-ActiveExchangeUsers.ps1 -HTTPProxyRequestsPerSec

Get backend related AVG latency for main components

.\Get-ActiveExchangeUsers.ps1 -E2EAVGLatency

Get time in GC for main components

.\Get-ActiveExchangeUsers.ps1 -TimeInGC

.NOTES

You need to run this script in the same AD site where the servers are for performance reasons.

#>
[CmdletBinding(DefaultParameterSetName = "ALL")]
param(
    [parameter(
        Mandatory=$false,
        Position=0)]
    [System.String[]]
    $ADSite,

    [parameter(
        Mandatory=$false,
        Position=1,
        ParameterSetName="Summary")]
    [System.Management.Automation.SwitchParameter]
    $Summary,

    [parameter(
        Mandatory=$false,
        Position=2,
        ParameterSetName="ASPDOTNET")]
    [System.Management.Automation.SwitchParameter]
    $UseASPDOTNET,

    [parameter(
        Mandatory=$false,
        Position=3,
        ParameterSetName="HTTPProxyAVGLatency")]
    [System.Management.Automation.SwitchParameter]
    $HTTPProxyAVGLatency,

    [parameter(
        Mandatory=$false,
        Position=4,
        ParameterSetName="HTTPProxyOutstandingRequests")]
    [System.Management.Automation.SwitchParameter]
    $HTTPProxyOutstandingRequests,

    [parameter(
        Mandatory=$false,
        Position=5,
        ParameterSetName="HTTPProxyRequestsPerSec")]
    [System.Management.Automation.SwitchParameter]
    $HTTPProxyRequestsPerSec,

    [parameter(
        Mandatory=$false,
        Position=6,
        ParameterSetName="E2EAVGLatency")]
    [System.Management.Automation.SwitchParameter]
    $E2EAVGLatency,

    [parameter(
        Mandatory=$false,
        Position=7,
        ParameterSetName="TimeInGC")]
    [System.Management.Automation.SwitchParameter]
    $TimeInGC,

    [parameter(
        Mandatory=$false,
        Position=8)]
    [System.Array]
    $SpecifiedServers,

    [parameter(
        Mandatory=$false,
        Position=9)]
    [System.Int32]
    $MaxSamples = 1,

    [parameter(
        Mandatory=$false,
        Position=10)]
    [System.Management.Automation.SwitchParameter]
    $SendMail,

    [parameter(
        Mandatory=$false,
        Position=11)]
    [System.String]
    $From,

    [parameter(
        Mandatory=$false,
        Position=12)]
    [System.String[]]
    $Recipients,

    [parameter(
        Mandatory=$false,
        Position=13)]
    [System.String]
    $SmtpServer,

    [parameter(
        Mandatory=$false,
        Position=14,
        ParameterSetName="IISMemory")]
    [System.Management.Automation.SwitchParameter]
    $IISMemoryUsage,

    [parameter(
        Mandatory=$false,
        Position=15,
        ParameterSetName="IISMemory")]
    [System.Management.Automation.SwitchParameter]
    $UseCIM,

    [parameter(
        Mandatory=$false,
        Position=16,
        ParameterSetName="IISMemory")]
    [System.Int32]
    $CimTimeoutSec = 30

)

$ErrorActionPreference = "silentlycontinue"
# function to get the Exchangeserver from AD site
Function GetExchServer
{
    [CmdLetBinding()]
    #http://technet.microsoft.com/en-us/library/bb123496(v=exchg.80).aspx on the bottom there is a list of values
    param(
        [System.Array]
        $Roles,
        [System.String[]]
        $ADSites
    )

    Process
    {
        $valid = @("2","4","16","20","32","36","38","54","64","16385","16439")
        ForEach ($Role in $Roles)
        {
            If (-not ($valid -contains $Role))
            {
                Write-Output "Please use the following numbers: MBX=2,CAS=4,UM=16,HT=32,Edge=64 multirole servers:CAS/HT=36,CAS/MBX/HT=38,CAS/UM=20,E2k13 MBX=54,E2K13 CAS=16385,E2k13 CAS/MBX=16439"
                return
            }
        }

        Function GetADSite
        {
            param(
                [System.String]
                $Name
            )

            If ($null -eq $Name)
            {
                [System.String]$Name = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).GetDirectoryEntry().Name
            }

            $FilterADSite = "(&(objectclass=site)(Name=$Name))"
            $RootADSite= ([ADSI]'LDAP://RootDse').configurationNamingContext
            $SearcherADSite = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"LDAP://$RootADSite")
            $SearcherADSite.Filter = "$FilterADSite"
            $SearcherADSite.pagesize = 1000
            $ResultsADSite = $SearcherADSite.FindOne()
            $ResultsADSite
        }

        $Filter = "(&(objectclass=msExchExchangeServer)(|"
        ForEach ($ADSite in $ADSites)
        {
            $Site=''
            $Site = GetADSite -Name $ADSite
            If ($null -eq $Site)
            {
                Write-Verbose "ADSite $($ADSite) could not be found!"
            }
            Else
            {
                Write-Verbose "Add ADSite $($ADSite) to filter!"
                $Filter += "(msExchServerSite=$((GetADSite -Name $ADSite).properties.distinguishedname))"
            }
        }

        $Filter += ")(|"
        ForEach ($Role in $Roles)
        {
            $Filter += "(msexchcurrentserverroles=$Role)"
        }

        $Filter += "))"
        $Root= ([ADSI]'LDAP://RootDse').configurationNamingContext
        $Searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"LDAP://$Root")
        $Searcher.Filter = "$Filter"
        $Searcher.pagesize = 1000
        $Results = $Searcher.FindAll()
        If ("0" -ne $Results.Count)
        {
            $Results
        }
        Else
        {
            Write-Verbose "No server found!"
        }
    }
}

Function GetCounterSum ()
{
param(
    [System.Boolean]
    $UseASPDOTNET
)

    If ($UseASPDOTNET)
    {
        $ActiveSyncFE = $((Get-Counter "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_Microsoft-Server-ActiveSync)\Requests Executing" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $ActiveSync = $((Get-Counter "\ASP.NET Apps v4.0.30319(_LM_W3SVC_2_ROOT_Microsoft-Server-ActiveSync)\Requests Executing" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $OWAFE = $((Get-Counter "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_OWA)\Requests Executing" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $OWA = $((Get-Counter "\ASP.NET Apps v4.0.30319(_LM_W3SVC_2_ROOT_OWA)\Requests Executing" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $OAFE = $((Get-Counter "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_RPC)\Requests Executing" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $OA = $((Get-Counter "\ASP.NET Apps v4.0.30319(_LM_W3SVC_2_ROOT_RPC)\Requests Executing" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $EWSFE = $((Get-Counter "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_EWS)\Requests Executing" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $EWS = $((Get-Counter "\ASP.NET Apps v4.0.30319(_LM_W3SVC_2_ROOT_EWS)\Requests Executing" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $MAPIFE = $((Get-Counter "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_MAPI)\Requests Executing" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $MAPIBE = $((Get-Counter "\ASP.NET Apps v4.0.30319(_LM_W3SVC_2_ROOT_MAPI_EMSMDB)\Requests Executing" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
    }
    Else
    {
        $ActiveSync = $((Get-Counter "\MSExchange ActiveSync\Current Requests" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $OWA = $((Get-Counter "\MSExchange OWA\Current Unique Users" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $OA = $((Get-Counter "\RPC/HTTP Proxy\Current Number of Unique Users" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $RPC = $((Get-Counter "\MSExchange RpcClientAccess\User Count" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $EWS = $((Get-Counter "\W3SVC_W3WP(*msexchangeservicesapppool)\Active Requests" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $MAPI = $((Get-Counter "\MSExchange MapiHttp Emsmdb\Active User Count" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $MAPIFE = $((Get-Counter "\W3SVC_W3WP(*MSExchangeMapiFront*)\Active Requests" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
        $MAPIBE = $((Get-Counter "\W3SVC_W3WP(*MSExchangeMapiMailbox*)\Active Requests" -ComputerName ($AvailableServers | ForEach-Object{$_}) -MaxSamples $MaxSamples | ForEach-Object{($_.CounterSamples | Select-Object -ExpandProperty CookedValue | Measure-Object -Sum).Sum} | Measure-Object -Average).Average)
    }

    $obj = New-Object PSObject
    If ($UseASPDOTNET)
    {
        $obj | Add-Member NoteProperty -Name "Outlook Web App Request FE" -Value $([int][math]::Round($OWAFE))
        $obj | Add-Member NoteProperty -Name "Outlook Web App Request BE" -Value $([int][math]::Round($OWA))
        $obj | Add-Member NoteProperty -Name "ActiveSync Request Count FE" -Value $([int][math]::Round($ActiveSyncFE))
        $obj | Add-Member NoteProperty -Name "ActiveSync Request Count BE" -Value $([int][math]::Round($ActiveSync))
        $obj | Add-Member NoteProperty -Name "OutlookAnywhere Request Count FE" -Value $([int][math]::Round($OAFE))
        $obj | Add-Member NoteProperty -Name "OutlookAnywhere Request Count BE" -Value $([int][math]::Round($OA))
        $obj | Add-Member NoteProperty -Name "EWS Request Count FE" -Value $([int][math]::Round($EWSFE))
        $obj | Add-Member NoteProperty -Name "EWS Request Count BE" -Value $([int][math]::Round($EWS))
        $obj | Add-Member NoteProperty -Name "MAPI Request Count FE" -Value $([int][math]::Round($MAPIFE))
        $obj | Add-Member NoteProperty -Name "MAPI Request Count BE" -Value $([int][math]::Round($MAPIBE))
    }
    Else
    {
        $obj | Add-Member NoteProperty -Name "Outlook Web App" -Value $([math]::Round($OWA))
        $obj | Add-Member NoteProperty -Name "ActiveSync" -Value $([math]::Round($ActiveSync))
        $obj | Add-Member NoteProperty -Name "OutlookAnywhere" -Value $([math]::Round($OA))
        $obj | Add-Member NoteProperty -Name "RPC User Count" -Value $([math]::Round($RPC))
        $obj | Add-Member NoteProperty -Name "EWS User Count" -Value $([math]::Round($EWS))
        $obj | Add-Member NoteProperty -Name "MAPI User Count" -Value $([math]::Round($MAPI))
        $obj | Add-Member NoteProperty -Name "MAPI FE AppPool" -Value $([math]::Round($MAPIFE))
        $obj | Add-Member NoteProperty -Name "MAPI BE AppPool" -Value $([math]::Round($MAPIBE))
    }

    $obj
}

If ($HTTPProxyAVGLatency -or $HTTPProxyOutstandingRequests -or $HTTPProxyRequestsPerSec -or $TimeInGC)
{
    [System.Array]$Servers = GetExchServer -Roles 16439,16385 -ADSites $ADSite
}
Else
{
    [System.Array]$Servers = GetExchServer -Roles 4,36,38,54,16439,16385 -ADSites $ADSite
}
If ($SpecifiedServers)
{
    $Servers = $Servers | Where-Object {$SpecifiedServers -contains $_.Properties.name}
}

ForEach ($Server in $Servers)
{
    If (Test-Connection -ComputerName $Server.properties.name -Count 1 -Quiet)
    {
        [System.Array]$AvailableServers += $Server
    }
}

# find available servers
$AvailableServers = $AvailableServers | ForEach-Object{$_.properties.name} | Sort-Object
Write-Verbose "`nFound the following available server:`n$([String]::Join("`n",$AvailableServers))"

If ($HTTPProxyAVGLatency)
{
    [string[]]$counters = "\MSExchange HttpProxy(autodiscover)\Average ClientAccess Server Processing Latency",
    "\MSExchange HttpProxy(eas)\Average ClientAccess Server Processing Latency",
    "\MSExchange HttpProxy(ecp)\Average ClientAccess Server Processing Latency",
    "\MSExchange HttpProxy(ews)\Average ClientAccess Server Processing Latency",
    "\MSExchange HttpProxy(mapi)\Average ClientAccess Server Processing Latency",
    "\MSExchange HttpProxy(oab)\Average ClientAccess Server Processing Latency",
    "\MSExchange HttpProxy(owa)\Average ClientAccess Server Processing Latency",
    "\MSExchange HttpProxy(owacalendar)\Average ClientAccess Server Processing Latency",
    "\MSExchange HttpProxy(powershell)\Average ClientAccess Server Processing Latency",
    "\MSExchange HttpProxy(rpchttp)\Average ClientAccess Server Processing Latency"

    $objcol = @()
    $CounterStats = Get-Counter -ComputerName $AvailableServers -Counter $counters -MaxSamples $MaxSamples
    ForEach ($Server in $AvailableServers)
    {
        [System.Array]$AutoD = $null
        [System.Array]$EAS = $null
        [System.Array]$EWS = $null
        [System.Array]$MAPI = $null
        [System.Array]$RPCHttp = $null
        [System.Array]$OWA = $null
        [System.Array]$OWACal = $null
        [System.Array]$ECP = $null
        [System.Array]$Powershell = $null
        [System.Array]$OAB = $null
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        ForEach ($Sample in $CounterStats.CounterSamples | Where-Object {$_.Path -match $Server})
        {
            switch -wildcard ($Sample.Path)
            {
            '*autodiscover*' {$AutoD +=$Sample}
            '*eas*' {$EAS +=$Sample}
            '*ews*' {$EWS +=$Sample}
            '*mapi*' {$MAPI +=$Sample}
            '*rpchttp*' {$RPCHttp +=$Sample}
            '*owa)*' {$OWA +=$Sample}
            '*owacal*' {$OWACal +=$Sample}
            '*ecp*' {$ECP +=$Sample}
            '*powershe*' {$Powershell +=$Sample}
            '*oab*' {$OAB +=$Sample}
            }
        }

        $obj | Add-Member NoteProperty -Name "Autodiscover" -Value $([math]::Round(($AutoD | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "EAS" -Value $([math]::Round(($EAS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "EWS" -Value $([math]::Round(($EWS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI over Http" -Value $([math]::Round(($MAPI | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OutlookAnywhere" -Value $([math]::Round(($RPCHttp | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OWA" -Value $([math]::Round(($OWA | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OWACalendar" -Value $([math]::Round(($OWACal | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "ECP" -Value $([math]::Round(($ECP | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Powershell" -Value $([math]::Round(($Powershell | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OAB" -Value $([math]::Round(($OAB | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $objcol +=$obj
    }
$objcol
}

ElseIf ($HTTPProxyOutstandingRequests)
{
    [string[]]$counters = "\MSExchange HttpProxy(autodiscover)\Outstanding Proxy Requests",
    "\MSExchange HttpProxy(eas)\Outstanding Proxy Requests",
    "\MSExchange HttpProxy(ecp)\Outstanding Proxy Requests",
    "\MSExchange HttpProxy(ews)\Outstanding Proxy Requests",
    "\MSExchange HttpProxy(mapi)\Outstanding Proxy Requests",
    "\MSExchange HttpProxy(oab)\Outstanding Proxy Requests",
    "\MSExchange HttpProxy(owa)\Outstanding Proxy Requests",
    "\MSExchange HttpProxy(owacalendar)\Outstanding Proxy Requests",
    "\MSExchange HttpProxy(powershell)\Outstanding Proxy Requests",
    "\MSExchange HttpProxy(rpchttp)\Outstanding Proxy Requests"

    $objcol = @()
    $CounterStats = Get-Counter -ComputerName $AvailableServers -Counter $counters -MaxSamples $MaxSamples
    ForEach ($Server in $AvailableServers)
    {
        [System.Array]$AutoD = $null
        [System.Array]$EAS = $null
        [System.Array]$EWS = $null
        [System.Array]$MAPI = $null
        [System.Array]$RPCHttp = $null
        [System.Array]$OWA = $null
        [System.Array]$OWACal = $null
        [System.Array]$ECP = $null
        [System.Array]$Powershell = $null
        [System.Array]$OAB = $null
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        ForEach ($Sample in $CounterStats.CounterSamples | Where-Object {$_.Path -match $Server})
        {
            switch -wildcard ($Sample.Path)
            {
            '*autodiscover*' {$AutoD +=$Sample}
            '*eas*' {$EAS +=$Sample}
            '*ews*' {$EWS +=$Sample}
            '*mapi*' {$MAPI +=$Sample}
            '*rpchttp*' {$RPCHttp +=$Sample}
            '*owa)*' {$OWA +=$Sample}
            '*owacal*' {$OWACal +=$Sample}
            '*ecp*' {$ECP +=$Sample}
            '*powershe*' {$Powershell +=$Sample}
            '*oab*' {$OAB +=$Sample}
            }
        }

        $obj | Add-Member NoteProperty -Name "Autodiscover" -Value $([math]::Round(($AutoD | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "EAS" -Value $([math]::Round(($EAS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "EWS" -Value $([math]::Round(($EWS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI over Http" -Value $([math]::Round(($MAPI | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OutlookAnywhere" -Value $([math]::Round(($RPCHttp | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OWA" -Value $([math]::Round(($OWA | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OWACalendar" -Value $([math]::Round(($OWACal | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "ECP" -Value $([math]::Round(($ECP | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Powershell" -Value $([math]::Round(($Powershell | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OAB" -Value $([math]::Round(($OAB | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $objcol +=$obj
    }
$objcol
}

ElseIf ($HTTPProxyRequestsPerSec)
{
    [string[]]$counters = "\MSExchange HttpProxy(autodiscover)\Proxy Requests/Sec",
    "\MSExchange HttpProxy(eas)\Proxy Requests/Sec",
    "\MSExchange HttpProxy(ecp)\Proxy Requests/Sec",
    "\MSExchange HttpProxy(ews)\Proxy Requests/Sec",
    "\MSExchange HttpProxy(mapi)\Proxy Requests/Sec",
    "\MSExchange HttpProxy(oab)\Proxy Requests/Sec",
    "\MSExchange HttpProxy(owa)\Proxy Requests/Sec",
    "\MSExchange HttpProxy(owacalendar)\Proxy Requests/Sec",
    "\MSExchange HttpProxy(powershell)\Proxy Requests/Sec",
    "\MSExchange HttpProxy(rpchttp)\Proxy Requests/Sec"

    $objcol = @()
    $CounterStats = Get-Counter -ComputerName $AvailableServers -Counter $counters -MaxSamples $MaxSamples
    ForEach ($Server in $AvailableServers)
    {
        [System.Array]$AutoD = $null
        [System.Array]$EAS = $null
        [System.Array]$EWS = $null
        [System.Array]$MAPI = $null
        [System.Array]$RPCHttp = $null
        [System.Array]$OWA = $null
        [System.Array]$OWACal = $null
        [System.Array]$ECP = $null
        [System.Array]$Powershell = $null
        [System.Array]$OAB  = $null
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        ForEach ($Sample in $CounterStats.CounterSamples | Where-Object {$_.Path -match $Server})
        {
            switch -wildcard ($Sample.Path)
            {
            '*autodiscover*' {$AutoD +=$Sample}
            '*eas*' {$EAS +=$Sample}
            '*ews*' {$EWS +=$Sample}
            '*mapi*' {$MAPI +=$Sample}
            '*rpchttp*' {$RPCHttp +=$Sample}
            '*owa)*' {$OWA +=$Sample}
            '*owacal*' {$OWACal +=$Sample}
            '*ecp*' {$ECP +=$Sample}
            '*powershe*' {$Powershell +=$Sample}
            '*oab*' {$OAB +=$Sample}
            }
        }

        $obj | Add-Member NoteProperty -Name "Autodiscover" -Value $([math]::Round(($AutoD.CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "EAS" -Value $([math]::Round(($EAS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "EWS" -Value $([math]::Round(($EWS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI over Http" -Value $([math]::Round(($MAPI | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OutlookAnywhere" -Value $([math]::Round(($RPCHttp | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OWA" -Value $([math]::Round(($OWA | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OWACalendar" -Value $([math]::Round(($OWACal | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "ECP" -Value $([math]::Round(($ECP | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Powershell" -Value $([math]::Round(($Powershell | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OAB" -Value $([math]::Round(($OAB | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $objcol +=$obj
    }
$objcol
}

ElseIf ($E2EAVGLatency)
{
    [string[]]$counters = "\MSExchangeIS Client Type(administrator)\RPC Average Latency",
    "\MSExchangeIS Client Type(airsync)\RPC Average Latency",
    "\MSExchangeIS Client Type(availabilityservice)\RPC Average Latency",
    "\MSExchangeIS Client Type(momt)\RPC Average Latency",
    "\MSExchangeIS Client Type(owa)\RPC Average Latency",
    "\MSExchangeIS Client Type(rpchttp)\RPC Average Latency",
    "\MSExchangeIS Client Type(webservices)\RPC Average Latency",
    "\MSExchangeIS Client Type(outlookservice)\RPC Average Latency",
    "\MSExchangeIS Client Type(simplemigration)\RPC Average Latency",
    "\MSExchangeIS Client Type(contentindexing)\RPC Average Latency",
    "\MSExchangeIS Client Type(eventbasedassistants)\RPC Average Latency",
    "\MSExchangeIS Client Type(transport)\RPC Average Latency",
    "\MSExchange RpcClientAccess\RPC Averaged Latency",
    "\MSExchange MapiHttp Emsmdb\Averaged Latency"

    $objcol = @()
    $CounterStats = Get-Counter -ComputerName $AvailableServers -Counter $counters -MaxSamples $MaxSamples
    ForEach ($Server in $AvailableServers)
    {
        [System.Array]$administrator =$null
        [System.Array]$airsync =$null
        [System.Array]$availabilityservice =$null
        [System.Array]$momt =$null
        [System.Array]$rpchttp =$null
        [System.Array]$owa =$null
        [System.Array]$outlookservice =$null
        [System.Array]$simplemigration =$null
        [System.Array]$contentindexing =$null
        [System.Array]$eventbasedassistants =$null
        [System.Array]$transport =$null
        [System.Array]$RPCAveragedLatency =$null
        [System.Array]$MAPIAveragedLatency =$null
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        ForEach ($Sample in $CounterStats.CounterSamples | Where-Object {$_.Path -match $Server})
        {
            switch -wildcard ($Sample.Path)
            {
            '*administrator*' {$administrator += $Sample}
            '*airsync*' {$airsync += $Sample}
            '*availabilityservice*' {$availabilityservice += $Sample}
            '*momt*' {$momt += $Sample}
            '*rpchttp*' {$rpchttp += $Sample}
            '*owa*' {$owa += $Sample}
            '*outlookservice*' {$outlookservice += $Sample}
            '*simplemigration*' {$simplemigration += $Sample}
            '*contentindexing*' {$contentindexing += $Sample}
            '*eventbasedassistants*' {$eventbasedassistants += $Sample}
            '*transport*' {$transport += $Sample}
            '*RPC Averaged Latency*' {$RPCAveragedLatency += $Sample}
            '*MapiHttp Emsmdb\Averag*' {$MAPIAveragedLatency += $Sample}
            }
        }

        $obj | Add-Member NoteProperty -Name "RPC Averaged Latency" -Value $([math]::Round(($RPCAveragedLatency | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MapiHttp Averaged Latency" -Value $([math]::Round(($MAPIAveragedLatency | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "administrator" -Value $([math]::Round(($administrator | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "airsync" -Value $([math]::Round(($airsync | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "availabilityservice" -Value $([math]::Round(($availabilityservice | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "momt" -Value $([math]::Round(($momt | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "owa" -Value $([math]::Round(($owa | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "rpchttp" -Value $([math]::Round(($rpchttp | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "webservices" -Value $([math]::Round(($webservices | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "outlookservice" -Value $([math]::Round(($outlookservice | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "simplemigration" -Value $([math]::Round(($simplemigration | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "contentindexing" -Value $([math]::Round(($contentindexing | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "eventbasedassistants" -Value $([math]::Round(($eventbasedassistants | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "transport" -Value $([math]::Round(($transport | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "contentindexing" -Value $([math]::Round(($contentindexing | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $objcol +=$obj
    }
$objcol
}

ElseIf ($TimeInGC)
{
    [System.String[]]$counters ="\.NET CLR Memory(w3w*)\% Time in GC","\.NET CLR Memory(w3*)\Process ID","\W3SVC_W3WP(*)\Active Requests"
    $objcol = @()
    $CounterStats = Get-Counter -ComputerName $AvailableServers -Counter $counters -MaxSamples $MaxSamples
    Write-Verbose "Done with counter collection!"
    ForEach ($Server in $AvailableServers)
    {
        $MAPIFEProcID  = $null
        $MAPIFEWorker  = $null
        [System.Array]$MAPIFE = $null
        $MAPIBEProcID  = $null
        $MAPIBEWorker  = $null
        [System.Array]$MAPIBE = $null
        $EWSProcID = $null
        $EWSWorker = $null
        [System.Array]$EWS = $null
        $EASProcID = $null
        $EASWorker = $null
        [System.Array]$EAS = $null
        $RPCFEProcID = $null
        $RPCFEWorker = $null
        #[System.Array]$RPCFE = $null
        $RPCBEProcID = $null
        $RPCBEWorker = $null
        #[System.Array]$RPCBE = $null

        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        #filter counters for server
        $ServerStats = $CounterStats | Select-Object -ExpandProperty CounterSamples | Where-Object {$_.Path -match $Server}

        $MAPIFEProcID = ($ServerStats | Where-Object {$_.path -Match "msexchangemapifrontendapppool"})[0].InstanceName.Split("_")[0]
        $MAPIFEWorker = ($ServerStats | Where-Object {$_.CookedValue -EQ $($MAPIFEProcID)}| Select-Object -ExpandProperty Path -Unique).Split("(")[1].Split(")")[0]

        $MAPIBEProcID = ($ServerStats | Where-Object {$_.path -Match "msexchangemapimailboxapppool"})[0].InstanceName.Split("_")[0]
        $MAPIBEWorker = ($ServerStats | Where-Object {$_.CookedValue -EQ $($MAPIBEProcID)}| Select-Object -ExpandProperty Path -Unique).Split("(")[1].Split(")")[0]

        $EWSProcID    = ($ServerStats | Where-Object {$_.path -Match "msexchangeservicesapppool"})[0].InstanceName.Split("_")[0]
        $EWSWorker    = ($ServerStats | Where-Object {$_.CookedValue -EQ $($EWSProcID)}| Select-Object -ExpandProperty Path -Unique).Split("(")[1].Split(")")[0]

        $EASProcID    = ($ServerStats | Where-Object {$_.path -Match "msexchangesyncapppool"})[0].InstanceName.Split("_")[0]
        $EASWorker    = ($ServerStats | Where-Object {$_.CookedValue -EQ $($EASProcID)}| Select-Object -ExpandProperty Path -Unique).Split("(")[1].Split(")")[0]

        $RPCFEProcID  = ($ServerStats | Where-Object {$_.path -Match "msexchangerpcproxyfrontendapppool"})[0].InstanceName.Split("_")[0]
        $RPCFEWorker  = ($ServerStats | Where-Object {$_.CookedValue -EQ $($RPCFEProcID)}| Select-Object -ExpandProperty Path -Unique).Split("(")[1].Split(")")[0]

        $RPCBEProcID  = ($ServerStats | Where-Object {$_.path -Match "msexchangerpcproxyapppool"})[0].InstanceName.Split("_")[0]
        $RPCBEWorker  = ($ServerStats | Where-Object {$_.CookedValue -EQ $($RPCBEProcID)}| Select-Object -ExpandProperty Path -Unique).Split("(")[1].Split(")")[0]

        $Info = "`nServer:`t`t`t`t`t`t`t`t`t$($Server)`nMAPI FE ProcID:`t$($MAPIFEProcID)`nMAPI FE Worker:`t$($MAPIFEWorker)`nMAPI BE ProcID:`t$($MAPIBEProcID)"
        $Info+= "`nMAPI BE Worker:`t$($MAPIBEWorker)`nEWS ProcID:`t`t`t`t`t$($EWSProcID)`nEWS Worker:`t`t`t`t`t$($EWSWorker)`nEAS ProcID:`t`t`t`t`t$($EASProcID)"
        $Info+= "`nEAS Worker:`t`t`t`t`t$($EASWorker)`nRPC FE ProcID:`t`t$($RPCFEProcID)`nRPC FE Worker:`t`t$($RPCFEWorker)`nRPC BE ProcID:`t`t$($RPCBEProcID)`nRPC BE Worker:`t`t$($RPCBEWorker)"
        Write-Verbose $Info
        $obj | Add-Member NoteProperty -Name "MAPI FE Avg" -Value $("{0:N5}" -f ($Serverstats | Where-Object {$_.Path -match "$MAPIFEWorker\)\\% Time"} | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average)
        $obj | Add-Member NoteProperty -Name "MAPI FE Max" -Value $("{0:N5}" -f ($Serverstats | Where-Object {$_.Path -match "$MAPIFEWorker\)\\% Time"} | Select-Object -ExpandProperty CookedValue | Measure-Object -Maximum).Maximum)
        $obj | Add-Member NoteProperty -Name "MAPI BE Avg" -Value $("{0:N5}" -f ($Serverstats | Where-Object {$_.Path -match "$MAPIBEWorker\)\\% Time"} | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average)
        $obj | Add-Member NoteProperty -Name "MAPI BE Max" -Value $("{0:N5}" -f ($Serverstats | Where-Object {$_.Path -match "$MAPIBEWorker\)\\% Time"} | Select-Object -ExpandProperty CookedValue | Measure-Object -Maximum).Maximum)
        $obj | Add-Member NoteProperty -Name "EWS Avg" -Value $("{0:N5}" -f ($Serverstats | Where-Object {$_.Path -match "$EWSWorker\)\\% Time"} | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average)
        $obj | Add-Member NoteProperty -Name "EWS Max" -Value $("{0:N5}" -f ($Serverstats | Where-Object {$_.Path -match "$EWSWorker\)\\% Time"} | Select-Object -ExpandProperty CookedValue | Measure-Object -Maximum).Maximum)
        $obj | Add-Member NoteProperty -Name "EAS Avg" -Value $("{0:N5}" -f ($Serverstats | Where-Object {$_.Path -match "$EASWorker\)\\% Time"} | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average)
        $obj | Add-Member NoteProperty -Name "EAS Max" -Value $("{0:N5}" -f ($Serverstats | Where-Object {$_.Path -match "$EASWorker\)\\% Time"} | Select-Object -ExpandProperty CookedValue | Measure-Object -Maximum).Maximum)
        $obj | Add-Member NoteProperty -Name "RPC FE Avg" -Value $("{0:N5}" -f ($Serverstats | Where-Object {$_.Path -match "$RPCFEWorker\)\\% Time"} | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average)
        $obj | Add-Member NoteProperty -Name "RPC FE Max" -Value $("{0:N5}" -f ($Serverstats | Where-Object {$_.Path -match "$RPCFEWorker\)\\% Time"} | Select-Object -ExpandProperty CookedValue | Measure-Object -Maximum).Maximum)
        $obj | Add-Member NoteProperty -Name "RPC BE Avg" -Value $("{0:N5}" -f ($Serverstats | Where-Object {$_.Path -match "$RPCBEWorker\)\\% Time"} | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average)
        $obj | Add-Member NoteProperty -Name "RPC BE Max" -Value $("{0:N5}" -f ($Serverstats | Where-Object {$_.Path -match "$RPCBEWorker\)\\% Time"} | Select-Object -ExpandProperty CookedValue | Measure-Object -Maximum).Maximum)
        $objcol +=$obj
    }
$objcol
}

ElseIf ($IISMemoryUsage)
{
    If($UseCIM)
    {
        Write-Verbose "Create CimSession for $($AvailableServers)"
        $SessionOpt= New-CimSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
        $CimSessions = New-CimSession -ComputerName $AvailableServers -SkipTestConnection -ErrorVariable CimSessionError -SessionOption $SessionOpt -OperationTimeoutSec $CimTimeoutSec
        Write-Verbose "Collect data from Win32_Process"
        $CimData = Get-CimInstance -CimSession $CimSessions -ClassName Win32_Process -KeyOnly -Property CSName,ProcessID,CommandLine -Filter "NAME='w3wp.exe'" -ErrorVariable CimDataError
        Write-Verbose "Collect data from Win32_PerfRawData_PerfProc_Process"
        $WorkingSet = Get-CimInstance -CimSession $CimSessions -ClassName Win32_PerfRawData_PerfProc_Process -KeyOnly -Property WorkingSetPrivate,IDProcess -ErrorVariable WorkingError
        Write-Verbose "Collecting CIM data done"
        $objcol = @()
        ForEach ($SRV in $AvailableServers)
        {
            Write-Verbose "Process CIM data of $($Srv)"
            $obj = New-Object PSObject
            $obj | Add-Member NoteProperty -Name "Server" -Value $Srv
            $MemoryData = foreach($Set in $WorkingSet) {if($Set.CimSystemProperties.ServerName -Match $SRV){$Set}}
            $CimDataSrv = foreach($Data in $CimData) {if($Data.CSName -match $SRV){$Data} }
            $CimErrorSrv = foreach($CimError in $CimDataError) {if($CimError.OriginInfo.PSComputerName -match $SRV){$CimError}}
            $WorkingErrorSrv = foreach($WorkErr in $WorkingError) {if($WorkErr.OriginInfo.PSComputerName -match $SRV){$WorkErr}}
            If (($null -ne $MemoryData) -and ($null -ne $CimDataSrv))
            {
                [System.Int64]$TotalMem = 0
                ForEach ($Data in $CimDataSrv)
                {
                    [System.Int64]$Mem = [System.Math]::Round( ($MemoryData | Where-Object IDProcess -EQ $Data.ProcessId).WorkingSetPrivate/1MB)
                    $TotalMem += $Mem
                    Write-Verbose "$($Data.CommandLine.Split('"')[1]) $([system.math]::Round( ($MemoryData | Where-Object IDProcess -EQ $Data.ProcessId).WorkingSetPrivate/1MB))MB"
                    $obj | Add-Member NoteProperty -Name "$($Data.CommandLine.Split('"')[1])" -Value $Mem #([int] ([system.math]::Round( ($MemoryData | Where-Object IDProcess -EQ $Data.ProcessId).WorkingSetPrivate/1MB)))
                }
                $obj | Add-Member NoteProperty -Name "TotalMemUsage" -Value $TotalMem
                $obj | Add-Member NoteProperty -Name "CimError" -Value $null
            }

            If($CimErrorSrv -or $WorkingErrorSrv)
            {
                If($CimErrorSrv)
                {
                    $obj | Add-Member NoteProperty -Name "CimError" -Value "$($CimErrorSrv.Exception.Message) for Win32_Process"
                }
                If($WorkingErrorSrv)
                {
                    $obj | Add-Member NoteProperty -Name "CimError" -Value "$($WorkingErrorSrv.Exception.Message) for Win32_PerfRawData_PerfProc_Process"
                }
            }

            If($CimErrorSrv -or $WorkingErrorSrv)
            {
                $obj | Add-Member NoteProperty -Name "CimError" -Value "Both CIM queries failed"
            }

        $objcol +=$obj
        }
        Write-Verbose "Removing CimSessions"
        Remove-CimSession -CimSession $CimSessions -Confirm:$false
    }
    Else{
        [System.String[]]$counters = "\Process(w3wp*)\working set - private"
        $objcol = @()
        $CounterStats = Get-Counter -ComputerName $AvailableServers -Counter $counters -MaxSamples $MaxSamples
        ForEach ($Server in $AvailableServers)
        {
            [System.Int64]$TotalMem = 0
            $obj = New-Object PSObject
            $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
            ForEach ($Sample in $CounterStats.CounterSamples | Where-Object {$_.Path -match $Server})
            {
                [System.Int64]$Mem = [System.Math]::Round(($Sample| Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average/1MB)
                $TotalMem += $Mem
                $obj | Add-Member NoteProperty -Name "IISWorker($($Sample.Path.Split('(')[1].Split(')')[0]))" -Value $Mem
            }

            $obj | Add-Member NoteProperty -Name "TotalMemUsage" -Value $TotalMem
            $objcol +=$obj
        }
    }
$objcol
}
Else {
If ($Summary)
{
    If($SendMail)
    {
        $objcol = @()
        $objcol += GetCounterSum -UseASPDOTNET:$UseASPDOTNET
        $objcol
    }
    Else {
        GetCounterSum -UseASPDOTNET:$UseASPDOTNET
    }
}
ElseIf ($UseASPDOTNET)
{
    [System.String[]]$counters = "\ASP.NET Apps v4.0.30319(*)\requests executing"
    $objcol = @()
    $CounterStats = Get-Counter -ComputerName $AvailableServers -Counter $counters -MaxSamples $MaxSamples
    ForEach ($Server in $AvailableServers)
    {
        [System.Array]$Total = $null
        [System.Array]$CafeAutoDiscover = $null
        [System.Array]$CafeECP = $null
        [System.Array]$CafeEWS = $null
        [System.Array]$CafeMAPI = $null
        [System.Array]$CafeEAS = $null
        [System.Array]$CafeOAB = $null
        [System.Array]$CafeOWA = $null
        [System.Array]$CafeOWACal = $null
        [System.Array]$CafePowerShell = $null
        [System.Array]$CafeRPC = $null
        [System.Array]$MBAutoDiscover = $null
        [System.Array]$MBECP = $null
        [System.Array]$MBEWS = $null
        [System.Array]$MBMAPI = $null
        [System.Array]$MBEAS = $null
        [System.Array]$MBOAB = $null
        [System.Array]$MBOWA = $null
        [System.Array]$MBOWACal = $null
        [System.Array]$MBPowerShell = $null
        [System.Array]$MBRPC = $null
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        ForEach ($Sample in $CounterStats.CounterSamples | Where-Object {$_.Path -match $Server})
        {
            switch -wildcard ($Sample.Path)
            {
                '*__total__*' {$Total += $Sample}
                '*1_root_autodiscover*' {$CafeAutoDiscover += $Sample}
                '*1_root_ecp*' {$CafeECP += $Sample}
                '*1_root_ews*' {$CafeEWS += $Sample}
                '*1_root_mapi)*' {$CafeMAPI += $Sample}
                '*1_root_micro*' {$CafeEAS += $Sample}
                '*1_root_oab*' {$CafeOAB += $Sample}
                '*1_root_owa*' {$CafeOWA += $Sample}
                '*1_root_owa_*' {$CafeOWACal += $Sample}
                '*1_root_powe*' {$CafePowerShell += $Sample}
                '*1_root_rpc*'  {$CafeRPC += $Sample}
                '*2_root_autodiscover*' {$MBAutoDiscover += $Sample}
                '*2_root_ecp*' {$MBECP += $Sample}
                '*2_root_ews*' {$MBEWS += $Sample}
                '*2_root_mapi_emsmdb*' {$MBMAPI += $Sample}
                '*2_root_mapi_nspi*' {$MBMAPINSPI += $Sample}
                '*2_root_micro*' {$MBEAS += $Sample}
                '*2_root_oab*' {$MBOAB += $Sample}
                '*2_root_owa*' {$MBOWA += $Sample}
                '*2_root_owa_*'{$MBOWACal += $Sample}
                '*2_root_powe*'{$MBPowerShell += $Sample}
                '*2_root_rpc*' {$MBRPC += $Sample}
            }
        }
        $obj | Add-Member NoteProperty -Name "Total Requests" -Value $([int][math]::Round(($Total | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Cafe Autodiscover" -Value $([int][math]::Round(($CafeAutoDiscover | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Cafe ECP" -Value $([int][math]::Round(($CafeECP | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Cafe EWS" -Value $([int][math]::Round(($CafeEWS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Cafe MAPI" -Value $([int][math]::Round(($CafeMAPI | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Cafe EAS" -Value $([int][math]::Round(($CafeEAS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Cafe OAB" -Value $([int][math]::Round(($CafeOAB | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Cafe OWA" -Value $([int][math]::Round(($CafeOWA | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Cafe OWACal" -Value $([int][math]::Round(($CafeOWACal | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Cafe PowerShell" -Value $([int][math]::Round(($CafePowerShell | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Cafe RPC" -Value $([int][math]::Round(($CafeRPC | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MB Autodiscover" -Value $([int][math]::Round(($MBAutoDiscover | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MB ECP" -Value $([int][math]::Round(($MBECP | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MB EWS" -Value $([int][math]::Round(($MBEWS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MB MAPI" -Value $([int][math]::Round(($MBMAPI | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MB MAPI NSPI" -Value $([int][math]::Round(($MBMAPINSPI | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MB EAS" -Value $([int][math]::Round(($MBEAS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MB OAB" -Value $([int][math]::Round(($MBOAB | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MB OWA" -Value $([int][math]::Round(($MBOWA | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MB OWACal" -Value $([int][math]::Round(($MBOWACal | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MB PowerShell" -Value $([int][math]::Round(($MBPowerShell | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MB RPC" -Value $([int][math]::Round(($MBRPC | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))

        $objcol +=$obj
    }
    $objcol
}
Else {
    [string[]]$counters = "\MSExchange RpcClientAccess\User Count",
    "\MSExchange RpcClientAccess\Connection Count",
    "\RPC/HTTP Proxy\Current Number of Unique Users",
    "\MSExchange OWA\Current Unique Users",
    "\MSExchange ActiveSync\Current Requests",
    "\W3SVC_W3WP(*msexchangeservicesapppool)\Active Requests",
    "\Web Service(_Total)\Current Connections",
    "\Web Service(_Total)\Maximum Connections",
    "\Netlogon(_Total)\Semaphore Timeouts",
    "\MSExchange MapiHttp Emsmdb\Active User Count",
    "\MSExchange MapiHttp Emsmdb\Connection Count",
    "\W3SVC_W3WP(*MSExchangeMapiFront*)\Active Requests",
    "\W3SVC_W3WP(*MSExchangeMapiMailbox*)\Active Requests"

    $objcol = @()
    $CounterStats = Get-Counter -ComputerName $AvailableServers -Counter $counters -MaxSamples $MaxSamples
    ForEach ($Server in $AvailableServers)
    {
        [System.Array]$RPC = $null
        [System.Array]$RPCConn = $null
        [System.Array]$OA = $null
        [System.Array]$OWA = $null
        [System.Array]$EAS = $null
        [System.Array]$EWS = $null
        [System.Array]$IIS  = $null
        [System.Array]$IISMax = $null
        [System.Array]$Semaphore = $null
        [System.Array]$MAPIUser = $null
        [System.Array]$MAPICon = $null
        [System.Array]$MAPIFront = $null
        [System.Array]$MAPIBE = $null
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        ForEach ($Sample in $CounterStats.CounterSamples | Where-Object {$_.Path -match $Server}) {
            switch -wildcard ($Sample.Path)
            {
            '*RpcClientAccess\User Count' {$RPC += $Sample}
            '*RpcClientAccess\Conne*' {$RPCConn += $Sample}
            '*HTTP Proxy*' {$OA += $Sample}
            '*OWA\Curre*' {$OWA += $Sample}
            '*ActiveSync*' {$EAS += $Sample}
            '*icesapppool)\Acti*' {$EWS += $Sample}
            '*ice(_Total)\Curren*' {$IIS += $Sample}
            '*ice(_Total)\Maxi*' {$IISMax += $Sample}
            '*al)\Semaph*' {$Semaphore += $Sample}
            '*Emsmdb\Active*' {$MAPIUser += $Sample}
            '*Emsmdb\Conn*' {$MAPICon += $Sample}
            '*MapiFrontEnd*' {$MAPIFront += $Sample}
            '*MapiMailbox*' {$MAPIBE += $Sample}
            }
        }
        $MAPIRatio= $([math]::Round(($MAPIBE | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))/$([math]::Round(($MAPIFront | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "RPC User Count" -Value $([math]::Round(($RPC | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "RPC Connection Count" -Value $([math]::Round(($RPCConn | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "OutlookAnywhere" -Value $([math]::Round(($OA | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "Outlook Web App" -Value $([math]::Round(($OWA | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "ActiveSync" -Value $([math]::Round(($EAS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "EWS" -Value $([math]::Round(($EWS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI User count" -Value $([math]::Round(($MAPIUser | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI Connection Count" -Value $([math]::Round(($MAPICon | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI FE AppPool" -Value $([math]::Round(($MAPIFront | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI BE AppPool" -Value $([math]::Round(($MAPIBE | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI Ratio BE|FE" -Value $("{0:N2}" -f $MAPIRatio)
        $obj | Add-Member NoteProperty -Name "IIS Current Connection Count" -Value $([math]::Round(($IIS | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "IIS Max Connection Count" -Value $([math]::Round(($IISMax | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
        $obj | Add-Member NoteProperty -Name "SemaphoreTimeouts" -Value $([math]::Round(($Semaphore | Select-Object -ExpandProperty CookedValue | Measure-Object -Average).Average))
    $objcol +=$obj
    }
$objcol
}
}

If ($SendMail)
{
    #Build stamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH-mm-ss"

    # Some CSS to get a pretty report
$head = @'
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($ReportTitle)</title>
<style type="text/css">
<!-
body {
    font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}
h2{
    font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
    clear: both;
    font-size: 100%;
    color:#354B5E;
}
h3{
    font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
    clear: both;
    font-size: 75%;
    margin-left: 20px;
    margin-top: 30px;
    color:#475F77;
}
table{
    border-collapse: collapse;
    border: 1px solid black;
    font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
    color: black;
    margin-bottom: 10px;
}

table td{
    border: 1px solid black;
    font-size: 12px;
    padding-left: 5px;
    padding-right: 5px;
    text-align: left;
}

table th {
    border: 1px solid black;
    font-size: 12px;
    font-weight: bold;
    padding-left: 5px;
    padding-right: 5px;
    text-align: left;
}

TR:Hover TD {Background-Color: #C1D5F8;}

->
</style>
'@

    If ($HTTPProxyAVGLatency)
    {
        $stamp = "HTTPProxyAVGLatency_$($timestamp)"
        $subject = "HTTP Proxy AVG Latency"
        [System.String]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String #| Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd
    }
    ElseIf($HTTPProxyOutstandingRequests)
    {
        $stamp = "HTTPProxyOutstandingRequests_$($timestamp)"
        $subject = "HTTP Proxy Outstanding Requests"
        [System.String]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }
    ElseIf($HTTPProxyRequestsPerSec)
    {
        $stamp = "HTTPProxyRequestsPerSec_$($timestamp)"
        $subject = "HTTP Proxy Requests Per Sec"
        [System.String]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }
    ElseIf($E2EAVGLatency)
    {
        $stamp = "E2EAVGLatency_$($timestamp)"
        $subject = "E2E AVG Latency"
        [System.String]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }
    ElseIf($TimeInGC)
    {
        $stamp = "TimeInGC_$($timestamp)"
        $subject = "Time in GC"
        [System.String]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }
    ElseIf($Summary)
    {
        $stamp = "Summary_$($timestamp)"
        $subject = "Summary"
        [System.String]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }
    ElseIf($IISMemoryUsage)
    {
        $stamp = "IISMemoryUsage_$($timestamp)"
        $subject = "IIS Memory Usage"
        [System.String]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }
    Else{
        $stamp = "Overview_$($timestamp)"
        $subject = "Overview"
        [System.String]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }

    $objcol | Export-Csv -NoTypeInformation "$($stamp).csv" -Force -Encoding UTF8

    Send-MailMessage -Subject "Active Exchange User $($timestamp) | $($subject)" -From $From -To $Recipients -SmtpServer $SmtpServer -Body $body -BodyAsHtml -Attachments "$($stamp).csv"
}