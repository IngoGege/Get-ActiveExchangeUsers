<#

.SYNOPSIS

Created by: https://ingogegenwarth.wordpress.com/
Version:    42 ("What do you get if you multiply six by nine?")
Changed:    29.06.2016

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

.PARAMETER HTTPProxyAVGLatency

the script will collect for each protocol the performance counter "\MSExchange HttpProxy(protocoll)\Average ClientAccess Server Processing Latency"

.PARAMETER HTTPProxyOutstandingRequests

the script will collect for each protocol the performance counter "\MSExchange HttpProxy(protocoll)\Outstanding Proxy Requests"

.PARAMETER HTTPProxyRequestsPerSec

the script will collect for each protocol the performance counter "\MSExchange HttpProxy(protocoll)\\Proxy Requests/Sec"

.PARAMETER E2EAVGLatency

the script will collect for main protocols counters like "\MSExchangeIS Client Type(*)\RPC Average Latency",\MSExchange RpcClientAccess\RPC Averaged Latency","\MSExchange MapiHttp Emsmdb\Averaged Latency"

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
    [parameter( Mandatory=$false, Position=0)]
    [string]$ADSite="$(([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).GetDirectoryEntry().Name)",
    
    [parameter( Mandatory=$false, Position=1, ParameterSetName="Summary")]
    [switch]$Summary,
    
    [parameter( Mandatory=$false, Position=2, ParameterSetName="HTTPProxyAVGLatency")]
    [switch]$HTTPProxyAVGLatency,
    
    [parameter( Mandatory=$false, Position=3, ParameterSetName="HTTPProxyOutstandingRequests")]
    [switch]$HTTPProxyOutstandingRequests,
    
    [parameter( Mandatory=$false, Position=4, ParameterSetName="HTTPProxyRequestsPerSec")]
    [switch]$HTTPProxyRequestsPerSec,
    
    [parameter( Mandatory=$false, Position=5, ParameterSetName="E2EAVGLatency")]
    [switch]$E2EAVGLatency,

    [parameter( Mandatory=$false, Position=6, ParameterSetName="TimeInGC")]
    [switch]$TimeInGC,

    [parameter( Mandatory=$false, Position=7)]
    [array]$SpecifiedServers,

    [parameter( Mandatory=$false, Position=8)]
    [int]$MaxSamples = 1,

    [parameter( Mandatory=$false, Position=9)]
    [switch]$SendMail,

    [parameter( Mandatory=$false, Position=10)]
    [String]$From,

    [parameter( Mandatory=$false, Position=11)]
    [String[]]$Recipients,

    [parameter( Mandatory=$false, Position=12)]
    [string]$SmtpServer

)
$ErrorActionPreference = "silentlycontinue"
# function to get the Exchangeserver from AD site
function GetExchServer {
    #http://technet.microsoft.com/en-us/library/bb123496(v=exchg.80).aspx on the bottom there is a list of values
    param([array]$Roles,[string]$ADSite)
    Process {
        $valid = @("2","4","16","20","32","36","38","54","64","16385","16439")
        ForEach ($Role in $Roles){
            If (!($valid -contains $Role)) {
                Write-Output -fore red "Please use the following numbers: MBX=2,CAS=4,UM=16,HT=32,Edge=64 multirole servers:CAS/HT=36,CAS/MBX/HT=38,CAS/UM=20,E2k13 MBX=54,E2K13 CAS=16385,E2k13 CAS/MBX=16439"
                Break
            }
        }
        Function GetADSite {
            param([string]$Name)
            If (!($Name)) {
                [string]$Name = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).GetDirectoryEntry().Name
            }
            $FilterADSite = "(&(objectclass=site)(Name=$Name))"
            $RootADSite= ([ADSI]'LDAP://RootDse').configurationNamingContext
            $SearcherADSite = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"LDAP://$RootADSite")
            $SearcherADSite.Filter = "$FilterADSite"
            $SearcherADSite.pagesize = 1000
            $ResultsADSite = $SearcherADSite.FindOne()
            $ResultsADSite
        }
        $Filter = "(&(objectclass=msExchExchangeServer)(msExchServerSite=$((GetADSite -Name $ADSite).properties.distinguishedname))(|"
        ForEach ($Role in $Roles){
            $Filter += "(msexchcurrentserverroles=$Role)"
        }
        $Filter += "))"
        $Root= ([ADSI]'LDAP://RootDse').configurationNamingContext
        $Searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"LDAP://$Root")
        $Searcher.Filter = "$Filter"
        $Searcher.pagesize = 1000
        $Results = $Searcher.FindAll()
        $Results
    }
}

function GetCounterSum ()
{
    $ActiveSync = $((Get-Counter "\MSExchange ActiveSync\Current Requests" -ComputerName ($AvailableServers | %{$_}) -MaxSamples $MaxSamples | %{($_.CounterSamples | select -ExpandProperty CookedValue | measure -Sum).Sum} | measure -Average).Average)
    $OWA        = $((Get-Counter "\MSExchange OWA\Current Unique Users" -ComputerName ($AvailableServers | %{$_}) -MaxSamples $MaxSamples | %{($_.CounterSamples | select -ExpandProperty CookedValue | measure -Sum).Sum} | measure -Average).Average)
    $OA         = $((Get-Counter "\RPC/HTTP Proxy\Current Number of Unique Users" -ComputerName ($AvailableServers | %{$_}) -MaxSamples $MaxSamples | %{($_.CounterSamples | select -ExpandProperty CookedValue | measure -Sum).Sum} | measure -Average).Average)
    $RPC        = $((Get-Counter "\MSExchange RpcClientAccess\User Count" -ComputerName ($AvailableServers | %{$_}) -MaxSamples $MaxSamples | %{($_.CounterSamples | select -ExpandProperty CookedValue | measure -Sum).Sum} | measure -Average).Average)
    $EWS        = $((Get-Counter "\W3SVC_W3WP(*msexchangeservicesapppool)\Active Requests" -ComputerName ($AvailableServers | %{$_}) -MaxSamples $MaxSamples | %{($_.CounterSamples | select -ExpandProperty CookedValue | measure -Sum).Sum} | measure -Average).Average)
    $MAPI       = $((Get-Counter "\MSExchange MapiHttp Emsmdb\Active User Count" -ComputerName ($AvailableServers | %{$_}) -MaxSamples $MaxSamples | %{($_.CounterSamples | select -ExpandProperty CookedValue | measure -Sum).Sum} | measure -Average).Average)
    $MAPIFE     = $((Get-Counter "\W3SVC_W3WP(*MSExchangeMapiFront*)\Active Requests" -ComputerName ($AvailableServers | %{$_}) -MaxSamples $MaxSamples | %{($_.CounterSamples | select -ExpandProperty CookedValue | measure -Sum).Sum} | measure -Average).Average)
    $MAPIBE     = $((Get-Counter "\W3SVC_W3WP(*MSExchangeMapiMailbox*)\Active Requests" -ComputerName ($AvailableServers | %{$_}) -MaxSamples $MaxSamples | %{($_.CounterSamples | select -ExpandProperty CookedValue | measure -Sum).Sum} | measure -Average).Average)

    $obj = New-Object PSObject
    $obj | Add-Member NoteProperty -Name "Outlook Web App" -Value $([math]::Round($OWA))
    $obj | Add-Member NoteProperty -Name "ActiveSync" -Value $([math]::Round($ActiveSync))
    $obj | Add-Member NoteProperty -Name "OutlookAnywhere" -Value $([math]::Round($OA))
    $obj | Add-Member NoteProperty -Name "RPC User Count" -Value $([math]::Round($RPC))
    $obj | Add-Member NoteProperty -Name "EWS User Count" -Value $([math]::Round($EWS))
    $obj | Add-Member NoteProperty -Name "MAPI User Count" -Value $([math]::Round($MAPI))
    $obj | Add-Member NoteProperty -Name "MAPI FE AppPool" -Value $([math]::Round($MAPIFE))
    $obj | Add-Member NoteProperty -Name "MAPI BE AppPool" -Value $([math]::Round($MAPIBE))
    $obj
}

If ($HTTPProxyAVGLatency -or $HTTPProxyOutstandingRequests -or $HTTPProxyRequestsPerSec -or $TimeInGC) {
    [array]$Servers = GetExchServer -Roles 16439,16385 -ADSite $ADSite
}
Else {
    [array]$Servers = GetExchServer -Roles 4,36,38,54,16439,16385 -ADSite $ADSite
}
If ($SpecifiedServers) {
    $Servers = $Servers | ? {$SpecifiedServers -contains $_.Properties.name}
}
ForEach ($Server in $Servers) {
    If (Test-Connection -ComputerName $Server.properties.name -Count 1 -Quiet) {
        [array]$AvailableServers += $Server
    }
}
#find available servers
$AvailableServers = $AvailableServers | %{$_.properties.name} | sort
Write-Verbose "`nFound the following available server:`n$([String]::Join("`n",$AvailableServers))"

If ($HTTPProxyAVGLatency) {
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
    ForEach ($Server in $AvailableServers) {
        [array]$AutoD       = ""
        [array]$EAS         = ""
        [array]$EWS         = ""
        [array]$MAPI        = ""
        [array]$RPCHttp     = ""
        [array]$OWA         = ""
        [array]$OWACal      = ""
        [array]$ECP         = ""
        [array]$Powershell  = ""
        [array]$OAB         = ""
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        ForEach ($Sample in $CounterStats.CounterSamples | ?{$_.Path -match $Server}) {
            switch -wildcard ($Sample.Path)
            {
            '*autodiscover*'    {$AutoD         +=$Sample}
            '*eas*'             {$EAS           +=$Sample}
            '*ews*'             {$EWS           +=$Sample}
            '*mapi*'            {$MAPI          +=$Sample}
            '*rpchttp*'         {$RPCHttp       +=$Sample}
            '*owa)*'            {$OWA           +=$Sample}
            '*owacal*'          {$OWACal        +=$Sample}
            '*ecp*'             {$ECP           +=$Sample}
            '*powershe*'        {$Powershell    +=$Sample}
            '*oab*'             {$OAB           +=$Sample}
            }
        }
        $obj | Add-Member NoteProperty -Name "Autodiscover" -Value $([math]::Round(($AutoD | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "EAS" -Value $([math]::Round(($EAS | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "EWS" -Value $([math]::Round(($EWS | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI over Http" -Value $([math]::Round(($MAPI | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OutlookAnywhere" -Value $([math]::Round(($RPCHttp | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OWA" -Value $([math]::Round(($OWA | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OWACalendar" -Value $([math]::Round(($OWACal | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "ECP" -Value $([math]::Round(($ECP | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "Powershell" -Value $([math]::Round(($Powershell | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OAB" -Value $([math]::Round(($OAB | select -ExpandProperty CookedValue | measure -Average).Average))
        $objcol +=$obj
    }
$objcol
}
ElseIf ($HTTPProxyOutstandingRequests) {
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
    ForEach ($Server in $AvailableServers) {
        [array]$AutoD       = ""
        [array]$EAS         = ""
        [array]$EWS         = ""
        [array]$MAPI        = ""
        [array]$RPCHttp     = ""
        [array]$OWA         = ""
        [array]$OWACal      = ""
        [array]$ECP         = ""
        [array]$Powershell  = ""
        [array]$OAB         = ""
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        ForEach ($Sample in $CounterStats.CounterSamples | ?{$_.Path -match $Server}) {
            switch -wildcard ($Sample.Path)
            {
            '*autodiscover*'    {$AutoD         +=$Sample}
            '*eas*'             {$EAS           +=$Sample}
            '*ews*'             {$EWS           +=$Sample}
            '*mapi*'            {$MAPI          +=$Sample}
            '*rpchttp*'         {$RPCHttp       +=$Sample}
            '*owa)*'            {$OWA           +=$Sample}
            '*owacal*'          {$OWACal        +=$Sample}
            '*ecp*'             {$ECP           +=$Sample}
            '*powershe*'        {$Powershell    +=$Sample}
            '*oab*'             {$OAB           +=$Sample}
            }
        }
        $obj | Add-Member NoteProperty -Name "Autodiscover" -Value $([math]::Round(($AutoD | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "EAS" -Value $([math]::Round(($EAS | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "EWS" -Value $([math]::Round(($EWS | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI over Http" -Value $([math]::Round(($MAPI | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OutlookAnywhere" -Value $([math]::Round(($RPCHttp | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OWA" -Value $([math]::Round(($OWA | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OWACalendar" -Value $([math]::Round(($OWACal | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "ECP" -Value $([math]::Round(($ECP | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "Powershell" -Value $([math]::Round(($Powershell | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OAB" -Value $([math]::Round(($OAB | select -ExpandProperty CookedValue | measure -Average).Average))
        $objcol +=$obj
    }
$objcol
}
ElseIf ($HTTPProxyRequestsPerSec) {
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
    ForEach ($Server in $AvailableServers) {
        [array]$AutoD       = ""
        [array]$EAS         = ""
        [array]$EWS         = ""
        [array]$MAPI        = ""
        [array]$RPCHttp     = ""
        [array]$OWA         = ""
        [array]$OWACal      = ""
        [array]$ECP         = ""
        [array]$Powershell  = ""
        [array]$OAB         = ""
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        ForEach ($Sample in $CounterStats.CounterSamples | ?{$_.Path -match $Server}) {
            switch -wildcard ($Sample.Path)
            {
            '*autodiscover*'    {$AutoD         +=$Sample}
            '*eas*'             {$EAS           +=$Sample}
            '*ews*'             {$EWS           +=$Sample}
            '*mapi*'            {$MAPI          +=$Sample}
            '*rpchttp*'         {$RPCHttp       +=$Sample}
            '*owa)*'            {$OWA           +=$Sample}
            '*owacal*'          {$OWACal        +=$Sample}
            '*ecp*'             {$ECP           +=$Sample}
            '*powershe*'        {$Powershell    +=$Sample}
            '*oab*'             {$OAB           +=$Sample}
            }
        }
        $obj | Add-Member NoteProperty -Name "Autodiscover" -Value $([math]::Round(($AutoD.CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "EAS" -Value $([math]::Round(($EAS | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "EWS" -Value $([math]::Round(($EWS | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI over Http" -Value $([math]::Round(($MAPI | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OutlookAnywhere" -Value $([math]::Round(($RPCHttp | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OWA" -Value $([math]::Round(($OWA | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OWACalendar" -Value $([math]::Round(($OWACal | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "ECP" -Value $([math]::Round(($ECP | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "Powershell" -Value $([math]::Round(($Powershell | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OAB" -Value $([math]::Round(($OAB | select -ExpandProperty CookedValue | measure -Average).Average))
    $objcol +=$obj
    }
$objcol
}
ElseIf ($E2EAVGLatency) {
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
    ForEach ($Server in $AvailableServers) {
        [array]$administrator       =""
        [array]$airsync             =""
        [array]$availabilityservice =""
        [array]$momt                =""
        [array]$rpchttp             =""
        [array]$owa                 =""
        [array]$outlookservice      =""
        [array]$simplemigration     =""
        [array]$contentindexing     =""
        [array]$eventbasedassistants=""
        [array]$transport           =""
        [array]$RPCAveragedLatency  =""
        [array]$MAPIAveragedLatency =""
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        ForEach ($Sample in $CounterStats.CounterSamples | ?{$_.Path -match $Server}) {
                switch -wildcard ($Sample.Path)
                {
                '*administrator*'           {$administrator         +=$Sample}
                '*airsync*'                 {$airsync               +=$Sample}
                '*availabilityservice*'     {$availabilityservice   +=$Sample}
                '*momt*'                    {$momt                  +=$Sample}
                '*rpchttp*'                 {$rpchttp               +=$Sample}
                '*owa*'                     {$owa                   +=$Sample}
                '*outlookservice*'          {$outlookservice        +=$Sample}
                '*simplemigration*'         {$simplemigration       +=$Sample}
                '*contentindexing*'         {$contentindexing       +=$Sample}
                '*eventbasedassistants*'    {$eventbasedassistants  +=$Sample}
                '*transport*'               {$transport             +=$Sample}
                '*RPC Averaged Latency*'    {$RPCAveragedLatency    +=$Sample}
                '*MapiHttp Emsmdb\Averag*'  {$MAPIAveragedLatency   +=$Sample}
                }
            }
            $obj | Add-Member NoteProperty -Name "RPC Averaged Latency" -Value $([math]::Round(($RPCAveragedLatency | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "MapiHttp Averaged Latency" -Value $([math]::Round(($MAPIAveragedLatency | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "administrator" -Value $([math]::Round(($administrator | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "airsync" -Value $([math]::Round(($airsync | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "availabilityservice" -Value $([math]::Round(($availabilityservice | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "momt" -Value $([math]::Round(($momt | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "owa" -Value $([math]::Round(($owa | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "rpchttp" -Value $([math]::Round(($rpchttp | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "webservices" -Value $([math]::Round(($webservices | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "outlookservice" -Value $([math]::Round(($outlookservice | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "simplemigration" -Value $([math]::Round(($simplemigration | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "contentindexing" -Value $([math]::Round(($contentindexing | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "eventbasedassistants" -Value $([math]::Round(($eventbasedassistants | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "transport" -Value $([math]::Round(($transport | select -ExpandProperty CookedValue | measure -Average).Average))
            $obj | Add-Member NoteProperty -Name "contentindexing" -Value $([math]::Round(($contentindexing | select -ExpandProperty CookedValue | measure -Average).Average))
        $objcol +=$obj
    }
$objcol
}
ElseIf ($TimeInGC) {
    [string[]]$counters ="\.NET CLR Memory(w3w*)\% Time in GC","\.NET CLR Memory(w3*)\Process ID","\W3SVC_W3WP(*)\Active Requests"
    $objcol = @()
    $CounterStats = Get-Counter -ComputerName $AvailableServers -Counter $counters -MaxSamples $MaxSamples
    Write-Verbose "Done with counter collection!"
    ForEach ($Server in $AvailableServers) {
        $MAPIFEProcID  = ""
        $MAPIFEWorker  = ""
        [array]$MAPIFE = ""
        $MAPIBEProcID  = ""
        $MAPIBEWorker  = ""
        [array]$MAPIBE = ""
        $EWSProcID     = ""
        $EWSWorker     = ""
        [array]$EWS    = ""
        $EASProcID     = ""
        $EASWorker     = ""
        [array]$EAS    = ""
        $RPCFEProcID   = ""
        $RPCFEWorker   = ""
        [array]$RPCFE  = ""
        $RPCBEProcID   = ""
        $RPCBEWorker   = ""
        [array]$RPCBE   = ""

        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        #filter counters for server
        $ServerStats = $CounterStats | select -ExpandProperty CounterSamples | ?{$_.Path -match $Server}

        $MAPIFEProcID = ($ServerStats | ?{$_.path -Match "msexchangemapifrontendapppool"})[0].InstanceName.Split("_")[0]
        $MAPIFEWorker = ($ServerStats | ?{$_.CookedValue -EQ $($MAPIFEProcID)}| select -ExpandProperty Path -Unique).Split("(")[1].Split(")")[0]

        $MAPIBEProcID = ($ServerStats | ?{$_.path -Match "msexchangemapimailboxapppool"})[0].InstanceName.Split("_")[0]
        $MAPIBEWorker = ($ServerStats | ?{$_.CookedValue -EQ $($MAPIBEProcID)}| select -ExpandProperty Path -Unique).Split("(")[1].Split(")")[0]

        $EWSProcID    = ($ServerStats | ?{$_.path -Match "msexchangeservicesapppool"})[0].InstanceName.Split("_")[0]
        $EWSWorker    = ($ServerStats | ?{$_.CookedValue -EQ $($EWSProcID)}| select -ExpandProperty Path -Unique).Split("(")[1].Split(")")[0]

        $EASProcID    = ($ServerStats | ?{$_.path -Match "msexchangesyncapppool"})[0].InstanceName.Split("_")[0]
        $EASWorker    = ($ServerStats | ?{$_.CookedValue -EQ $($EASProcID)}| select -ExpandProperty Path -Unique).Split("(")[1].Split(")")[0]

        $RPCFEProcID  = ($ServerStats | ?{$_.path -Match "msexchangerpcproxyfrontendapppool"})[0].InstanceName.Split("_")[0]
        $RPCFEWorker  = ($ServerStats | ?{$_.CookedValue -EQ $($RPCFEProcID)}| select -ExpandProperty Path -Unique).Split("(")[1].Split(")")[0]

        $RPCBEProcID  = ($ServerStats | ?{$_.path -Match "msexchangerpcproxyapppool"})[0].InstanceName.Split("_")[0]
        $RPCBEWorker  = ($ServerStats | ?{$_.CookedValue -EQ $($RPCBEProcID)}| select -ExpandProperty Path -Unique).Split("(")[1].Split(")")[0]

        $Info = "`nServer:`t`t`t`t`t`t`t`t`t$($Server)`nMAPI FE ProcID:`t$($MAPIFEProcID)`nMAPI FE Worker:`t$($MAPIFEWorker)`nMAPI BE ProcID:`t$($MAPIBEProcID)"
        $Info+= "`nMAPI BE Worker:`t$($MAPIBEWorker)`nEWS ProcID:`t`t`t`t`t$($EWSProcID)`nEWS Worker:`t`t`t`t`t$($EWSWorker)`nEAS ProcID:`t`t`t`t`t$($EASProcID)"
        $Info+= "`nEAS Worker:`t`t`t`t`t$($EASWorker)`nRPC FE ProcID:`t`t$($RPCFEProcID)`nRPC FE Worker:`t`t$($RPCFEWorker)`nRPC BE ProcID:`t`t$($RPCBEProcID)`nRPC BE Worker:`t`t$($RPCBEWorker)"
        Write-Verbose $Info
        $obj | Add-Member NoteProperty -Name "MAPI FE Avg" -Value $("{0:N5}" -f ($Serverstats | ?{$_.Path -match "$MAPIFEWorker\)\\% Time"} | select -ExpandProperty CookedValue | measure -Average).Average)
        $obj | Add-Member NoteProperty -Name "MAPI FE Max" -Value $("{0:N5}" -f ($Serverstats | ?{$_.Path -match "$MAPIFEWorker\)\\% Time"} | select -ExpandProperty CookedValue | measure -Maximum).Maximum)
        $obj | Add-Member NoteProperty -Name "MAPI BE Avg" -Value $("{0:N5}" -f ($Serverstats | ?{$_.Path -match "$MAPIBEWorker\)\\% Time"} | select -ExpandProperty CookedValue | measure -Average).Average)
        $obj | Add-Member NoteProperty -Name "MAPI BE Max" -Value $("{0:N5}" -f ($Serverstats | ?{$_.Path -match "$MAPIBEWorker\)\\% Time"} | select -ExpandProperty CookedValue | measure -Maximum).Maximum)
        $obj | Add-Member NoteProperty -Name "EWS Avg" -Value $("{0:N5}" -f ($Serverstats | ?{$_.Path -match "$EWSWorker\)\\% Time"} | select -ExpandProperty CookedValue | measure -Average).Average)
        $obj | Add-Member NoteProperty -Name "EWS Max" -Value $("{0:N5}" -f ($Serverstats | ?{$_.Path -match "$EWSWorker\)\\% Time"} | select -ExpandProperty CookedValue | measure -Maximum).Maximum)
        $obj | Add-Member NoteProperty -Name "EAS Avg" -Value $("{0:N5}" -f ($Serverstats | ?{$_.Path -match "$EASWorker\)\\% Time"} | select -ExpandProperty CookedValue | measure -Average).Average)
        $obj | Add-Member NoteProperty -Name "EAS Max" -Value $("{0:N5}" -f ($Serverstats | ?{$_.Path -match "$EASWorker\)\\% Time"} | select -ExpandProperty CookedValue | measure -Maximum).Maximum)
        $obj | Add-Member NoteProperty -Name "RPC FE Avg" -Value $("{0:N5}" -f ($Serverstats | ?{$_.Path -match "$RPCFEWorker\)\\% Time"} | select -ExpandProperty CookedValue | measure -Average).Average)
        $obj | Add-Member NoteProperty -Name "RPC FE Max" -Value $("{0:N5}" -f ($Serverstats | ?{$_.Path -match "$RPCFEWorker\)\\% Time"} | select -ExpandProperty CookedValue | measure -Maximum).Maximum)
        $obj | Add-Member NoteProperty -Name "RPC BE Avg" -Value $("{0:N5}" -f ($Serverstats | ?{$_.Path -match "$RPCBEWorker\)\\% Time"} | select -ExpandProperty CookedValue | measure -Average).Average)
        $obj | Add-Member NoteProperty -Name "RPC BE Max" -Value $("{0:N5}" -f ($Serverstats | ?{$_.Path -match "$RPCBEWorker\)\\% Time"} | select -ExpandProperty CookedValue | measure -Maximum).Maximum)

        $objcol +=$obj
    }
$objcol
}
Else {
If ($Summary) {
    if($SendMail) {
        $objcol = @()
        $objcol += GetCounterSum
        $objcol | ft -a
    }
    else {
        GetCounterSum | ft -a
    }
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
    ForEach ($Server in $AvailableServers) {
        [array]$RPC         =""
        [array]$RPCConn     =""
        [array]$OA          =""
        [array]$OWA         =""
        [array]$EAS         =""
        [array]$EWS         =""
        [array]$IIS         =""
        [array]$IISMax      =""
        [array]$Semaphore   =""
        [array]$MAPIUser    =""
        [array]$MAPICon     =""
        [array]$MAPIFront   =""
        [array]$MAPIBE      =""
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "Server" -Value $($Server)
        ForEach ($Sample in $CounterStats.CounterSamples | ?{$_.Path -match $Server}) {
            switch -wildcard ($Sample.Path)
            {
            '*RpcClientAccess\User Count'   {$RPC       +=$Sample}
            '*RpcClientAccess\Conne*'       {$RPCConn   +=$Sample}
            '*HTTP Proxy*'                  {$OA        +=$Sample}
            '*OWA\Curre*'                   {$OWA       +=$Sample}
            '*ActiveSync*'                  {$EAS       +=$Sample}
            '*icesapppool)\Acti*'           {$EWS       +=$Sample}
            '*ice(_Total)\Curren*'          {$IIS       +=$Sample}
            '*ice(_Total)\Maxi*'            {$IISMax    +=$Sample}
            '*al)\Semaph*'                  {$Semaphore +=$Sample}
            '*Emsmdb\Active*'               {$MAPIUser  +=$Sample}
            '*Emsmdb\Conn*'                 {$MAPICon   +=$Sample}
            '*MapiFrontEnd*'                {$MAPIFront +=$Sample}
            '*MapiMailbox*'                 {$MAPIBE +=$Sample}
            }
        }
        $MAPIRatio= $([math]::Round(($MAPIBE | select -ExpandProperty CookedValue | measure -Average).Average))/$([math]::Round(($MAPIFront | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "RPC User Count" -Value $([math]::Round(($RPC | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "RPC Connection Count" -Value $([math]::Round(($RPCConn | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "OutlookAnywhere" -Value $([math]::Round(($OA | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "Outlook Web App" -Value $([math]::Round(($OWA | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "ActiveSync" -Value $([math]::Round(($EAS | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "EWS" -Value $([math]::Round(($EWS | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI User count" -Value $([math]::Round(($MAPIUser | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI Connection Count" -Value $([math]::Round(($MAPICon | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI FE AppPool" -Value $([math]::Round(($MAPIFront | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI BE AppPool" -Value $([math]::Round(($MAPIBE | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "MAPI Ratio BE|FE" -Value $("{0:N2}" -f $MAPIRatio) #$("{0:N5}" -f $([math]::Round(($MAPIBE | select -ExpandProperty CookedValue | measure -Average).Average))/$([math]::Round(($MAPIFront | select -ExpandProperty CookedValue | measure -Average).Average)))
        $obj | Add-Member NoteProperty -Name "IIS Current Connection Count" -Value $([math]::Round(($IIS | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "IIS Max Connection Count" -Value $([math]::Round(($IISMax | select -ExpandProperty CookedValue | measure -Average).Average))
        $obj | Add-Member NoteProperty -Name "SemaphoreTimeouts" -Value $([math]::Round(($Semaphore | select -ExpandProperty CookedValue | measure -Average).Average))
    $objcol +=$obj
    }
$objcol
}
}

If ($SendMail) {
    #Build stamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH-mm-ss"

    # Some CSS to get a pretty report
$head = @'
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($ReportTitle)</title>
<style type=”text/css”>
<!–
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

    If ($HTTPProxyAVGLatency){
        $stamp = "HTTPProxyAVGLatency_$($timestamp)"
        $subject = "HTTP Proxy AVG Latency"
        [string]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String #| Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd
    }
    ElseIf($HTTPProxyOutstandingRequests){
        $stamp = "HTTPProxyOutstandingRequests_$($timestamp)"
        $subject = "HTTP Proxy Outstanding Requests"
        [string]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }
    ElseIf($HTTPProxyRequestsPerSec){
        $stamp = "HTTPProxyRequestsPerSec_$($timestamp)"
        $subject = "HTTP Proxy Requests Per Sec"
        [string]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }
    ElseIf($E2EAVGLatency){
        $stamp = "E2EAVGLatency_$($timestamp)"
        $subject = "E2E AVG Latency"
        [string]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }
    ElseIf($TimeInGC){
        $stamp = "TimeInGC_$($timestamp)"
        $subject = "Time in GC"
        [string]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }
    ElseIf($Summary){
        $stamp = "Summary_$($timestamp)"
        $subject = "Summary"
        [string]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }
    Else{
        $stamp = "Overview_$($timestamp)"
        $subject = "Overview"
        [string]$body = $objcol | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($subject)</h2>" | Out-String
    }

    $objcol | Export-Csv -NoTypeInformation "$($stamp).csv" -Force -Encoding UTF8

    Send-MailMessage -Subject "Active Exchange User $($timestamp) | $($subject)" -From $From -To $Recipients -SmtpServer $SmtpServer -Body $body -BodyAsHtml -Attachments "$($stamp).csv" 
}