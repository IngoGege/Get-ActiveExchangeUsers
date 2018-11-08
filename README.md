# Get-ActiveExchangeUsers

Retrieves multiple kinds of KPI of Exchange servers using Get-Counter.

### Prerequisites

# must be run in the context of an administrative account of Exchange servers as it pulls the data from performance counters

## Examples

# get HTTPProxy latency and send report to multiple recipients

```
.\Get-ActiveExchangeUser.ps1 -HTTPProxyAVGLatency -SendMail -From admin@contoso.com -Recipients rob@contoso.com,peter@contoso.com -SmtpServer smtp.contoso.local
```

# get IIS memory usage for application pools using CIM

```
.\Get-ActiveExchangeUser.ps1 -IISMemoryUsage -UseCIM
```

# get current requests for application pools, which is more reliable than using Exchange performance counters

```
.\Get-ActiveExchangeUser.ps1 -UseASPDOTNET
```

## Parameters

### -ADSite

here you can define in which ADSite is searched for Exchange server. If omitted current AD site will be used.

### -Summary

if used the script will sum up the active user count across all servers per protocol

### -UseASPDOTNET

switch to use IIS performance counters (ASP.NET Apps v4.0.30319\Requests Executing) for gathering current requests per protocol

### -HTTPProxyAVGLatency

the script will collect for each protocol the performance counter "\MSExchange HttpProxy(protocoll)\Average ClientAccess Server Processing Latency"

### -HTTPProxyOutstandingRequests

the script will collect for each protocol the performance counter "\MSExchange HttpProxy(protocoll)\Outstanding Proxy Requests"

### -HTTPProxyRequestsPerSec

the script will collect for each protocol the performance counter "\MSExchange HttpProxy(protocoll)\\Proxy Requests/Sec"

### -E2EAVGLatency

the script will collect for main protocols counters like "\MSExchangeIS Client Type(*)\RPC Average Latency",\MSExchange RpcClientAccess\RPC Averaged Latency","\MSExchange MapiHttp Emsmdb\Averaged Latency"

### -TimeInGC

collects and compute the average of the following GC performance counters "\.NET CLR Memory(w3w*)\% Time in GC","\.NET CLR Memory(w3*)\Process ID","\W3SVC_W3WP(*)\Active Requests"

### -SpecifiedServers

filtering for specific servers, which were found in given AD site

### -MaxSamples

as the script uses the CmdLet Get-Counter you can define the number of MaxSamples. Default is 1

### -SendMail

switch to send an e-mail with a CSV attached

### -From

define the sender address

### -Recipients

define the recipients

### -SmtpServer

which SmtpServer to use

### -IISMemoryUsage

collects the following performance counters "\Process(w3wp*)\working set - private". Note: This will return only the workes with no hint, which worker process it is. Use UseCIM switch for details.

### -UseCIM

Uses CIM for gathering IIS memory usage for application pools, which returns friendly name of application pools.

### -CimTimeoutSec

Timeout for CIM connection. Default is 30 seconds.

### About

For more information on this script, as well as usage and examples, see
the related blog article on [The Clueless Guy](https://ingogegenwarth.wordpress.com/2016/05/09/get-activeexchangeusers-2-0/).

## License

This project is licensed under the MIT License - see the LICENSE.md for details.