# Get-ActiveExchangeUsers

Retrieves multiple kinds of KPI of Exchange servers using Get-Counter.

### Prerequisites

* must be run in the context of an administrative account of Exchange servers

### Usage

The script has multiple parameters. Only a few are shown in this example:

```
.\Get-ActiveExchangeUser.ps1 -HTTPProxyAVGLatency -SendMail -From admin@contoso.com -Recipients rob@contoso.com,peter@contoso.com -SmtpServer smtp.contoso.local
```

### About

For more information on this script, as well as usage and examples, see
the related blog article on [The Clueless Guy](https://ingogegenwarth.wordpress.com/2016/05/09/get-activeexchangeusers-2-0/).

## License

This project is licensed under the MIT License - see the LICENSE.md for details.