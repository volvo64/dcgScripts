@{
    ModuleVersion = '1.0'
    NestedModules = @(
        '.\functions\Import-HostedExchangeGroupsFromCSV.ps1'
        '.\functions\Import-HostedExchangeMailboxesFromCSV.ps1'
        '.\functions\New-HostedExchangeCompany.ps1'
        '.\functions\New-HostedExchangeMailbox.ps1'
        '.\functions\New-HostedExchangeMailboxExportRequest.ps1'
        '.\functions\New-HostedExchangeMailContact.ps1'
        )

    FunctionsToExport = @('*')
    }