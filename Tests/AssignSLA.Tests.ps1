BeforeAll {
    $script:ScriptPath = (Resolve-Path (Join-Path $PSScriptRoot '..' 'AssignSLA.ps1')).Path
    $script:StubModulePath = (Resolve-Path (Join-Path $PSScriptRoot 'TestHelpers' 'RubrikPolarisStub' 'RubrikPolarisStub.psd1')).Path
    $script:BulkCsvPath = (Resolve-Path (Join-Path $PSScriptRoot 'Fixtures' 'bulk-sites.csv')).Path
    $script:BulkCsvAllValidPath = (Resolve-Path (Join-Path $PSScriptRoot 'Fixtures' 'bulk-sites-all-valid.csv')).Path

    $script:Subscriptions = @([pscustomobject]@{ name = 'Contoso'; subscriptionId = 'sub-1' })
    $script:Slas = @(
        [pscustomobject]@{ name = 'Gold'; id = 'sla-gold' },
        [pscustomobject]@{ name = 'Gold-Extended'; id = 'sla-gold-ext' }
    )
    $script:SharePointSites = @(
        [pscustomobject]@{ name = 'Site A'; url = 'https://contoso.sharepoint.com/sites/A'; id = 'site-a-id' },
        [pscustomobject]@{ name = 'Site B'; url = 'https://contoso.sharepoint.com/sites/B'; id = 'site-b-id' }
    )
    $script:OneDrives = @(
        [pscustomobject]@{ name = 'Alice'; userPrincipalName = 'alice@contoso.com'; id = 'od-alice-id' }
    )
    $script:Mailboxes = @(
        [pscustomobject]@{ name = 'Bob'; userPrincipalName = 'bob@contoso.com'; id = 'mbx-bob-id' }
    )
}

Describe 'AssignSLA.ps1' {
    BeforeEach {
        Import-Module $script:StubModulePath -Force
        Set-RubrikPolarisStubData -Subscriptions $script:Subscriptions -Slas $script:Slas `
            -SharePointSites $script:SharePointSites -OneDrives $script:OneDrives -Mailboxes $script:Mailboxes
    }

    AfterEach {
        Remove-Module RubrikPolarisStub -Force -ErrorAction SilentlyContinue
    }

    Context 'Parameter validation' {
        It 'throws when neither -SearchByUrl nor -InputFile is supplied' {
            {
                & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold'
            } | Should -Throw
        }

        It 'throws when both -SearchByUrl and -InputFile are supplied' {
            {
                & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                    -SearchByUrl 'https://contoso.sharepoint.com/sites/A' -InputFile $script:BulkCsvPath
            } | Should -Throw
        }

        It 'throws when -InputFile does not exist' {
            {
                & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                    -InputFile (Join-Path $PSScriptRoot 'Fixtures' 'does-not-exist.csv')
            } | Should -Throw
        }
    }

    Context 'Connection and lookup failures' {
        It 'throws a clear error when Connect-Polaris fails' {
            Set-RubrikPolarisStubData -Subscriptions $script:Subscriptions -Slas $script:Slas `
                -SharePointSites $script:SharePointSites -ConnectShouldThrow

            {
                & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                    -SearchByUrl 'https://contoso.sharepoint.com/sites/A'
            } | Should -Throw '*connect*'
        }

        It 'throws when the subscription is not found' {
            {
                & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'DoesNotExist' -SlaDomain 'Gold' `
                    -SearchByUrl 'https://contoso.sharepoint.com/sites/A'
            } | Should -Throw '*subscription*'
        }

        It 'throws when the SLA Domain is not found' {
            {
                & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'DoesNotExist' `
                    -SearchByUrl 'https://contoso.sharepoint.com/sites/A'
            } | Should -Throw '*SLA*'
        }

        It 'throws when the single-object target is not found' {
            {
                & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                    -SearchByUrl 'https://contoso.sharepoint.com/sites/Missing'
            } | Should -Throw '*No SharePoint object matching*'
        }

        It 'throws when the single-object identifier matches more than one object' {
            Set-RubrikPolarisStubData -Subscriptions $script:Subscriptions -Slas $script:Slas -SharePointSites @(
                [pscustomobject]@{ name = 'Site A'; url = 'https://contoso.sharepoint.com/sites/A'; id = 'site-a-id' },
                [pscustomobject]@{ name = 'Site A (duplicate)'; url = 'https://contoso.sharepoint.com/sites/A'; id = 'site-a-dup-id' }
            )

            {
                & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                    -SearchByUrl 'https://contoso.sharepoint.com/sites/A'
            } | Should -Throw '*Multiple SharePoint objects matching*'
        }
    }

    Context 'Single-object mode (-SearchByUrl)' {
        It 'assigns the resolved SLA Domain id to the matched SharePoint site' {
            & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                -SearchByUrl 'https://contoso.sharepoint.com/sites/A'

            $assignments = @(Get-RubrikPolarisStubAssignments)
            $assignments.Count | Should -Be 1
            $assignments[0].ObjectID | Should -Be 'site-a-id'
            $assignments[0].SlaID | Should -Be 'sla-gold'
        }

        It 'passes the literal UNPROTECTED string through as the SlaID' {
            & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'UNPROTECTED' `
                -SearchByUrl 'https://contoso.sharepoint.com/sites/A'

            $assignments = @(Get-RubrikPolarisStubAssignments)
            $assignments.Count | Should -Be 1
            $assignments[0].SlaID | Should -Be 'UNPROTECTED'
        }

        It 'does not assign anything under -WhatIf' {
            & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                -SearchByUrl 'https://contoso.sharepoint.com/sites/A' -WhatIf

            @(Get-RubrikPolarisStubAssignments).Count | Should -Be 0
        }
    }

    Context 'Bulk mode (-InputFile)' {
        It 'assigns the SLA Domain to every matched row in the CSV' {
            & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                -InputFile $script:BulkCsvAllValidPath

            $assignments = @(Get-RubrikPolarisStubAssignments)
            $assignments.Count | Should -Be 2
            $assignments.ObjectID | Should -Contain 'site-a-id'
            $assignments.ObjectID | Should -Contain 'site-b-id'
        }

        It 'skips a row whose identifier cannot be resolved, still assigns the other rows, then reports the overall failure' {
            $Error.Clear()

            {
                & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                    -InputFile $script:BulkCsvPath 2>$null
            } | Should -Throw '*could not be assigned*'

            @(Get-RubrikPolarisStubAssignments).Count | Should -Be 2
            @($Error).Count | Should -BeGreaterThan 0
            (($Error | ForEach-Object { $_.ToString() }) -join ' ') | Should -Match 'Missing'
        }

        It 'skips (without aborting) a row whose identifier matches more than one object' {
            Set-RubrikPolarisStubData -Subscriptions $script:Subscriptions -Slas $script:Slas -SharePointSites @(
                [pscustomobject]@{ name = 'Site A'; url = 'https://contoso.sharepoint.com/sites/A'; id = 'site-a-id' },
                [pscustomobject]@{ name = 'Site A (duplicate)'; url = 'https://contoso.sharepoint.com/sites/A'; id = 'site-a-dup-id' },
                [pscustomobject]@{ name = 'Site B'; url = 'https://contoso.sharepoint.com/sites/B'; id = 'site-b-id' }
            )

            {
                & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                    -InputFile $script:BulkCsvAllValidPath 2>$null
            } | Should -Throw '*could not be assigned*'

            $assignments = @(Get-RubrikPolarisStubAssignments)
            $assignments.Count | Should -Be 1
            $assignments[0].ObjectID | Should -Be 'site-b-id'
        }

        It 'does not assign anything under -WhatIf' {
            & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                -InputFile $script:BulkCsvAllValidPath -WhatIf

            @(Get-RubrikPolarisStubAssignments).Count | Should -Be 0
        }
    }

    Context '-ObjectType dispatch' {
        It 'routes -ObjectType OneDrive to Get-PolarisM365OneDrives, matching by user principal name' {
            & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                -ObjectType 'OneDrive' -SearchByUrl 'alice@contoso.com'

            $called = Get-RubrikPolarisStubCalledFunctions
            $called | Should -Contain 'Get-PolarisM365OneDrives'
            $called | Should -Not -Contain 'Get-PolarisM365SharePoint'
            $called | Should -Not -Contain 'Get-PolarisM365Mailboxes'

            $assignments = @(Get-RubrikPolarisStubAssignments)
            $assignments[0].ObjectID | Should -Be 'od-alice-id'
        }

        It 'routes -ObjectType Mailbox to Get-PolarisM365Mailboxes, matching by user principal name' {
            & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                -ObjectType 'Mailbox' -SearchByUrl 'bob@contoso.com'

            $called = Get-RubrikPolarisStubCalledFunctions
            $called | Should -Contain 'Get-PolarisM365Mailboxes'
            $called | Should -Not -Contain 'Get-PolarisM365SharePoint'
            $called | Should -Not -Contain 'Get-PolarisM365OneDrives'

            $assignments = @(Get-RubrikPolarisStubAssignments)
            $assignments[0].ObjectID | Should -Be 'mbx-bob-id'
        }

        It 'defaults to -ObjectType SharePoint when not specified' {
            & $script:ScriptPath -PathToM365Module $script:StubModulePath -SubName 'Contoso' -SlaDomain 'Gold' `
                -SearchByUrl 'https://contoso.sharepoint.com/sites/A'

            Get-RubrikPolarisStubCalledFunctions | Should -Contain 'Get-PolarisM365SharePoint'
        }
    }
}
