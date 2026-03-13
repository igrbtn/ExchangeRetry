BeforeAll {
    # Dot-source the main script to import functions
    . "$PSScriptRoot/../ExchangeRetry.ps1"
}

Describe 'ExchangeRetry Module' {

    Context 'Script loads without errors' {
        It 'Should define Show-ExchangeRetryGUI function' {
            Get-Command -Name Show-ExchangeRetryGUI -ErrorAction SilentlyContinue |
                Should -Not -BeNullOrEmpty
        }

        It 'Should define Connect-ExchangeRemote function' {
            Get-Command -Name Connect-ExchangeRemote -ErrorAction SilentlyContinue |
                Should -Not -BeNullOrEmpty
        }

        It 'Should define Get-ExchangeQueues function' {
            Get-Command -Name Get-ExchangeQueues -ErrorAction SilentlyContinue |
                Should -Not -BeNullOrEmpty
        }

        It 'Should define Invoke-QueueRetry function' {
            Get-Command -Name Invoke-QueueRetry -ErrorAction SilentlyContinue |
                Should -Not -BeNullOrEmpty
        }

        It 'Should define Invoke-MessageRetry function' {
            Get-Command -Name Invoke-MessageRetry -ErrorAction SilentlyContinue |
                Should -Not -BeNullOrEmpty
        }
    }

    Context 'Configuration' {
        It 'Should have Config hashtable with required keys' {
            $script:Config | Should -Not -BeNullOrEmpty
            $script:Config.Keys | Should -Contain 'ExchangeServer'
            $script:Config.Keys | Should -Contain 'DefaultPageSize'
            $script:Config.Keys | Should -Contain 'RefreshIntervalSec'
        }

        It 'DefaultPageSize should be a positive number' {
            $script:Config.DefaultPageSize | Should -BeGreaterThan 0
        }
    }

    Context 'Invoke-MessageRetry' {
        BeforeAll {
            # Mock Resume-Message to simulate Exchange cmdlet
            Mock Resume-Message { } -Verifiable
        }

        It 'Should return results for each message identity' {
            $results = Invoke-MessageRetry -MessageIdentity @('msg1', 'msg2')
            $results.Count | Should -Be 2
        }
    }
}
