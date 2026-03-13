BeforeAll {
    # Dot-source the trace script functions only (skip param block execution)
    # We source only the function definitions
    $scriptContent = Get-Content -Path "$PSScriptRoot/../ExchangeTrace.ps1" -Raw

    # Extract and define functions
    $functionBlocks = [regex]::Matches($scriptContent, '(?ms)(function\s+[\w-]+\s*\{.+?\n\})')
    foreach ($block in $functionBlocks) {
        Invoke-Expression $block.Value
    }
}

Describe 'Parse-EmailHeaders' {

    Context 'Basic header parsing' {
        BeforeAll {
            $testHeaders = @"
Received: from mail-out.example.com (mail-out.example.com [203.0.113.10])
 by mx.target.com (Postfix) with ESMTPS id ABC123;
 Thu, 12 Mar 2026 10:30:15 +0000
Received: from internal.example.com (internal.example.com [10.0.0.5])
 by mail-out.example.com with ESMTP id DEF456;
 Thu, 12 Mar 2026 10:30:10 +0000
Message-ID: <test123@example.com>
From: sender@example.com
To: recipient@target.com
Subject: Test Message
Date: Thu, 12 Mar 2026 10:30:00 +0000
Authentication-Results: mx.target.com; spf=pass; dkim=pass; dmarc=pass
X-MS-Exchange-Organization-SCL: 0
X-MS-Exchange-Organization-AuthSource: mail-out.example.com
"@
            $script:parsed = Parse-EmailHeaders -RawHeaders $testHeaders
        }

        It 'Should extract Message-ID' {
            $parsed.MessageId | Should -Be '<test123@example.com>'
        }

        It 'Should extract From' {
            $parsed.From | Should -Be 'sender@example.com'
        }

        It 'Should extract To' {
            $parsed.To | Should -Be 'recipient@target.com'
        }

        It 'Should extract Subject' {
            $parsed.Subject | Should -Be 'Test Message'
        }

        It 'Should extract SPF result' {
            $parsed.SPF | Should -Be 'pass'
        }

        It 'Should extract DKIM result' {
            $parsed.DKIM | Should -Be 'pass'
        }

        It 'Should extract DMARC result' {
            $parsed.DMARC | Should -Be 'pass'
        }

        It 'Should parse Received hops' {
            $parsed.TotalHops | Should -BeGreaterOrEqual 2
        }

        It 'Should extract X-Headers' {
            $parsed.XHeaders.Keys | Should -Contain 'X-MS-Exchange-Organization-SCL'
        }
    }

    Context 'Empty headers' {
        It 'Should handle empty input gracefully' {
            $result = Parse-EmailHeaders -RawHeaders ''
            $result.TotalHops | Should -Be 0
            $result.MessageId | Should -BeNullOrEmpty
        }
    }
}

Describe 'Search-TransportLogs' {

    Context 'Non-existent path' {
        It 'Should produce error for missing path' {
            { Search-TransportLogs -LogPath '/nonexistent/path' -Pattern 'test' -ErrorAction Stop } |
                Should -Throw
        }
    }

    Context 'Search in temp directory with test log' {
        BeforeAll {
            $script:tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "exchange-trace-test-$(Get-Random)"
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

            @"
2026-03-12T10:00:00 SMTP SEND user@example.com -> recipient@target.com OK
2026-03-12T10:01:00 SMTP SEND other@example.com -> someone@target.com OK
2026-03-12T10:02:00 SMTP FAIL user@example.com -> bad@target.com 550 User unknown
"@ | Set-Content -Path (Join-Path $tempDir 'test.log')
        }

        AfterAll {
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'Should find matching lines' {
            $results = Search-TransportLogs -LogPath $tempDir -Pattern 'user@example.com'
            $results.Count | Should -Be 2
        }

        It 'Should return correct properties' {
            $results = Search-TransportLogs -LogPath $tempDir -Pattern 'FAIL'
            $results[0].File | Should -Be 'test.log'
            $results[0].Line | Should -BeGreaterThan 0
            $results[0].Match | Should -Match 'FAIL'
        }

        It 'Should return empty for non-matching pattern' {
            $results = Search-TransportLogs -LogPath $tempDir -Pattern 'nonexistent_pattern_xyz'
            $results.Count | Should -Be 0
        }
    }
}

Describe 'Export-Results' {

    Context 'CSV export' {
        BeforeAll {
            $script:tempFile = Join-Path ([System.IO.Path]::GetTempPath()) "export-test-$(Get-Random).csv"
        }

        AfterAll {
            Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
        }

        It 'Should create CSV file' {
            $data = @(
                [PSCustomObject]@{ Name = 'Test1'; Value = 'A' }
                [PSCustomObject]@{ Name = 'Test2'; Value = 'B' }
            )
            Export-Results -Data $data -FilePath $tempFile -Format 'CSV'
            Test-Path $tempFile | Should -BeTrue
        }
    }

    Context 'JSON export' {
        BeforeAll {
            $script:tempFile = Join-Path ([System.IO.Path]::GetTempPath()) "export-test-$(Get-Random).json"
        }

        AfterAll {
            Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
        }

        It 'Should create valid JSON file' {
            $data = @(
                [PSCustomObject]@{ Name = 'Test1'; Value = 'A' }
            )
            Export-Results -Data $data -FilePath $tempFile -Format 'JSON'
            $content = Get-Content -Path $tempFile -Raw
            { $content | ConvertFrom-Json } | Should -Not -Throw
        }
    }
}
