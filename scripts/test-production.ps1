#Requires -Version 7.0
<#
.SYNOPSIS
    Post-deploy smoke test: checks health/readiness/version and OAuth
    discovery endpoints against a live deployment, without touching any
    real Microsoft account data.

.PARAMETER BaseUrl
    e.g. https://todo.h-aa.dk
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$BaseUrl
)

$ErrorActionPreference = 'Stop'
$BaseUrl = $BaseUrl.TrimEnd('/')
$failures = New-Object System.Collections.Generic.List[string]

function Test-Endpoint {
    param(
        [string]$Path,
        [int]$ExpectedStatus = 200,
        [scriptblock]$Validate
    )
    $uri = "$BaseUrl$Path"
    try {
        $response = Invoke-WebRequest -Uri $uri -UseBasicParsing -TimeoutSec 10 -SkipHttpErrorCheck
        if ($response.StatusCode -ne $ExpectedStatus) {
            $failures.Add("$Path -> expected HTTP $ExpectedStatus, got $($response.StatusCode)")
            return
        }
        if ($Validate) {
            $json = $response.Content | ConvertFrom-Json
            & $Validate $json
        }
        Write-Host "[ok] $Path" -ForegroundColor Green
    } catch {
        $failures.Add("$Path -> request failed: $($_.Exception.Message)")
    }
}

Write-Host "== Smoke-testing $BaseUrl ==" -ForegroundColor Cyan

Test-Endpoint -Path '/health' -Validate { param($json) if ($json.status -ne 'ok') { $failures.Add('/health did not report status=ok') } }
Test-Endpoint -Path '/ready'
Test-Endpoint -Path '/version' -Validate {
    param($json)
    if (-not $json.version) { $failures.Add('/version missing "version" field') }
    Write-Host "    version: $($json.version), node: $($json.nodeVersion)" -ForegroundColor DarkGray
}
Test-Endpoint -Path '/.well-known/oauth-protected-resource' -Validate {
    param($json)
    if (-not $json.authorization_servers) { $failures.Add('Protected Resource Metadata missing authorization_servers') }
}
Test-Endpoint -Path '/.well-known/oauth-authorization-server' -Validate {
    param($json)
    foreach ($field in @('authorization_endpoint', 'token_endpoint', 'jwks_uri')) {
        if (-not $json.$field) { $failures.Add("Authorization Server Metadata missing $field") }
    }
}
Test-Endpoint -Path '/.well-known/jwks.json' -Validate {
    param($json)
    if (-not $json.keys -or $json.keys.Count -eq 0) { $failures.Add('JWKS has no keys') }
}

# /mcp should require auth (401), never silently succeed unauthenticated.
Test-Endpoint -Path '/mcp' -ExpectedStatus 401

# /health must not leak secrets/tenant/account data — a cheap content check.
$healthRaw = (Invoke-WebRequest -Uri "$BaseUrl/health" -UseBasicParsing).Content
foreach ($sensitivePattern in @('tenant', 'secret', 'token', 'connectionstring', 'password')) {
    if ($healthRaw -match $sensitivePattern) {
        $failures.Add("/health response contains suspicious term '$sensitivePattern' — review src/http/health.ts")
    }
}

if ($failures.Count -gt 0) {
    Write-Host "`n== Smoke test FAILED ==" -ForegroundColor Red
    foreach ($f in $failures) { Write-Host " - $f" -ForegroundColor Red }
    exit 1
}

Write-Host "`n== Smoke test passed ==" -ForegroundColor Green
