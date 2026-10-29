#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Seed-GraphTestData.ps1" -DelayMilliseconds 0 -Force

$originalSave = ${function:Save-SeedManifest}
$scratch = Join-Path ([IO.Path]::GetTempPath()) "intranet-seed-tests-$([guid]::NewGuid().ToString('N'))"
$null = [IO.Directory]::CreateDirectory($scratch)
$requests = [System.Collections.Generic.List[object]]::new()
$sleeps = [System.Collections.Generic.List[double]]::new()
$checks = 0
$transport = $null
$connected = 0
$retry = 4

function Assert {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:checks++
}

function Assert-Throws {
    param([scriptblock] $Action, [string] $Pattern)
    $caught = $null
    try { & $Action | Out-Null }
    catch { $caught = $_ }
    Assert ($null -ne $caught -and $caught.ToString() -like "*$Pattern*") "Expected error matching '$Pattern'; got '$caught'"
}

function New-TestResponse {
    param([int] $Status = 200, [object] $Data, [string] $RetryAfter)
    $response = [System.Net.Http.HttpResponseMessage]::new($Status)
    $text = if ($null -eq $Data) { '' } else { ConvertTo-Json -InputObject $Data -Depth 20 -Compress }
    $response.Content = [System.Net.Http.StringContent]::new($text)
    if ($RetryAfter) { $null = $response.Headers.TryAddWithoutValidation('Retry-After', $RetryAfter) }
    return $response
}

# All authentication and HTTP commands are local fakes, including module loading.
function Import-Module { param($Name) }
function Get-Module { param([switch] $ListAvailable, $Name) return @{ Name = 'Microsoft.Graph.Authentication' } }
function Connect-MgGraph {
    param($Scopes, $ContextScope, [switch] $NoWelcome, $TenantId, $ClientId, [switch] $UseDeviceCode)
    $script:connected++
    $script:requestedScopes = $Scopes
}
function Get-MgContext { return @{ AuthType = 'Delegated'; Environment = 'Global'; TenantId = 'tenant-a' } }
function Get-MgRequestContext { return @{ MaxRetry = $script:retry } }
function Set-MgRequestContext { param([int] $MaxRetry) $script:retry = $MaxRetry }
function Start-Sleep { param([double] $Seconds, [double] $Milliseconds) $script:sleeps.Add($Seconds + $Milliseconds / 1000) }
function Write-Progress { param($Activity, $Status, $PercentComplete, [switch] $Completed) }
function Read-Host { throw 'Unexpected prompt in automated test.' }
function Invoke-MgGraphRequest {
    [CmdletBinding()]
    param($Method, [uri] $Uri, $Body, $Headers, $OutputType, [switch] $SkipHttpErrorCheck, $ContentType)
    Assert ($Uri.Host -eq 'graph.microsoft.com') 'Only the public Graph host is used'
    Assert ($OutputType -eq 'HttpResponseMessage' -and $SkipHttpErrorCheck) 'Transport exposes HTTP status for controlled retries'
    $request = @{
        Method = [string] $Method
        Path = $Uri.PathAndQuery.Substring('/v1.0'.Length)
        Body = if ($Body) { ConvertFrom-Json -InputObject $Body -AsHashtable } else { $null }
        Headers = $Headers
    }
    $script:requests.Add($request)
    if (-not $script:transport) { throw 'No fake Graph transport installed.' }
    return & $script:transport $request
}

try {
    # Dry-run is dependency-free, with no files, prompts or authentication.
    $DryRun = $true
    $null = Invoke-GraphSeed
    Assert ($connected -eq 0 -and $requests.Count -eq 0) 'Dry run never signs in or sends requests'
    Assert (@(Get-ChildItem -LiteralPath $scratch).Count -eq 0) 'Dry run creates no manifest'
    $DryRun = $false
    Assert ($StartDate.Offset -eq [timespan]::Zero -and $StartDate.Hour -eq 8) 'Default start is 08:00 UTC, not local time'

    $start = [datetimeoffset] '2026-11-02T10:00:00+02:00'
    $titles = [System.Collections.Generic.HashSet[string]]::new()
    $transactions = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($index in 1..1101) {
        $planner = Get-SeedPayload Planner $index 'run-a' 'user-a' 'plan-a' 'bucket-a' $start 7
        $todo = Get-SeedPayload Todo $index 'run-a' 'user-a' '' '' $start 7
        $event = Get-SeedPayload Calendar $index 'run-a' 'user-a' '' '' $start 1
        Assert ($planner.assignments['user-a']['@odata.type'] -eq '#microsoft.graph.plannerAssignment') 'Planner assigns the signed-in user'
        Assert ($planner.planId -eq 'plan-a' -and $planner.bucketId -eq 'bucket-a') 'Planner destination is preserved'
        Assert ($planner.percentComplete -lt 100 -and $todo.status -ne 'completed') 'All seeded tasks remain visible as open tasks'
        Assert ($todo.body.content -like '*Lorem ipsum*' -and $todo.dueDateTime.timeZone -eq 'UTC') 'To Do has Lorem ipsum content and UTC due date'
        Assert ($event.attendees.Count -eq 0 -and -not $event.isReminderOn -and $event.showAs -eq 'free' -and
            -not $event.isOnlineMeeting -and -not $event.responseRequested) 'Calendar creates do not notify invitees or block time'
        Assert ($event.start.dateTime.StartsWith('2026-11-02T')) 'Days=1 keeps 1,101 calendar events on the requested date'
        Assert ($titles.Add($planner.title) -and $transactions.Add($event.transactionId)) 'Titles and event transaction IDs are unique above 1,000 items'
        Assert (([datetime] $event.end.dateTime - [datetime] $event.start.dateTime).TotalMinutes -eq 30) 'Meetings last 30 minutes'
    }

    # HTTP retry policy: 429 honors seconds/date headers; ambiguous creates never retry.
    $script:transport = {
        param($request)
        if ($script:requests.Count -eq 1) { return New-TestResponse 429 @{ error = 'throttled' } '2' }
        return New-TestResponse 201 @{ id = 'accepted' }
    }
    $requests.Clear(); $sleeps.Clear()
    $accepted = Invoke-SeedGraph POST '/me/events' @{ subject = 'test' }
    Assert ($accepted.id -eq 'accepted' -and $requests.Count -eq 2 -and $sleeps[0] -eq 2) '429 POST retries after Retry-After seconds'
    Assert ($requests[0].Headers.Prefer -eq 'IdType="ImmutableId"') 'Calendar create requests immutable IDs'

    $script:transport = {
        param($request)
        if ($script:requests.Count -eq 1) {
            return New-TestResponse 429 @{} ([datetimeoffset]::UtcNow.AddSeconds(30).ToString('r'))
        }
        return New-TestResponse 200 @{ id = 'accepted' }
    }
    $requests.Clear(); $sleeps.Clear()
    $null = Invoke-SeedGraph GET '/me'
    Assert ($sleeps[0] -ge 28 -and $sleeps[0] -le 30) 'Retry-After HTTP dates are respected'

    $script:transport = { param($request) New-TestResponse 503 @{ error = 'unavailable' } }
    $requests.Clear()
    Assert-Throws { Invoke-SeedGraph POST '/planner/tasks' @{ title = 'test' } } 'HTTP 503'
    Assert ($requests.Count -eq 1) 'Ambiguous POST is not retried'
    $script:transport = { param($request) New-TestResponse 429 @{ error = 'still throttled' } }
    $requests.Clear(); $sleeps.Clear()
    Assert-Throws { Invoke-SeedGraph POST '/me/events' @{ subject = 'test' } } 'HTTP 429'
    Assert ($requests.Count -eq 7 -and $sleeps[0] -eq 2 -and $sleeps[-1] -eq 64) '429 without Retry-After uses bounded exponential backoff'
    $script:transport = { param($request) New-TestResponse 503 @{ error = 'unavailable' } }
    $requests.Clear(); $sleeps.Clear()
    Assert-Throws { Invoke-SeedGraph GET '/me' } 'HTTP 503'
    Assert ($requests.Count -eq 7 -and $sleeps.Count -eq 6) 'GET transient retries are bounded'

    $script:transport = { param($request) New-TestResponse 403 @{ error = 'denied' } }
    $requests.Clear()
    Assert-Throws { Invoke-SeedGraph GET '/me' } 'HTTP 403'
    Assert ($requests.Count -eq 1) 'Permission errors fail explicitly without retries'

    $script:transport = {
        param($request)
        if ($request.Path -eq '/me/planner/tasks') {
            return New-TestResponse 200 @{ value = @(@{ id = 'first' }); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/me/planner/tasks?cursor=2' }
        }
        return New-TestResponse 200 @{ value = @(@{ id = 'second' }) }
    }
    $pages = @(Get-SeedCollection '/me/planner/tasks')
    Assert ($pages.Count -eq 2 -and $pages[1].id -eq 'second') 'Preflight follows pagination'
    $script:transport = { param($request) New-TestResponse 200 @{ value = @(); '@odata.nextLink' = 'https://example.invalid/me' } }
    Assert-Throws { Get-SeedCollection '/me/planner/tasks' } 'unexpected pagination'
    $script:transport = { param($request) New-TestResponse 200 @{ value = @(); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/me/planner/tasks' } }
    Assert-Throws { Get-SeedCollection '/me/planner/tasks' } 'repeated page'

    # Exercise the entire large-run loop, recording every checkpoint in memory.
    # Atomic disk persistence is tested separately below to avoid gigabytes of test writes.
    $checkpoints = [System.Collections.Generic.List[int]]::new()
    function Save-SeedManifest {
        param([hashtable] $Manifest, [string] $Path, [switch] $Create)
        $script:checkpoints.Add($Manifest.items.Count)
        $script:lastManifest = $Manifest
    }
    $Count = 1101; $Types = @('Planner', 'Todo', 'Calendar')
    $PlanId = 'plan-a'; $BucketId = 'bucket-a'
    $ManifestPath = Join-Path $scratch 'large.json'
    $script:transport = {
        param($request)
        switch -Wildcard ($request.Path) {
            '/me' { return New-TestResponse 200 @{ id = 'user-a'; userPrincipalName = 'test@example.invalid' } }
            '/planner/plans/plan-a' { return New-TestResponse 200 @{ id = 'plan-a'; title = 'Test plan' } }
            '/planner/plans/plan-a/buckets' { return New-TestResponse 200 @{ value = @(@{ id = 'bucket-a'; name = 'Test bucket' }) } }
            '/planner/plans/plan-a/tasks' { return New-TestResponse 200 @{ value = @() } }
            '/me/planner/tasks' { return New-TestResponse 200 @{ value = @() } }
            '*/details' {
                if ($request.Method -eq 'PATCH') {
                    Assert ($request.Headers['If-Match'] -eq 'fresh-etag' -and $request.Body.description -like '*Lorem ipsum*') 'Planner descriptions use ETag and Lorem ipsum'
                    return New-TestResponse 204
                }
                Assert ($script:lastManifest.items[-1].kind -eq 'Planner' -and
                    $script:checkpoints[-1] -eq $script:lastManifest.items.Count) 'Planner ID persisted before details request'
                return New-TestResponse 200 @{ '@odata.etag' = 'fresh-etag' }
            }
            default {
                if ($request.Method -eq 'POST') { return New-TestResponse 201 @{ id = "resource-$($script:requests.Count)" } }
                throw "Unexpected request: $($request.Method) $($request.Path)"
            }
        }
    }
    $requests.Clear()
    Invoke-GraphSeed
    Assert ($lastManifest.items.Count -eq 3304) 'Full run records 3,303 tasks/events plus one To Do list'
    foreach ($kind in 'Planner', 'Todo', 'Calendar') {
        Assert (@($lastManifest.items | Where-Object kind -eq $kind).Count -eq 1101) "Full run creates exactly 1,101 $kind items"
    }
    Assert ($checkpoints.Count -eq 3305 -and $checkpoints[0] -eq 0 -and $checkpoints[-1] -eq 3304) 'Every create is checkpointed immediately'
    Assert ($retry -eq 4) 'SDK retry context is restored after success'
    Assert ('Tasks.ReadWrite' -in $requestedScopes -and 'Calendars.ReadWrite' -in $requestedScopes) 'Both task and calendar write scopes requested'
    ${function:Save-SeedManifest} = $originalSave
    Save-SeedManifest $lastManifest $ManifestPath -Create
    $roundTrip = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json -AsHashtable
    Assert ($roundTrip.items.Count -eq 3304) 'Large manifest survives JSON persistence'
    Assert-Throws { Save-SeedManifest $lastManifest $ManifestPath -Create } '*'
    Assert (@(Get-ChildItem -LiteralPath $scratch -Filter '*.tmp').Count -eq 0) 'Atomic save cleans temporary files on failure'

    # Small real-manifest cleanup: reverse order, ETags, escaped IDs, 404, and reruns.
    $ManifestPath = Join-Path $scratch 'cleanup.json'
    $manifest = @{
        schemaVersion = 1; runId = 'run-cleanup'; userId = 'user-a'; tenantId = 'tenant-a'
        types = @('Planner', 'Todo', 'Calendar')
        items = @(
            @{ kind = 'TodoList'; id = 'list-a'; deleted = $false },
            @{ kind = 'Todo'; id = 'task-a'; parentId = 'list-a'; deleted = $false },
            @{ kind = 'Planner'; id = 'planner-a'; deleted = $false },
            @{ kind = 'Calendar'; id = 'event+a='; deleted = $false },
            @{ kind = 'Calendar'; id = 'already-gone'; deleted = $false }
        )
    }
    Save-SeedManifest $manifest $ManifestPath -Create
    $script:transport = {
        param($request)
        if ($request.Path -eq '/me') { return New-TestResponse 200 @{ id = 'user-a' } }
        if ($request.Path -like '*/already-gone') { return New-TestResponse 404 }
        if ($request.Method -eq 'DELETE') {
            if ($request.Path -like '/planner/*') { Assert ($request.Headers['If-Match'] -eq 'fresh-etag') 'Planner deletion uses fresh ETag' }
            return New-TestResponse 204
        }
        if ($request.Path -like '*/tasks' -and $request.Path -like '/me/todo/*') { return New-TestResponse 200 @{ value = @() } }
        return New-TestResponse 200 @{ id = 'found'; '@odata.etag' = 'fresh-etag' }
    }
    $Cleanup = $true
    $requests.Clear()
    Invoke-GraphSeed
    $deletes = @($requests | Where-Object Method -eq 'DELETE')
    Assert ($deletes.Count -eq 4 -and $deletes[-1].Path -eq '/me/todo/lists/list-a') 'Cleanup removes tasks before their dedicated list'
    Assert ($deletes[0].Path -eq '/me/events/event%2Ba%3D' -and $deletes[0].Headers.Prefer -eq 'IdType="ImmutableId"') 'Cleanup escapes and uses immutable calendar IDs'
    $saved = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json -AsHashtable
    Assert (@($saved.items | Where-Object { -not $_.deleted }).Count -eq 0) 'Cleanup checkpoints every deletion including already absent items'
    $requests.Clear()
    Invoke-GraphSeed
    Assert (@($requests | Where-Object Method -eq 'DELETE').Count -eq 0) 'Cleanup is rerunnable without repeat deletes'

    $saved.userId = 'different-user'
    Save-SeedManifest $saved $ManifestPath
    $requests.Clear()
    Assert-Throws { Invoke-GraphSeed } 'different user or tenant'
    Assert (@($requests | Where-Object Method -eq 'DELETE').Count -eq 0) 'Wrong-user cleanup never deletes'
    $saved.userId = 'user-a'; $saved.tenantId = 'different-tenant'
    Save-SeedManifest $saved $ManifestPath
    Assert-Throws { Invoke-GraphSeed } 'different user or tenant'

    $saved.tenantId = 'tenant-a'
    $saved.items = @(@{ kind = 'TodoList'; id = 'list-a'; deleted = $false })
    Save-SeedManifest $saved $ManifestPath
    $script:transport = {
        param($request)
        if ($request.Path -eq '/me') { return New-TestResponse 200 @{ id = 'user-a' } }
        if ($request.Path -like '*/tasks') { return New-TestResponse 200 @{ value = @(@{ id = 'unrecorded-task' }) } }
        return New-TestResponse 200 @{ id = 'list-a' }
    }
    $requests.Clear()
    Assert-Throws { Invoke-GraphSeed } 'Refusing to delete unrecorded tasks'
    Assert (@($requests | Where-Object Method -eq 'DELETE').Count -eq 0) 'Cleanup preserves user-added tasks'

    $manifest.items = @(@{ kind = 'Arbitrary'; id = 'unsafe'; deleted = $false })
    Assert-Throws { Remove-SeedResources $manifest } 'Unsupported manifest kind'

    # A details failure must leave the preceding create recorded for cleanup.
    $Cleanup = $false; $Types = @('Planner'); $Count = 1
    $ManifestPath = Join-Path $scratch 'partial.json'
    $script:transport = {
        param($request)
        switch ($request.Path) {
            '/me' { return New-TestResponse 200 @{ id = 'user-a' } }
            '/planner/plans/plan-a' { return New-TestResponse 200 @{ title = 'Test plan' } }
            '/planner/plans/plan-a/buckets' { return New-TestResponse 200 @{ value = @(@{ id = 'bucket-a' }) } }
            '/planner/plans/plan-a/tasks' { return New-TestResponse 200 @{ value = @() } }
            '/me/planner/tasks' { return New-TestResponse 200 @{ value = @() } }
            '/planner/tasks' { return New-TestResponse 201 @{ id = 'partial-task' } }
            default { return New-TestResponse 403 @{ error = 'details denied' } }
        }
    }
    Assert-Throws { Invoke-GraphSeed } 'HTTP 403'
    $partial = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json -AsHashtable
    Assert ($partial.items.Count -eq 1 -and $partial.items[0].id -eq 'partial-task') 'Partial run retains created ID before a later failure'
    Assert ($retry -eq 4) 'SDK retry context is restored after failure'
    $requests.Clear()
    Assert-Throws { Invoke-GraphSeed } 'Manifest already exists'
    Assert ($requests.Count -eq 0) 'Existing manifest blocks reseeding before authentication or writes'

    $ManifestPath = Join-Path $scratch 'limit.json'
    $Count = 2
    $script:transport = {
        param($request)
        switch ($request.Path) {
            '/me' { return New-TestResponse 200 @{ id = 'user-a' } }
            '/planner/plans/plan-a' { return New-TestResponse 200 @{ title = 'Nearly full test plan' } }
            '/planner/plans/plan-a/buckets' { return New-TestResponse 200 @{ value = @(@{ id = 'bucket-a' }) } }
            '/planner/plans/plan-a/tasks' { return New-TestResponse 200 @{ value = @(1..2999 | ForEach-Object { @{ percentComplete = 0 } }) } }
            default { throw "Unexpected request in capacity preflight: $($request.Path)" }
        }
    }
    $requests.Clear()
    Assert-Throws { Invoke-GraphSeed } 'exceed Planner plan limits'
    Assert (@($requests | Where-Object Method -eq 'POST').Count -eq 0 -and
        -not (Test-Path -LiteralPath $ManifestPath)) 'Planner capacity preflight stops before creating data or a manifest'

    $Count = 3001
    Assert-Throws { Invoke-GraphSeed } '3,000 active tasks'
    Write-Host "PASS: $checks assertions; no network calls or tenant writes."
}
finally {
    # The only recursively deleted directory is this test's resolved, unique scratch directory.
    if ([IO.Directory]::Exists($scratch)) { [IO.Directory]::Delete($scratch, $true) }
}
