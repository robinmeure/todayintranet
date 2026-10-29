#Requires -Version 7.0
<#
.SYNOPSIS
Creates current-user Microsoft 365 load-test data, or removes a recorded run.
.EXAMPLE
.\scripts\Seed-GraphTestData.ps1 -DryRun
.EXAMPLE
.\scripts\Seed-GraphTestData.ps1 -Count 1100 -PlanId '<test-plan-id>'
.EXAMPLE
.\scripts\Seed-GraphTestData.ps1 -Cleanup -ManifestPath '.\temp\graph-seed\<run>.json'
#>
[CmdletBinding()]
param(
    [ValidateSet('Planner', 'Todo', 'Calendar')]
    [ValidateNotNullOrEmpty()]
    [string[]] $Types = @('Planner', 'Todo', 'Calendar'),
    [ValidateRange(1, 10000)]
    [int] $Count = 1100,
    [string] $PlanId,
    [string] $BucketId,
    [datetimeoffset] $StartDate = ([datetime]::UtcNow.Date.AddDays(1).AddHours(8)),
    [ValidateRange(1, 365)]
    [int] $Days = 7,
    [ValidateRange(0, 60000)]
    [int] $DelayMilliseconds = 200,
    [string] $TenantId,
    [string] $ClientId,
    [switch] $UseDeviceCode,
    [string] $ManifestPath,
    [switch] $DryRun,
    [switch] $Cleanup,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

function Invoke-SeedGraph {
    param(
        [string] $Method, [string] $Path, [hashtable] $Body,
        [hashtable] $Headers = @{}, [switch] $AllowNotFound
    )
    if ($Path -notmatch '^/(me|planner)(/|$)') {
        throw "Unexpected Graph path: $Path"
    }
    if ($Path -like '/me/events*') { $Headers['Prefer'] = 'IdType="ImmutableId"' }
    $request = @{
        Method = $Method
        Uri = "https://graph.microsoft.com/v1.0$Path"
        Headers = $Headers
        OutputType = 'HttpResponseMessage'
        SkipHttpErrorCheck = $true
        ErrorAction = 'Stop'
    }
    if ($null -ne $Body) {
        $request.Body = ConvertTo-Json -InputObject $Body -Depth 20 -Compress
        $request.ContentType = 'application/json'
    }
    for ($attempt = 0; $attempt -le 6; $attempt++) {
        if ($DelayMilliseconds) { Start-Sleep -Milliseconds $DelayMilliseconds }
        # A failed POST may already have committed. Retry only explicit throttle rejections.
        $response = Invoke-MgGraphRequest @request
        try {
            $status = [int] $response.StatusCode
            $text = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            if ($response.IsSuccessStatusCode) {
                if ($text) { return ConvertFrom-Json -InputObject $text -AsHashtable }
                return
            }
            if ($status -eq 404 -and $AllowNotFound) { return }
            $retryable = $status -eq 429 -or ($Method -eq 'GET' -and $status -in 502, 503, 504)
            if ($retryable -and $attempt -lt 6) {
                $wait = [math]::Pow(2, $attempt + 1)
                $retryAfter = $response.Headers.RetryAfter
                if ($null -ne $retryAfter.Delta) { $wait = $retryAfter.Delta.TotalSeconds }
                elseif ($null -ne $retryAfter.Date) {
                    $wait = [math]::Max(0, ($retryAfter.Date - [datetimeoffset]::UtcNow).TotalSeconds)
                }
                Write-Warning "Graph HTTP $status; retrying $Method after $([math]::Ceiling($wait)) seconds."
                Start-Sleep -Seconds ([math]::Ceiling($wait))
                continue
            }
            throw "Graph $Method $Path failed (HTTP $status): $text"
        }
        finally { $response.Dispose() }
    }
}

function Get-SeedCollection {
    param([string] $Path)
    $visited = [System.Collections.Generic.HashSet[string]]::new()
    while ($Path) {
        if (-not $visited.Add($Path)) { throw 'Graph returned a repeated page.' }
        $page = Invoke-SeedGraph GET $Path
        if ($null -eq $page -or $page.value -isnot [array]) { throw "Invalid Graph collection: $Path" }
        foreach ($item in $page.value) { $item }
        $next = $page['@odata.nextLink']
        $Path = $null
        if ($next) {
            # Never follow a server-supplied link to another host or API version.
            if ($next -isnot [string] -or -not $next.StartsWith('https://graph.microsoft.com/v1.0/', [StringComparison]::Ordinal)) {
                throw 'Graph returned an unexpected pagination URL.'
            }
            $Path = $next.Substring('https://graph.microsoft.com/v1.0'.Length)
        }
    }
}

function Save-SeedManifest {
    param([hashtable] $Manifest, [string] $Path, [switch] $Create)
    $temporary = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        [IO.File]::WriteAllText($temporary, (ConvertTo-Json -InputObject $Manifest -Depth 20), [Text.UTF8Encoding]::new($false))
        [IO.File]::Move($temporary, $Path, -not $Create)
    }
    finally {
        if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
    }
}

function Get-SeedPayload {
    param(
        [string] $Type, [int] $Index, [string] $RunId, [string] $UserId,
        [string] $TargetPlan, [string] $TargetBucket, [datetimeoffset] $Start, [int] $Span
    )
    $topics = @('Operational resilience review', 'Quarterly delivery planning', 'Customer onboarding',
        'Risk and compliance assessment', 'Platform migration', 'Service performance review',
        'Procurement approval', 'Data quality workshop', 'Security awareness', 'Budget forecast',
        'Stakeholder alignment', 'Business continuity exercise')
    $topic = $topics[($Index - 1) % $topics.Length]
    $title = "[IntranetSeed:$RunId] $topic $($Index.ToString('D5')) - Lorem ipsum dolor sit amet"
    $description = "$title`n`nLorem ipsum dolor sit amet, consectetur adipiscing elit. " +
        'Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ' +
        'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.' +
        "`n`nSynthetic enterprise test data. Review scope, confirm ownership, record dependencies, and agree next steps."
    $day = $Start.ToUniversalTime().AddDays(($Index - 1) % $Span)
    $due = $day.AddDays(-3).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
    switch ($Type) {
        'Planner' {
            return @{
                planId = $TargetPlan; bucketId = $TargetBucket; title = $title
                assignments = @{ $UserId = @{ '@odata.type' = '#microsoft.graph.plannerAssignment'; orderHint = ' !' } }
                dueDateTime = $due
                percentComplete = @(0, 50)[($Index - 1) % 2]
                priority = @(1, 3, 5, 9)[($Index - 1) % 4]
            }
        }
        'Todo' {
            return @{
                title = $title; body = @{ contentType = 'text'; content = $description }
                status = @('notStarted', 'inProgress')[($Index - 1) % 2]
                importance = @('normal', 'high', 'low')[($Index - 1) % 3]
                dueDateTime = @{ dateTime = $due.TrimEnd('Z'); timeZone = 'UTC' }
                isReminderOn = $false
            }
        }
        'Calendar' {
            $time = $day.AddMinutes(([math]::Floor(($Index - 1) / $Span) % 16) * 30)
            return @{
                subject = $title; body = @{ contentType = 'text'; content = $description }
                start = @{ dateTime = $time.ToString("yyyy-MM-dd'T'HH:mm:ss"); timeZone = 'UTC' }
                end = @{ dateTime = $time.AddMinutes(30).ToString("yyyy-MM-dd'T'HH:mm:ss"); timeZone = 'UTC' }
                location = @{ displayName = @('Amsterdam HQ', 'Project room', 'Innovation hub', 'Remote workspace')[($Index - 1) % 4] }
                attendees = @(); isReminderOn = $false; showAs = 'free'
                responseRequested = $false; isOnlineMeeting = $false
                transactionId = "$RunId-$Index"
            }
        }
        default { throw "Unknown seed type: $Type" }
    }
}

function Add-SeedResource {
    param([hashtable] $Manifest, [string] $Path, [string] $Kind, [hashtable] $Body, [string] $ParentId)
    $created = Invoke-SeedGraph POST $Path $Body
    if (-not $created.id) {
        throw "Graph did not return an ID for $Kind. Inspect resources tagged $($Manifest.runId) before running again."
    }
    $entry = @{ kind = $Kind; id = $created.id; parentId = $ParentId; deleted = $false }
    $Manifest.items.Add($entry)
    # Persist immediately, before descriptions or subsequent creates can fail.
    Save-SeedManifest $Manifest $ManifestPath
    return $entry
}

function Get-SeedResourcePath {
    param([hashtable] $Entry)
    if ([string]::IsNullOrWhiteSpace($Entry.id)) { throw 'Manifest contains an empty resource ID.' }
    $id = [uri]::EscapeDataString($Entry.id)
    switch ($Entry.kind) {
        'Planner' { return "/planner/tasks/$id" }
        'Calendar' { return "/me/events/$id" }
        'TodoList' { return "/me/todo/lists/$id" }
        'Todo' {
            if ([string]::IsNullOrWhiteSpace($Entry.parentId)) { throw 'Manifest contains an empty To Do list ID.' }
            return "/me/todo/lists/$([uri]::EscapeDataString($Entry.parentId))/tasks/$id"
        }
        default { throw "Unsupported manifest kind: $($Entry.kind)" }
    }
}

function Remove-SeedResources {
    param([hashtable] $Manifest)
    # Validate all entries before the first deletion; never execute arbitrary URLs from a manifest.
    foreach ($entry in $Manifest.items) { $null = Get-SeedResourcePath $entry }
    for ($index = $Manifest.items.Count - 1; $index -ge 0; $index--) {
        $entry = $Manifest.items[$index]
        if ($entry.deleted) { continue }
        $path = Get-SeedResourcePath $entry
        $headers = @{}
        $resource = Invoke-SeedGraph GET $path -AllowNotFound
        if ($null -ne $resource) {
            if ($entry.kind -eq 'Planner') {
                if (-not $resource['@odata.etag']) { throw "Planner task has no ETag: $path" }
                $headers['If-Match'] = $resource['@odata.etag']
            }
            if ($entry.kind -eq 'TodoList') {
                $remaining = Invoke-SeedGraph GET "$path/tasks"
                if ($null -eq $remaining -or $remaining.value -isnot [array]) { throw 'Invalid To Do list response during cleanup.' }
                if ($remaining.value.Count -or $remaining['@odata.nextLink']) {
                    throw 'The seed To Do list is not empty. Refusing to delete unrecorded tasks; move them out and rerun cleanup.'
                }
            }
            $null = Invoke-SeedGraph DELETE $path -Headers $headers -AllowNotFound
        }
        $entry.deleted = $true
        Save-SeedManifest $Manifest $ManifestPath
        Write-Progress -Activity 'Removing Graph test data' -Status "$($Manifest.items.Count - $index) / $($Manifest.items.Count)" `
            -PercentComplete ((($Manifest.items.Count - $index) / $Manifest.items.Count) * 100)
    }
    Write-Progress -Activity 'Removing Graph test data' -Completed
    Write-Host "Cleanup complete. Manifest retained at $ManifestPath"
}

function Invoke-GraphSeed {
    $selected = @($Types | Select-Object -Unique)
    if ($Cleanup -and $DryRun) { throw 'Use -Cleanup or -DryRun, not both.' }
    if (-not $Cleanup -and 'Planner' -in $selected -and $Count -gt 3000) {
        throw 'Planner allows at most 3,000 active tasks per plan. Select fewer tasks or omit Planner.'
    }
    $runId = [guid]::NewGuid().ToString('N')
    if ($DryRun) {
        Write-Host "DRY RUN: $Count items per type ($($Count * $selected.Count) total); no sign-in, files, or Graph requests."
        foreach ($type in $selected) {
            $payload = Get-SeedPayload $type 1 $runId '<signed-in-user-id>' '<test-plan-id>' '<test-bucket-id>' $StartDate $Days
            Write-Host "$type sample:"
            ConvertTo-Json -InputObject $payload -Depth 10
        }
        return
    }
    $manifest = $null
    if ($Cleanup) {
        if (-not $ManifestPath) { throw '-Cleanup requires -ManifestPath.' }
        $script:ManifestPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ManifestPath)
        $manifest = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json -AsHashtable
        if ($manifest.schemaVersion -ne 1 -or -not $manifest.runId -or $manifest.items -isnot [array] -or
            -not $manifest.userId -or -not $manifest.tenantId) { throw 'Invalid seed manifest.' }
        $runId = $manifest.runId
        $selected = @($manifest.types)
        if (-not $selected.Count -or @($selected | Where-Object { $_ -notin 'Planner', 'Todo', 'Calendar' }).Count) {
            throw 'Invalid seed types in manifest.'
        }
    }
    else {
        if (-not $ManifestPath) { $script:ManifestPath = Join-Path $PSScriptRoot "..\temp\graph-seed\$runId.json" }
        $script:ManifestPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ManifestPath)
        if (Test-Path -LiteralPath $ManifestPath) { throw "Manifest already exists; choose a new path: $ManifestPath" }
    }
    if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
        throw 'Install the authentication module first: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
    Import-Module Microsoft.Graph.Authentication
    $scopes = @('User.Read')
    if ('Planner' -in $selected -or 'Todo' -in $selected) { $scopes += 'Tasks.ReadWrite' }
    if ('Calendar' -in $selected) { $scopes += 'Calendars.ReadWrite' }
    $connect = @{ Scopes = $scopes; ContextScope = 'Process'; NoWelcome = $true }
    if ($TenantId) { $connect.TenantId = $TenantId }
    if ($ClientId) { $connect.ClientId = $ClientId }
    if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
    Connect-MgGraph @connect
    $previousRetry = (Get-MgRequestContext).MaxRetry
    try {
        $null = Set-MgRequestContext -MaxRetry 0
        $context = Get-MgContext
        if ($context.AuthType -ne 'Delegated' -or $context.Environment -ne 'Global' -or -not $context.TenantId) {
            throw 'This script requires delegated sign-in to the Global Microsoft Graph environment.'
        }
        $me = Invoke-SeedGraph GET '/me'
        if (-not $me.id) { throw 'Graph did not return the signed-in user ID.' }
        Write-Host "Account: $($me.userPrincipalName) | User: $($me.id) | Tenant: $($context.TenantId)"
        if ($Cleanup) {
            if ($manifest.userId -ne $me.id -or $manifest.tenantId -ne $context.TenantId) {
                throw 'The manifest belongs to a different user or tenant. No items were deleted.'
            }
            Write-Host "Delete only resources recorded in $ManifestPath"
            if (-not $Force -and (Read-Host 'Type DELETE to confirm cleanup') -cne 'DELETE') { throw 'Cleanup cancelled.' }
            Remove-SeedResources $manifest
            return
        }
        if ('Planner' -in $selected) {
            if (-not $PlanId) { $script:PlanId = Read-Host 'Existing TEST Planner plan ID (visible to other plan members)' }
            if ([string]::IsNullOrWhiteSpace($PlanId)) { throw 'Planner requires an existing test plan ID.' }
            $planPath = "/planner/plans/$([uri]::EscapeDataString($PlanId))"
            $plan = Invoke-SeedGraph GET $planPath
            $buckets = @(Get-SeedCollection "$planPath/buckets")
            if (-not $buckets.Count) { throw 'Create a bucket in the test plan first.' }
            if (-not $BucketId) {
                $buckets | ForEach-Object { Write-Host "$($_.name): $($_.id)" }
                $script:BucketId = Read-Host 'Destination test bucket ID'
            }
            if ($BucketId -notin $buckets.id) { throw 'The selected bucket does not belong to this plan.' }
            $tasks = @(Get-SeedCollection "$planPath/tasks")
            $active = @($tasks | Where-Object { $_.percentComplete -lt 100 }).Count
            if ($active + $Count -gt 3000 -or $tasks.Count + $Count -gt 9000) {
                throw 'This seed would exceed Planner plan limits (3,000 active / 9,000 total tasks).'
            }
            $assigned = @(Get-SeedCollection '/me/planner/tasks')
            if ($assigned.Count + $Count -gt 9000) { throw 'This seed would exceed 9,000 assigned Planner tasks.' }
            Write-Host "Planner: $($plan.title) | Bucket: $BucketId | All new tasks assigned to $($me.id)"
        }
        Write-Host "Create $Count items per type: $($selected -join ', '). Calendar: $($StartDate.ToUniversalTime()) UTC, spread over $Days days."
        Write-Host 'Use a TEST account. Calendar events overlap, are free, and have no attendees or reminders.'
        if (-not $Force -and (Read-Host 'Type SEED to confirm creation') -cne 'SEED') { throw 'Seeding cancelled.' }
        $manifest = @{
            schemaVersion = 1; runId = $runId; userId = $me.id; tenantId = $context.TenantId
            types = $selected; countPerType = $Count; createdUtc = [datetimeoffset]::UtcNow.ToString('o')
            items = [System.Collections.Generic.List[object]]::new()
        }
        $null = [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($ManifestPath))
        Save-SeedManifest $manifest $ManifestPath -Create
        Write-Host "Cleanup manifest: $ManifestPath | Run tag: [IntranetSeed:$runId]"
        $listId = $null
        foreach ($type in $selected) {
            if ($type -eq 'Todo') {
                $list = Add-SeedResource $manifest '/me/todo/lists' 'TodoList' @{ displayName = "[IntranetSeed:$runId] Enterprise test tasks" }
                $listId = $list.id
            }
            for ($index = 1; $index -le $Count; $index++) {
                $body = Get-SeedPayload $type $index $runId $me.id $PlanId $BucketId $StartDate $Days
                $path = switch ($type) {
                    'Planner' { '/planner/tasks' }
                    'Todo' { "/me/todo/lists/$([uri]::EscapeDataString($listId))/tasks" }
                    'Calendar' { '/me/events' }
                }
                $entry = Add-SeedResource $manifest $path $type $body $listId
                if ($type -eq 'Planner') {
                    $detailsPath = "$(Get-SeedResourcePath $entry)/details"
                    $details = Invoke-SeedGraph GET $detailsPath
                    if (-not $details['@odata.etag']) { throw 'Planner details did not return an ETag.' }
                    $description = (Get-SeedPayload Todo $index $runId $me.id $PlanId $BucketId $StartDate $Days).body.content
                    $null = Invoke-SeedGraph PATCH $detailsPath @{ description = $description } -Headers @{ 'If-Match' = $details['@odata.etag'] }
                }
                Write-Progress -Activity "Seeding $type" -Status "$index / $Count" -PercentComplete ($index / $Count * 100)
                if ($index % 100 -eq 0 -or $index -eq $Count) { Write-Host "${type}: $index / $Count recorded" }
            }
            Write-Progress -Activity "Seeding $type" -Completed
        }
        Write-Host "Seeding complete: $($Count * $selected.Count) items. Cleanup manifest: $ManifestPath"
    }
    catch {
        if ($Cleanup) {
            Write-Warning "Cleanup stopped. Progress is retained at $ManifestPath; resolve the error and rerun cleanup with the same account."
        }
        else {
            Write-Warning "Stopped. Keep the manifest at $ManifestPath for cleanup. A timed-out create may have committed without a recorded ID; inspect run tag [IntranetSeed:$runId] before retrying. Do not blindly rerun seeding."
        }
        throw
    }
    finally { $null = Set-MgRequestContext -MaxRetry $previousRetry }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-GraphSeed }
