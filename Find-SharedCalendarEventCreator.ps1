#Requires -Version 7.0
<#
.SYNOPSIS
  Find who created an event on a shared Exchange Online calendar.

.DESCRIPTION
  Signs in with app-only client credentials and reads the shared mailbox
  calendar through Microsoft Graph. PidTagCreatorName is the usual creator
  signal when the organizer is the shared mailbox. -Audit searches the
  Exchange mailbox audit log for Create operations.

  Keep TENANT_ID, CLIENT_ID, and CLIENT_SECRET on your machine. Do not paste
  the client secret into chat or commit .env.

.PARAMETER Calendar
  Shared mailbox SMTP address. The primary SMTP address is the safest value.

.PARAMETER Subject
  Event subject to find.

.PARAMETER Match
  Contains (default) is a case-insensitive substring applied in this script.
  Graph does not support contains() on event.subject. Equals is a full-subject
  match and is also sent as subject eq.

.PARAMETER Start
  Window start. YYYY-MM-DD or an ISO-8601 datetime. Default: 90 days ago (UTC).

.PARAMETER End
  Window end. YYYY-MM-DD or an ISO-8601 datetime. Default: 180 days ahead (UTC).

.PARAMETER Audit
  Search Purview for Exchange Create records. Requires the application
  permission AuditLogsQuery-Exchange.Read.All.

.PARAMETER AuditKeyword
  Keyword sent to the audit query. Default: the mailbox SMTP address.

.PARAMETER Masters
  Read /calendar/events (series masters and single instances) instead of
  calendarView occurrences.

.PARAMETER Json
  Print the report as JSON.

.EXAMPLE
  pwsh ./Find-SharedCalendarEventCreator.ps1 -Calendar shared-calendar@yourtenant.org -Subject "staff meeting"

.EXAMPLE
  pwsh ./Find-SharedCalendarEventCreator.ps1 -Calendar shared-calendar@yourtenant.org -Subject "Staff meeting" -Match Equals -Audit
#>
[CmdletBinding()]
param(
    [string]$Calendar,
    [string]$Subject,
    [ValidateSet('Contains', 'Equals')]
    [string]$Match = 'Contains',
    [string]$Start,
    [string]$End,
    [switch]$Audit,
    [string]$AuditKeyword,
    [switch]$Masters,
    [switch]$Json,
    [ValidateRange(1, 100)]
    [int]$Top = 50,
    [ValidateRange(1, 100)]
    [int]$MaxPages = 20,
    [ValidateRange(1, 100)]
    [int]$MaxHydrate = 25,
    [string]$EnvFile,
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret
)

class GraphRequestException : System.Exception {
    [int] $StatusCode
    GraphRequestException([string] $message, [int] $statusCode) : base($message) {
        $this.StatusCode = $statusCode
    }
}

$script:CreatorProperties = @(
    @{ GraphId = 'String 0x3FF8'; Name = 'PidTagCreatorName' }
    @{ GraphId = 'String 0x3FFA'; Name = 'PidTagLastModifierName' }
    @{ GraphId = 'String 0x5D01'; Name = 'PidTagSenderSmtpAddress' }
    @{ GraphId = 'String 0x5D02'; Name = 'PidTagSentRepresentingSmtpAddress' }
    @{ GraphId = 'String 0x0C1A'; Name = 'PidTagSenderName' }
    @{ GraphId = 'String 0x0042'; Name = 'PidTagSentRepresentingName' }
    @{ GraphId = 'String 0x0C1F'; Name = 'PidTagSenderEmailAddress' }
    @{ GraphId = 'String 0x0065'; Name = 'PidTagSentRepresentingEmailAddress' }
)

function Import-DotEnvFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }
    foreach ($rawLine in [System.IO.File]::ReadAllLines($Path)) {
        $line = $rawLine.Trim()
        if (-not $line -or $line.StartsWith('#')) {
            continue
        }
        if ($line.StartsWith('export ')) {
            $line = $line.Substring(7).Trim()
        }
        $separator = $line.IndexOf('=')
        if ($separator -lt 1) {
            continue
        }
        $key = $line.Substring(0, $separator).Trim()
        $value = $line.Substring($separator + 1).Trim()
        if ($value.Length -ge 2 -and (
                ($value.StartsWith('"') -and $value.EndsWith('"')) -or
                ($value.StartsWith("'") -and $value.EndsWith("'"))
            )) {
            $value = $value.Substring(1, $value.Length - 2)
        }
        if ($key -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') {
            continue
        }
        $existing = [Environment]::GetEnvironmentVariable($key)
        if ([string]::IsNullOrEmpty($existing)) {
            Set-Item -Path "Env:$key" -Value $value
        }
    }
}

function ConvertTo-TimeBoundary {
    param(
        [Parameter(Mandatory)][string]$Value,
        [Parameter(Mandatory)][ValidateSet('start', 'end')][string]$Which
    )
    $trimmed = $Value.Trim()
    if ($trimmed -match '^\d{4}-\d{2}-\d{2}$') {
        if ($Which -eq 'start') {
            return "${trimmed}T00:00:00.000Z"
        }
        return "${trimmed}T23:59:59.000Z"
    }
    try {
        $parsed = [datetimeoffset]::Parse(
            $trimmed,
            [cultureinfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind
        )
    }
    catch {
        throw "Invalid -$Which value '$Value'. Use YYYY-MM-DD or an ISO-8601 datetime."
    }
    return $parsed.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
}

function Get-DefaultCalendarWindow {
    $now = [datetimeoffset]::UtcNow
    return [pscustomobject]@{
        Start = $now.AddDays(-90).ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        End   = $now.AddDays(180).ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    }
}

function Get-AuditQueryWindow {
    param(
        [string]$WindowStart,
        [string]$WindowEnd,
        [datetimeoffset]$Now = [datetimeoffset]::UtcNow
    )
    $start = [datetimeoffset]::Parse($WindowStart)
    $end = [datetimeoffset]::Parse($WindowEnd)
    $retentionStart = $Now.AddDays(-180)
    $truncatedFuture = $false
    $truncatedRetention = $false
    if ($end -gt $Now) {
        $end = $Now
        $truncatedFuture = $true
    }
    if ($start -lt $retentionStart) {
        $start = $retentionStart
        $truncatedRetention = $true
    }
    $notes = @()
    if ($truncatedFuture) {
        $notes += 'The audit query end was clamped to the current time.'
    }
    if ($truncatedRetention) {
        $notes += 'The audit query start was clamped to the past 180 days (Microsoft Purview Audit Standard retention). If this tenant has a longer audit retention policy, pass a -Start/-End window inside that retention period.'
    }
    $empty = $start -ge $end
    if ($empty) {
        $notes += 'The calendar window does not overlap audit history that this script will query.'
    }
    return [pscustomobject]@{
        Start               = $start.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        End                 = $end.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        TruncatedFuture     = $truncatedFuture
        TruncatedRetention  = $truncatedRetention
        Empty               = $empty
        Note                = $(if ($notes.Count) { $notes -join ' ' } else { $null })
    }
}

function Test-SubjectMatch {
    param(
        [AllowNull()][string]$Haystack,
        [string]$Needle,
        [ValidateSet('Contains', 'Equals')][string]$Mode = 'Contains'
    )
    $left = $(if ($null -eq $Haystack) { '' } else { $Haystack }).Trim().ToLowerInvariant()
    $right = $Needle.Trim().ToLowerInvariant()
    if (-not $right) {
        return $false
    }
    if ($Mode -eq 'Equals') {
        return $left -eq $right
    }
    return $left.Contains($right)
}

function New-CalendarListUri {
    param(
        [string]$GraphBase,
        [string]$CalendarEmail,
        [string]$WindowStart,
        [string]$WindowEnd,
        [int]$PageSize,
        [bool]$UseMasters,
        [string]$SubjectEquals,
        [bool]$IncludeStartFilter = $true
    )
    $root = $GraphBase.TrimEnd('/')
    $user = [uri]::EscapeDataString($CalendarEmail)
    $path = if ($UseMasters) {
        "$root/users/$user/calendar/events"
    }
    else {
        "$root/users/$user/calendar/calendarView"
    }
    $query = [System.Collections.Generic.List[string]]::new()
    if (-not $UseMasters) {
        $query.Add("startDateTime=$([uri]::EscapeDataString($WindowStart))")
        $query.Add("endDateTime=$([uri]::EscapeDataString($WindowEnd))")
    }
    else {
        $select = 'id,subject,start,end,organizer,isOrganizer,changeKey,type,iCalUId,isCancelled,webLink'
        $query.Add('$select=' + [uri]::EscapeDataString($select))
    }
    $query.Add('$top=' + $PageSize)
    if ($SubjectEquals) {
        $literal = "'" + $SubjectEquals.Replace("'", "''") + "'"
        $query.Add('$filter=' + [uri]::EscapeDataString("subject eq $literal"))
    }
    elseif ($UseMasters -and $IncludeStartFilter) {
        $odataDate = ($WindowStart -replace '\.\d+', '') -replace 'Z$', ''
        $query.Add('$filter=' + [uri]::EscapeDataString("start/dateTime ge '$odataDate'"))
    }
    return "${path}?$($query -join '&')"
}

function Get-CreatorPropertyExpand {
    param([string[]]$GraphIds)
    $clauses = foreach ($id in $GraphIds) { "id eq '$id'" }
    $filter = $clauses -join ' or '
    return "singleValueExtendedProperties(`$filter=$filter)"
}

function New-EventHydrateUri {
    param(
        [string]$GraphBase,
        [string]$CalendarEmail,
        [string]$EventId,
        [string]$Expand
    )
    $root = $GraphBase.TrimEnd('/')
    $user = [uri]::EscapeDataString($CalendarEmail)
    $id = [uri]::EscapeDataString($EventId)
    $uri = "$root/users/$user/events/$id"
    if ($Expand) {
        $uri += '?$expand=' + [uri]::EscapeDataString($Expand)
    }
    return $uri
}

function ConvertTo-ExtendedProperties {
    param($Properties)
    $lookup = @{}
    foreach ($item in $script:CreatorProperties) {
        $lookup[$item.GraphId.ToLowerInvariant()] = $item.Name
    }
    $list = @()
    $byName = @{}
    foreach ($prop in @($Properties)) {
        if ($null -eq $prop -or [string]::IsNullOrWhiteSpace([string]$prop.id) -or $null -eq $prop.value) {
            continue
        }
        $id = [string]$prop.id
        $name = $lookup[$id.ToLowerInvariant()]
        $list += [pscustomobject]@{
            id            = $id
            canonicalName = $name
            value         = [string]$prop.value
        }
        if ($name) {
            $byName[$name] = [string]$prop.value
        }
    }
    return [pscustomobject]@{
        List   = @($list)
        ByName = $byName
    }
}

function Test-SameText {
    param([AllowNull()][string]$Left, [AllowNull()][string]$Right)
    $a = $(if ($null -eq $Left) { '' } else { $Left }).Trim().ToLowerInvariant()
    $b = $(if ($null -eq $Right) { '' } else { $Right }).Trim().ToLowerInvariant()
    return ($a.Length -gt 0 -and $a -eq $b)
}

function Get-BlankOrNull {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }
    return $Value.Trim()
}

function Get-InferredCreator {
    param(
        [string]$CalendarEmail,
        [AllowNull()][string]$OrganizerName,
        [AllowNull()][string]$OrganizerEmail,
        [hashtable]$Extended
    )
    $calendarEmail = $CalendarEmail.Trim()
    $organizerName = Get-BlankOrNull $OrganizerName
    $organizerEmail = Get-BlankOrNull $OrganizerEmail
    $organizerIsMailbox = $null
    if ($organizerEmail) {
        $organizerIsMailbox = Test-SameText $organizerEmail $calendarEmail
    }
    $creatorName = Get-BlankOrNull $Extended['PidTagCreatorName']
    $senderSmtp = $null
    foreach ($candidate in @($Extended['PidTagSenderSmtpAddress'], $Extended['PidTagSentRepresentingSmtpAddress'])) {
        $trimmed = Get-BlankOrNull $candidate
        if ($trimmed -and -not (Test-SameText $trimmed $calendarEmail)) {
            $senderSmtp = $trimmed
            break
        }
    }
    $senderName = Get-BlankOrNull $Extended['PidTagSenderName']
    if (-not $senderName) {
        $senderName = Get-BlankOrNull $Extended['PidTagSentRepresentingName']
    }
    $creatorIsMailbox = $false
    if ($creatorName -and ((Test-SameText $creatorName $calendarEmail) -or ($organizerIsMailbox -eq $true -and (Test-SameText $creatorName $organizerName)))) {
        $creatorIsMailbox = $true
    }

    if ($creatorName -and -not $creatorIsMailbox) {
        $email = $null
        if ((Test-SameText $creatorName $organizerName) -and $organizerIsMailbox -eq $false) {
            $email = $organizerEmail
        }
        elseif ($senderSmtp -and (-not $senderName -or (Test-SameText $senderName $creatorName))) {
            $email = $senderSmtp
        }
        $confidence = $(if ($email) { 'high' } else { 'medium' })
        $note = if ($email) {
            'PidTagCreatorName (MAPI 0x3FF8) names someone other than the shared mailbox, and an SMTP address on the item matches that person.'
        }
        else {
            'PidTagCreatorName (MAPI 0x3FF8) differs from the shared mailbox. On a shared calendar this is usually the display name of the person who created the item. Graph does not provide that person''s SMTP address here.'
        }
        return [pscustomobject]@{
            name                        = $creatorName
            email                       = $email
            source                      = 'PidTagCreatorName'
            confidence                  = $confidence
            organizerIsCalendarMailbox  = $organizerIsMailbox
            note                        = $note
        }
    }

    if ($senderSmtp) {
        return [pscustomobject]@{
            name                       = $senderName
            email                      = $senderSmtp
            source                     = 'PidTagSenderSmtpAddress'
            confidence                 = 'medium'
            organizerIsCalendarMailbox = $organizerIsMailbox
            note                       = 'PidTagSenderSmtpAddress or PidTagSentRepresentingSmtpAddress differs from the calendar mailbox. For a meeting this is often the sending mailbox, which may be the person who created the item.'
        }
    }

    if ($organizerEmail -and $organizerIsMailbox -eq $false) {
        return [pscustomobject]@{
            name                       = $organizerName
            email                      = $organizerEmail
            source                     = 'organizer'
            confidence                 = 'medium'
            organizerIsCalendarMailbox = $false
            note                       = 'The organizer is not the shared calendar mailbox. This looks like a meeting organized by that person rather than an appointment created directly on the shared calendar.'
        }
    }

    if ($creatorName) {
        return [pscustomobject]@{
            name                       = $creatorName
            email                      = $(if ($organizerIsMailbox) { $organizerEmail } else { $null })
            source                     = 'PidTagCreatorName'
            confidence                 = 'low'
            organizerIsCalendarMailbox = $organizerIsMailbox
            note                       = 'PidTagCreatorName matches the shared mailbox. Exchange often stores the mailbox as both organizer and creator when someone works in that calendar. A mailbox audit Create record can show whether a delegate did it.'
        }
    }

    $note = if ($organizerIsMailbox) {
        'Graph returned the shared mailbox as organizer and no distinct creator property. The event resource has no createdBy field. Run again with -Audit to search the Exchange mailbox audit log.'
    }
    else {
        'Graph did not return PidTagCreatorName or a distinct organizer. The event resource has no createdBy field. Run again with -Audit to search the Exchange mailbox audit log.'
    }
    return [pscustomobject]@{
        name                       = $null
        email                      = $null
        source                     = 'unresolved'
        confidence                 = 'none'
        organizerIsCalendarMailbox = $organizerIsMailbox
        note                       = $note
    }
}

function Get-BestCreator {
    param($Inferred, $AuditMatches)
    $matches = @($AuditMatches)
    if ($Inferred.confidence -in @('high', 'medium')) {
        return [pscustomobject]@{
            name       = $Inferred.name
            email      = $Inferred.email
            source     = $Inferred.source
            confidence = $Inferred.confidence
            detail     = $Inferred.note
        }
    }
    $users = [System.Collections.Generic.List[string]]::new()
    foreach ($hit in $matches) {
        $user = Get-BlankOrNull $hit.userId
        if (-not $user) { $user = Get-BlankOrNull $hit.userPrincipalName }
        if ($user -and -not $users.Contains($user)) {
            $users.Add($user)
        }
    }
    if ($users.Count -eq 1) {
        $hit = $matches | Where-Object { $_.userId -eq $users[0] -or $_.userPrincipalName -eq $users[0] } | Select-Object -First 1
        $confidence = $(if ($matches.Count -gt 1) { 'low' } else { 'medium' })
        $detail = if ($matches.Count -gt 1) {
            "Several Exchange Create audit records match this subject. They all name $($users[0])."
        }
        else {
            'Taken from the Exchange mailbox audit Create record. UserId is the account that performed the create.'
        }
        return [pscustomobject]@{
            name       = $hit.logonUserDisplayName
            email      = $users[0]
            source     = 'mailbox audit Create'
            confidence = $confidence
            detail     = $detail
        }
    }
    if ($users.Count -gt 1) {
        return [pscustomobject]@{
            name       = $null
            email      = $null
            source     = 'mailbox audit Create'
            confidence = 'low'
            detail     = "Several accounts created matching items in this window: $($users -join ', '). Narrow -Start/-End or use -Match Equals."
        }
    }
    if ($Inferred.confidence -eq 'low') {
        return [pscustomobject]@{
            name       = $Inferred.name
            email      = $Inferred.email
            source     = $Inferred.source
            confidence = 'low'
            detail     = $Inferred.note
        }
    }
    return [pscustomobject]@{
        name       = $null
        email      = $null
        source     = 'unresolved'
        confidence = 'none'
        detail     = $Inferred.note
    }
}

function ConvertTo-LogonTypeName {
    param($Value)
    if ($Value -is [string] -and -not [string]::IsNullOrWhiteSpace($Value)) {
        return $Value.Trim()
    }
    if ($null -ne $Value -and "$Value" -match '^\d+$') {
        switch ([int]$Value) {
            0 { return 'Owner' }
            1 { return 'Admin' }
            2 { return 'Delegate' }
            default { return "$Value" }
        }
    }
    return $null
}

function Get-ObjectProperty {
    param($Object, [string[]]$Names)
    if ($null -eq $Object) {
        return $null
    }
    if ($Object -is [string]) {
        return $null
    }
    foreach ($name in $Names) {
        $prop = $Object.PSObject.Properties[$name]
        if ($prop) {
            return $prop.Value
        }
    }
    return $null
}

function ConvertTo-AuditHit {
    param($Record)
    $data = Get-ObjectProperty $Record @('auditData')
    if ($data -is [string]) {
        try {
            $data = $data | ConvertFrom-Json
        }
        catch {
            $data = $null
        }
    }
    $item = Get-ObjectProperty $data @('Item', 'item')
    $parent = Get-ObjectProperty $item @('ParentFolder', 'parentFolder')
    $folderPath = Get-ObjectProperty $parent @('Path', 'path')
    if (-not $folderPath) {
        $folderPath = Get-ObjectProperty $data @('FolderPathName', 'folderPathName')
    }
    $subject = Get-ObjectProperty $item @('Subject', 'subject')
    if (-not $subject) {
        $subject = Get-ObjectProperty $data @('ItemSubject', 'itemSubject')
    }
    $created = Get-ObjectProperty $Record @('createdDateTime')
    if (-not $created) {
        $created = Get-ObjectProperty $data @('CreationTime', 'creationTime')
    }
    $operation = Get-ObjectProperty $Record @('operation')
    if (-not $operation) {
        $operation = Get-ObjectProperty $data @('Operation')
    }
    $userId = Get-ObjectProperty $Record @('userId')
    if (-not $userId) {
        $userId = Get-ObjectProperty $data @('UserId')
    }
    $folderText = Get-BlankOrNull ([string]$folderPath)
    $folderLooks = $null
    if ($folderText) {
        $folderLooks = $folderText -match 'calendar'
    }
    return [pscustomobject]@{
        id                    = Get-BlankOrNull ([string](Get-ObjectProperty $Record @('id')))
        createdDateTime       = Get-BlankOrNull ([string]$created)
        operation             = Get-BlankOrNull ([string]$operation)
        userId                = Get-BlankOrNull ([string]$userId)
        userPrincipalName     = Get-BlankOrNull ([string](Get-ObjectProperty $Record @('userPrincipalName')))
        logonType             = ConvertTo-LogonTypeName (Get-ObjectProperty $data @('LogonType', 'logonType'))
        logonUserDisplayName  = Get-BlankOrNull ([string](Get-ObjectProperty $data @('LogonUserDisplayName', 'logonUserDisplayName')))
        mailboxOwnerUpn       = Get-BlankOrNull ([string](Get-ObjectProperty $data @('MailboxOwnerUPN', 'mailboxOwnerUpn')))
        subject               = Get-BlankOrNull ([string]$subject)
        folderPath            = $folderText
        folderLooksLikeCalendar = $folderLooks
        clientInfo            = Get-BlankOrNull ([string](Get-ObjectProperty $data @('ClientInfoString', 'clientInfoString')))
        clientIp              = Get-BlankOrNull ([string]$(
                $ip = Get-ObjectProperty $Record @('clientIp')
                if (-not $ip) { $ip = Get-ObjectProperty $data @('ClientIP', 'clientIp') }
                $ip
            ))
    }
}

function Select-MatchingAuditHits {
    param(
        $Hits,
        [AllowNull()][string]$EventSubject,
        [AllowNull()][string]$EventCreated,
        [string]$CalendarEmail
    )
    $matches = foreach ($hit in @($Hits)) {
        if ($null -eq $hit) { continue }
        if ($hit.operation -and $hit.operation.ToLowerInvariant() -ne 'create') { continue }
        if ($hit.mailboxOwnerUpn -and -not (Test-SameText $hit.mailboxOwnerUpn $CalendarEmail)) { continue }
        if ($hit.folderLooksLikeCalendar -eq $false) { continue }
        if (-not (Test-SubjectMatch $hit.subject $EventSubject 'Equals')) { continue }
        $hit
    }
    $eventTime = $null
    if ($EventCreated) {
        try { $eventTime = [datetimeoffset]::Parse($EventCreated) } catch { $eventTime = $null }
    }
    return @($matches | Sort-Object `
        @{ Expression = { if ($_.folderLooksLikeCalendar -eq $true) { 0 } else { 1 } } }, `
        @{ Expression = {
                if (-not $eventTime -or -not $_.createdDateTime) { return [double]::MaxValue }
                try {
                    $hitTime = [datetimeoffset]::Parse($_.createdDateTime)
                    return [math]::Abs(($hitTime - $eventTime).TotalSeconds)
                }
                catch { return [double]::MaxValue }
            } })
}

function ConvertFrom-GraphDateTime {
    param($Value)
    $dateTime = Get-ObjectProperty $Value @('dateTime')
    if (-not $dateTime) {
        return $null
    }
    $text = [regex]::Replace([string]$dateTime, '(\.\d{3})\d+', '$1')
    $zone = [string](Get-ObjectProperty $Value @('timeZone'))
    $hasZone = $text -match '(?:Z|[+-]\d{2}:\d{2})$'
    if (-not $hasZone) {
        if ([string]::IsNullOrWhiteSpace($zone) -or $zone.Trim().ToUpperInvariant() -in @('UTC', 'GMT')) {
            $text = "${text}Z"
        }
        else {
            return $null
        }
    }
    try {
        return [datetimeoffset]::Parse($text, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
    }
    catch {
        return $null
    }
}

function Test-EventInWindow {
    param($Event, [string]$WindowStart, [string]$WindowEnd)
    $start = ConvertFrom-GraphDateTime $Event.start
    if ($null -eq $start) {
        return $true
    }
    $windowStart = [datetimeoffset]::Parse($WindowStart)
    $windowEnd = [datetimeoffset]::Parse($WindowEnd)
    return ($start -ge $windowStart -and $start -le $windowEnd)
}

function Format-PersonLine {
    param($Person)
    if ($Person.name -and $Person.email) {
        return "$($Person.name) <$($Person.email)>"
    }
    if ($Person.name) { return [string]$Person.name }
    if ($Person.email) { return [string]$Person.email }
    return '(unknown)'
}

function Format-OptionalBool {
    param($Value)
    if ($null -eq $Value) { return 'unknown' }
    if ($Value) { return 'yes' }
    return 'no'
}

function Format-LookupReport {
    param($Report)
    $sourcePath = if ($Report.source -eq 'calendarView') { 'calendar/calendarView' } else { 'calendar/events' }
    $matchDescription = if ($Report.match -eq 'Equals') {
        'case-insensitive equals. The script asks Graph for subject eq when that filter is accepted, then checks again locally.'
    }
    else {
        'case-insensitive contains, applied in this script. Graph does not support contains() on event.subject.'
    }
    $lines = @(
        "Calendar: $($Report.calendar)"
        "Subject query: $($Report.subjectQuery)"
        "Match: $($Report.match.ToLowerInvariant()) ($matchDescription)"
        "Source: /users/{calendar}/$sourcePath"
        "Window: $($Report.window.start) .. $($Report.window.end)"
        "Scanned events: $($Report.scannedEventCount) across $($Report.pagesFetched) page(s)"
        "Matches: $($Report.matchCount)"
    )
    if ($Report.truncated) {
        $lines += 'The event list stopped at -MaxPages. There may be more events in the window.'
    }
    $audit = $Report.audit
    $auditLine = switch ($audit.status) {
        'not_requested' { $audit.guidance }
        'skipped' { $audit.guidance }
        'error' { "failed. $($audit.error)" }
        'ok' {
            $range = if ($audit.window) { "$($audit.window.start) .. $($audit.window.end)" } else { 'unspecified' }
            $parts = @(
                "ok, query $($audit.queryId)"
                "keyword $($audit.keyword)"
                "window $range"
                "scanned $($audit.scannedRecordCount) Create record(s)"
            )
            if ($audit.windowNote) { $parts += $audit.windowNote }
            $parts -join '; '
        }
        default { $audit.status }
    }
    $lines += "Audit: $auditLine"
    if ($Report.matchCount -eq 0) {
        $lines += ''
        $lines += 'No matching events.'
        return ($lines -join "`n")
    }
    foreach ($event in @($Report.events)) {
        $lines += ''
        $lines += '---'
        $lines += "Subject: $(if ($event.subject) { $event.subject } else { '(none)' })"
        $lines += "Id: $(if ($event.id) { $event.id } else { '(none)' })"
        $lines += "Type: $(if ($event.type) { $event.type } else { '(none)' })"
        $lines += "iCalUId: $(if ($event.iCalUId) { $event.iCalUId } else { '(none)' })"
        $startText = if ($event.start.dateTime) {
            if ($event.start.timeZone) { "$($event.start.dateTime) ($($event.start.timeZone))" } else { $event.start.dateTime }
        }
        else { '(none)' }
        $endText = if ($event.end.dateTime) {
            if ($event.end.timeZone) { "$($event.end.dateTime) ($($event.end.timeZone))" } else { $event.end.dateTime }
        }
        else { '(none)' }
        $lines += "Start: $startText"
        $lines += "End: $endText"
        $lines += "Created: $(if ($event.createdDateTime) { $event.createdDateTime } else { '(none)' })"
        $lines += "Last modified: $(if ($event.lastModifiedDateTime) { $event.lastModifiedDateTime } else { '(none)' })"
        $lines += "Organizer: $(Format-PersonLine $event.organizer)"
        $createdBy = if ($event.createdBy) { Format-PersonLine $event.createdBy } else { 'not returned by the Graph event resource' }
        $modifiedBy = if ($event.lastModifiedBy) { Format-PersonLine $event.lastModifiedBy } else { 'not returned by the Graph event resource' }
        $lines += "createdBy: $createdBy"
        $lines += "lastModifiedBy: $modifiedBy"
        if ($event.hydrateError) {
            $lines += "Extended properties: unavailable ($($event.hydrateError))"
        }
        elseif (-not $event.singleValueExtendedProperties -or @($event.singleValueExtendedProperties).Count -eq 0) {
            $lines += 'singleValueExtendedProperties: none returned'
        }
        else {
            $lines += 'singleValueExtendedProperties:'
            foreach ($property in @($event.singleValueExtendedProperties)) {
                $label = if ($property.canonicalName) { $property.canonicalName } else { $property.id }
                $lines += "  ${label}: $($property.value)"
            }
        }
        $lines += "Inferred creator: $(Format-PersonLine $event.inferredCreator)"
        $lines += "  source: $($event.inferredCreator.source)"
        $lines += "  confidence: $($event.inferredCreator.confidence)"
        $lines += "  organizer is calendar mailbox: $(Format-OptionalBool $event.inferredCreator.organizerIsCalendarMailbox)"
        $lines += "  note: $($event.inferredCreator.note)"
        if ($null -ne $event.audit) {
            if (@($event.audit).Count -eq 0) {
                $lines += 'Audit matches: none for this subject'
            }
            else {
                $lines += "Audit matches: $(@($event.audit).Count)"
                $shown = 0
                foreach ($hit in @($event.audit)) {
                    if ($shown -ge 5) { break }
                    $who = if ($hit.logonUserDisplayName) {
                        "$($hit.logonUserDisplayName) <$($(if ($hit.userId) { $hit.userId } else { $hit.userPrincipalName }))>"
                    }
                    elseif ($hit.userId) { $hit.userId }
                    else { 'unknown user' }
                    $bits = @($who)
                    if ($hit.logonType) { $bits += "logon $($hit.logonType)" }
                    $bits += $(if ($hit.createdDateTime) { $hit.createdDateTime } else { 'time unknown' })
                    if ($hit.subject) { $bits += "subject `"$($hit.subject)`"" }
                    if ($hit.folderPath) { $bits += "folder $($hit.folderPath)" }
                    $lines += "  - $($bits -join '; ')"
                    $shown++
                }
            }
        }
        $lines += "Best creator: $(Format-PersonLine $event.bestCreator)"
        $lines += "  source: $($event.bestCreator.source)"
        $lines += "  confidence: $($event.bestCreator.confidence)"
        $lines += "  detail: $($event.bestCreator.detail)"
    }
    return ($lines -join "`n")
}

function Get-ActorRef {
    param($Value)
    if ($null -eq $Value) { return $null }
    $user = Get-ObjectProperty $Value @('user')
    if (-not $user) { $user = $Value }
    $email = Get-ObjectProperty $user @('email', 'address')
    if (-not $email) {
        $emailAddress = Get-ObjectProperty $user @('emailAddress')
        $email = Get-ObjectProperty $emailAddress @('address')
    }
    $name = Get-ObjectProperty $user @('displayName', 'name')
    if (-not $name) {
        $emailAddress = Get-ObjectProperty $user @('emailAddress')
        $name = Get-ObjectProperty $emailAddress @('name')
    }
    $email = Get-BlankOrNull ([string]$email)
    $name = Get-BlankOrNull ([string]$name)
    if (-not $email -and -not $name) { return $null }
    return [pscustomobject]@{ name = $name; email = $email }
}

function Invoke-GraphRequest {
    param(
        [ValidateSet('GET', 'POST')][string]$Method,
        [string]$Uri,
        [string]$Token,
        $Body,
        [string]$AllowedHost,
        [int]$MaxAttempts = 4
    )
    $target = [uri]$Uri
    if ($target.Scheme -ne 'https' -or $target.Host -ne $AllowedHost) {
        throw "Refusing to send the Graph token to $($target.Host)."
    }
    $attempt = 0
    while ($true) {
        $attempt++
        $headers = @{
            Authorization = "Bearer $Token"
            Accept        = 'application/json'
            Prefer        = 'outlook.timezone="UTC"'
        }
        $requestParams = @{
            Method             = $Method
            Uri                = $Uri
            Headers            = $headers
            SkipHttpErrorCheck = $true
        }
        if ($null -ne $Body) {
            $requestParams.ContentType = 'application/json'
            $requestParams.Body = ($Body | ConvertTo-Json -Depth 6 -Compress)
        }
        $response = Invoke-WebRequest @requestParams
        if ($response.StatusCode -in 429, 503 -and $attempt -lt $MaxAttempts) {
            $retryAfter = 0
            $retryHeader = [string]$response.Headers['Retry-After']
            if ($retryHeader) {
                [void][int]::TryParse(($retryHeader -split '[,\s]')[0], [ref]$retryAfter)
            }
            $delaySeconds = if ($retryAfter -gt 0) { [math]::Min($retryAfter, 10) } else { [math]::Min($attempt, 10) }
            Start-Sleep -Seconds $delaySeconds
            continue
        }
        $content = [string]$response.Content
        if ($response.StatusCode -ge 400) {
            $message = $content
            try {
                $parsed = $content | ConvertFrom-Json
                if ($parsed.error.message) {
                    $code = $parsed.error.code
                    $message = $(if ($code) { "${code}: $($parsed.error.message)" } else { [string]$parsed.error.message })
                }
            }
            catch { }
            if ($message.Length -gt 500) {
                $message = $message.Substring(0, 500)
            }
            throw [GraphRequestException]::new("Graph $($response.StatusCode) $message".Trim(), [int]$response.StatusCode)
        }
        if ([string]::IsNullOrWhiteSpace($content)) {
            return $null
        }
        return $content | ConvertFrom-Json
    }
}

function Get-GraphToken {
    param(
        [string]$AuthorityHost,
        [string]$Tenant,
        [string]$AppId,
        [string]$Secret,
        [string]$Scope
    )
    $tokenUri = "$($AuthorityHost.TrimEnd('/'))/$([uri]::EscapeDataString($Tenant))/oauth2/v2.0/token"
    $form = "client_id=$([uri]::EscapeDataString($AppId))&client_secret=$([uri]::EscapeDataString($Secret))&scope=$([uri]::EscapeDataString($Scope))&grant_type=client_credentials"
    $response = Invoke-WebRequest -Method Post -Uri $tokenUri -Body $form -ContentType 'application/x-www-form-urlencoded' -SkipHttpErrorCheck
    $content = [string]$response.Content
    $redacted = $content.Replace($Secret, '[redacted]')
    if ($response.StatusCode -ge 400) {
        throw "Token request failed ($($response.StatusCode)). Check TENANT_ID, CLIENT_ID, and CLIENT_SECRET. $redacted"
    }
    $parsed = $content | ConvertFrom-Json
    if (-not $parsed.access_token) {
        throw "Token request failed. Check TENANT_ID, CLIENT_ID, and CLIENT_SECRET. $($redacted.Substring(0, [math]::Min(400, $redacted.Length)))"
    }
    return [string]$parsed.access_token
}

function Get-PagedGraphValues {
    param(
        [string]$FirstUri,
        [string]$Token,
        [string]$AllowedHost,
        [int]$PageLimit
    )
    $items = @()
    $uri = $FirstUri
    $pages = 0
    $truncated = $false
    while ($uri -and $pages -lt $PageLimit) {
        $page = Invoke-GraphRequest -Method GET -Uri $uri -Token $Token -AllowedHost $AllowedHost
        $items += @($page.value)
        $uri = [string](Get-ObjectProperty $page @('@odata.nextLink'))
        $pages++
    }
    if ($uri) { $truncated = $true }
    return [pscustomobject]@{
        Items     = @($items | Where-Object { $null -ne $_ })
        Pages     = $pages
        Truncated = $truncated
    }
}

function Get-CalendarEventPage {
    param(
        [string]$GraphBase,
        [string]$Token,
        [string]$CalendarEmail,
        [string]$WindowStart,
        [string]$WindowEnd,
        [int]$PageSize,
        [int]$PageLimit,
        [bool]$UseMasters,
        [string]$SubjectEquals,
        [bool]$IncludeStartFilter,
        [bool]$AllowServerFilterRetry = $true
    )
    $uri = New-CalendarListUri -GraphBase $GraphBase -CalendarEmail $CalendarEmail -WindowStart $WindowStart -WindowEnd $WindowEnd -PageSize $PageSize -UseMasters $UseMasters -SubjectEquals $SubjectEquals -IncludeStartFilter $IncludeStartFilter
    $hostName = ([uri]$GraphBase).Host
    try {
        $page = Get-PagedGraphValues -FirstUri $uri -Token $Token -AllowedHost $hostName -PageLimit $PageLimit
        return [pscustomobject]@{
            Events             = $page.Items
            Source             = $(if ($UseMasters) { 'events' } else { 'calendarView' })
            PagesFetched       = $page.Pages
            Truncated          = $page.Truncated
            ServerSubjectFilter = [bool]$SubjectEquals
        }
    }
    catch {
        $status = $_.Exception.StatusCode
        $canRetry = $AllowServerFilterRetry -and ($status -eq 400) -and ($SubjectEquals -or ($UseMasters -and $IncludeStartFilter))
        if (-not $canRetry) { throw }
        return Get-CalendarEventPage -GraphBase $GraphBase -Token $Token -CalendarEmail $CalendarEmail -WindowStart $WindowStart -WindowEnd $WindowEnd -PageSize $PageSize -PageLimit $PageLimit -UseMasters $UseMasters -SubjectEquals $null -IncludeStartFilter $false -AllowServerFilterRetry $false
    }
}

function Get-HydratedEvent {
    param(
        [string]$GraphBase,
        [string]$Token,
        [string]$CalendarEmail,
        $Event
    )
    if (-not $Event.id) {
        return [pscustomobject]@{ Event = $Event; Error = 'Event has no id, so creator properties were not loaded.' }
    }
    $fullIds = @($script:CreatorProperties | ForEach-Object { $_.GraphId })
    $expands = @(
        (Get-CreatorPropertyExpand -GraphIds $fullIds)
        (Get-CreatorPropertyExpand -GraphIds @('String 0x3FF8', 'String 0x3FFA'))
        $null
    )
    $hostName = ([uri]$GraphBase).Host
    $lastError = $null
    for ($i = 0; $i -lt $expands.Count; $i++) {
        $uri = New-EventHydrateUri -GraphBase $GraphBase -CalendarEmail $CalendarEmail -EventId ([string]$Event.id) -Expand $expands[$i]
        try {
            $full = Invoke-GraphRequest -Method GET -Uri $uri -Token $Token -AllowedHost $hostName
            return [pscustomobject]@{ Event = $full; Error = $null }
        }
        catch {
            $lastError = $_.Exception.Message
            $retryable = ($_.Exception.StatusCode -eq 400 -and $i -lt ($expands.Count - 1))
            if (-not $retryable) {
                return [pscustomobject]@{ Event = $Event; Error = $lastError }
            }
        }
    }
    return [pscustomobject]@{ Event = $Event; Error = $lastError }
}

function Search-CalendarCreates {
    param(
        [string]$GraphBase,
        [string]$Token,
        [string]$CalendarEmail,
        $AuditWindow,
        [string]$Keyword
    )
    $hostName = ([uri]$GraphBase).Host
    $root = $GraphBase.TrimEnd('/')
    $created = Invoke-GraphRequest -Method POST -Uri "$root/security/auditLog/queries" -Token $Token -AllowedHost $hostName -Body ([ordered]@{
            displayName         = 'shared-calendar-creator Create lookup'
            filterStartDateTime = $AuditWindow.Start
            filterEndDateTime   = $AuditWindow.End
            recordTypeFilters   = @('exchangeItem')
            operationFilters    = @('Create')
            keywordFilter       = $Keyword
        })
    if (-not $created.id) {
        throw 'Audit query was created without an id.'
    }
    $deadline = [datetimeoffset]::UtcNow.AddSeconds(90)
    $status = if ($created.status) { [string]$created.status } else { 'notStarted' }
    while ($status.ToLowerInvariant() -notin @('succeeded', 'failed', 'cancelled')) {
        if ([datetimeoffset]::UtcNow -ge $deadline) {
            throw "Audit query $($created.id) was still $status after 90s."
        }
        Start-Sleep -Seconds 2
        $current = Invoke-GraphRequest -Method GET -Uri "$root/security/auditLog/queries/$([uri]::EscapeDataString([string]$created.id))" -Token $Token -AllowedHost $hostName
        if ($current.status) { $status = [string]$current.status }
    }
    if ($status.ToLowerInvariant() -ne 'succeeded') {
        throw "Audit query $($created.id) finished with status $status."
    }
    $hits = @()
    $scanned = 0
    $next = "$root/security/auditLog/queries/$([uri]::EscapeDataString([string]$created.id))/records"
    $pages = 0
    while ($next -and $pages -lt 10) {
        $page = Invoke-GraphRequest -Method GET -Uri $next -Token $Token -AllowedHost $hostName
        foreach ($record in @($page.value)) {
            if ($null -eq $record) { continue }
            $scanned++
            $hit = ConvertTo-AuditHit $record
            if (-not $hit.operation -or $hit.operation.ToLowerInvariant() -eq 'create') {
                $hits += $hit
            }
        }
        $next = [string](Get-ObjectProperty $page @('@odata.nextLink'))
        $pages++
    }
    return [pscustomobject]@{
        QueryId             = [string]$created.id
        Hits                = @($hits)
        ScannedRecordCount  = $scanned
    }
}

function Build-EventReport {
    param(
        $Event,
        [string]$CalendarEmail,
        $AuditHits,
        [string]$HydrateError
    )
    $extended = ConvertTo-ExtendedProperties $Event.singleValueExtendedProperties
    $organizerName = Get-ObjectProperty (Get-ObjectProperty $Event.organizer @('emailAddress')) @('name')
    $organizerEmail = Get-ObjectProperty (Get-ObjectProperty $Event.organizer @('emailAddress')) @('address')
    $inferred = Get-InferredCreator -CalendarEmail $CalendarEmail -OrganizerName ([string]$organizerName) -OrganizerEmail ([string]$organizerEmail) -Extended $extended.ByName
    $ranked = $null
    if ($null -ne $AuditHits) {
        $ranked = @(Select-MatchingAuditHits -Hits $AuditHits -EventSubject ([string]$Event.subject) -EventCreated ([string]$Event.createdDateTime) -CalendarEmail $CalendarEmail)
    }
    $best = Get-BestCreator -Inferred $inferred -AuditMatches $(if ($ranked) { $ranked } else { @() })
    return [pscustomobject]@{
        id                            = Get-BlankOrNull ([string]$Event.id)
        subject                       = Get-BlankOrNull ([string]$Event.subject)
        start                         = $Event.start
        end                           = $Event.end
        type                          = Get-BlankOrNull ([string]$Event.type)
        iCalUId                       = Get-BlankOrNull ([string]$Event.iCalUId)
        isCancelled                   = $Event.isCancelled
        createdDateTime               = Get-BlankOrNull ([string]$Event.createdDateTime)
        lastModifiedDateTime          = Get-BlankOrNull ([string]$Event.lastModifiedDateTime)
        organizer                     = [pscustomobject]@{ name = (Get-BlankOrNull ([string]$organizerName)); email = (Get-BlankOrNull ([string]$organizerEmail)) }
        createdBy                     = Get-ActorRef $Event.createdBy
        lastModifiedBy                = Get-ActorRef $Event.lastModifiedBy
        singleValueExtendedProperties = @($extended.List)
        extendedProperties            = $extended.ByName
        inferredCreator               = $inferred
        audit                         = $ranked
        bestCreator                   = $best
        hydrateError                  = $(if ($HydrateError) { $HydrateError } else { $null })
    }
}

function Start-SharedCalendarLookup {
    if ([string]::IsNullOrWhiteSpace($Calendar) -or [string]::IsNullOrWhiteSpace($Subject)) {
        throw 'Both -Calendar and -Subject are required.'
    }
    if ($Calendar -notmatch '@') {
        throw '-Calendar must be an email address.'
    }
    $Calendar = $Calendar.Trim()
    $Subject = $Subject.Trim()

    if ($EnvFile) {
        if (-not (Test-Path -LiteralPath $EnvFile)) {
            throw "Env file not found: $EnvFile"
        }
        Import-DotEnvFile $EnvFile
    }
    else {
        Import-DotEnvFile (Join-Path $PSScriptRoot '.env')
        $cwdEnv = Join-Path (Get-Location) '.env'
        $scriptEnv = Join-Path $PSScriptRoot '.env'
        if ((Test-Path -LiteralPath $cwdEnv) -and ((Resolve-Path -LiteralPath $cwdEnv).Path -ne (Resolve-Path -LiteralPath $scriptEnv -ErrorAction SilentlyContinue).Path)) {
            Import-DotEnvFile $cwdEnv
        }
    }

    $resolvedTenant = if ($TenantId) { $TenantId.Trim() } else { [string]$env:TENANT_ID }
    $resolvedClient = if ($ClientId) { $ClientId.Trim() } else { [string]$env:CLIENT_ID }
    $resolvedSecret = if ($ClientSecret) { $ClientSecret.Trim() } else { [string]$env:CLIENT_SECRET }
    $missing = @()
    if ([string]::IsNullOrWhiteSpace($resolvedTenant)) { $missing += 'TENANT_ID' }
    if ([string]::IsNullOrWhiteSpace($resolvedClient)) { $missing += 'CLIENT_ID' }
    if ([string]::IsNullOrWhiteSpace($resolvedSecret)) { $missing += 'CLIENT_SECRET' }
    if ($missing.Count) {
        throw "Missing $($missing -join ', '). Set them in the environment or a local .env file. Do not commit the secret or paste it into chat."
    }

    $graphBase = if ($env:GRAPH_BASE_URL) { $env:GRAPH_BASE_URL.Trim().TrimEnd('/') } else { 'https://graph.microsoft.com/v1.0' }
    $authority = if ($env:AUTHORITY_HOST) { $env:AUTHORITY_HOST.Trim().TrimEnd('/') } else { 'https://login.microsoftonline.com' }
    foreach ($pair in @(@{ Name = 'GRAPH_BASE_URL'; Value = $graphBase }, @{ Name = 'AUTHORITY_HOST'; Value = $authority })) {
        $parsedUri = $null
        if (-not [uri]::TryCreate($pair.Value, [UriKind]::Absolute, [ref]$parsedUri) -or $parsedUri.Scheme -ne 'https') {
            throw "$($pair.Name) must be an https URL."
        }
    }

    $defaults = Get-DefaultCalendarWindow
    $windowStart = if ($Start) { ConvertTo-TimeBoundary -Value $Start -Which start } else { $defaults.Start }
    $windowEnd = if ($End) { ConvertTo-TimeBoundary -Value $End -Which end } else { $defaults.End }
    if ([datetimeoffset]::Parse($windowStart) -ge [datetimeoffset]::Parse($windowEnd)) {
        throw '-Start must be earlier than -End.'
    }

    $auditEnabled = $false
    if ($script:ScriptBoundParameters.ContainsKey('Audit')) {
        $auditEnabled = [bool]$Audit
    }
    elseif ($env:AUDIT_FALLBACK -match '^(1|true|yes|on)$') {
        $auditEnabled = $true
    }

    $scope = ([uri]$graphBase).GetLeftPart([System.UriPartial]::Authority) + '/.default'
    $token = Get-GraphToken -AuthorityHost $authority -Tenant $resolvedTenant -AppId $resolvedClient -Secret $resolvedSecret -Scope $scope
    $allowedHost = ([uri]$graphBase).Host
    $subjectEquals = if ($Match -eq 'Equals') { $Subject } else { $null }

    try {
        $listed = Get-CalendarEventPage -GraphBase $graphBase -Token $token -CalendarEmail $Calendar -WindowStart $windowStart -WindowEnd $windowEnd -PageSize $Top -PageLimit $MaxPages -UseMasters ([bool]$Masters) -SubjectEquals $subjectEquals -IncludeStartFilter $true
    }
    catch {
        $status = $_.Exception.StatusCode
        $message = $_.Exception.Message
        if ($status -eq 403) {
            throw "$message Application Calendars.Read is required, with admin consent. If an Exchange application access policy is in place, $Calendar must be in the policy group. Check with: Test-ApplicationAccessPolicy -Identity $Calendar -AppId <client-id>"
        }
        if ($status -eq 404) {
            throw "$message The mailbox or calendar was not found. Use the shared mailbox primary SMTP address for -Calendar."
        }
        if ($status -eq 401) {
            throw "$message The app token was rejected. Confirm admin consent was granted and the secret is current."
        }
        throw
    }

    $matched = @(foreach ($event in @($listed.Events)) {
            if ($null -eq $event) { continue }
            if (-not (Test-SubjectMatch ([string]$event.subject) $Subject $Match)) { continue }
            if ($listed.Source -ne 'calendarView' -and -not (Test-EventInWindow $event $windowStart $windowEnd)) { continue }
            $event
        })
    $rows = @()
    $hydratedCount = 0
    $index = 0
    foreach ($event in $matched) {
        if ($index -ge $MaxHydrate) {
            $rows += [pscustomobject]@{
                Event = $event
                Error = "Skipped extended-property lookup because the match count exceeded -MaxHydrate $MaxHydrate."
            }
        }
        else {
            $hydrated = Get-HydratedEvent -GraphBase $graphBase -Token $token -CalendarEmail $Calendar -Event $event
            if (-not $hydrated.Error) { $hydratedCount++ }
            $rows += $hydrated
        }
        $index++
    }

    $unresolved = $false
    foreach ($row in $rows) {
        $preview = Build-EventReport -Event $row.Event -CalendarEmail $Calendar -AuditHits $null -HydrateError $row.Error
        if ($preview.inferredCreator.confidence -in @('low', 'none')) {
            $unresolved = $true
        }
    }

    $auditSection = [pscustomobject]@{
        status              = 'not_requested'
        guidance            = $(if ($unresolved) {
                'Mailbox audit fallback is off. Re-run with -Audit (or AUDIT_FALLBACK=true) after an admin grants the application permission AuditLogsQuery-Exchange.Read.All and consents to it. Purview audits a calendar item Create for the mailbox owner, a delegate, and an admin when mailbox auditing is enabled. Shared-mailbox searches use the mailbox SMTP address as the keyword, not as the user filter. Audit (Standard) retention is 180 days.'
            }
            else {
                'Not requested. Pass -Audit to compare these events with Exchange Create audit records.'
            })
        queryId             = $null
        window              = $null
        windowNote          = $null
        keyword             = $null
        scannedRecordCount  = $null
        error               = $null
    }
    $auditHits = $null
    $exitCode = 0
    if ($auditEnabled -and $rows.Count -eq 0) {
        $auditSection.status = 'skipped'
        $auditSection.guidance = 'No calendar events matched, so no audit query was sent.'
    }
    elseif ($auditEnabled) {
        $auditWindow = Get-AuditQueryWindow -WindowStart $windowStart -WindowEnd $windowEnd
        $keyword = if ($AuditKeyword) { $AuditKeyword } else { $Calendar }
        $auditSection.keyword = $keyword
        $auditSection.window = [pscustomobject]@{ start = $auditWindow.Start; end = $auditWindow.End }
        $auditSection.windowNote = $auditWindow.Note
        if ($auditWindow.Empty) {
            $auditSection.status = 'skipped'
            $auditSection.guidance = $auditWindow.Note
        }
        else {
            try {
                $auditResult = Search-CalendarCreates -GraphBase $graphBase -Token $token -CalendarEmail $Calendar -AuditWindow $auditWindow -Keyword $keyword
                $auditSection.status = 'ok'
                $auditSection.guidance = $null
                $auditSection.queryId = $auditResult.QueryId
                $auditSection.scannedRecordCount = $auditResult.ScannedRecordCount
                $auditHits = @($auditResult.Hits)
            }
            catch {
                $auditSection.status = 'error'
                $auditSection.guidance = $null
                $detail = $_.Exception.Message
                if ($_.Exception.StatusCode -in 401, 403) {
                    $detail = "$detail Grant application permission AuditLogsQuery-Exchange.Read.All (admin consent). The Graph audit-log query API is POST /security/auditLog/queries. Confirm mailbox auditing is enabled and Create is still audited for the shared mailbox."
                }
                $auditSection.error = $detail
                $exitCode = 2
            }
        }
    }

    $eventReports = @(foreach ($row in $rows) {
            Build-EventReport -Event $row.Event -CalendarEmail $Calendar -AuditHits $auditHits -HydrateError $row.Error
        })
    $report = [pscustomobject]@{
        calendar           = $Calendar
        subjectQuery       = $Subject
        match              = $Match
        source             = $listed.Source
        window             = [pscustomobject]@{ start = $windowStart; end = $windowEnd }
        truncated          = [bool]$listed.Truncated
        pagesFetched       = $listed.PagesFetched
        scannedEventCount  = @($listed.Events).Count
        matchCount         = $eventReports.Count
        hydratedCount      = $hydratedCount
        audit              = $auditSection
        events             = @($eventReports)
    }

    if ($Json) {
        [Console]::Out.WriteLine(($report | ConvertTo-Json -Depth 8))
    }
    else {
        [Console]::Out.WriteLine((Format-LookupReport $report))
    }
    return $exitCode
}

if ($MyInvocation.InvocationName -ne '.') {
    $script:ScriptBoundParameters = $PSBoundParameters
    try {
        exit (Start-SharedCalendarLookup)
    }
    catch {
        [Console]::Error.WriteLine($_.Exception.Message)
        exit 1
    }
}
