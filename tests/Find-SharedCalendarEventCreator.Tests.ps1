#Requires -Version 7.0
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..' 'Find-SharedCalendarEventCreator.ps1')

$failures = [System.Collections.Generic.List[string]]::new()
function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { $script:failures.Add($Message) }
}
function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    if ($Actual -ne $Expected) {
        $script:failures.Add("$Message expected [$Expected] actual [$Actual]")
    }
}

Assert-True (Test-SubjectMatch 'Weekly Staff Meeting' 'staff' 'Contains') 'contains matches case-insensitively'
Assert-True (-not (Test-SubjectMatch 'Weekly Staff Meeting' 'budget' 'Contains')) 'contains rejects other text'
Assert-True (Test-SubjectMatch 'Staff Meeting' 'staff meeting' 'Equals') 'equals ignores case'
Assert-True (-not (Test-SubjectMatch 'Weekly Staff Meeting' 'staff' 'Equals')) 'equals requires the full subject'

$calendar = 'shared-calendar@example.com'
$uri = [uri](New-CalendarListUri -GraphBase 'https://graph.microsoft.com/v1.0' -CalendarEmail $calendar -WindowStart '2026-01-01T00:00:00.000Z' -WindowEnd '2026-02-01T00:00:00.000Z' -PageSize 50 -UseMasters $false -SubjectEquals $null)
Assert-True ($uri.AbsolutePath -eq "/v1.0/users/$([uri]::EscapeDataString($calendar))/calendar/calendarView") 'calendarView path'
Assert-True ($uri.Query -match 'startDateTime=') 'calendarView has a start'
Assert-True ($uri.Query -notmatch '\$select|%24select') 'calendarView omits $select so createdDateTime is returned'

$equals = [uri](New-CalendarListUri -GraphBase 'https://graph.microsoft.com/v1.0' -CalendarEmail $calendar -WindowStart '2026-01-01T00:00:00.000Z' -WindowEnd '2026-02-01T00:00:00.000Z' -PageSize 25 -UseMasters $false -SubjectEquals "O'Brien review")
$filter = $null
foreach ($part in $equals.Query.TrimStart('?').Split('&')) {
    $separator = $part.IndexOf('=')
    $key = [uri]::UnescapeDataString($part.Substring(0, $separator))
    if ($key -eq '$filter') {
        $filter = [uri]::UnescapeDataString($part.Substring($separator + 1))
    }
}
Assert-Equal $filter "subject eq 'O''Brien review'" 'equals filter escapes quotes'

$hydrate = New-EventHydrateUri -GraphBase 'https://graph.microsoft.com/v1.0' -CalendarEmail $calendar -EventId 'event id/with slash' -Expand (Get-CreatorPropertyExpand -GraphIds @('String 0x3FF8', 'String 0x3FFA'))
Assert-True ($hydrate -match 'String(%20|\+)0x3FF8') 'hydrate expand requests PidTagCreatorName'
Assert-True ($hydrate -match 'event%20id%2Fwith%20slash') 'event id keeps the slash encoded'

$inferred = Get-InferredCreator -CalendarEmail $calendar -OrganizerName 'Community Calendar' -OrganizerEmail $calendar -Extended @{ PidTagCreatorName = 'Lance Smith' }
Assert-Equal $inferred.source 'PidTagCreatorName' 'creator source'
Assert-Equal $inferred.confidence 'medium' 'display name without smtp is medium'
Assert-Equal $inferred.organizerIsCalendarMailbox $true 'organizer is the mailbox'

$withSmtp = Get-InferredCreator -CalendarEmail $calendar -OrganizerName 'Community Calendar' -OrganizerEmail $calendar -Extended @{
    PidTagCreatorName        = 'Lance Smith'
    PidTagSenderName         = 'Lance Smith'
    PidTagSenderSmtpAddress  = 'lance@example.com'
}
Assert-Equal $withSmtp.confidence 'high' 'matching smtp raises confidence'
Assert-Equal $withSmtp.email 'lance@example.com' 'smtp is kept'

$organizer = Get-InferredCreator -CalendarEmail $calendar -OrganizerName 'Ada Lovelace' -OrganizerEmail 'ada@example.com' -Extended @{}
Assert-Equal $organizer.source 'organizer' 'different organizer is used'
Assert-Equal $organizer.organizerIsCalendarMailbox $false 'organizer is not the mailbox'

$low = Get-InferredCreator -CalendarEmail $calendar -OrganizerName 'Community Calendar' -OrganizerEmail $calendar -Extended @{ PidTagCreatorName = 'Community Calendar' }
Assert-Equal $low.confidence 'low' 'mailbox creator name stays low'
$best = Get-BestCreator -Inferred $low -AuditMatches @([pscustomobject]@{
        userId = 'delegate@example.com'; userPrincipalName = $null; logonUserDisplayName = 'Dee Delegate'
    })
Assert-Equal $best.source 'mailbox audit Create' 'audit user replaces a mailbox-only creator'
Assert-Equal $best.email 'delegate@example.com' 'audit user id'

$hit = ConvertTo-AuditHit ([pscustomobject]@{
        id              = 'rec-1'
        createdDateTime = '2026-09-01T15:00:01Z'
        operation       = 'Create'
        userId          = 'delegate@example.com'
        auditData       = [pscustomobject]@{
            LogonType = 2
            MailboxOwnerUPN = $calendar
            Item = [pscustomobject]@{
                Subject = 'Staff meeting'
                ParentFolder = [pscustomobject]@{ Path = '\Calendar' }
            }
        }
    })
Assert-Equal $hit.logonType 'Delegate' 'logon type 2 is Delegate'
Assert-Equal $hit.subject 'Staff meeting' 'audit subject'
Assert-Equal $hit.folderLooksLikeCalendar $true 'calendar folder'

$ranked = @(Select-MatchingAuditHits -Hits @(
        $hit
        (ConvertTo-AuditHit ([pscustomobject]@{
                operation = 'Create'
                userId = 'other@example.com'
                auditData = [pscustomobject]@{
                    MailboxOwnerUPN = 'someone-else@example.com'
                    Item = [pscustomobject]@{ Subject = 'Staff meeting'; ParentFolder = [pscustomobject]@{ Path = '\Calendar' } }
                }
            }))
        (ConvertTo-AuditHit ([pscustomobject]@{
                operation = 'Update'
                userId = 'pat@example.com'
                auditData = [pscustomobject]@{
                    MailboxOwnerUPN = $calendar
                    Item = [pscustomobject]@{ Subject = 'Staff meeting'; ParentFolder = [pscustomobject]@{ Path = '\Calendar' } }
                }
            }))
    ) -EventSubject 'Staff meeting' -EventCreated '2026-09-01T15:00:02Z' -CalendarEmail $calendar)
Assert-Equal $ranked.Count 1 'only the matching Create on this mailbox is kept'
Assert-Equal $ranked[0].userId 'delegate@example.com' 'matched user'

Assert-Equal (ConvertTo-TimeBoundary -Value '2026-09-01' -Which start) '2026-09-01T00:00:00.000Z' 'date-only start'
Assert-Equal (ConvertTo-TimeBoundary -Value '2026-09-01' -Which end) '2026-09-01T23:59:59.000Z' 'date-only end'
$graphDate = ConvertFrom-GraphDateTime ([pscustomobject]@{ dateTime = '2026-09-23T15:00:00.0000000'; timeZone = 'UTC' })
Assert-Equal $graphDate.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss') '2026-09-23T15:00:00' 'Graph UTC dateTime without an offset'

$now = [datetimeoffset]::Parse('2026-09-23T00:00:00Z')
$auditWindow = Get-AuditQueryWindow -WindowStart '2020-01-01T00:00:00.000Z' -WindowEnd '2027-01-01T00:00:00.000Z' -Now $now
Assert-True $auditWindow.TruncatedFuture 'future audit end is clamped'
Assert-True $auditWindow.TruncatedRetention 'old audit start is clamped'
Assert-Equal $auditWindow.End $now.ToString('yyyy-MM-ddTHH:mm:ss.fffZ') 'audit end is now'

$envFile = New-TemporaryFile
Set-Content -LiteralPath $envFile -Value "export SAMPLE_TENANT_ID=`"from-file`"`nSAMPLE_SECRET='s'"
$env:SAMPLE_TENANT_ID = 'already-set'
Import-DotEnvFile $envFile
Assert-Equal $env:SAMPLE_TENANT_ID 'already-set' 'dotenv does not override the environment'
Assert-Equal $env:SAMPLE_SECRET 's' 'dotenv loads a missing key'
Remove-Item Env:SAMPLE_TENANT_ID, Env:SAMPLE_SECRET -ErrorAction SilentlyContinue
Remove-Item $envFile

$extended = ConvertTo-ExtendedProperties @(
    [pscustomobject]@{ id = 'String 0x3ff8'; value = 'Lance Smith' }
    [pscustomobject]@{ id = 'String 0x3FFA'; value = 'Lance Smith' }
)
Assert-Equal $extended.ByName['PidTagCreatorName'] 'Lance Smith' 'property id case is ignored'

if ($failures.Count) {
    $failures | ForEach-Object { [Console]::Error.WriteLine($_) }
    exit 1
}
Write-Output "PowerShell tests passed ($((Get-Command Test-SubjectMatch).Name) and helpers)."
