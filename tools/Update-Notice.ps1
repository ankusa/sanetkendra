[CmdletBinding()]
param(
    [string]$LatestJobsUrl = "https://www.sarkariresult.com/latestjob/",
    [string]$AdmissionUrl = "https://www.sarkariresult.com/admission/",
    [ValidateRange(1, 500)][int]$MaxItems = 500,
    [ValidateRange(4, 4)][int]$NoticeItems = 4,
    [string]$WebJobsUrl = "",
    [string]$SourceFile,
    [switch]$NoOpen
)

$ErrorActionPreference = "Stop"
$OutputFile = Join-Path $PSScriptRoot "SANET-KENDRA-NOTICE.html"
$CacheFile = Join-Path $PSScriptRoot "SANET-KENDRA-NOTICE-CACHE.html"
$LinkIndexFile = Join-Path $PSScriptRoot "JOB-LINKS.json"
$VacancyIndexFile = Join-Path $PSScriptRoot "JOB-VACANCIES.json"
$WebJobsDir = Join-Path $PSScriptRoot "WEB-JOBS"
$WebJobsFile = Join-Path $WebJobsDir "index.html"
$QrFile = Join-Path $PSScriptRoot "QR-JOBS.png"
$WebJobsUrlFile = Join-Path $PSScriptRoot "WEB-JOBS-URL.txt"
if ([string]::IsNullOrWhiteSpace($WebJobsUrl)) {
    if (Test-Path $WebJobsUrlFile) { $WebJobsUrl = (Get-Content $WebJobsUrlFile -Raw -Encoding UTF8).Trim() }
    if ([string]::IsNullOrWhiteSpace($WebJobsUrl)) { $WebJobsUrl = "https://sanetkendra.in/jobs/" }
}
$UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/151 Safari/537.36"

function Get-WebText([string]$Url) {
    $lastError = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -UserAgent $UserAgent -TimeoutSec 15
            if (-not [string]::IsNullOrWhiteSpace([string]$response.Content)) { return [string]$response.Content }
        }
        catch { $lastError = $_ }
        if ($attempt -lt 3) { Start-Sleep -Seconds 2 }
    }
    $curl = Get-Command curl.exe -ErrorAction SilentlyContinue
    if ($curl) {
        $text = & $curl.Source -L --fail --silent --show-error --max-time 20 --retry 2 --retry-delay 2 -A $UserAgent $Url
        $joined = ($text -join "`n")
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($joined)) { return $joined }
    }
    if ($lastError) { throw $lastError }
    throw "Website se data download nahi hua: $Url"
}

function Open-Notice([string]$Path) {
    $edgeCandidates = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { $_ -and (Test-Path $_) }
    if ($edgeCandidates.Count -gt 0) { Start-Process $edgeCandidates[0] -ArgumentList @("--new-window", $Path) } else { Start-Process $Path }
}

function Clean-Title([string]$Text) {
    $value = [regex]::Replace($Text, '<[^>]+>', ' ')
    $value = [System.Net.WebUtility]::HtmlDecode($value)
    return [regex]::Replace($value, '\s+', ' ').Trim()
}

function Convert-SourceHtml([string]$Html) {
    if ($Html -match 'class=["'']html-tag["'']') {
        for ($decodePass = 1; $decodePass -le 3; $decodePass++) {
            $Html = [regex]::Replace($Html, '(?is)<a\b[^>]*class=["''][^"'']*html-attribute-value[^"'']*["''][^>]*>(?<value>.*?)</a>', '${value}')
            $Html = [regex]::Replace($Html, '(?is)<span\b[^>]*>(.*?)</span>', '$1')
            $Html = [System.Net.WebUtility]::HtmlDecode($Html)
        }
    }
    return $Html
}

function Resolve-JobUrl([string]$Url) {
    if ([string]::IsNullOrWhiteSpace($Url)) { return "" }
    $value = [System.Net.WebUtility]::HtmlDecode($Url.Trim())
    try { return ([uri]::new([uri]$LatestJobsUrl, $value)).AbsoluteUri } catch { return $value }
}

function Get-EntryType([string]$Title) {
    if ($Title -match '(?i)Scholarship|छात्रवृत्ति') { return "छात्रवृत्ति" }
    if ($Title -match '(?i)Admission|Entrance|Counselling|Counseling|NEET|JEE|CUET|CTET|\bPET\b|\bTET\b|\bSTET\b|Eligibility\s+Test|UGC\s*NET|CAT\b|GATE\b|Exam\s+Form|University\s+Form|School\s+Admission') { return "परीक्षा/प्रवेश" }
    return "भर्ती"
}

function New-JobEntry([string]$Title, [string]$Url, [string]$LastDate = "", [string]$Vacancies = "") {
    return [pscustomobject]@{ Title = $Title; Url = (Resolve-JobUrl $Url); LastDate = $LastDate; Vacancies = $Vacancies; EntryType = (Get-EntryType $Title) }
}

function Get-AligarhAudienceScore($Entry) {
    # Every verified live item still goes to the QR web page; this score affects
    # only the four large printed cards at SANET KENDRA, Pala Fatak, Aligarh.
    $text = (([string]$Entry.Title) + ' ' + ([string]$Entry.Url)).ToLowerInvariant()
    $score = 0

    if ($Entry.EntryType -eq 'भर्ती') { $score += 1000 } else { $score -= 1000 }

    # Exact local opportunities first, then Uttar Pradesh, then all-India jobs.
    if ($text -match 'aligarh|amu\b') { $score += 10000 }
    if ($text -match 'uttar\s*pradesh|\bupsssc\b|\buppsc\b|\bupprpb\b|\bup\s+police\b|\bup\s+anganwadi\b|\bup\s+home\s*guard|\bup\s+seva') { $score += 5000 }
    if ($text -match '\bssc\b|\bupsc\b|railway|\brrb\b|india\s+post|postal|\bibps\b|state\s+bank|\bsbi\b|bank\s+of\s+india|central\s+bank|union\s+bank|indian\s+bank|army|navy|air\s*force|agniveer|coast\s+guard|cisf|crpf|bsf|itbp|ssb\b|assam\s+rifles|fci\b|lic\b|epfo\b|\baai\b|drdo|isro|barc\b|ntpc|ongc|gail\b|sail\b|indian\s+oil|iocl\b|bhel\b|bel\b|esic\b|kvs\b|nvs\b') { $score += 3500 }

    # Broad-entry roles attract more local students and job-seekers.
    if ($text -match '10th|matric|high\s*school|12th|intermediate|10\+2|iti\b|graduate|graduation') { $score += 1000 }
    if ($text -match 'group\s*[cd]|clerk|constable|mts\b|chsl\b|cgl\b|apprentice|assistant|technician|driver|helper|anganwadi|teacher|teaching|junior\s+assistant|data\s+entry|stenographer|auditor|accountant') { $score += 1000 }
    if ($text -match 'female|women|mahila') { $score += 120 }

    # Vacancy volume is a strong proxy for how many passers-by may be eligible.
    $vacancyNumber = 0
    if ([int]::TryParse(([string]$Entry.Vacancies).Replace(',', ''), [ref]$vacancyNumber)) {
        if ($vacancyNumber -ge 5000) { $score += 1200 }
        elseif ($vacancyNumber -ge 1000) { $score += 800 }
        elseif ($vacancyNumber -ge 200) { $score += 400 }
        elseif ($vacancyNumber -gt 0 -and $vacancyNumber -lt 50) { $score -= 500 }
    }

    # Highly specialised posts stay available on the QR page. They can still
    # appear on print when truly local, but broad-entry jobs normally rank above.
    if ($text -match 'professor|faculty|architect|scientist|research|specialist|medical\s+officer|dental|veterinary|law\s+officer') { $score -= 1500 }

    # State-bound vacancies outside UP remain on the QR page but normally do
    # not occupy one of the four printed cards.
    if ($text -match 'bihar|jharkhand|jssc\b|bpsc\b|rajasthan|rsmssb|rssb\b|madhya\s*pradesh|mpesb|mppsc|chhattisgarh|cgpsc|odisha|opsc\b|west\s+bengal|wbpsc|gujarat|gpsc\b|maharashtra|mpsc\b|karnataka|kerala|tamil\s*nadu|telangana|andhra\s*pradesh|haryana|hssc\b|punjab|psssb|uttarakhand|uksssc|himachal|hppsc|assam|arunachal|tripura|manipur|mizoram|nagaland|meghalaya') { $score -= 5000 }

    return $score
}

function Get-NoticeAudienceBucket($Entry) {
    $text = (([string]$Entry.Title) + ' ' + ([string]$Entry.Url)).ToLowerInvariant()
    if ($text -match 'aligarh|amu\b') { return 'ALIGARH' }
    if ($text -match 'uttar\s*pradesh|\bupsssc\b|\buppsc\b|\bupprpb\b|\bup\s+police\b|\bup\s+anganwadi\b|\bup\s+home\s*guard|\bup\s+seva') { return 'UP' }
    if ($text -match '\bssc\b|\bupsc\b|railway|\brrb\b|india\s+post|postal|\bibps\b|state\s+bank|\bsbi\b|bank\s+of\s+india|central\s+bank|union\s+bank|indian\s+bank|army|navy|air\s*force|agniveer|coast\s+guard|cisf|crpf|bsf|itbp|ssb\b|assam\s+rifles|fci\b|lic\b|epfo\b|\baai\b|drdo|isro|barc\b|ntpc|ongc|gail\b|sail\b|indian\s+oil|iocl\b|bhel\b|bel\b|esic\b|kvs\b|nvs\b') { return 'INDIA' }
    if ($text -match 'bihar|jharkhand|jssc\b|bpsc\b|rajasthan|rsmssb|rssb\b|madhya\s*pradesh|mpesb|mppsc|chhattisgarh|cgpsc|odisha|opsc\b|west\s+bengal|wbpsc|gujarat|gpsc\b|maharashtra|mpsc\b|karnataka|kerala|tamil\s*nadu|telangana|andhra\s*pradesh|haryana|hssc\b|punjab|psssb|uttarakhand|uksssc|himachal|hppsc|assam|arunachal|tripura|manipur|mizoram|nagaland|meghalaya') { return 'OTHER-STATE' }
    return 'OTHER'
}

function ConvertTo-CardHtml($Entry) {
    $title = [System.Net.WebUtility]::HtmlEncode([string]$Entry.Title)
    $lastDate = [System.Net.WebUtility]::HtmlEncode([string]$Entry.LastDate)
    $vacancies = [System.Net.WebUtility]::HtmlEncode([string]$Entry.Vacancies)
    $entryType = [System.Net.WebUtility]::HtmlEncode([string]$Entry.EntryType)
    $url = [System.Net.WebUtility]::HtmlEncode([string]$Entry.Url)
    $meta = @()
    if ($entryType) { $meta += $entryType }
    if ($vacancies -and $Entry.EntryType -eq 'भर्ती') { $meta += "कुल पद: $vacancies" }
    if ($lastDate) { $meta += "अंतिम तिथि: $lastDate" }
    $text = if ($meta.Count -gt 0) { "$title <span class=`"last-date`">$($meta -join ' | ')</span>" } else { $title }
    if ($url) { return "<li><a href=`"$url`" target=`"_blank`" rel=`"noopener noreferrer`" title=`"Job page खोलें`">$text</a></li>" }
    return "<li>$text</li>"
}

function Read-LinkIndex {
    $index = @{}
    if (Test-Path $LinkIndexFile) {
        $saved = Get-Content $LinkIndexFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($property in $saved.PSObject.Properties) { $index[$property.Name] = [string]$property.Value }
    }
    return $index
}

function Save-LinkIndex($Entries) {
    $index = Read-LinkIndex
    foreach ($entry in @($Entries)) {
        if ($entry.Title -and $entry.Url) { $index[[string]$entry.Title] = [string]$entry.Url }
    }
    $json = $index | ConvertTo-Json -Depth 3
    [System.IO.File]::WriteAllText($LinkIndexFile, $json, [System.Text.UTF8Encoding]::new($true))
}

function Get-VacancyCount([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $plain = Clean-Title $Text
    $patterns = @(
        '(?i)\b(?:total\s+(?:no\.?\s+of\s+)?(?:vacanc(?:y|ies)|posts?|seats?|positions?)|vacancy\s+details\s*[:|-]?\s*total|number\s+of\s+(?:vacanc(?:y|ies)|posts?))\s*[:=-]?\s*(?<count>\d[\d,]{0,8})\b',
        '(?i)\b(?:for|of)\s+(?<count>\d[\d,]{0,8})\s+(?:vacanc(?:y|ies)|posts?|seats?|positions?)\b'
    )
    foreach ($pattern in $patterns) {
        $match = [regex]::Match($plain, $pattern)
        if ($match.Success) {
            $number = 0
            if ([int]::TryParse($match.Groups['count'].Value.Replace(',', ''), [ref]$number) -and $number -gt 0 -and $number -le 10000000) {
                return $number.ToString('N0', [System.Globalization.CultureInfo]::GetCultureInfo('en-IN'))
            }
        }
    }
    return ""
}

function Get-VerifiedJobInfo([string]$Html) {
    $decoded = Convert-SourceHtml $Html
    $plain = Clean-Title $decoded
    $datePatterns = @(
        '(?i)\bRe\s*-?\s*Open(?:ed)?\s+(?:Apply\s+Online|Online\s+Application|Registration).{0,160}?\d{1,2}[/-]\d{1,2}[/-]\d{4}\s*(?:to|till|through|[-–—])\s*(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\b(?:Application|Registration)\s+(?:Re\s*-?\s*Open(?:ed)?|Window\s+Re\s*-?\s*Open(?:ed)?).{0,160}?\d{1,2}[/-]\d{1,2}[/-]\d{4}\s*(?:to|till|through|[-–—])\s*(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\b(?:Apply\s+Online|Online\s+Application|Application\s+Form|Registration|Admission).{0,160}?\d{1,2}[/-]\d{1,2}[/-]\d{4}\s*(?:to|till|through|[-–—])\s*(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\b(?:Last\s+Date\s+Extended|Extended\s+Last\s+Date|New\s+Last\s+Date|Re\s*-?\s*Open\s+Last\s+Date).{0,120}?(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\b(?:Apply\s+Online|Application|Registration)\s+(?:Re\s*-?\s*Start(?:ed)?|Restart(?:ed)?).{0,160}?\d{1,2}[/-]\d{1,2}[/-]\d{4}\s*(?:to|till|through|[-–—])\s*(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\bRe\s*-?\s*Open(?:ed)?\s+(?:Apply\s+Online|Online\s+Application|Registration).{0,120}?(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\bLast\s+Date\s+for\s+Apply\s+Online.{0,120}?(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\bApply\s+Online\s+Last\s+Date.{0,120}?(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\bOnline\s+Application\s+Last\s+Date.{0,120}?(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\bRegistration\s+Last\s+Date.{0,120}?(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\bLast\s+Date\s+for\s+(?:Online\s+)?Registration.{0,120}?(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\bLast\s+Date\s+for\s+(?:Online\s+)?(?:Form|Admission).{0,120}?(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\bLast\s+Date\s+to\s+(?:Apply|Register)(?:\s+Online)?.{0,120}?(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})',
        '(?i)\bClosing\s+Date.{0,120}?(?<date>\d{1,2}[/-]\d{1,2}[/-]\d{4})'
    )
    $dates = [System.Collections.Generic.List[datetime]]::new()
    foreach ($pattern in $datePatterns) {
        foreach ($match in [regex]::Matches($plain, $pattern)) {
            $value = $match.Groups['date'].Value
            $parts = [regex]::Match($value, '^(?<d>\d{1,2})[/-](?<m>\d{1,2})[/-](?<y>\d{4})$')
            try {
                $dates.Add((Get-Date -Year ([int]$parts.Groups['y'].Value) -Month ([int]$parts.Groups['m'].Value) -Day ([int]$parts.Groups['d'].Value) -Hour 23 -Minute 59 -Second 59))
            } catch { }
        }
    }
    $verifiedDate = if ($dates.Count -gt 0) { @($dates | Sort-Object -Descending)[0] } else { $null }
    return [pscustomobject]@{
        LastDate = $verifiedDate
        Vacancies = (Get-VacancyCount $plain)
    }
}

function Get-JobPageText([string]$Url) {
    try {
        $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -UserAgent $UserAgent -TimeoutSec 10
        if (-not [string]::IsNullOrWhiteSpace([string]$response.Content)) { return [string]$response.Content }
    } catch { $firstError = $_ }
    $curl = Get-Command curl.exe -ErrorAction SilentlyContinue
    if ($curl) {
        $text = & $curl.Source -L --fail --silent --show-error --max-time 12 -A $UserAgent $Url
        $joined = ($text -join "`n")
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($joined)) { return $joined }
    }
    if ($firstError) { throw $firstError }
    throw "Job page download nahi hua: $Url"
}

function Get-VerifiedLiveJobs($Entries) {
    $index = @{}
    if (Test-Path $VacancyIndexFile) {
        $saved = Get-Content $VacancyIndexFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($property in $saved.PSObject.Properties) { $index[$property.Name] = [string]$property.Value }
    }
    $verified = [System.Collections.Generic.List[object]]::new()
    $today = (Get-Date).Date
    foreach ($entry in @($Entries)) {
        try {
            Write-Host "Last date verify ($($entry.EntryType)): $($entry.Title)" -ForegroundColor DarkCyan
            $detailHtml = Get-JobPageText $entry.Url
            $info = Get-VerifiedJobInfo $detailHtml
            if (-not $info.LastDate) {
                Write-Host "SKIP (job page par exact last date nahi mili): $($entry.Title)" -ForegroundColor Yellow
                continue
            }
            if ($info.LastDate.Date -lt $today) {
                Write-Host "SKIP (last date over): $($entry.Title)" -ForegroundColor Yellow
                continue
            }
            $entry.LastDate = $info.LastDate.ToString('dd-MM-yyyy')
            if ($entry.EntryType -eq 'भर्ती') {
                if ($info.Vacancies) { $entry.Vacancies = $info.Vacancies }
                elseif (-not $entry.Vacancies -and $index.ContainsKey($entry.Url)) { $entry.Vacancies = $index[$entry.Url] }
                if ($entry.Vacancies) { $index[$entry.Url] = $entry.Vacancies }
            }
            else { $entry.Vacancies = "" }
            $verified.Add($entry)
        }
        catch {
            Write-Host "SKIP (job page verify nahi hua): $($entry.Title)" -ForegroundColor Yellow
        }
    }
    $json = $index | ConvertTo-Json -Depth 3
    [System.IO.File]::WriteAllText($VacancyIndexFile, $json, [System.Text.UTF8Encoding]::new($true))
    return @($verified)
}

function Repair-NoticeLinks([string]$Path) {
    $index = Read-LinkIndex
    if ($index.Count -eq 0) { return }
    $html = [System.IO.File]::ReadAllText($Path)
    $evaluator = [System.Text.RegularExpressions.MatchEvaluator]{
        param($match)
        $title = Clean-Title $match.Groups['title'].Value
        if ($index.ContainsKey($title)) { return (ConvertTo-CardHtml (New-JobEntry $title $index[$title])) }
        return $match.Value
    }
    $html = [regex]::Replace($html, '(?is)<li>(?<title>[^<]+)</li>', $evaluator)
    if ($html -notmatch '\.updates li a,\.featured li a') {
        $html = $html.Replace('</style>', '.updates li a,.featured li a{display:flex;align-items:center;width:100%;height:100%;color:inherit;text-decoration:none;overflow-wrap:anywhere}</style>')
    }
    [System.IO.File]::WriteAllText($Path, $html, [System.Text.UTF8Encoding]::new($true))
}

function Get-LiveLatestJobs([string]$Html) {
    $results = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    # Chrome ka "view-source:" save actual markup ko escaped spans mein rakhta hai.
    # Decode karne ke baad wahi parser local test file aur live page dono par chalta hai.
    if ($Html -match 'class=["'']html-tag["'']') {
        for ($decodePass = 1; $decodePass -le 3; $decodePass++) {
            $Html = [regex]::Replace($Html, '(?is)<a\b[^>]*class=["''][^"'']*html-attribute-value[^"'']*["''][^>]*>(?<value>.*?)</a>', '${value}')
            $Html = [regex]::Replace($Html, '(?is)<span\b[^>]*>(.*?)</span>', '$1')
            $Html = [System.Net.WebUtility]::HtmlDecode($Html)
        }
    }
    $pattern = '(?is)<a\b(?<attrs>[^>]*)>(?<title>[^<]*?(?:\bLast\s*Date\s*:|\bOnline\s+Form\b|\bApply\s+Online\b|\bAdmission\s+Form\b|\bEntrance\s+Exam\b)[^<]*)</a>'
    foreach ($match in [regex]::Matches($Html, $pattern)) {
        $fullTitle = Clean-Title $match.Groups['title'].Value
        $parts = [regex]::Match($fullTitle, '^(?<job>.*?)\s*\|?\s*Last\s*Date\s*:\s*(?<date>[^|]+)', 'IgnoreCase')
        $title = if ($parts.Success) { $parts.Groups['job'].Value.Trim(' ', '|', '-', ':') } else { $fullTitle.Trim(' ', '|', '-', ':') }
        $lastDateText = if ($parts.Success) { $parts.Groups['date'].Value.Trim() } else { "" }
        $dateMatch = [regex]::Match($lastDateText, '\b(?<d>\d{1,2})[/-](?<m>\d{1,2})[/-](?<y>\d{4})\b')
        if ($dateMatch.Success) {
            try {
                $lastDate = Get-Date -Year ([int]$dateMatch.Groups['y'].Value) -Month ([int]$dateMatch.Groups['m'].Value) -Day ([int]$dateMatch.Groups['d'].Value) -Hour 23 -Minute 59 -Second 59
                $lastDateText = $lastDate.ToString('dd-MM-yyyy')
            } catch { continue }
        }
        $hrefMatch = [regex]::Match($match.Groups['attrs'].Value, '(?is)\bhref\s*=\s*["''](?<url>[^"'']+)["'']')
        $url = if ($hrefMatch.Success) { $hrefMatch.Groups['url'].Value } else { "" }
        if ($title -match '(?i)^(Apply Online|Online Form|Admission Form|Entrance Exam)$') { continue }
        if ($title -and $url -and $seen.Add($title)) { $results.Add((New-JobEntry $title $url $lastDateText (Get-VacancyCount $title))) }
        if ($results.Count -ge $MaxItems) { break }
    }
    return @($results)
}

function ConvertTo-WebJobHtml($Entry) {
    $title = [System.Net.WebUtility]::HtmlEncode([string]$Entry.Title)
    $lastDate = [System.Net.WebUtility]::HtmlEncode([string]$Entry.LastDate)
    $vacancies = [System.Net.WebUtility]::HtmlEncode([string]$Entry.Vacancies)
    $entryType = [System.Net.WebUtility]::HtmlEncode([string]$Entry.EntryType)
    $url = [System.Net.WebUtility]::HtmlEncode([string]$Entry.Url)
    $searchText = [System.Net.WebUtility]::HtmlEncode(("$title $entryType $vacancies $lastDate").ToLowerInvariant())
    $meta = @()
    if ($entryType) { $meta += $entryType }
    if ($vacancies -and $Entry.EntryType -eq 'भर्ती') { $meta += "कुल पद: $vacancies" }
    if ($lastDate) { $meta += "अंतिम तिथि: $lastDate" }
    $metaHtml = [System.Net.WebUtility]::HtmlEncode(($meta -join ' • '))
    return "<article class=`"job`" data-search=`"$searchText`" data-type=`"$entryType`"><div><h2>$title</h2><p>$metaHtml</p></div><a class=`"view`" href=`"$url`" target=`"_blank`" rel=`"noopener noreferrer`">विवरण देखें</a></article>"
}

function New-WebJobsPage($Entries, [string]$Today) {
    $cards = @(foreach ($entry in @($Entries)) { ConvertTo-WebJobHtml $entry })
    $count = @($Entries).Count
    return @"
<!doctype html><html lang="hi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>SANET KENDRA - Live Jobs & Exams</title><meta name="description" content="SANET KENDRA पर आज की live सरकारी jobs, exams, admission और scholarship forms.">
<style>*{box-sizing:border-box}body{margin:0;background:#f5f6f8;color:#151515;font-family:"Nirmala UI","Segoe UI",Arial,sans-serif}.wrap{max-width:900px;margin:auto;padding:18px}.hero{background:#fff;border:1px solid #ddd;border-radius:14px;padding:18px;margin-bottom:14px}.hero h1{margin:0 0 4px;font-size:28px}.hero p{margin:4px 0;color:#555}.tools{position:sticky;top:0;background:#f5f6f8;padding:8px 0 12px;z-index:2}.search{width:100%;font-size:18px;padding:13px 14px;border:1px solid #bbb;border-radius:10px;background:#fff}.filters{display:flex;gap:8px;overflow:auto;padding-top:9px}.filters button{white-space:nowrap;border:1px solid #bbb;background:#fff;border-radius:999px;padding:9px 13px;font-weight:700}.filters button.active{background:#111;color:#fff;border-color:#111}.list{display:grid;gap:10px}.job{display:flex;gap:14px;justify-content:space-between;align-items:center;background:#fff;border:1px solid #ddd;border-radius:12px;padding:14px}.job h2{font-size:19px;line-height:1.25;margin:0 0 6px}.job p{margin:0;color:#7a0000;font-weight:800}.view{flex:none;text-decoration:none;background:#0b57d0;color:#fff;padding:10px 12px;border-radius:8px;font-weight:800}.empty{display:none;text-align:center;padding:35px;background:#fff;border-radius:12px}.foot{text-align:center;color:#666;font-size:13px;padding:20px 4px}@media(max-width:560px){.wrap{padding:10px}.hero h1{font-size:24px}.job{align-items:flex-start;flex-direction:column}.view{width:100%;text-align:center}.job h2{font-size:18px}}</style></head><body><main class="wrap"><section class="hero"><h1>📢 SANET KENDRA</h1><p><b>Live Jobs, Exams, Admission & Scholarship Forms</b></p><p>अपडेट: $Today • कुल $count चालू फॉर्म</p></section><section class="tools"><input id="q" class="search" type="search" placeholder="Job / Exam खोजें…" aria-label="Job या exam खोजें"><div class="filters"><button class="active" data-filter="all">सभी</button><button data-filter="भर्ती">भर्ती</button><button data-filter="परीक्षा/प्रवेश">परीक्षा/प्रवेश</button><button data-filter="छात्रवृत्ति">छात्रवृत्ति</button></div></section><section id="list" class="list">$($cards -join "`r`n")</section><div id="empty" class="empty">कोई matching form नहीं मिला।</div><div class="foot">आवेदन से पहले official notification अवश्य जाँचें।</div></main><script>(()=>{const q=document.getElementById('q'),jobs=[...document.querySelectorAll('.job')],empty=document.getElementById('empty');let filter='all';function apply(){const text=q.value.trim().toLowerCase();let shown=0;jobs.forEach(j=>{const okText=!text||j.dataset.search.includes(text);const okType=filter==='all'||j.dataset.type===filter;const show=okText&&okType;j.style.display=show?'flex':'none';if(show)shown++});empty.style.display=shown?'none':'block'}q.addEventListener('input',apply);document.querySelectorAll('[data-filter]').forEach(b=>b.addEventListener('click',()=>{document.querySelectorAll('[data-filter]').forEach(x=>x.classList.remove('active'));b.classList.add('active');filter=b.dataset.filter;apply()}));})();</script></body></html>
"@
}

function Update-QrCode([string]$Url) {
    try {
        $encoded = [uri]::EscapeDataString($Url)
        $qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=600x600&margin=10&data=$encoded"
        Invoke-WebRequest -Uri $qrUrl -UseBasicParsing -OutFile $QrFile -TimeoutSec 15
        return (Test-Path $QrFile)
    } catch {
        Write-Host "QR image download nahi hui; custom URL text phir bhi notice par rahega." -ForegroundColor Yellow
        return $false
    }
}

try {
    $sourceLabel = "SarkariResult.com Latest Jobs + Admission"
    $titles = @()
    if ($SourceFile) {
        $latestHtml = [System.IO.File]::ReadAllText((Resolve-Path $SourceFile))
        $sourceLabel = "Local test source"
        $titles = @(Get-LiveLatestJobs $latestHtml)
    }
    else {
        $allEntries = [System.Collections.Generic.List[object]]::new()
        foreach ($source in @(
            [pscustomobject]@{ Name = 'Latest Jobs'; Url = $LatestJobsUrl; Type = '' },
            [pscustomobject]@{ Name = 'Entrance / Admission Exams'; Url = $AdmissionUrl; Type = 'परीक्षा/प्रवेश' }
        )) {
            try {
                Write-Host "$($source.Name) page se live forms liye ja rahe hain..." -ForegroundColor Cyan
                $sourceHtml = Get-WebText $source.Url
                foreach ($entry in @(Get-LiveLatestJobs $sourceHtml)) {
                    if ($source.Type -and $entry.EntryType -eq 'भर्ती') { $entry.EntryType = $source.Type }
                    $allEntries.Add($entry)
                }
            }
            catch {
                Write-Host "$($source.Name) page abhi uplabdh nahi hai." -ForegroundColor Yellow
                Write-Host $_.Exception.Message -ForegroundColor DarkYellow
            }
        }
        $titles = @($allEntries | Group-Object Url | ForEach-Object { $_.Group[0] } | Select-Object -First $MaxItems)
    }
    if ($titles.Count -eq 0) {
        $savedNotice = if (Test-Path $CacheFile) { $CacheFile } elseif (Test-Path $OutputFile) { $OutputFile } else { $null }
        if ($savedNotice) {
            if ($savedNotice -ne $OutputFile) { Copy-Item $savedNotice $OutputFile -Force }
            Repair-NoticeLinks $OutputFile
            if (-not $NoOpen) { Open-Notice $OutputFile }
            Write-Host "Website timeout: pichhla saved notice khola gaya." -ForegroundColor Yellow
            exit 0
        }
        throw "Live entries nahi mili aur saved notice bhi nahi hai."
    }
    if (-not $SourceFile) { $titles = @(Get-VerifiedLiveJobs $titles) }
    if ($titles.Count -eq 0) { throw "Job pages verify karne ke baad koi live form nahi mila." }

    $today = Get-Date -Format "dd-MM-yyyy"
    $count = $titles.Count
    Save-LinkIndex $titles
    New-Item -ItemType Directory -Force -Path $WebJobsDir | Out-Null
    $webHtml = New-WebJobsPage $titles $today
    [System.IO.File]::WriteAllText($WebJobsFile, $webHtml, [System.Text.UTF8Encoding]::new($true))
    $qrReady = Update-QrCode $WebJobsUrl

    # Notice board: normally 2 popular All-India + 2 Aligarh/UP recruitments.
    # If one group has fewer live forms, the best eligible job fills the space.
    # The QR web page above still contains every verified live job/exam/form.
    $sortedRecruitments = @($titles | Where-Object { $_.EntryType -eq 'भर्ती' } | Sort-Object @{Expression={ Get-AligarhAudienceScore $_ }; Descending=$true}, @{Expression={
        try { [datetime]::ParseExact($_.LastDate,'dd-MM-yyyy',[System.Globalization.CultureInfo]::InvariantCulture) } catch { [datetime]::MaxValue }
    }}, Title)
    $indiaJobs = @($sortedRecruitments | Where-Object { (Get-NoticeAudienceBucket $_) -eq 'INDIA' } | Select-Object -First 2)
    $upJobs = @($sortedRecruitments | Where-Object { (Get-NoticeAudienceBucket $_) -in @('ALIGARH','UP') } | Select-Object -First 2)
    $noticeList = [System.Collections.Generic.List[object]]::new()
    for ($slot = 0; $slot -lt 2; $slot++) {
        if ($slot -lt $indiaJobs.Count) { $noticeList.Add($indiaJobs[$slot]) }
        if ($slot -lt $upJobs.Count) { $noticeList.Add($upJobs[$slot]) }
    }
    foreach ($entry in $sortedRecruitments) {
        if ($noticeList.Count -ge $NoticeItems) { break }
        if ((Get-NoticeAudienceBucket $entry) -eq 'OTHER-STATE') { continue }
        if (-not @($noticeList | Where-Object { $_.Url -eq $entry.Url }).Count) { $noticeList.Add($entry) }
    }
    $noticeEntries = @($noticeList | Select-Object -First $NoticeItems)
    foreach ($entry in $noticeEntries) {
        Write-Host "NOTICE [$((Get-NoticeAudienceBucket $entry))]: $($entry.Title)" -ForegroundColor Green
    }
    $noticeCards = @(foreach ($entry in $noticeEntries) { ConvertTo-CardHtml $entry })
    $qrHtml = if ($qrReady) { '<img class="qr" src="QR-JOBS.png" alt="सभी jobs और exams खोलने का QR code">' } else { '<div class="qr-fallback">QR</div>' }
    $safeWebUrl = [System.Net.WebUtility]::HtmlEncode($WebJobsUrl)

    $html = @"
<!doctype html><html lang="hi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>SANET KENDRA Notice - $today</title>
<style>:root{--item-font:28pt}@page{size:A4 landscape;margin:5mm}*{box-sizing:border-box}html,body{margin:0;padding:0;background:#e9edf3;color:#000}body{font-family:"Nirmala UI","Segoe UI",Arial,sans-serif}.toolbar{width:287mm;margin:8px auto;padding:9px 14px;background:#fff;font:14px "Segoe UI",Arial,sans-serif}.toolbar button{padding:10px 17px;border:0;border-radius:5px;background:#0b57d0;color:#fff;font-weight:800}.sheet{width:287mm;height:200mm;margin:auto;background:#fff;padding:3mm 4mm;display:grid;grid-template-columns:minmax(0,1fr) 43mm;grid-template-rows:auto minmax(0,1fr) auto;gap:2mm;overflow:hidden}.head{grid-column:1/-1;display:flex;justify-content:space-between;align-items:end;border-bottom:2px solid #000;padding-bottom:1.4mm}.head h1{font-size:24pt;line-height:1;margin:0}.head .date{font-size:10.5pt;font-weight:800;white-space:nowrap}.content{min-width:0;display:grid;grid-template-rows:auto minmax(0,1fr);min-height:0}.section{font-size:14pt;font-weight:900;margin:0 0 1.4mm}.jobs{margin:0;padding:0;list-style:none;display:grid;grid-template-columns:repeat(2,minmax(0,1fr));grid-template-rows:repeat(2,minmax(0,1fr));gap:2.5mm;min-height:0}.jobs li{position:relative;padding:3.2mm;border:1.7px solid #222;display:flex;align-items:center;overflow:hidden;font-size:var(--item-font);font-weight:900;line-height:1.08;min-width:0;min-height:0}.jobs li a{display:flex;flex-direction:column;justify-content:center;width:100%;height:100%;color:#000;text-decoration:none;overflow-wrap:anywhere}.last-date{display:block;margin-top:1.6mm;font-size:.54em;line-height:1.13;color:#7a0000;font-weight:900}.side{border-left:1.5px solid #000;padding-left:2.2mm;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;overflow:hidden}.side h2{font-size:14pt;line-height:1.12;margin:0 0 2mm}.qr{width:35mm;height:35mm;object-fit:contain}.qr-fallback{width:35mm;height:35mm;border:3px solid #000;display:grid;place-items:center;font-size:22pt;font-weight:900}.scan{font-size:12.5pt;line-height:1.15;font-weight:900;margin:2mm 0}.url{display:none}.all{margin-top:1.5mm;font-size:9pt;line-height:1.18;font-weight:700}.foot{grid-column:1/-1;text-align:center;border-top:1px solid #777;padding-top:.8mm;font-size:8.5pt;white-space:nowrap}.font-status{float:right;font-weight:800;padding:8px}@media print{html,body{background:#fff;width:287mm;height:200mm}.toolbar{display:none}.sheet{width:287mm;height:200mm;margin:0;padding:3mm 4mm;break-inside:avoid;page-break-inside:avoid}}</style></head><body><div class="toolbar"><button onclick="fitNotice();window.print()">A4 Print</button><span id="fontStatus" class="font-status">Auto-fit</span></div><main class="sheet"><header class="head"><h1>📢 भारत की 4 लोकप्रिय सरकारी भर्तियाँ</h1><div class="date">SANET KENDRA • $today</div></header><section class="content"><div class="section">अलीगढ़ से आवेदन योग्य • All India + UP अवसर</div><ol class="jobs">$($noticeCards -join "`r`n")</ol></section><aside class="side"><h2>सभी Jobs, Exams और Forms</h2>$qrHtml<div class="scan">QR कोड स्कैन करें</div><div class="url">$safeWebUrl</div><div class="all">मोबाइल पर पूरी verified live list देखें।</div></aside><footer class="foot">SANET KENDRA • पाला फाटक, अलीगढ़ • आवेदन से पहले official notification अवश्य जाँचें।</footer></main><script>function fitNotice(){const cards=[...document.querySelectorAll('.jobs li')].slice(0,4),status=document.getElementById('fontStatus'),sizes=[];cards.forEach(card=>{const t=card.querySelector('a')||card;let lo=19,hi=40;while(hi-lo>.25){const m=(lo+hi)/2;card.style.fontSize=m+'pt';if(t.scrollHeight>t.clientHeight+1||t.scrollWidth>t.clientWidth+1)hi=m;else lo=m}const f=Math.floor(lo*4)/4;card.style.fontSize=f+'pt';sizes.push(f)});if(sizes.length)status.textContent='Auto-fit '+Math.min(...sizes)+'–'+Math.max(...sizes)+'pt'}window.addEventListener('load',()=>{document.fonts&&document.fonts.ready?document.fonts.ready.then(fitNotice):fitNotice()});window.addEventListener('beforeprint',fitNotice);window.addEventListener('resize',fitNotice);</script></body></html>
"@
    $html = $html.Replace('📢 भारत की 4 लोकप्रिय सरकारी भर्तियाँ', '📢 अभी आवेदन चालू — 4 प्रमुख सरकारी भर्तियाँ')
    if ($html -match '</script></body></html>') {
        $html = $html.Replace('</script></body></html>', "if(new URLSearchParams(location.search).has('autoprint')){window.addEventListener('load',()=>setTimeout(()=>window.print(),1200));}</script></body></html>")
    }
    $portraitPrintFix = @'
<style id="print-layout-v23">
@page{size:297mm 210mm;margin:5mm}
.toolbar{width:287mm}
.sheet{width:287mm;height:200mm;padding:3mm 4mm;grid-template-columns:minmax(0,1fr) 43mm;grid-template-rows:auto minmax(0,1fr) auto;gap:2mm}
.section{display:none}
.content{grid-template-rows:minmax(0,1fr)}
.jobs{height:100%;grid-template-columns:repeat(2,minmax(0,1fr));grid-template-rows:repeat(2,minmax(0,1fr));gap:3mm}
.jobs li{padding:4mm;font-size:30pt;line-height:1.08}.last-date{font-size:.52em}
.side{border-left:1.5px solid #000;border-top:0;padding:0 0 0 2.5mm;display:flex;text-align:center}
.side h2{font-size:14pt;margin:0 0 2mm}.qr,.qr-fallback{width:36mm;height:36mm}.scan{font-size:13pt;margin:2mm 0}.all{font-size:9.5pt;margin-top:1.5mm}.foot{grid-column:1/-1;font-size:9pt}
@media print{html,body{width:287mm;height:200mm}.sheet{width:287mm;height:200mm;padding:3mm 4mm;margin:0}}
</style>
'@
    $html = $html.Replace('</head>', "$portraitPrintFix</head>")
    [System.IO.File]::WriteAllText($OutputFile, $html, [System.Text.UTF8Encoding]::new($true))
    [System.IO.File]::WriteAllText($CacheFile, $html, [System.Text.UTF8Encoding]::new($true))
    if (-not $NoOpen) { Open-Notice $OutputFile }
    Write-Host "v20 India Popular + Aligarh Eligible Notice taiyar: $OutputFile" -ForegroundColor Green
    Write-Host "QR mobile web page: $WebJobsFile" -ForegroundColor Green
    Write-Host "Hosting par WEB-JOBS/index.html ko /jobs/ path par upload karein." -ForegroundColor Cyan
    Write-Host "QR target: $WebJobsUrl" -ForegroundColor Cyan
}
catch {
    Write-Host "ERROR: Notice nahi ban saka." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    exit 1
}
