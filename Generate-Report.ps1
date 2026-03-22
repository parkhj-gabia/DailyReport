# Generate-Report.ps1
param (
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$WorkerName,

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$InputFile
)

$K_GS = "가산"
$K_SC = "서초"
$K_GN = "강남"
$K_GC = "과천"
$K_HI = "하이웍스"
$K_WH = "웹훅"
$K_NW = "야간근무자"
$K_OU = "운영유닛"
$K_N1 = "없음"

if ([string]::IsNullOrWhiteSpace($WorkerName) -or [string]::IsNullOrWhiteSpace($InputFile)) {
    Write-Host "Usage: .\Generate-Report.ps1 <WorkerName> <InputFile>"
    Exit
}

$dailyworkPath = "C:\Users\javarange\DailyReport\$InputFile"
$userListPath = "C:\Users\javarange\DailyReport\userlist.txt"
$templatePath = "C:\Users\javarange\DailyReport\template.txt"

$targetDCs = @()
$koreanNames = @()
$primaryWorkers = $WorkerName -split "," | ForEach-Object { $_.Trim() }

if (Test-Path $userListPath) {
    foreach ($line in (Get-Content $userListPath -Encoding UTF8)) {
        foreach ($pw in $primaryWorkers) {
            if ($line -match "^([^\(]+)\(.*$pw.*:\s*(.+)$") {
                $koreanNames += $matches[1].Trim()
                $dc = $matches[2].Trim()
                if ($targetDCs -notcontains $dc) { $targetDCs += $dc }
            }
        }
    }
}

if ($targetDCs.Count -eq 0) {
    Write-Host "Error: Worker '$WorkerName' not found in userlist.txt."
    Exit
}

function Get-NormalizeCenter($n) {
    if ($n -eq $K_GN -or $n -eq $K_SC) { return $K_SC }
    if ($n -eq $K_GC -or $n -eq $K_GS) { return $K_GS }
    return $n
}

$yesterday = (Get-Date).AddDays(-1)
$dateString = $yesterday.ToString("yyyy-MM-dd")
$dayOfWeekIdx = (Get-Date).AddDays(-1).DayOfWeek.value__
$daysKor = @("일", "월", "화", "수", "목", "금", "토")
$dayOfWeek = $daysKor[$dayOfWeekIdx]

if ($koreanNames.Count -eq 0) { Exit }

$targetDC = (Get-NormalizeCenter $targetDCs[0]).Trim()
$headerLine = "$dateString ($dayOfWeek) $K_NW - IDC$K_OU $WorkerName"

$workItems = @()
if (Test-Path $dailyworkPath) {
    foreach ($line in (Get-Content $dailyworkPath -Encoding UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $cols = $line -split "`t"
        if ($cols.Count -ge 8) {
            $w = $cols[1].Trim()
            $center = Get-NormalizeCenter $cols[2].Trim()
            
            $wList = $w -split "," | ForEach-Object { $_.Trim() }
            $found = $false
            foreach ($wSingle in $wList) {
                if ($koreanNames -contains $wSingle) { $found = $true; break }
            }
            if (-not $found) { continue }
            
            $timeVal = 0
            $timeStr = ""
            if ($cols[0] -match "\s(\d{2}:\d{2})") {
                $timeStr = $matches[1]
                $h = [int]$timeStr.Substring(0,2)
                $m = [int]$timeStr.Substring(3,2)
                if ($h -lt 8) { $h += 24 }
                $timeVal = ($h * 60) + $m
            }
            $workItems += [PSCustomObject]@{
                Worker   = $cols[1].Trim()
                TimeStr  = $timeStr
                MemberId = $cols[5].Trim()
                IP       = $cols[6].Trim()
                Content  = $cols[7].Trim()
                Service  = $cols[4].Trim()
                SortVal  = $timeVal
            }
        }
    }
}

$outputFileName = "dailyreport_$($targetDC)_$($yesterday.ToString('yyyyMMdd')).txt"
$outputPath = "C:\Users\javarange\DailyReport\$outputFileName"

$itemsG2 = $workItems | Where-Object { ($_.Service -notmatch $K_HI) -and ($_.Content -notmatch $K_WH) -and ($_.Content -notmatch "83") } | Sort-Object Worker, SortVal
$itemsG4 = @($workItems | Where-Object { ($_.Service -match $K_HI) -or ($_.Content -match $K_WH) -or ($_.Content -match "83") })
$itemsG4 = $itemsG4 | Sort-Object Worker, SortVal

$output = @()
if (Test-Path $templatePath) {
    $inG3 = $false; $inG4 = $false; $inG6 = $false
    $lineIdx = 0
    foreach ($line in (Get-Content $templatePath -Encoding UTF8)) {
        if ($lineIdx -eq 0) { $output += $headerLine; $lineIdx++; continue }
        $lineIdx++
        if ($line -match "^3\)") {
            $output += $line; $inG3 = $true
            if ($itemsG2.Count -eq 0) { $output += "- $K_N1" }
            else { foreach ($i in $itemsG2) { $output += "- [$($i.TimeStr)] $($i.MemberId) / $($i.IP) - $($i.Content)" } }
            continue
        }
        if ($line -match "^4\)") { $output += ""; $inG3 = $false }
        if ($line -match "^4\)") {
            $output += $line; $inG4 = $true
            if ($itemsG4.Count -eq 0) { $output += "- $K_N1" }
            else { foreach ($i in $itemsG4) { $output += "- [$($i.TimeStr)] $($i.Content)" } }
            continue
        }
        if ($line -match "^5\)") { $output += ""; $inG4 = $false }
        if ($line -match "^6\)") {
            $inG6 = $true
            $output += $line
            continue
        }
        if ($line -match "INTMON ARP" -or $line -match "육안점검") { continue }
        if ($inG6 -and $line.Trim() -eq "") { continue }

        if ($line -match "^감사" -or $line -match "^7\)") { 
            if ($inG6) { 
                if ($targetDC -eq $K_GS) {
                    $output += "- INTMON ARP 체크 확인 - 가산U+ 정상"
                    $output += "- INTMON ARP 체크 확인 - 과천KINX 정상/가산KINX 정상"
                } else {
                    $output += "- INTMON ARP 체크 확인 - 서초U+/강남 정상"
                }
                if ($targetDC -eq $K_SC) {
                    $output += "- 카버코리아 육안점검 특이사항 없음 메일 발송 완료"
                    $output += "- 씨젠 육안점검 특이사항 없음 메일 발송 완료"
                } elseif ($targetDC -eq $K_GS) {
                    $output += "- 캘러웨이골프 코리아 특이사항 "
                    $output += "  - 육안점검 메일 발송완료"
                }
                $output += ""
            }
            $inG6 = $false 
        }
        if (-not $inG3 -and -not $inG4) { $output += $line }
    }
}
$output | Out-File -FilePath $outputPath -Encoding UTF8
Write-Host "Success: $outputFileName"

