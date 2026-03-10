<#
.SYNOPSIS
    Generate daily report from template and dailywork data.
#>

param (
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$WorkerName,

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$InputFile
)

if ([string]::IsNullOrWhiteSpace($WorkerName) -or [string]::IsNullOrWhiteSpace($InputFile)) {
    $scriptName = $MyInvocation.MyCommand.Name
    Write-Host "사용법: .\$scriptName <야간근무자이름> <dailywork.txt>"
    Write-Host "  예시: .\$scriptName Benedict dailywork.txt"
    Exit
}

$ErrorActionPreference = "Stop"

$templatePath = "C:\Users\javarange\DailyReport\template.txt"
$dailyworkPath = Resolve-Path $InputFile -ErrorAction Stop

# 1. 실행일 기준 어제 날짜 구하기
$yesterday = (Get-Date).AddDays(-1)
$dateString = $yesterday.ToString("yyyy-MM-dd")

# 요일 구하기 (한글)
$culture = New-Object System.Globalization.CultureInfo("ko-KR")
$dayOfWeek = $culture.DateTimeFormat.GetAbbreviatedDayName($yesterday.DayOfWeek)

# 2. 야간근무자 이름 (첫 번째 명령줄 인자로 받음)
$headerLine = "$dateString ($dayOfWeek) 야간근무자 - IDC운영유닛 $WorkerName"

# 3. dailywork.txt 데이터 읽기 및 파싱
# 데이터 양식: YYYY-MM-DD HH:MM 근무자 데이터 센터 그룹 서비스 회원아이디 아이피 처리내용 [작업]
$workItems = @()
if (Test-Path $dailyworkPath) {
    # 탭 구분자로 파싱. 인코딩 명시
    $lines = Get-Content $dailyworkPath -Encoding UTF8
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        
        $cols = $line -split "`t"
        if ($cols.Count -ge 8) {
            $dateTimeStr = $cols[0].Trim()
            $worker = $cols[1].Trim()
            $center = $cols[2].Trim()
            $group = $cols[3].Trim()
            $service = $cols[4].Trim()
            $memberId = $cols[5].Trim()
            $ip = $cols[6].Trim()
            $content = $cols[7].Trim()
            
            # 시간 추출 (HH:MM)
            $timePart = ""
            if ($dateTimeStr -match "\d{4}-\d{2}-\d{2}\s+(\d{2}:\d{2})") {
                $timePart = $matches[1]
            }

            # 정렬을 위한 가상 시간 계산 (00:00~08:00은 다음날로 간주하여 정렬 우선순위 뒤로)
            # 00~08시를 24~32시로 변환하여 수치 정렬에 사용
            $sortTimeValue = 0
            if ($timePart -match "^(\d{2}):(\d{2})$") {
                $hour = [int]$matches[1]
                $minute = [int]$matches[2]
                
                if ($hour -lt 8) {
                    $hour += 24
                }
                
                $sortTimeValue = ($hour * 60) + $minute
            }

            $workItems += [PSCustomObject]@{
                TimeStr   = $timePart
                Center    = $center
                Service   = $service
                MemberId  = $memberId
                IP        = $ip
                Content   = $content
                SortValue = $sortTimeValue
            }
        }
    }
}

# 4. 분류 및 정렬
# "하이웍스" -> 4번
# 나머지 -> 2번
$serviceGroup4 = @("하이웍스")

# 데이터센터(Center) 목록 추출 및 각각 파일 생성
$dataCenters = $workItems | Select-Object -ExpandProperty Center -Unique

foreach ($dc in $dataCenters) {
    if ([string]::IsNullOrWhiteSpace($dc)) { continue }
    
    # 출력 파일명 생성 (dailyreport_데이터센터_YYYYMMDD.txt -> dailyreport_서초_YYYYMMDD.txt)
    $outputFileName = "dailyreport_$($dc)_$($yesterday.ToString('yyyyMMdd')).txt"
    $outputPath = "C:\Users\javarange\DailyReport\$outputFileName"
    
    $dcItems = $workItems | Where-Object { $_.Center -eq $dc }
    
    $itemsForGroup2 = $dcItems | Where-Object { $serviceGroup4 -notcontains $_.Service } | Sort-Object SortValue
    
    # 4번 그룹 항목 추출 (기본)
    $group4Raw = @($dcItems | Where-Object { $serviceGroup4 -contains $_.Service })
    
    # 데이터센터별 고정 수동 항목 객체로 추가 (정렬을 위해)
    if ($dc -match "가산") {
        $group4Raw += [PSCustomObject]@{
            TimeStr   = "04:10"
            Content   = "INTMON ARP 체크 확인 - 가산U+ 정상"
            SortValue = (4 + 24) * 60 + 10 # 28 * 60 + 10 = 1690
        }
        $group4Raw += [PSCustomObject]@{
            TimeStr   = "07:10"
            Content   = "INTMON ARP 체크 확인 - 과천KINX 정상/가산KINX 정상"
            SortValue = (7 + 24) * 60 + 10 # 31 * 60 + 10 = 1870
        }
    }
    elseif ($dc -match "서초") {
        $group4Raw += [PSCustomObject]@{
            TimeStr   = "07:10"
            Content   = "INTMON ARP 체크 확인 - 서초U+/강남 정상"
            SortValue = (7 + 24) * 60 + 10 # 31 * 60 + 10 = 1870
        }
    }

    # 최종 병합 후 시간순 정렬
    $itemsForGroup4 = $group4Raw | Sort-Object SortValue

    # 5. template 텍스트 치환하여 결과물 생성
    $outputContent = @()
    if (Test-Path $templatePath) {
        $templateLines = Get-Content $templatePath -Encoding UTF8
        
        $inGroup2 = $false
        $inGroup4 = $false
        $inGroup6 = $false
        
        foreach ($line in $templateLines) {
            # 첫 번째 라인 (날짜 근무자) 교체
            if ($line -match "^\d{4}-\d{2}-\d{2}.+야간근무자") {
                $outputContent += $headerLine
                continue
            }
            
            # 2) 작업내역 위치 찾기
            if ($line -match "^2\)\s*작업내역") {
                $outputContent += $line
                $inGroup2 = $true
                # 다음 라인부터는 기존 데이터 무시하고 생성된 데이터 삽입
                
                if ($itemsForGroup2.Count -eq 0) {
                    $outputContent += "- 없음"
                }
                else {
                    foreach ($item in $itemsForGroup2) {
                        $outputContent += "- [$($item.TimeStr)] $($item.MemberId) / $($item.IP) - $($item.Content)"
                    }
                }
                continue
            }
            
            # 3)장애처리 를 만나면 2번 그룹 주입 종료
            if ($line -match "^3\)") {
                $outputContent += "" # 3번 항목 시작 전 공백
                $inGroup2 = $false
            }
            
            # 4) 83/가비아시스템 위치 찾기
            if ($line -match "^4\)\s*83[/\s]*가비아시스템") {
                $outputContent += $line
                $inGroup4 = $true
                
                # 고정 라인 추가 (항상 같은 내용 - 시간 정렬 제외 항목)
                $outputContent += "- [20:00~08:00] VOC 장애관련 문의 : 0건"
                $outputContent += "- [20:00~08:00] gabia.com 로그인 및 도메인 검색 기능 정상 확인"
                
                foreach ($item in $itemsForGroup4) {
                    # 4번은 "처리내용" 만 출력
                    $outputContent += "- [$($item.TimeStr)] $($item.Content)"
                }
                
                continue
            }
            
            # 5) 공,클 수신메일 을 만나면 4번 그룹 주입 종료
            if ($line -match "^5\)") {
                $outputContent += "" # 5번 항목 시작 전 공백
                $inGroup4 = $false
            }

            # 6) 항목 위치 찾기
            if ($line -match "^6\)") {
                $inGroup6 = $true
                if ($dc -match "서초") {
                    $outputContent += "6) 카버코리아 육안점검 특이사항 없음 메일 발송 완료"
                    $outputContent += "   씨젠 육안점검 특이사항 없음 메일 발송 완료"
                }
                elseif ($dc -match "가산") {
                    $outputContent += "6) 캘러웨이골프 코리아 특이사항 "
                    $outputContent += "- 육안점검 메일 발송완료"
                }
                else {
                    $outputContent += $line
                    $inGroup6 = $false
                }
                continue
            }
            
            # 7) 항목 시작 시 6번 그룹 주입 종료
            if ($line -match "^7\)") {
                if ($inGroup6) {
                    $outputContent += "" # 6번 그룹에서 생략된 공백 보완
                }
                $inGroup6 = $false
            }
            
            # 그룹 2, 4, 6 주입 중에는 템플릿의 기존 라인 무시
            if (-not $inGroup2 -and -not $inGroup4 -and -not $inGroup6) {
                $outputContent += $line
            }
        }
    }
    else {
        Write-Warning "template.txt 파일을 찾을 수 없습니다."
        Exit
    }

    # 6. 결과 파일 저장 (UTF8 with BOM)
    $outputContent | Out-File -FilePath $outputPath -Encoding UTF8
    Write-Host "보고서 생성이 완료되었습니다: $outputPath"
}
