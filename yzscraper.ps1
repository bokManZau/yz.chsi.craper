<#
.SYNOPSIS
  研招网硕士专业目录爬虫 — v2.3（详情翻页 + 研究方向前缀 + 合并单元格）
#>

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11
$USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
$BASE = "https://yz.chsi.com.cn"
$REQUEST_DELAY_SECONDS = 2.0

$script:DetailReferer = ""
$script:CookieHeader = ""

# ==========================  HTTP 请求  =================================

function Invoke-YZ {
    param([string]$Url, [string]$Method="GET", $Body = $null)

    $webReq = [System.Net.WebRequest]::Create($Url)
    $webReq.Method = $Method
    $webReq.UserAgent = $USER_AGENT
    $webReq.Timeout = 30000

    if ($Method -eq "POST") {
        $webReq.ContentType = "application/x-www-form-urlencoded"
        $webReq.Headers.Add("X-Requested-With", "XMLHttpRequest")
        $webReq.Accept = "application/json, text/javascript, */*; q=0.01"
    } else {
        $webReq.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
    }
    $webReq.Headers.Add("Accept-Language", "zh-CN,zh;q=0.9,en;q=0.8")
    $webReq.Headers.Add("Origin", $BASE)
    $webReq.Referer = if ($script:DetailReferer) { $script:DetailReferer } else { "$BASE/zsml/" }

    if ($script:CookieHeader) {
        $webReq.Headers.Add("Cookie", $script:CookieHeader)
    }

    if ($Body -and $Method -eq "POST") {
        $parts = @()
        foreach ($kv in $Body.GetEnumerator()) {
            $key = [System.Uri]::EscapeDataString($kv.Key)
            $val = $kv.Value
            if ($val -is [array]) {
                for ($i = 0; $i -lt $val.Count; $i++) {
                    $parts += "$([System.Uri]::EscapeDataString("$key[$i]"))=$([System.Uri]::EscapeDataString("$($val[$i])"))"
                }
            } else {
                $parts += "$key=$([System.Uri]::EscapeDataString("$val"))"
            }
        }
        $formBody = $parts -join "&"
        $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($formBody)
        $webReq.ContentLength = $bodyBytes.Length
        $webReq.GetRequestStream().Write($bodyBytes, 0, $bodyBytes.Length)
    }

    try {
        $resp = $webReq.GetResponse()
        $respStream = $resp.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($respStream)
        $content = $reader.ReadToEnd()
        $reader.Close()

        $setCookie = $resp.Headers["Set-Cookie"]
        if ($setCookie) {
            $cookieMap = @{}
            if ($script:CookieHeader) {
                $script:CookieHeader -split "; " | ForEach-Object {
                    $p = $_ -split "=", 2
                    if ($p.Count -eq 2) { $cookieMap[$p[0].Trim()] = $p[1].Trim() }
                }
            }
            $setCookie -split "," | ForEach-Object {
                $m = [regex]::Match($_, '([^=]+)=([^;]+)')
                if ($m.Success) { $cookieMap[$m.Groups[1].Value.Trim()] = $m.Groups[2].Value.Trim() }
            }
            $pairs = $cookieMap.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }
            $script:CookieHeader = $pairs -join "; "
        }

        Start-Sleep -Seconds $REQUEST_DELAY_SECONDS
        return $content
    } catch {
        $ex = $_.Exception
        if ($ex.Response) {
            Write-Host "  [✗] HTTP $($ex.Response.StatusCode.value__)" -ForegroundColor Red
            Write-Host "      $Url" -ForegroundColor Gray
            try {
                $errBody = [System.IO.StreamReader]::new($ex.Response.GetResponseStream()).ReadToEnd()
                if ($errBody.Length -lt 500) { Write-Host "      返回: $errBody" -ForegroundColor Yellow }
            } catch {}
        } else {
            Write-Host "  [✗] 请求失败: $($_.Exception.Message)" -ForegroundColor Red
        }
        return $null
    }
}

# ==========================  URL 解析  ==================================

function Parse-DetailUrl {
    param([string]$Url)
    if (-not $Url) { return $null }
    $result = @{}
    if ($Url -match 'zydm=(\d+)') { $result.zydm = $matches[1] }
    if ($Url -match 'zymc=([^&]+)') { $result.zymc = [System.Uri]::UnescapeDataString($matches[1]) }
    if ($Url -match 'mldm=([^&]+)') { $result.mldm = $matches[1] }
    if ($Url -match 'mlmc=([^&]+)') { $result.mlmc = [System.Uri]::UnescapeDataString($matches[1]) }
    if ($Url -match 'yjxkdm=([^&]+)') { $result.yjxkdm = $matches[1] }
    if ($Url -match 'yjxkmc=([^&]+)') { $result.yjxkmc = [System.Uri]::UnescapeDataString($matches[1]) }
    if ($Url -match 'xwlx=([^&]+)') { $result.xwlx = $matches[1] }
    if ($Url -match 'sign=([^&]+)') { $result.sign = $matches[1] }
    if ($result.zydm) { return $result }
    return $null
}

# ==========================  获取招生单位列表  ===========================

function Get-SchoolList {
    param([hashtable]$Params)

    $api = "$BASE/zsml/rs/zydws.do"
    Write-Host "  → 正在获取招生单位列表 ..." -ForegroundColor Yellow

    $allSchools = @()
    $seenCodes = @{}
    $curPage = 1
    $totalPages = 1
    $actualPageSize = 10
    $maxEmptyPages = 3
    $emptyPageCount = 0

    do {
        $start = ($curPage - 1) * $actualPageSize
        $body = @{
            zydm = $Params.zydm; zymc = $Params.zymc
            dwmc = ""; dwdm = ""
            ssdm = ""; xxfs = ""; dwlxs = @("all"); tydxs = ""; jsggjh = ""
            start = $start; curPage = $curPage; pageSize = $actualPageSize
        }

        $jsonText = Invoke-YZ -Url $api -Method POST -Body $body
        if (-not $jsonText) { break }

        try { $json = $jsonText | ConvertFrom-Json } catch {
            Write-Host "  [✗] 解析 JSON 失败" -ForegroundColor Red; break
        }
        if (-not $json.flag -or -not $json.msg -or -not $json.msg.list) { break }

        if ($json.msg.pageCount -and $json.msg.pageCount -gt 0) { $actualPageSize = $json.msg.pageCount }
        if ($json.msg.totalPage) { $totalPages = $json.msg.totalPage }

        $newCount = 0; $dupCount = 0
        foreach ($item in $json.msg.list) {
            if (-not $item.dwdm) { continue }
            if ($seenCodes.ContainsKey($item.dwdm)) { $dupCount++; continue }
            $seenCodes[$item.dwdm] = $true
            $allSchools += @{
                dwdm = $item.dwdm; dwmc = $item.dwmc; ssmc = $item.ssmc
                mxxfs = if ($item.mxxfs) { $item.mxxfs } else { "" }
                mdwlxs = if ($item.mdwlxs) { $item.mdwlxs } else { @("all") }
                mtydxs = if ($item.mtydxs) { $item.mtydxs } else { "" }
                mjsggjh = if ($item.mjsggjh) { $item.mjsggjh } else { "" }
            }
            $newCount++
        }

        Write-Host "    [页 $curPage] 新增 $newCount 所（跳过 $dupCount 重复）累计 $($allSchools.Count) 所" -ForegroundColor Gray

        if ($json.msg.list.Count -lt $actualPageSize) { break }
        if ($newCount -eq 0) {
            $emptyPageCount++
            if ($emptyPageCount -ge $maxEmptyPages) { break }
        } else { $emptyPageCount = 0 }
        $curPage++
        if ($curPage -gt $totalPages -and $totalPages -gt 0) { break }
    } while ($true)

    Write-Host "  [✓] 共获取 $($allSchools.Count) 个招生单位（唯一）" -ForegroundColor Green
    return $allSchools
}

# ==========================  学习方式映射  ===============================

function Get-XxfsText {
    param([string]$Val)
    switch ($Val) {
        "1" { return "全日制" }
        "2" { return "非全日制" }
        default { return $Val }
    }
}

# ==========================  获取研究方向详情（含翻页）=====================

function Get-Detail {
    param([hashtable]$Params, [hashtable]$School)

    $api = "$BASE/zsml/rs/yjfxs.do"
    $allList = @()
    $curPage = 1
    $totalPages = 1
    $pageSize = 10     # 与服务端默认一致
    $totalCount = 0

    do {
        $start = ($curPage - 1) * $pageSize
        $body = @{
            zydm = $Params.zydm; zymc = $Params.zymc
            dwdm = $School.dwdm; xxfs = $School.mxxfs
            dwlxs = $School.mdwlxs; tydxs = $School.mtydxs; jsggjh = $School.mjsggjh
            start = $start; curPage = $curPage; pageSize = $pageSize; totalCount = $totalCount
        }

        if ($curPage -gt 1) {
            Write-Host "    [详情翻页 $curPage] start=$start" -ForegroundColor DarkGray
        }

        $jsonText = Invoke-YZ -Url $api -Method POST -Body $body
        if (-not $jsonText) { break }

        try { $json = $jsonText | ConvertFrom-Json } catch { break }
        if (-not $json.flag -or -not $json.msg) { break }

        $list = $json.msg.list
        if (-not $list -or $list.Count -eq 0) { break }

        # 更新分页信息
        if ($json.msg.pageCount -and $json.msg.pageCount -gt 0) { $pageSize = $json.msg.pageCount }
        if ($json.msg.totalPage) { $totalPages = $json.msg.totalPage }
        if ($json.msg.totalCount) { $totalCount = $json.msg.totalCount }

        $allList += $list

        if ($list.Count -lt $pageSize) { break }
        $curPage++
        if ($curPage -gt $totalPages -and $totalPages -gt 0) { break }
    } while ($true)

    if ($allList.Count -eq 0) { return @() }

    $records = @()
    foreach ($item in $allList) {
        # ------ 拟招生人数 ------
        $nzsrs = ""
        if ($item.nzsrs -ne $null -and "$($item.nzsrs)" -ne "0") {
            $nzsrs = "$($item.nzsrs)"
        } elseif ($item.ssjstmrs -ne $null -and "$($item.ssjstmrs)" -ne "0") {
            $nzsrs = "$($item.ssjstmrs)"
        } elseif ($item.nzsrsstr -ne $null -and "$($item.nzsrsstr)" -ne "") {
            $m = [regex]::Match($item.nzsrsstr, '(\d+)')
            if ($m.Success) { $nzsrs = $m.Groups[1].Value }
        } else {
            $nzsrs = if ($item.nzsrs -ne $null) { "$($item.nzsrs)" } else { "" }
        }

        # ------ 所在地 ------
        $ssmc = $School.ssmc
        if (-not $ssmc -and $item.szss -ne $null) { $ssmc = "$($item.szss)" }

        # ------ 学习方式 ------
        $xxfsmc = ""
        if ($item.xxfsmc -and "$($item.xxfsmc)" -ne "") {
            $xxfsmc = "$($item.xxfsmc)"
        } elseif ($item.xxfs -ne $null -and "$($item.xxfs)" -ne "") {
            $xxfsmc = Get-XxfsText -Val "$($item.xxfs)"
        }

        # ------ 研究方向（含前缀代码）------
        $yjfxDisplay = ""
        $yjfxCode = if ($item.yjfxdm -ne $null) { "$($item.yjfxdm)" } else { "" }
        $yjfxName = if ($item.yjfxmc -ne $null) { "$($item.yjfxmc)" } else { "" }
        if ($yjfxCode -and $yjfxCode -ne "00") {
            $yjfxDisplay = "($yjfxCode)$yjfxName"
        } else {
            $yjfxDisplay = $yjfxName
        }

        # ------ 考试科目（支持多个可选组合，用"或"连接）------
        $examEnglish = @(); $examCourse1 = @(); $examCourse2 = @()

        if ($item.kskmz -ne $null -and @($item.kskmz).Count -gt 0) {
            $kmMap = @{ "km2Vo" = "外语"; "km3Vo" = "业务课一"; "km4Vo" = "业务课二" }
            foreach ($examGroup in @($item.kskmz)) {
                foreach ($key in @("km2Vo", "km3Vo", "km4Vo")) {
                    if ($examGroup.$key -ne $null) {
                        $vo = $examGroup.$key
                        $code = if ($vo.kskmdm -ne $null) { "$($vo.kskmdm)" } else { "" }
                        $name = if ($vo.kskmmc -ne $null) { "$($vo.kskmmc)" } else { "" }
                        $formatted = if ($code) { "($code)$name" } else { $name }
                        if ($key -eq "km2Vo") { if ($examEnglish -notcontains $formatted) { $examEnglish += $formatted } }
                        elseif ($key -eq "km3Vo") { if ($examCourse1 -notcontains $formatted) { $examCourse1 += $formatted } }
                        elseif ($key -eq "km4Vo") { if ($examCourse2 -notcontains $formatted) { $examCourse2 += $formatted } }
                    }
                }
            }
        }
        $examEnglish = $examEnglish -join " 或 "
        $examCourse1 = $examCourse1 -join " 或 "
        $examCourse2 = $examCourse2 -join " 或 "

        $records += @{
            院校代码 = $School.dwdm
            招生单位 = $School.dwmc
            所在地 = $ssmc
            院系所 = if ($item.yxsmc -ne $null) { "$($item.yxsmc)" } else { "" }
            专业 = if ($item.zymc -ne $null) { "$($item.zymc)" } else { "" }
            学习方式 = $xxfsmc
            研究方向 = $yjfxDisplay
            拟招生人数 = $nzsrs
            外语 = $examEnglish
            业务课一 = $examCourse1
            业务课二 = $examCourse2
            备注 = if ($item.zybz -ne $null) { "$($item.zybz)" } else { "" }
        }
    }

    if ($allList.Count -gt 10) {
        Write-Host "    → $($records.Count) 条研究方向记录（翻页 $curPage 页）" -ForegroundColor Gray
    } else {
        Write-Host "    → $($records.Count) 条研究方向记录" -ForegroundColor Gray
    }
    return $records
}

# ==========================  CSV 导出（含合并相同单元格）==================

$CSV_HEADERS = @("院校代码","招生单位","所在地","院系所","专业","学习方式","研究方向","拟招生人数","外语","业务课一","业务课二","备注")

function Export-Csv {
    param($Records, $FilePath)

    if ($Records.Count -eq 0) { "无数据" | Out-File -FilePath $FilePath -Encoding utf8; return }

    # 需要合并的列（相同值仅保留首次出现）
    $mergeCols = @("院校代码", "招生单位", "院系所")

    $lines = @()
    $lines += $CSV_HEADERS -join ","

    $prevValues = @{}
    foreach ($h in $mergeCols) { $prevValues[$h] = $null }

    foreach ($rec in $Records) {
        $vals = @()
        foreach ($h in $CSV_HEADERS) {
            $v = $rec.$h
            if (-not $v) { $v = "" }

            # 合并逻辑：相同值仅第一次显示
            if ($mergeCols -contains $h) {
                if ($v -eq $prevValues[$h]) {
                    $v = ""  # 与前一行相同，留空
                } else {
                    $prevValues[$h] = $v  # 记录新值
                }
            }

            if ($v -match '[,"\n]') { $v = '"' + $v.Replace('"', '""') + '"' }
            $vals += $v
        }
        $lines += $vals -join ","
    }

    $utf8Bom = [System.Text.Encoding]::UTF8.GetPreamble()
    $content = [System.Text.Encoding]::UTF8.GetBytes(($lines -join "`r`n"))
    $fs = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Create)
    $fs.Write($utf8Bom, 0, $utf8Bom.Length)
    $fs.Write($content, 0, $content.Length)
    $fs.Close()
}

# ==========================  入口  =======================================

Write-Host ""
Write-Host "╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║       研招网 · 硕士专业目录爬取工具 v2.3     ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

$rawUrl = Read-Host "  粘贴专业详情页 URL"

$params = Parse-DetailUrl -Url $rawUrl
if (-not $params) {
    Write-Host "`n  [✗] URL 解析失败" -ForegroundColor Red
    Read-Host; exit
}

Write-Host "`n  [✓] 专业: $($params.zymc)（$($params.zydm)）" -ForegroundColor Green
Write-Host "      门类/学科: $($params.mlmc) > $($params.yjxkmc)"

Write-Host "`n▸ 如需完整数据，请粘贴已登录的 Cookie" -ForegroundColor Yellow
$userCookie = Read-Host "  粘贴 Cookie（直接回车跳过）"
if ($userCookie) {
    $script:CookieHeader = $userCookie
    Write-Host "  [✓] Cookie 已设置" -ForegroundColor Green
}

Write-Host "`n▸ 正在连接研招网 ..." -ForegroundColor Yellow
$detailHtml = Invoke-YZ -Url $rawUrl
if (-not $detailHtml) {
    Write-Host "  [✗] 无法连接" -ForegroundColor Red; Read-Host; exit
}
Write-Host "  [✓] 连接成功" -ForegroundColor Green
$script:DetailReferer = $rawUrl

$schools = Get-SchoolList -Params $params
if ($schools.Count -eq 0) {
    Write-Host "`n  [✗] 未找到招生单位" -ForegroundColor Red; Read-Host; exit
}
Write-Host "`n  [✓] 共找到 $($schools.Count) 个招生单位" -ForegroundColor Green

$allRecords = @()
$i = 0
foreach ($school in $schools) {
    $i++
    Write-Host "`n  [$i/$($schools.Count)] [$($school.dwdm)] $($school.dwmc)（$($school.ssmc)）..." -ForegroundColor Cyan
    $details = Get-Detail -Params $params -School $school
    $allRecords += $details
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  共获取 $($allRecords.Count) 条记录" -ForegroundColor Green

if ($allRecords.Count -gt 0) {
    $desktop = [Environment]::GetFolderPath("Desktop")
    $safeName = "$($params.zydm)$($params.zymc)"
    $xwlxText = if ($params.xwlx -eq "ssxw") { "学硕" } elseif ($params.xwlx -eq "zyxw") { "专硕" } else { "" }
    $filename = "${safeName}${xwlxText}_专业目录.csv"
    $filepath = Join-Path $desktop $filename
    Export-Csv -Records $allRecords -FilePath $filepath
    Write-Host "`n  [✓] 已保存到桌面: $filename" -ForegroundColor Green
    Write-Host "  用 Excel 直接打开即可（相同院校/院系已合并显示）" -ForegroundColor White
} else {
    Write-Host "  [i] 无数据可导出" -ForegroundColor Yellow
}

Write-Host "`n  按回车退出..."
Read-Host

