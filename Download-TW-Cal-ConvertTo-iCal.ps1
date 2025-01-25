
$dgpaPre = 'https://www.dgpa.gov.tw/'
$dgpaPattern1 = '(中華民國.*年.*日曆表)\s*$'
$dgpaurl1 = $($dgpaPre + 'informationlist?uid=41')
$TWHlsFltPtn2 = '\.xls\s*.*$'

function Filter-PatternInHTMLLink {
    param (
        [Parameter(Mandatory=$true)]
        [object]$document,

        [Parameter(Mandatory=$true)]
        [string]$pattern
    )

    $linkmatches = @()
    foreach ($link in $document.getElementsByTagName("a")) {
        $linkProperties = @{}
        $linkinnerText = $link.innerText
        if (-not $linkinnerText -match $pattern) {
            $linkinnerText = $link.parentNode.innerText
        }
        $linkinnerText | Select-String $pattern -AllMatches | ForEach-Object {
            $_.Matches | ForEach-Object {
                $_ | ForEach-Object {
                    $linkProperties = @{
                        'Title' = $_.Groups[0].Value
                        'Url' = $link.href
                    }
                    $linkmatches += [PSCustomObject]$linkProperties
                }
            }
        }
    }
    return $linkmatches
}

if ($PSCommandPath -eq $null) {
	function GetPSCommandPath() {
		return "$($MyInvocation.PSCommandPath)"
	}
	$PSCommandPath = GetPSCommandPath
}

#設定路徑變數
$CurrentPS1File = $(Get-Item -Path "$PSCommandPath")
$currentDirectory = "$("$($CurrentPS1File.PSParentPath)" + '\')"
$OldDirectory = "$($currentDirectory + $($CurrentPS1File.BaseName) + '_Old\')"
Set-Location $currentDirectory
New-Item -ItemType Directory -Force -Path $OldDirectory

Add-Type -AssemblyName System.Web

$dgpaResp1 = Invoke-WebRequest -Uri $dgpaurl1
$dgpaDoc1 = $dgpaResp1.ParsedHtml
$TWHldLst1 = Filter-PatternInHTMLLink -document $dgpaDoc1 -pattern $dgpaPattern1
if(-not $TWHldLst1.item){
    $TWHldURL1 = $TWHldLst1.Url -replace '^about:',''
} else{
    $TWHldURL1 = $TWHldLst1.item(0).Url -replace '^about:',''
}

$dgpaurl2 = $( $dgpaPre + $TWHldURL1 )
$dgpaurl2
$dgpaResp2 = Invoke-WebRequest -Uri $dgpaurl2
$dgpaDoc2 = $dgpaResp2.ParsedHtml
$TWHldLst2 = Filter-PatternInHTMLLink -document $dgpaDoc2 -pattern $TWHlsFltPtn2
if(-not $TWHldLst2.item){
    $TWHldURL2 = $TWHldLst2.Url -replace '^about:',''
} else {
    $TWHldURL2 = $TWHldLst2.item(0).Url -replace '^about:',''
}

$dgpaurl3 = $( $dgpaPre + $TWHldURL2 )
$dgpaurl3

# 解析 URL
$uri = [System.Uri]::new($dgpaurl3)
$queryParams = [System.Web.HttpUtility]::ParseQueryString($uri.Query)

# 提取並解碼檔名
$encodedFilename = $queryParams["name"]
$decodedFilename = [System.Web.HttpUtility]::UrlDecode($encodedFilename)

# 設置目標檔案路徑
$targetFile = Join-Path -Path $currentDirectory -ChildPath $decodedFilename

# 檢查檔案是否已存在
if (-Not (Test-Path -Path $targetFile)) {
    # 下載文件
    Invoke-WebRequest -Uri $dgpaurl3 -OutFile $targetFile
    Write-Output "文件已下載到: $targetFile"
} else {
    Write-Output "檔案已存在: $targetFile，未下載。"
}

# 宣告雜湊表變數
[hashtable]$matches = @{}

# 獲取最新修改的 .xls 檔案
$latestXlsFile = Get-ChildItem -Path $currentDirectory -Filter *.xls | Sort-Object LastWriteTime -Descending | Select-Object -First 1

# 檢查是否找到 .xls 檔案
if ($latestXlsFile) {
    Write-Output "最新的 .xls 檔案: $($latestXlsFile.Name)"
    Write-Output "最後修改時間: $($latestXlsFile.LastWriteTime.ToString('yyyy/MM/dd HH:mm:ss'))"

    # 初始化 Excel COM 物件
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($latestXlsFile.FullName)
    $sheet = $workbook.Sheets.Item("Sheet1")

    # 取得儲存格的邊框設定屬性
    #$tmpcell = $sheet.Cells.Item(3, 15)
    #$tmpcell.Borders | select * | ft -AutoSize

    # 初始化一個 List 變數來存儲匹配的儲存格
    $matchedCells = New-Object System.Collections.Generic.List[PSObject]

    $cellyear = ''
    $yearPtn = '西元[[:space:]]*([0-9]+)[[:space:]]*年'
    # 遍歷工作表的所有儲存格
    foreach ($cell in $sheet.UsedRange.Cells) {
        #$cellyear = (Get-Date).Year
        if ($cell.Text -match $yearPtn) {
            $cell.Text | Select-String $yearPtn -AllMatches | ForEach-Object {
                $_.Matches | ForEach-Object {
                    $_ | ForEach-Object {
                        if($_.Groups[1].Value){
                            $cellyear = $_.Groups[1].Value
                            break
                        }
                    }
                }
            }
        }
    }
    foreach ($cell in $sheet.UsedRange.Cells) {
        if ($cell.Text -match '^[[:space:]]*月[[:space:]]*$') {
            # 向左搜尋第一個有左邊邊框的儲存格
            $leftCell = $null
            $colIndex = $cell.Column
            while ($colIndex -gt 0) {
                $potentialCell = $sheet.Cells.Item($cell.Row, $colIndex)
                $pclinestyle1 = $potentialCell.Borders.Item(1).LineStyle
                $pclinestyle2 = $potentialCell.Borders.Item(2).LineStyle
                if ( ($pclinestyle1 -ne $null) -and ($pclinestyle1 -ge 1) ) {
                    $leftCell = $sheet.Cells.Item($cell.Row, $colIndex)
                    break
                } elseif ( ($pclinestyle2 -ne $null) -and ($pclinestyle2 -ge 1) -and ($colIndex -ne $cell.Column) ) {
                    $leftCell = $sheet.Cells.Item($cell.Row, $colIndex+1)
                    break
                }
                $colIndex--
            }

            # 向右搜尋第一個有右邊邊框的儲存格
            $rightCell = $null
            $colIndex = $cell.Column
            while ($colIndex -le $sheet.UsedRange.Columns.Count) {
                $potentialCell = $sheet.Cells.Item($cell.Row, $colIndex)
                $pclinestyle2 = $potentialCell.Borders.Item(2).LineStyle
                $pclinestyle1 = $potentialCell.Borders.Item(1).LineStyle
                if ( ($pclinestyle2 -ne $null) -and ($pclinestyle2 -ge 1) ) {
                    $rightCell = $sheet.Cells.Item($cell.Row, $colIndex)
                    break
                } elseif ( ($pclinestyle1 -ne $null) -and ($pclinestyle1 -ge 1) -and ($colIndex -ne $cell.Column) ) {
                    $rightCell = $sheet.Cells.Item($cell.Row, $colIndex-1)
                    break
                }
                $colIndex++
            }

            # 合併 Row 為 $cell 的儲存格內容
            $combinedText = ""
            if ($leftCell -ne $null -and $rightCell -ne $null) {
                for ($i = $leftCell.Column; $i -le $rightCell.Column; $i++) {
                    $combinedText += $sheet.Cells.Item($cell.Row, $i).Text
                }
            }

            # 向上搜尋第一個有上邊框的儲存格
            $topCell = $null
            $rowIndex = $cell.Row
            while ($rowIndex -gt 0) {
                $potentialCell = $sheet.Cells.Item($rowIndex, $cell.Column)
                $pclinestyle3 = $potentialCell.Borders.Item(3).LineStyle
                $pclinestyle4 = $potentialCell.Borders.Item(4).LineStyle
                if ( ($pclinestyle3 -ne $null) -and ($pclinestyle3 -ge 1) ) {
                    $topCell = $sheet.Cells.Item($rowIndex, $cell.Column)
                    break
                } elseif ( ($pclinestyle4 -ne $null) -and ($pclinestyle4 -ge 1) -and ($rowIndex -ne $cell.Row) ) {
                    $topCell = $sheet.Cells.Item($rowIndex+1, $cell.Column)
                    break
                }
                $rowIndex--
            }

            # 向下搜尋第一個有下邊框的儲存格
            $bottomCell = $null
            $rowIndex = $cell.Row
            while ($rowIndex -le $sheet.UsedRange.Rows.Count) {
                $potentialCell = $sheet.Cells.Item($rowIndex, $cell.Column)
                $pclinestyle4 = $potentialCell.Borders.Item(4).LineStyle
                $pclinestyle3 = $potentialCell.Borders.Item(3).LineStyle
                if ( ($pclinestyle4 -ne $null) -and ($pclinestyle4 -ge 1) ) {
                    $bottomCell = $sheet.Cells.Item($rowIndex, $cell.Column)
                    break
                } elseif ( ($pclinestyle3 -ne $null) -and ($pclinestyle3 -ge 1) -and ($rowIndex -ne $cell.Row) ) {
                    $bottomCell = $sheet.Cells.Item($rowIndex-1, $cell.Column)
                    break
                }
                $rowIndex++
            }

            if ($combinedText -ne $null) {
                # 創建並填充自定義物件
                # 循環範圍內的儲存格，將 Text 內容加入到 '星期、日期或節日' 欄位
                for ($row = $cell.Row+2; $row -le $bottomCell.Row; $row=$row+2) {
                    for ($col = $leftCell.Column; $col -le $rightCell.Column; $col++) {
                        $celldate = $sheet.Cells.Item($row, $col)
                        $cellcolor = $celldate.Interior.Color
                        $cellHld = '上班日'
                        switch ($cellcolor) {
                            16751103 { $cellHld = '放假日'; break }
                            16777215 { $cellHld = '上班日'; break }
                        }
                        if ($celldate.Text) {
                            $cellpsco = [PSCustomObject]@{
                                '年份' = $cellyear
                                '月份' = $combinedText
                                '日期' = $celldate.Text
                                '星期' = $($sheet.Cells.Item($cell.Row + 1, $col).Text)
                                '節日' = $($sheet.Cells.Item($row + 1, $col).Text)
                                '上班或放假' = $cellHld
                            }
                            # 將匹配的儲存格添加到 List 中
                            $matchedCells.Add($cellpsco)
                        }
                    }
                }
            }
        }
    }

    # 將 List 添加到雜湊表
    $matches["MatchedCells"] = $matchedCells

    # 關閉工作簿和 Excel 應用程式
    $workbook.Close($false)
    $excel.Quit()

    # 釋放 COM 物件
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

    # 輸出匹配到的儲存格地址
    Write-Output "匹配到的儲存格地址:"
    $matches["MatchedCells"] | Format-Table -AutoSize

    <#
    $matches["MatchedCells"] 輸出內容前9筆：
    年份   月份  日期 星期 節日     上班或放假
    --   --  -- -- --     -----
    2025 一月  1  三  初二     放假日  
    2025 一月  2  四  初三     上班日  
    2025 一月  3  五  初四     上班日  
    2025 一月  4  六  初五     放假日  
    2025 一月  5  日  小寒     放假日  
    2025 一月  6  一  初七     上班日  
    2025 一月  7  二  初八     上班日  
    2025 一月  8  三  初九     上班日  
    2025 一月  9  四  初十     上班日 
    ...
    #>

    $FiltedCells = New-Object System.Collections.Generic.List[PSObject]
    foreach ($cell in $matches["MatchedCells"]) {
        # 排除星期一到星期五的上班日和星期六到星期日的休假日
        if (( $($cell.'星期' -in @('一', '二', '三', '四', '五')) -and $($cell.'上班或放假' -eq '放假日') ) -or 
            ( $($cell.'星期' -in @('六', '日')) -and $($cell.'上班或放假' -eq '上班日')) ) {
            $FiltedCells += $cell
        }
    }
    $matches["FiltedCells"] = $FiltedCells

    # 定義一個雜湊表來替換月份
    $monthMapping = @{
        '一月' = 1
        '二月' = 2
        '三月' = 3
        '四月' = 4
        '五月' = 5
        '六月' = 6
        '七月' = 7
        '八月' = 8
        '九月' = 9
        '十月' = 10
        '十一月' = 11
        '十二月' = 12
    }

    # 轉換為 iCal 格式並輸出到 .ics 檔案
    $icalContent = "BEGIN:VCALENDAR`nVERSION:2.0"
    $icalContent += "`nTZID:Asia/Taipei"
    $icalContent += "`nX-WR-CALNAME:中華民國政府行政機關辦公日曆表"
    foreach ($cell in $matches["FiltedCells"]) {
        $monthNumber = $monthMapping[$cell.月份]
        $date = "$($cell.年份)$(($monthNumber.ToString()).PadLeft(2, '0'))$(($cell.日期).PadLeft(2, '0'))"
        $uid = [guid]::NewGuid().ToString()
        $icalContent += "`nBEGIN:VEVENT"
        $icalContent += "`nUID:$uid"
        $icalContent += "`nDTSTART;VALUE=DATE:$date"
        $icalContent += "`nDTEND;VALUE=DATE:$date"
        $icalContent += "`nSUMMARY:$($cell.'節日' -replace '[\r\n]',' ') $($cell.'上班或放假' -replace '[\r\n]',' ')"
        $icalContent += "`nEND:VEVENT"
    }
    $icalContent += "`nEND:VCALENDAR"

    # 輸出 iCal 檔案
    $icalFilePath = "$($OldDirectory + $($CurrentPS1File.BaseName) + '-' + $(Get-Date).ToString('yyyyMMdd-HHmmss') + '.ics')"
    Set-Content -Path $icalFilePath -Value $icalContent -Force
    Copy-Item -Path $icalFilePath -Destination "$($currentDirectory + $($CurrentPS1File.BaseName) + '.ics')" -Force

} else {
    Write-Output "當前目錄中未找到 .xls 檔案。"
}
