$txt1 = @"
To: 
Cc: 
Subject: 今週の週報の報告です

さん

お疲れ様です。片桐です。
掲題の件、下記のとおり、ご報告いたします。

報告期間: {0}年{1}月{2}日（月）～{3}年{4}月{5}日（金）
◆業務内容:
"@

$txt2 = @"
◎{0}月{1}日（{2}）
・CHIPS
"@

$txt3 = @"
◆ 成果・ 報告事項：


ご確認お願いいたします。
"@


# ================ main process ================
Write-Host "`n-----------------------------------------------------------`n"

$today = Get-Date
$friday = $null
$monday = $null

$todayWeekday = $today.DayOfWeek.value__
if ($todayWeekday -eq 5) {
    $friday = $today
    $monday = $friday.AddDays(-4)

} else {
    for ($i = 1; $i -le 6; $i++) {
        $tmp_day = $today.AddDays($i*-1)
        if ($tmp_day.DayOfWeek.value__ -eq 5) {
            $friday = $tmp_day
            $monday = $friday.AddDays(-4)
            break
        }
    }
}

Write-Host ($txt1 -f $monday.Year, $monday.Month, $monday.Day, $friday.Year, $friday.Month, $friday.day)

for ($i=0; $i -lt 5; $i++) {
    $day = $monday.AddDays($i)
    Write-Host ($txt2 -f $day.Month, $day.Day, $day.ToString("ddd"))
}

Write-Host "`n$txt3"
Write-Host "`n-----------------------------------------------------------`n"

$null = Read-Host 'Press Enter Key...'
