# メモ
# 指定されたパスの直下ファイルを全件確認して
# 指定されたシート・レンジの値を読み取りコンソールに出力するだけ
# コンソールをコピーしてテキストに貼り付けるなどしてあとはご自由に
# UTF-8で出力されるっぽいので注意（Shift-JIS）に変更する必要あり
#
# シート名'xxxxx'が存在しないとExceptionになる（未処理）

# 参照するパス
$path = 'C:\Users\xxxx\Desktop\work'

# 参照するシート名
$sheetName = 'xxxxx'

$items = Get-ChildItem $path -File
$excel = New-Object -ComObject Excel.Application

# ヘッダーを出力
$aryStr = @('ファイル名', 'xxxx', 'xxxx')
$joinstr = $aryStr -join ","

foreach ($item in $items) {
    $excel.Visible = $True
    $excel.DisplayAlerts = $Flase

    $book = $excel.Workbooks.Open($item.FullName)
    $sheet = $excel.Sheets($sheetName)

    $readData = $sheet.Range("A011:A014")

    $aryStr = @(
        $item.Name,
        -join ('"' + $readData[1].Text + '"')
        -join ('"' + $readData[2].Text + '"')
        -join ('"' + $readData[3].Text + '"')
    )

    #Excel終了
    $excel.Quit()
}

#プロセス解放
$excel = $Null
[GC]::collect()

