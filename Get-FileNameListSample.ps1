# Get-FileNameListSample
# 指定されたパスに格納されたファイルのファイル名を取得します。

$path = 'C:\Users\xxxxx\xxxxx\xxxxx' # 参照するパス

$items = Get-childItem $path -File
foreach($item in $items) {
    $item.name
}