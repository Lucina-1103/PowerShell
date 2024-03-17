# 変換元フォルダ
$folder1 = "C:\Users\Desktop\画像\"
$folder2 = "C:\Users\Desktop\リサイズ\"
$sizeT = 50
$sizeY = 50

[void][Reflecttion.Assembly]::LoadWithPartialName("System.Drawing")

$itemList = Get-ChildItem $folder1;
foreach($item in $itemList) {
    if($item.PSIContainer) {
        # フォルダ名を取得した場合はスキップ
    } else {
        $image = New-Object System.Drawing.Bitmap($folder1 + $item.name)

        $canvas = New-Object System.Drawing.Bitmap($sizeT, $sizeY)

        $graphics = [System.Drawing.Graphics]::FromImage($canvas)
        $graphics.DrawImage($image, (New-Object System.Drawing.Rectangle(0, 0, $canvas.Width, $canvas.Height)))

        $canvas.Save($folder2 + $item.name, [System.Drawing.Imageing.ImageFormat]::Jpeg)

        $graphics.Dispose()
        $canvas.Dispose()
        $image.Dispose()
    }
}