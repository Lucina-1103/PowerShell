# �ϊ����t�H���_
$folder1 = "C:\Users\Desktop\�摜\"
$folder2 = "C:\Users\Desktop\���T�C�Y\"
$sizeT = 50
$sizeY = 50

[void][Reflecttion.Assembly]::LoadWithPartialName("System.Drawing")

$itemList = Get-ChildItem $folder1;
foreach($item in $itemList) {
    if($item.PSIContainer) {
        # �t�H���_�����擾�����ꍇ�̓X�L�b�v
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