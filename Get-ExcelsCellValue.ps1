# ����
# �w�肳�ꂽ�p�X�̒����t�@�C����S���m�F����
# �w�肳�ꂽ�V�[�g�E�����W�̒l��ǂݎ��R���\�[���ɏo�͂��邾��
# �R���\�[�����R�s�[���ăe�L�X�g�ɓ\��t����Ȃǂ��Ă��Ƃ͂����R��
# UTF-8�ŏo�͂������ۂ��̂Œ��ӁiShift-JIS�j�ɕύX����K�v����
#
# �V�[�g��'xxxxx'�����݂��Ȃ���Exception�ɂȂ�i�������j

# �Q�Ƃ���p�X
$path = 'C:\Users\xxxx\Desktop\work'

# �Q�Ƃ���V�[�g��
$sheetName = 'xxxxx'

$items = Get-ChildItem $path -File
$excel = New-Object -ComObject Excel.Application

# �w�b�_�[���o��
$aryStr = @('�t�@�C����', 'xxxx', 'xxxx')
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

    #Excel�I��
    $excel.Quit()
}

#�v���Z�X���
$excel = $Null
[GC]::collect()

