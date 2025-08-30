# ���ӓ_
# - Shift JIS �ŕۑ����Ȃ��ƃt�@�C���p�X�Ȃǂ̕����񏈗������܂����삵�Ȃ��̂ŋC������B
# - �A�j���[�V����������X���C�h�� png ���͂ł��Ȃ��̂ŁA���̏ꍇ�͌Â������ł��K�v������B

if ($Args.Length -ne 1) {
  Write-Output "You need to give the path of target Powerpoint File"
  exit
}

if (-Not (Test-Path $Args[0] -PathType Leaf)) {
  Write-Output "File '${Args[0]}' not exist."
  exit
}

$targetPath = "sample.pptx"
$targetPath = Resolve-Path -Path $Args[0]
# $targetBase = (Get-Item $targetPath).Basename
$targetBase = [System.IO.Path]::GetFileNameWithoutExtension((Get-Item $targetPath))

$baseFolder = Split-Path -Path $targetPath
# $baseFolder = "tmp_out"
$outputFolder = "${baseFolder}\${targetBase}"

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Open("${targetPath}")

# �X�V���͌Ăяo�����Ńt�H���_���ƍ��Ȃ����z��Ȃ̂ŁA�Ƃ肠�����폜�͕s�v
# $pngFiles = Get-ChildItem -Path $outputFolder -Filter *.png | Select-Object -ExpandProperty Name
# foreach ($file in $pngFiles) {
#   Remove-Item -Path "${outputFolder}\${file}"
# }

# �t�H���_���Ƀh�b�g���܂܂��Ə���Ɋg���q�Ƃ��Ĕ��f����č폜����Ă��܂����߁A
# ������ targetPath ���w�肵�Ă���B
# $pres.SaveAs("$targetPath", [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::PpSaveAsPNG)

$width = 480
$height = 270
# �o�̓C���[�W�̃T�C�Y���w�肷��ɂ́A�e�X���C�h�ʁX�ɃG�N�X�|�[�g����K�v������B
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $slide = $pres.Slides.Item($i)
    $file  = Join-Path $outputFolder "$i.png"

    if (-not (Test-Path $outputFolder)) {
        New-Item -ItemType Directory -Path $outputFolder | Out-Null
    }

    $slide.Export($file, "PNG", $width, $height)
    Write-Output "Exported: $file"
}

$pres.Close()
$ppt.Quit()

# # Renaming
  # $outputFiles = Get-ChildItem "${outputFolder}\*.png"
# 
# for ($i = 1; $i -le $outputFiles.Count; $i++) {
#   Rename-Item -Path "${outputFolder}\�X���C�h$i.PNG" -NewName "${outputFolder}\$i.png"
# }
