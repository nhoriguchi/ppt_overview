# 留意点
# - Shift JIS で保存しないとファイルパスなどの文字列処理がうまく動作しないので気をつける。
# - アニメーションがあるスライドの png 化はできないので、その場合は古いやり方でやる必要がある。

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

# 更新時は呼び出し側でフォルダごと作りなおす想定なので、とりあえず削除は不要
# $pngFiles = Get-ChildItem -Path $outputFolder -Filter *.png | Select-Object -ExpandProperty Name
# foreach ($file in $pngFiles) {
#   Remove-Item -Path "${outputFolder}\${file}"
# }

# フォルダ名にドットが含まれると勝手に拡張子として判断されて削除されてしまうため、
# あえて targetPath を指定している。
# $pres.SaveAs("$targetPath", [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::PpSaveAsPNG)

$width = 480
$height = 270
# 出力イメージのサイズを指定するには、各スライド別々にエクスポートする必要がある。
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
#   Rename-Item -Path "${outputFolder}\スライド$i.PNG" -NewName "${outputFolder}\$i.png"
# }
