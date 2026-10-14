# 第一步：将所有.png文件重命名为.jpg
Get-ChildItem *.png | ForEach-Object {
    $newName = $_.BaseName + ".jpg"
    Rename-Item $_ -NewName $newName
}

# 第二步：给所有.pdf和.jpg文件添加序号
$i = 1
Get-ChildItem *.pdf, *.jpg | Sort-Object Name | ForEach-Object {
    Rename-Item $_ -NewName ("$i. " + $_.Name)
    $i++
}