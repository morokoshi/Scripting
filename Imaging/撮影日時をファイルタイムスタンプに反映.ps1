Param([String]$DirectoryPath)
$DateTimeOriginalTitle = "撮影日時"

Function Check-DirectoryPath{
    If (!$DirectoryPath){
        $Script:DirectoryPath = Read-Host "フォルダーのパス"
        Check-DirectoryPath
        Exit
    }
    If (!(Test-Path $DirectoryPath)){
        Write-Warning "$DirectoryPath が見つかりませんので、パスを指定しなおしてください"
        $Script:DirectoryPath = Read-Host "フォルダーのパス"
        Check-DirectoryPath
        Exit
    }
}
Check-DirectoryPath

$ShellApplication = New-Object -COMObject Shell.Application
$ShellParentDirectory = $ShellApplication.Namespace($DirectoryPath)
$DateTimeOriginalId = 0..400 | Select @{Name="Id";Expression={($_)}}, @{Name="PropertyName";Expression={($ShellParentDirectory.GetDetailsOf($null, $_))}} | Where-Object {$_.PropertyName -eq $DateTimeOriginalTitle}

Get-ChildItem $DirectoryPath -Filter "*.jpg" | ForEach-Object {
    $ShellFile = $ShellParentDirectory.ParseName($_.Name)
    $DateTimeOriginal = $ShellParentDirectory.GetDetailsOf($ShellFile, $DateTimeOriginalId.Id)
    #Photoshop Lightroom Classic, Canon EOS 5D Mark IVで出力された撮影日時には双方向テキスト U+200E を含むため削除する
    $ParsedDateTimeOriginal = $DateTimeOriginal.Replace("‎","")
    $ParsedDateTimeOriginal = [DateTime]$DateTimeOriginal
    Write-Host ($_.Name + ": $ParsedDateTimeOriginal")

    Set-ItemProperty $_ -Name CreationTime -Value $ParsedDateTimeOriginal
    Set-ItemProperty $_ -Name LastWriteTime -Value $ParsedDateTimeOriginal
    Set-ItemProperty $_ -Name LastAccessTime -Value $ParsedDateTimeOriginal
}
