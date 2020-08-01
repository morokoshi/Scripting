Add-Type -AssemblyName System.Runtime.InteropServices
$signature = @"
[DllImport("shlwapi.dll", BestFitMapping = false, CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = false, ThrowOnUnmappableChar = true)]
public static extern int SHLoadIndirectString(string pszSource, System.Text.StringBuilder pszOutBuf, int cchOutBuf, IntPtr ppvReserved); 
"@
$ShellLightweightUtilityFunctions = Add-Type -memberDefinition $signature -name "Win32SHLoadIndirectString" -passThru
[System.Text.StringBuilder]$outBuff = 1024


Function Get-AppxPackageDisplayName($ApplicationInformation){
    $DisplayNameInformation = "" | Select-Object Name, DisplayName, PackageFullName, DisplayNameByManifest, InstallLocation

    $DisplayNameInformation.Name = $ApplicationInformation.Name
    $DisplayNameInformation.DisplayName = $ApplicationInformation.DisplayName
    $DisplayNameInformation.PackageFullName = $ApplicationInformation.PackageFullName
    $DisplayNameInformation.InstallLocation = $ApplicationInformation.InstallLocation

    If ($ApplicationInformation.IsFramework){
        $DisplayNameInformation.DisplayNameByManifest = ([Xml](Get-Content -Path (Join-Path $_.InstallLocation "AppxManifest.xml") -Encoding utf8)).Package.Properties.DisplayName
    }
    Else{
        $DisplayNameInformation.DisplayNameByManifest = (Get-AppxPackageManifest -Package $DisplayNameInformation.PackageFullName).Package.Applications.Application.VisualElements.DisplayName
    }
    
    $DisplayNameInformation.DisplayNameByManifest | ForEach-Object{
        $DisplayNameByManifest = $_
        If (![String]::IsNullOrEmpty($DisplayNameByManifest)){
            $GetName = ""
            $CandidateName = ""

            # ms-resource:Resources/DisplayName のパターン
            If (![String]::IsNullOrEmpty($DisplayNameInformation.PackageFullName) -and ($DisplayNameByManifest.IndexOf("ms-resource:") -ne -1)){
                $Source = ((($DisplayNameByManifest -replace "ms-resource://", "") -replace "ms-resource:/", "") -replace "ms-resource:", "")
                $Source = "@{$($ApplicationInformation.PackageFullName)`?ms-resource:Resources/$Source}"
                If ($ShellLightweightUtilityFunctions::SHLoadIndirectString($Source, $outBuff ,$outBuff.Capacity, [System.IntPtr]::Zero) -eq 0){
                    $GetName = $outBuff.ToString()
                    If ($GetName.IndexOf("ms-resource:") -eq -1){
                        $CandidateName = $GetName
                    }
                }
            }

            # ms-resource:DisplayName のパターン
            If ([String]::IsNullOrEmpty($CandidateName) -and ![String]::IsNullOrEmpty($DisplayNameInformation.PackageFullName) -and ($DisplayNameByManifest.IndexOf("ms-resource:") -ne -1)){
                $Source = "@{$($ApplicationInformation.PackageFullName)`?$DisplayNameByManifest}"
                If ($ShellLightweightUtilityFunctions::SHLoadIndirectString($Source, $outBuff ,$outBuff.Capacity, [System.IntPtr]::Zero) -eq 0){
                    $GetName = $outBuff.ToString()
                    If ($GetName.IndexOf("ms-resource:") -eq -1){
                        $CandidateName = $GetName
                    }
                }
            }

            # 次の処理で CandidateName に DisplayNameByManifest を使用するので、ms-resource: を含む場合は空にしておく
            If ($DisplayNameByManifest.IndexOf("ms-resource:") -ne -1){
                $DisplayNameByManifest = $null
            }

            # CandidateName が空の場合はマニフェストから取得した文字列を使用する
            If ([String]::IsNullOrEmpty($CandidateName)){
                $CandidateName += $DisplayNameByManifest
            }

            If ($DisplayNameInformation.DisplayName -ne $CandidateName){
                If ($DisplayNameInformation.DisplayName.Count -eq 1){
                    # DisplayName が複数ある場合は配列にする
                    $TemporaryDisplayName = $DisplayNameInformation.DisplayName
                    $DisplayNameInformation.DisplayName = @()
                    $DisplayNameInformation.DisplayName += $TemporaryDisplayName
                    $DisplayNameInformation.DisplayName += $CandidateName
                }
                Else{
                    $DisplayNameInformation.DisplayName += $CandidateName
                }
            }
        }
    }

    # 配列からカンマ区切りのテキストにする
    $DisplayNameInformation.DisplayName = $DisplayNameInformation.DisplayName -Join ", "
    Return $DisplayNameInformation
}


Function Get-AppxProvisionedPackageDisplayName{
    $AppxPackageDisplayName = Get-AppxPackage | ForEach-Object {Get-AppxPackageDisplayName $_}
    Get-AppxProvisionedPackage -Online | ForEach-Object{
        $AppxPackageDisplayName | Where-Object Name -eq $_.DisplayName
    }
}


# 実行サンプル
# Usage1: Get-AppxPackage の結果
Get-AppxPackage | ForEach-Object {Get-AppxPackageDisplayName $_} | Select-Object Name, DisplayName | Format-Table -AutoSize

# Usage2: Get-AppxProvisionedPackage の結果
$AppxPackageDisplayName = Get-AppxPackage | ForEach-Object {Get-AppxPackageDisplayName $_}
Get-AppxProvisionedPackage -Online | ForEach-Object{
    $AppxPackageDisplayName | Where-Object Name -eq $_.DisplayName
} | Select-Object Name, DisplayName | Format-Table -AutoSize
