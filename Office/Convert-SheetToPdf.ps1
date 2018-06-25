$XlFixedFormatType = "Microsoft.Office.Interop.Excel.XlFixedFormatType" -as [type]
$XlFixedFormatQuality = "Microsoft.Office.Interop.Excel.XlFixedFormatQuality" -as [type]

$Type = $XlFixedFormatType::xlTypePDF
$Quality = $XlFixedFormatQuality::xlQualityStandard
$IncludeDocProperties = $True
$IgnorePrintAreas = $False
#$From = 1
#$To = 1
#$OpenAfterPublish = $False


$Index = 0
$Script:Logs = @()

Function ConvertWorkbooks{
    Process {
        $SourceFileName = $_.FullName
        Write-Progress -Activity "PDFに変換しています..." -Status $SourceFileName -PercentComplete ($Index / $SourceFiles.Count * 100)
        Try{
            $ExcelApplication = New-Object -ComObject Excel.Application
            $Workbook = $ExcelApplication.Workbooks.Open($SourceFileName, $False, $True)
            # https://msdn.microsoft.com/ja-jp/library/microsoft.office.tools.Excel.workbook.exportasfixedformat.aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
            Try{
                $Workbook.ExportAsFixedFormat($Type, (GetDestinationFileName $SourceFileName), $Quality, $IncludeDocProperties, $IgnorePrintAreas)#, , , $OpenAfterPublish)
            }
            Catch{
                $Log = "" | Select SourceFileName, Message
                $Log.SourceFileName = $SourceFileName
                $Log.Message = $Error[0].Exception.Message
                $Script:Logs += $Log
                #Write-Host ("保存できません: " + $SourceFileName) -ForegroundColor Red
                #Write-Host ("詳細: " + $Error[0].Exception.Message) -ForegroundColor Red
            }
            $Workbook.Close($False)
            $ExcelApplication.Quit()
        }
        Catch{
            $Log = "" | Select SourceFileName, Message
            $Log.SourceFileName = $SourceFileName
            $Log.Message = $Error[0].Exception.Message
            $Script:Logs += $Log
            #Write-Host ("保存できません: " + $SourceFileName) -ForegroundColor Red
            #Write-Host ("詳細: " + $Error[0].Exception.Message) -ForegroundColor Red
        }
        $Index += 1
    }
}

function GetDestinationFileName($SourceFileName){
    $DestinationFileName = [System.IO.Path]::ChangeExtension($SourceFileName, ".pdf")
    return $DestinationFileName
}

$ExcelApplication.Visible = $False

# 変換するファイルを指定する
$SourceFiles = Get-ChildItem -Recurse -Include "*.xls", "*.xlsx"
$SourceFiles | ConvertWorkbooks



if ($Script:Logs.Count -ne 0){
    Write-Host ""
    Write-Host "エラーが発生したファイル" -ForegroundColor Red
    $Script:Logs | Format-Table -AutoSize
}

Pause
