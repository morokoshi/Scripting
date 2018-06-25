$WordApplication = New-Object -ComObject Word.Application
$WdExportFormat = "Microsoft.Office.Interop.Word.WdExportFormat" -as [type]
$WdExportOptimizeFor = "Microsoft.Office.Interop.Word.WdExportOptimizeFor" -as [type]
$WdExportRange = "Microsoft.Office.Interop.Word.WdExportRange" -as [type]
$WdExportItem = "Microsoft.Office.Interop.Word.WdExportItem" -as [type]
$WdExportCreateBookmarks = "Microsoft.Office.Interop.Word.WdExportCreateBookmarks" -as [type]

$ExportFormat = $WdExportFormat::wdExportFormatPDF
$OpenAfterExport = $False
$OptimizeFor = $WdExportOptimizeFor::wdExportOptimizeForPrint
$Range = $WdExportRange::wdExportAllDocument
$Form = 1
$To = 1
$Item = $WdExportItem::wdExportDocumentContent
$IncludeDocProps = $True
$KeepIRM = $True
$CreateBookmarks = $WdExportCreateBookmarks::wdExportCreateNoBookmarks
$DocStructureTags = $True
$BitmapMissingFonts = $True
$UseISO19005_1 = $False

$Index = 0
$Script:Logs = @()

Function ConvertDocuments{
    Process {
        $SourceFileName = $_.FullName
        Write-Progress -Activity "PDFに変換しています..." -Status $SourceFileName -PercentComplete ($Index / $SourceFiles.Count * 100)
        Try{
            $Documents = $WordApplication.Documents.Open($SourceFileName, $False, $True)
            # https://msdn.microsoft.com/ja-jp/library/microsoft.office.tools.word.document.exportasfixedformat.aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
            Try{
                $Documents.ExportAsFixedFormat((GetDestinationFileName $SourceFileName), $ExportFormat, $OpenAfterExport, $OptimizeFor, $Range, $From, $To, $Item, $IncludeDocProps, $KeepIRM, $CreateBookmarks, $DocStructureTags, $BitmapMissingFonts, $UseISO19005_1)
            }
            Catch{
                $Log = "" | Select SourceFileName, Message
                $Log.SourceFileName = $SourceFileName
                $Log.Message = $Error[0].Exception.Message
                $Script:Logs += $Log
                #Write-Host ("保存できません: " + $SourceFileName) -ForegroundColor Red
                #Write-Host ("詳細: " + $Error[0].Exception.Message) -ForegroundColor Red
            }
            $Documents.Close($False)
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

$WordApplication.Visible = $False

# 変換するファイルを指定する
$SourceFiles = Get-ChildItem -Recurse -Include "*.doc", "*.docx"
$SourceFiles | ConvertDocuments

$WordApplication.Quit()

if ($Script:Logs.Count -ne 0){
    Write-Host ""
    Write-Host "エラーが発生したファイル" -ForegroundColor Red
    $Script:Logs | Format-Table -AutoSize
}

Pause
