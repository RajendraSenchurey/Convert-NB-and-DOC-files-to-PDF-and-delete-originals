#Converts both .nb and .docx to pdf and deletes originals
#It is just a one-step process. Set your main parent folder here. It will search this folder and all subfolders inside it, and undertakes conversions and deletion.  

$folderPath = "C:\Users\Rajendra\Desktop\FolderName"

#Get all target document types across all subfolders (-Recurse)
$targetFiles = Get-ChildItem -Path $folderPath -Recurse | Where-Object { 
    $_.Extension -match "^\.(nb|docx|doc|docs)$" -and $_.Name -notlike "~$*"
}

if ($targetFiles.Count -eq 0) {
    Write-Host "No matching documents found in '$folderPath' or its subfolders." -ForegroundColor Red
    return
}

Write-Host "Found $($targetFiles.Count) documents (.nb, .docx, .doc, .docs). Starting conversion..." -ForegroundColor Cyan

foreach ($file in $targetFiles) {
    $targetFolder = $file.DirectoryName
    $pdfPath      = [System.IO.Path]::ChangeExtension($file.FullName, ".pdf")
    
    Write-Host "Processing ($($file.Extension)): $($file.Name)" -ForegroundColor Yellow

    try {
        # -------------------------------------------------------------
        # CASE 1: NOTA BENE (.nb) FILES
        # -------------------------------------------------------------
        if ($file.Extension -like ".nb") {
            # Read and clean Nota Bene codes
            $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
            $text  = [System.Text.Encoding]::GetEncoding(1252).GetString($bytes)

            $text = $text -replace '[\x00-\x09\x0B\x0C\x0E-\x1F\x7F-\x9F]', ''
            $text = $text -replace '®[A-Z0-9\+\-]+\¯', ''
            $text = $text -replace '[®¯]', ''
            $text = $text -replace '[ｮｯョッツヨ]', ''
            $text = $text -replace '鍍he', '"the' -replace '的', '"' -replace '痴', "'s"

            # Create temporary text file
            $tempTxt = Join-Path $targetFolder "$($file.BaseName)_temp.txt"
            [System.IO.File]::WriteAllText($tempTxt, $text, [System.Text.Encoding]::UTF8)

            # Convert temp text to PDF via Word
            $word = New-Object -ComObject Word.Application
            $word.DisplayAlerts = 0
            $word.Visible = $false

            $doc = $word.Documents.Open($tempTxt, $false, $true)
            $doc.ExportAsFixedFormat($pdfPath, 17) # 17 = wdExportFormatPDF
            $doc.Close(0)
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

            Remove-Item $tempTxt -Force
        }
        # -------------------------------------------------------------
        # CASE 2: WORD DOCUMENTS (.docx, .doc, .docs)
        # -------------------------------------------------------------
        else {
            $word = New-Object -ComObject Word.Application
            $word.DisplayAlerts = 0
            $word.Visible = $false

            $doc = $word.Documents.Open($file.FullName, $false, $true)
            $doc.ExportAsFixedFormat($pdfPath, 17)
            $doc.Close(0)
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }

        # -------------------------------------------------------------
        # SAFETY CHECK & CLEANUP
        # -------------------------------------------------------------
        if (Test-Path $pdfPath) {
            Remove-Item $file.FullName -Force
            Write-Host "  [OK] Created PDF & deleted original: $($file.Name)" -ForegroundColor Green
        } else {
            Write-Host "  [WARNING] PDF creation unverified. Kept original: $($file.Name)" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "  [ERROR] Failed to convert $($file.Name): $_" -ForegroundColor Red
    }
}

Write-Host "`nAll done! All target files converted to PDF and original files cleaned up!" -ForegroundColor Cyan
