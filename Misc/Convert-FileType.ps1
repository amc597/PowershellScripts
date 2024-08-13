function Convert-FileType {
    param (
        $FilePath,
        $CurrentFileType,
        $FileTypeWanted
    )
    $Counter = 0
    $Start = (Get-Date)

    $CheckFile = Get-ChildItem -Path $FilePath -filter "*$CurrentFileType*" -Recurse
    $CounterMax = $CheckFile.Count

    foreach ($File in $CheckFile) {
        $Counter++
        Write-Host "$($File.FullName) - $Counter/$CounterMax`n" -ForegroundColor Magenta

        try {        
            convert.exe $File.FullName "$($File.FullName -replace $File.Extension, $FileTypeWanted)"
        }
        catch {
            Write-Host "broke"
        }
        finally {
            Write-Host "Done"
        }
    }

    $End = (Get-Date)
    $Elapsed = $End - $Start
    $ElapsedTime = 'Duration:  {0:hr} hour {0:mm} min {0:ss} sec' -f $Elapsed
    Write-Host  "$ElapsedTime" -ForegroundColor Magenta
}
Convert-FileType -FilePath "" -CurrentFileType ".tif" -FileTypeWanted ".pdf"
