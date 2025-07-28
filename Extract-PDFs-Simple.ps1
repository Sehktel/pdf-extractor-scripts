#Requires -Version 5.0

<#
.LICENSE
    MIT License
    
    Copyright (c) 2024 Sehktel
    GitHub: https://github.com/Sehktel/pdf-extractor-scripts
    
    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:
    
    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.
    
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

.SYNOPSIS
    Упрощенная версия для селективного извлечения PDF файлов из ZIP архива
.PARAMETER ZipFile
    Путь к ZIP архиву
.PARAMETER OutputDir
    Папка назначения для PDF файлов
.EXAMPLE
    .\Extract-PDFs-Simple.ps1 -ZipFile "C:\Downloads\MyArchive.zip" -OutputDir "C:\ExtractedPDFs"
#>

param(
    [Parameter(Mandatory = $true, HelpMessage = "Путь к ZIP архиву")]
    [string]$ZipFile,
    
    [Parameter(Mandatory = $true, HelpMessage = "Папка назначения")]
    [string]$OutputDir
)

# Импорт .NET классов для работы с ZIP
Add-Type -AssemblyName System.IO.Compression

function Extract-PDFsSelectively {
    param($ArchivePath, $TargetDir)
    
    try {
        # Валидация входных данных
        if (-not (Test-Path $ArchivePath)) {
            throw "Архив не найден: $ArchivePath"
        }
        
        # Создание целевой папки
        if (-not (Test-Path $TargetDir)) {
            New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
        }
        
        Write-Host "🔍 Анализируем архив..." -ForegroundColor Cyan
        
        # Открытие архива без полной распаковки
        $fileStream = [System.IO.File]::OpenRead($ArchivePath)
        $zipArchive = [System.IO.Compression.ZipArchive]::new($fileStream)
        
        try {
            # Поиск PDF файлов в архиве
            $pdfEntries = $zipArchive.Entries | Where-Object { 
                $_.Name -match '\.pdf$' -and $_.Length -gt 0 
            }
            
            Write-Host "📄 Найдено PDF файлов: $($pdfEntries.Count)" -ForegroundColor Green
            
            if ($pdfEntries.Count -eq 0) {
                Write-Host "⚠️ PDF файлы в архиве не обнаружены" -ForegroundColor Yellow
                return
            }
            
            $counter = 0
            foreach ($entry in $pdfEntries) {
                $counter++
                
                # Формирование пути с сохранением структуры папок
                $relativePath = $entry.FullName
                $outputPath = Join-Path $TargetDir $relativePath
                $outputDirectory = Split-Path $outputPath -Parent
                
                # Создание папки если нужно
                if ($outputDirectory -and -not (Test-Path $outputDirectory)) {
                    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
                }
                
                # Прямое извлечение файла из ZIP
                $entryStream = $entry.Open()
                $outputFileStream = [System.IO.File]::Create($outputPath)
                
                try {
                    $entryStream.CopyTo($outputFileStream)
                    $sizeKB = [math]::Round($entry.Length / 1KB, 1)
                    Write-Host "✅ [$counter/$($pdfEntries.Count)] $relativePath ($sizeKB КБ)" -ForegroundColor Green
                }
                finally {
                    $entryStream.Dispose()
                    $outputFileStream.Dispose()
                }
            }
            
            Write-Host "`n🎉 Готово! PDF файлы извлечены в: $TargetDir" -ForegroundColor Yellow
            
            # Статистика
            $totalSize = ($pdfEntries | Measure-Object -Property Length -Sum).Sum
            Write-Host "📊 Общий размер: $([math]::Round($totalSize / 1MB, 2)) МБ" -ForegroundColor Cyan
        }
        finally {
            $zipArchive.Dispose()
            $fileStream.Dispose()
        }
    }
    catch {
        Write-Host "❌ Ошибка: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# Запуск основной функции
Write-Host "🚀 Начинаем селективное извлечение PDF файлов" -ForegroundColor Green
Extract-PDFsSelectively -ArchivePath $ZipFile -TargetDir $OutputDir 