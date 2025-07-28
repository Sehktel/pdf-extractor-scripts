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
    Упрощенная версия для селективного извлечения PDF файлов из архивов (ZIP, RAR, 7Z)
.PARAMETER SourceArchive
    Путь к архиву (поддерживаются .zip, .rar, .7z)
.PARAMETER OutputDir
    Папка назначения для PDF файлов
.EXAMPLE
    .\Extract-PDFs-Simple.ps1 -SourceArchive "C:\Downloads\MyArchive.zip" -OutputDir "C:\ExtractedPDFs"
.EXAMPLE
    .\Extract-PDFs-Simple.ps1 -SourceArchive "C:\Downloads\Course.rar" -OutputDir "C:\ExtractedPDFs"
#>

param(
    [Parameter(Mandatory = $true, HelpMessage = "Путь к архиву (ZIP, RAR, 7Z)")]
    [string]$SourceArchive,
    
    [Parameter(Mandatory = $true, HelpMessage = "Папка назначения")]
    [string]$OutputDir
)

# Импорт .NET классов для работы с ZIP
Add-Type -AssemblyName System.IO.Compression

# Функция определения типа архива
function Get-ArchiveType {
    param([string]$FilePath)
    
    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    
    switch ($extension) {
        ".zip" { return "ZIP" }
        ".rar" { return "RAR" }
        ".7z"  { return "7Z" }
        default { throw "Неподдерживаемый тип архива: $extension. Поддерживаются: .zip, .rar, .7z" }
    }
}

# Функция поиска утилит
function Find-ArchiveTools {
    $tools = @{
        WinRAR = $null
        SevenZip = $null
    }
    
    # Поиск WinRAR
    $winrarPaths = @(
        "${env:ProgramFiles}\WinRAR\unrar.exe",
        "${env:ProgramFiles(x86)}\WinRAR\unrar.exe",
        "${env:ProgramFiles}\WinRAR\WinRAR.exe",
        "${env:ProgramFiles(x86)}\WinRAR\WinRAR.exe"
    )
    
    foreach ($path in $winrarPaths) {
        if (Test-Path $path) {
            $tools.WinRAR = $path
            break
        }
    }
    
    # Поиск 7-Zip
    $sevenZipPaths = @(
        "${env:ProgramFiles}\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
    )
    
    foreach ($path in $sevenZipPaths) {
        if (Test-Path $path) {
            $tools.SevenZip = $path
            break
        }
    }
    
    return $tools
}

function Extract-PDFsSelectively {
    param($ArchivePath, $TargetDir)
    
    try {
        # Валидация входных данных
        if (-not (Test-Path $ArchivePath)) {
            throw "Архив не найден: $ArchivePath"
        }
        
        # Определение типа архива
        $archiveType = Get-ArchiveType $ArchivePath
        Write-Host "🔍 Обрабатываем $archiveType архив..." -ForegroundColor Cyan
        
        # Создание целевой папки
        if (-not (Test-Path $TargetDir)) {
            New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
        }
        
        # Получение списка PDF файлов в зависимости от типа архива
        $pdfFiles = @()
        
        switch ($archiveType) {
            "ZIP" {
                $pdfFiles = Get-ZipPdfFiles $ArchivePath
            }
            "RAR" {
                $pdfFiles = Get-RarPdfFiles $ArchivePath
            }
            "7Z" {
                $pdfFiles = Get-SevenZipPdfFiles $ArchivePath
            }
        }
        
        Write-Host "📄 Найдено PDF файлов: $($pdfFiles.Count)" -ForegroundColor Green
        
        if ($pdfFiles.Count -eq 0) {
            Write-Host "⚠️ PDF файлы в архиве не обнаружены" -ForegroundColor Yellow
            return
        }
        
        # Извлечение PDF файлов
        $counter = 0
        foreach ($fileName in $pdfFiles) {
            $counter++
            
            $outputPath = Join-Path $TargetDir $fileName
            $outputDirectory = Split-Path $outputPath -Parent
            
            # Создание папки если нужно
            if ($outputDirectory -and -not (Test-Path $outputDirectory)) {
                New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
            }
            
            # Извлечение файла в зависимости от типа архива
            switch ($archiveType) {
                "ZIP" {
                    Extract-ZipFile $ArchivePath $fileName $outputPath
                }
                "RAR" {
                    Extract-RarFile $ArchivePath $fileName $outputPath
                }
                "7Z" {
                    Extract-SevenZipFile $ArchivePath $fileName $outputPath
                }
            }
            
            if (Test-Path $outputPath) {
                $sizeKB = [math]::Round((Get-Item $outputPath).Length / 1KB, 1)
                Write-Host "✅ [$counter/$($pdfFiles.Count)] $fileName ($sizeKB КБ)" -ForegroundColor Green
            }
        }
        
        Write-Host "`n🎉 Готово! PDF файлы извлечены в: $TargetDir" -ForegroundColor Yellow
        
        # Статистика
        $extractedFiles = Get-ChildItem $TargetDir -Recurse -Filter "*.pdf"
        $totalSize = ($extractedFiles | Measure-Object -Property Length -Sum).Sum
        Write-Host "📊 Извлечено файлов: $($extractedFiles.Count)" -ForegroundColor Cyan
        Write-Host "📊 Общий размер: $([math]::Round($totalSize / 1MB, 2)) МБ" -ForegroundColor Cyan
    }
    catch {
        Write-Host "❌ Ошибка: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# Функции для работы с ZIP архивами
function Get-ZipPdfFiles {
    param($ArchivePath)
    
    $fileStream = [System.IO.File]::OpenRead($ArchivePath)
    $zipArchive = [System.IO.Compression.ZipArchive]::new($fileStream)
    
    try {
        $pdfEntries = $zipArchive.Entries | Where-Object { 
            $_.Name -match '\.pdf$' -and $_.Length -gt 0 
        }
        return $pdfEntries | ForEach-Object { $_.FullName }
    }
    finally {
        $zipArchive.Dispose()
        $fileStream.Dispose()
    }
}

function Extract-ZipFile {
    param($ArchivePath, $FileName, $OutputPath)
    
    $fileStream = [System.IO.File]::OpenRead($ArchivePath)
    $zipArchive = [System.IO.Compression.ZipArchive]::new($fileStream)
    
    try {
        $entry = $zipArchive.Entries | Where-Object { $_.FullName -eq $FileName } | Select-Object -First 1
        
        if ($entry) {
            $entryStream = $entry.Open()
            $outputFileStream = [System.IO.File]::Create($OutputPath)
            
            try {
                $entryStream.CopyTo($outputFileStream)
            }
            finally {
                $entryStream.Dispose()
                $outputFileStream.Dispose()
            }
        }
    }
    finally {
        $zipArchive.Dispose()
        $fileStream.Dispose()
    }
}

# Функции для работы с RAR архивами
function Get-RarPdfFiles {
    param($ArchivePath)
    
    $tools = Find-ArchiveTools
    if (-not $tools.WinRAR) {
        throw "WinRAR не найден. Установите WinRAR для работы с .rar файлами"
    }
    
    $listResult = & $tools.WinRAR "l" "-cfg-" "$ArchivePath" 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Ошибка чтения RAR архива"
    }
    
    $pdfFiles = @()
    $inFileList = $false
    
    foreach ($line in $listResult) {
        if ($line -match "^-{15,}") {
            $inFileList = -not $inFileList
            continue
        }
        
        if ($inFileList -and $line -match "\s+(\S.*\.pdf)\s*$") {
            $pdfFiles += $matches[1].Trim()
        }
    }
    
    return $pdfFiles
}

function Extract-RarFile {
    param($ArchivePath, $FileName, $OutputPath)
    
    $tools = Find-ArchiveTools
    $outputDir = Split-Path $OutputPath -Parent
    
    $extractResult = & $tools.WinRAR "e" "-cfg-" "-o+" "$ArchivePath" "$FileName" "$outputDir" 2>&1
    
    if ($LASTEXITCODE -eq 0) {
        $extractedFile = Join-Path $outputDir (Split-Path $FileName -Leaf)
        if ((Test-Path $extractedFile) -and ($extractedFile -ne $OutputPath)) {
            Move-Item $extractedFile $OutputPath -Force
        }
    }
}

# Функции для работы с 7-Zip архивами
function Get-SevenZipPdfFiles {
    param($ArchivePath)
    
    $tools = Find-ArchiveTools
    if (-not $tools.SevenZip) {
        throw "7-Zip не найден. Установите 7-Zip для работы с .7z файлами"
    }
    
    $listResult = & $tools.SevenZip "l" "-slt" "$ArchivePath" 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Ошибка чтения 7-Zip архива"
    }
    
    $pdfFiles = @()
    $currentFile = ""
    
    foreach ($line in $listResult) {
        if ($line -match "^Path = (.+)$") {
            $currentFile = $matches[1]
        }
        elseif ($line -match "^Folder = -$" -and $currentFile -match "\.pdf$") {
            $pdfFiles += $currentFile
            $currentFile = ""
        }
    }
    
    return $pdfFiles
}

function Extract-SevenZipFile {
    param($ArchivePath, $FileName, $OutputPath)
    
    $tools = Find-ArchiveTools
    $tempDir = Join-Path $env:TEMP "7zip_simple_$(Get-Random)"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    
    try {
        $extractResult = & $tools.SevenZip "e" "$ArchivePath" "-o$tempDir" "$FileName" 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            $extractedFile = Join-Path $tempDir (Split-Path $FileName -Leaf)
            if (Test-Path $extractedFile) {
                Move-Item $extractedFile $OutputPath -Force
            }
        }
    }
    finally {
        if (Test-Path $tempDir) {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# Запуск основной функции
Write-Host "🚀 Начинаем селективное извлечение PDF файлов" -ForegroundColor Green
Extract-PDFsSelectively -ArchivePath $SourceArchive -TargetDir $OutputDir 