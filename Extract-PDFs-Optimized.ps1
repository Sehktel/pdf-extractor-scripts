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
    Архитектурно правильный скрипт для селективного извлечения PDF файлов из ZIP архива
.DESCRIPTION
    Использует прямой доступ к ZIP архиву без полной распаковки.
    Извлекает только PDF файлы с сохранением структуры папок.
    Максимально эффективен по памяти и дисковому пространству.
.PARAMETER SourceArchive
    Полный путь к исходному ZIP архиву (обязательный параметр)
.PARAMETER DestinationDirectory
    Целевая директория для извлеченных PDF файлов (обязательный параметр)
.PARAMETER IncludeSubdirectories
    Сохранять ли структуру поддиректорий (по умолчанию: true)
.PARAMETER OverwriteExisting
    Перезаписывать ли существующие файлы (по умолчанию: false)
.PARAMETER ShowProgress
    Показывать ли прогресс выполнения (по умолчанию: true)
.EXAMPLE
    .\Extract-PDFs-Optimized.ps1 -SourceArchive "C:\Downloads\CourseArchive.zip" -DestinationDirectory "C:\ExtractedPDFs"
.EXAMPLE
    .\Extract-PDFs-Optimized.ps1 -SourceArchive ".\MyArchive.zip" -DestinationDirectory ".\PDFs" -OverwriteExisting $true -LogLevel "Verbose"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Путь к архиву (.zip, .rar, .7z)")]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Leaf)) {
            throw "Файл архива не найден: $_"
        }
        if (-not ($_ -match '\.(zip|rar|7z)$')) {
            throw "Поддерживаемые форматы: .zip, .rar, .7z. Получен: $_"
        }
        return $true
    })]
    [string]$SourceArchive,
    
    [Parameter(Mandatory = $true, Position = 1, HelpMessage = "Целевая директория")]
    [string]$DestinationDirectory,
    
    [Parameter(Mandatory = $false)]
    [bool]$IncludeSubdirectories = $true,
    
    [Parameter(Mandatory = $false)]
    [bool]$OverwriteExisting = $false,
    
    [Parameter(Mandatory = $false)]
    [bool]$ShowProgress = $true,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Quiet", "Normal", "Verbose")]
    [string]$LogLevel = "Normal"
)

# Импорт необходимых .NET сборок для работы с ZIP архивами
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Перечисление поддерживаемых типов архивов
enum ArchiveType {
    Zip
    Rar
    SevenZip
}

# Класс для конфигурации утилит архивирования
class ArchiveTools {
    static [string]$WinRarPath
    static [string]$SevenZipPath
    static [bool]$Initialized = $false
    
    static [void] Initialize() {
        if ([ArchiveTools]::Initialized) { return }
        
        # Поиск WinRAR
        $winrarPaths = @(
            "${env:ProgramFiles}\WinRAR\unrar.exe",
            "${env:ProgramFiles(x86)}\WinRAR\unrar.exe",
            "${env:ProgramFiles}\WinRAR\WinRAR.exe",
            "${env:ProgramFiles(x86)}\WinRAR\WinRAR.exe"
        )
        
        foreach ($path in $winrarPaths) {
            if (Test-Path $path) {
                [ArchiveTools]::WinRarPath = $path
                break
            }
        }
        
        # Поиск 7-Zip
        $sevenZipPaths = @(
            "${env:ProgramFiles}\7-Zip\7z.exe",
            "${env:ProgramFiles(x86)}\7-Zip\7z.exe",
            "${env:ProgramData}\chocolatey\bin\7z.exe"
        )
        
        foreach ($path in $sevenZipPaths) {
            if (Test-Path $path) {
                [ArchiveTools]::SevenZipPath = $path
                break
            }
        }
        
        [ArchiveTools]::Initialized = $true
    }
}

# Класс для результатов операции
class ExtractionResult {
    [int]$TotalEntriesInArchive
    [int]$PdfFilesFound
    [int]$PdfFilesExtracted
    [int]$ErrorsOccurred
    [long]$TotalArchiveSize
    [long]$ExtractedPdfSize
    [string[]]$ExtractedFiles
    [string[]]$Errors
    
    ExtractionResult() {
        $this.ExtractedFiles = @()
        $this.Errors = @()
    }
}

# Функция логирования с учетом уровня детализации
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Success", "Warning", "Error", "Verbose")]
        [string]$Level = "Info",
        [string]$LogLevel = "Normal"
    )
    
    # Фильтрация сообщений по уровню логирования
    if ($LogLevel -eq "Quiet") { return }
    if ($LogLevel -eq "Normal" -and $Level -eq "Verbose") { return }
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $colors = @{
        "Info" = "Cyan"
        "Success" = "Green" 
        "Warning" = "Yellow"
        "Error" = "Red"
        "Verbose" = "Gray"
    }
    
    $prefix = switch($Level) {
        "Success" { "✅" }
        "Warning" { "⚠️" }
        "Error" { "❌" }
        "Verbose" { "🔍" }
        default { "ℹ️" }
    }
    
    Write-Host "[$timestamp] $prefix $Message" -ForegroundColor $colors[$Level]
}

# Функция определения типа архива
function Get-ArchiveType {
    param([string]$FilePath)
    
    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    
    switch ($extension) {
        ".zip" { return [ArchiveType]::Zip }
        ".rar" { return [ArchiveType]::Rar }
        ".7z"  { return [ArchiveType]::SevenZip }
        default { throw "Неподдерживаемый тип архива: $extension" }
    }
}

# Функция получения списка PDF файлов из RAR архива
function Get-RarPdfFiles {
    param(
        [string]$ArchivePath,
        [string]$LogLevel
    )
    
    [ArchiveTools]::Initialize()
    
    if (-not [ArchiveTools]::WinRarPath) {
        throw "WinRAR не найден. Установите WinRAR для работы с .rar файлами"
    }
    
    Write-Log "Анализ RAR архива с помощью: $([ArchiveTools]::WinRarPath)" "Verbose" $LogLevel
    
    # Получение списка файлов в архиве
    $listResult = & ([ArchiveTools]::WinRarPath) "l" "-cfg-" "$ArchivePath" 2>&1
    
    if ($LASTEXITCODE -ne 0) {
        throw "Ошибка чтения RAR архива: $($listResult -join "`n")"
    }
    
    # Парсинг вывода и поиск PDF файлов
    $pdfFiles = @()
    $inFileList = $false
    
    foreach ($line in $listResult) {
        if ($line -match "^-{15,}") {
            $inFileList = -not $inFileList
            continue
        }
        
        if ($inFileList -and $line -match "\s+(\S.*\.pdf)\s*$") {
            $fileName = $matches[1].Trim()
            $pdfFiles += $fileName
        }
    }
    
    return $pdfFiles
}

# Функция получения списка PDF файлов из 7-Zip архива
function Get-SevenZipPdfFiles {
    param(
        [string]$ArchivePath,
        [string]$LogLevel
    )
    
    [ArchiveTools]::Initialize()
    
    if (-not [ArchiveTools]::SevenZipPath) {
        throw "7-Zip не найден. Установите 7-Zip для работы с .7z файлами"
    }
    
    Write-Log "Анализ 7-Zip архива с помощью: $([ArchiveTools]::SevenZipPath)" "Verbose" $LogLevel
    
    # Получение списка файлов в архиве
    $listResult = & ([ArchiveTools]::SevenZipPath) "l" "-slt" "$ArchivePath" 2>&1
    
    if ($LASTEXITCODE -ne 0) {
        throw "Ошибка чтения 7-Zip архива: $($listResult -join "`n")"
    }
    
    # Парсинг вывода и поиск PDF файлов
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

# Функция извлечения файла из RAR архива
function Extract-RarFile {
    param(
        [string]$ArchivePath,
        [string]$FileName,
        [string]$OutputPath,
        [string]$LogLevel
    )
    
    $outputDir = Split-Path $OutputPath -Parent
    if (-not (New-DirectorySafe $outputDir $LogLevel)) {
        throw "Не удалось создать выходную директорию"
    }
    
    # Извлечение конкретного файла
    $extractResult = & ([ArchiveTools]::WinRarPath) "e" "-cfg-" "-o+" "$ArchivePath" "$FileName" "$outputDir" 2>&1
    
    if ($LASTEXITCODE -ne 0) {
        throw "Ошибка извлечения файла '$FileName': $($extractResult -join "`n")"
    }
    
    # Перемещение файла в правильное место с сохранением структуры
    $extractedFile = Join-Path $outputDir (Split-Path $FileName -Leaf)
    if (Test-Path $extractedFile) {
        if ($extractedFile -ne $OutputPath) {
            Move-Item $extractedFile $OutputPath -Force
        }
    }
}

# Функция извлечения файла из 7-Zip архива
function Extract-SevenZipFile {
    param(
        [string]$ArchivePath,
        [string]$FileName,
        [string]$OutputPath,
        [string]$LogLevel
    )
    
    $outputDir = Split-Path $OutputPath -Parent
    if (-not (New-DirectorySafe $outputDir $LogLevel)) {
        throw "Не удалось создать выходную директорию"
    }
    
    # Извлечение с сохранением структуры папок
    $tempDir = Join-Path $env:TEMP "7zip_extract_$(Get-Random)"
    New-DirectorySafe $tempDir $LogLevel | Out-Null
    
    try {
        $extractResult = & ([ArchiveTools]::SevenZipPath) "e" "$ArchivePath" "-o$tempDir" "$FileName" 2>&1
        
        if ($LASTEXITCODE -ne 0) {
            throw "Ошибка извлечения файла '$FileName': $($extractResult -join "`n")"
        }
        
        # Перемещение файла в правильное место
        $extractedFile = Join-Path $tempDir (Split-Path $FileName -Leaf)
        if (Test-Path $extractedFile) {
            Move-Item $extractedFile $OutputPath -Force
        }
    }
    finally {
        if (Test-Path $tempDir) {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# Функция для безопасного создания директории с обработкой ошибок
function New-DirectorySafe {
    param(
        [string]$Path,
        [string]$LogLevel
    )
    
    try {
        if (-not (Test-Path $Path)) {
            $null = New-Item -ItemType Directory -Path $Path -Force
            Write-Log "Создана директория: $Path" "Verbose" $LogLevel
        }
        return $true
    }
    catch {
        Write-Log "Ошибка создания директории '$Path': $($_.Exception.Message)" "Error" $LogLevel
        return $false
    }
}

# Основная функция селективного извлечения PDF файлов (универсальная)
function Invoke-SelectivePdfExtraction {
    param(
        [string]$ArchivePath,
        [string]$OutputDirectory,
        [bool]$PreserveStructure,
        [bool]$Overwrite,
        [bool]$ShowProgressBar,
        [string]$LogLevel
    )
    
    $result = [ExtractionResult]::new()
    
    try {
        Write-Log "Начинаем селективное извлечение PDF файлов" "Info" $LogLevel
        Write-Log "Источник: $ArchivePath" "Verbose" $LogLevel
        Write-Log "Назначение: $OutputDirectory" "Verbose" $LogLevel
        
        # Определение типа архива
        $archiveType = Get-ArchiveType $ArchivePath
        Write-Log "Тип архива: $archiveType" "Info" $LogLevel
        
        # Создание выходной директории
        if (-not (New-DirectorySafe $OutputDirectory $LogLevel)) {
            throw "Не удалось создать выходную директорию"
        }
        
        $result.TotalArchiveSize = (Get-Item $ArchivePath).Length
        
        # Получение списка PDF файлов в зависимости от типа архива
        $pdfFiles = @()
        switch ($archiveType) {
            ([ArchiveType]::Zip) {
                $pdfFiles = Get-ZipPdfFiles $ArchivePath $LogLevel
            }
            ([ArchiveType]::Rar) {
                $pdfFiles = Get-RarPdfFiles $ArchivePath $LogLevel
            }
            ([ArchiveType]::SevenZip) {
                $pdfFiles = Get-SevenZipPdfFiles $ArchivePath $LogLevel
            }
        }
        
        $result.PdfFilesFound = $pdfFiles.Count
        Write-Log "Обнаружено PDF файлов: $($result.PdfFilesFound)" "Success" $LogLevel
        
        if ($result.PdfFilesFound -eq 0) {
            Write-Log "В архиве не найдено PDF файлов для извлечения" "Warning" $LogLevel
            return $result
        }
        
        # Селективное извлечение PDF файлов
        $extractedCount = 0
        foreach ($fileName in $pdfFiles) {
            try {
                $extractedCount++
                
                # Определение пути назначения
                $relativePath = if ($PreserveStructure) {
                    $fileName
                } else {
                    Split-Path $fileName -Leaf
                }
                
                $outputPath = Join-Path $OutputDirectory $relativePath
                
                # Проверка на перезапись
                if ((Test-Path $outputPath) -and -not $Overwrite) {
                    Write-Log "Пропуск существующего файла: $relativePath" "Warning" $LogLevel
                    continue
                }
                
                # Извлечение файла в зависимости от типа архива
                switch ($archiveType) {
                    ([ArchiveType]::Zip) {
                        Extract-ZipFile $ArchivePath $fileName $outputPath $LogLevel
                    }
                    ([ArchiveType]::Rar) {
                        Extract-RarFile $ArchivePath $fileName $outputPath $LogLevel
                    }
                    ([ArchiveType]::SevenZip) {
                        Extract-SevenZipFile $ArchivePath $fileName $outputPath $LogLevel
                    }
                }
                
                # Подсчет статистики
                if (Test-Path $outputPath) {
                    $fileSize = (Get-Item $outputPath).Length
                    $result.ExtractedPdfSize += $fileSize
                    $result.ExtractedFiles += $relativePath
                    $result.PdfFilesExtracted++
                    
                    Write-Log "Извлечен: $relativePath ($([math]::Round($fileSize/1KB, 1)) КБ)" "Success" $LogLevel
                }
                
                # Обновление прогресса
                if ($ShowProgressBar) {
                    $percentComplete = [math]::Round(($extractedCount / $result.PdfFilesFound) * 100)
                    Write-Progress -Activity "Извлечение PDF файлов" -Status "Обработано $extractedCount из $($result.PdfFilesFound)" -PercentComplete $percentComplete
                }
            }
            catch {
                $result.ErrorsOccurred++
                $errorMsg = "Ошибка извлечения '$fileName': $($_.Exception.Message)"
                $result.Errors += $errorMsg
                Write-Log $errorMsg "Error" $LogLevel
            }
        }
        
        if ($ShowProgressBar) {
            Write-Progress -Activity "Извлечение PDF файлов" -Completed
        }
    }
    catch {
        $result.ErrorsOccurred++
        $errorMsg = "Критическая ошибка: $($_.Exception.Message)"
        $result.Errors += $errorMsg
        Write-Log $errorMsg "Error" $LogLevel
        throw
    }
    
    return $result
}

# Функция получения списка PDF файлов из ZIP архива
function Get-ZipPdfFiles {
    param(
        [string]$ArchivePath,
        [string]$LogLevel
    )
    
    $archiveStream = [System.IO.File]::OpenRead($ArchivePath)
    $archive = [System.IO.Compression.ZipArchive]::new($archiveStream, [System.IO.Compression.ZipArchiveMode]::Read)
    
    try {
        $pdfEntries = $archive.Entries | Where-Object { 
            $_.Name -match '\.pdf$' -and $_.Length -gt 0 
        }
        
        return $pdfEntries | ForEach-Object { $_.FullName }
    }
    finally {
        $archive.Dispose()
        $archiveStream.Dispose()
    }
}

# Функция извлечения файла из ZIP архива
function Extract-ZipFile {
    param(
        [string]$ArchivePath,
        [string]$FileName,
        [string]$OutputPath,
        [string]$LogLevel
    )
    
    $outputDir = Split-Path $OutputPath -Parent
    if (-not (New-DirectorySafe $outputDir $LogLevel)) {
        throw "Не удалось создать выходную директорию"
    }
    
    $archiveStream = [System.IO.File]::OpenRead($ArchivePath)
    $archive = [System.IO.Compression.ZipArchive]::new($archiveStream, [System.IO.Compression.ZipArchiveMode]::Read)
    
    try {
        $entry = $archive.Entries | Where-Object { $_.FullName -eq $FileName } | Select-Object -First 1
        
        if (-not $entry) {
            throw "Файл '$FileName' не найден в архиве"
        }
        
        $entryStream = $entry.Open()
        try {
            $outputFileStream = [System.IO.File]::Create($OutputPath)
            try {
                $entryStream.CopyTo($outputFileStream)
            }
            finally {
                $outputFileStream.Dispose()
            }
        }
        finally {
            $entryStream.Dispose()
        }
    }
    finally {
        $archive.Dispose()
        $archiveStream.Dispose()
    }
}

# Функция вывода финального отчета
function Show-ExtractionReport {
    param(
        [ExtractionResult]$Result,
        [string]$LogLevel
    )
    
    Write-Log "`n=== ОТЧЕТ О ВЫПОЛНЕНИИ ОПЕРАЦИИ ===" "Success" $LogLevel
    Write-Log "Записей в архиве: $($Result.TotalEntriesInArchive)" "Info" $LogLevel
    Write-Log "PDF файлов найдено: $($Result.PdfFilesFound)" "Info" $LogLevel
    Write-Log "PDF файлов извлечено: $($Result.PdfFilesExtracted)" "Success" $LogLevel
    
    if ($Result.ErrorsOccurred -gt 0) {
        Write-Log "Ошибок: $($Result.ErrorsOccurred)" "Error" $LogLevel
    }
    
    # Анализ эффективности
    $compressionRatio = if ($Result.TotalArchiveSize -gt 0) {
        [math]::Round((1 - ($Result.ExtractedPdfSize / $Result.TotalArchiveSize)) * 100, 1)
    } else { 0 }
    
    Write-Log "`n=== АНАЛИЗ ЭФФЕКТИВНОСТИ ===" "Info" $LogLevel
    Write-Log "Размер архива: $([math]::Round($Result.TotalArchiveSize / 1MB, 2)) МБ" "Info" $LogLevel
    Write-Log "Размер извлеченных PDF: $([math]::Round($Result.ExtractedPdfSize / 1MB, 2)) МБ" "Success" $LogLevel
    Write-Log "Эффективность фильтрации: $compressionRatio%" "Success" $LogLevel
    
    if ($LogLevel -eq "Verbose" -and $Result.ExtractedFiles.Count -gt 0) {
        Write-Log "`n=== СПИСОК ИЗВЛЕЧЕННЫХ ФАЙЛОВ ===" "Verbose" $LogLevel
        $Result.ExtractedFiles | ForEach-Object {
            Write-Log "  • $_" "Verbose" $LogLevel
        }
    }
    
    if ($Result.Errors.Count -gt 0) {
        Write-Log "`n=== ОШИБКИ ===" "Error" $LogLevel
        $Result.Errors | ForEach-Object {
            Write-Log "  • $_" "Error" $LogLevel
        }
    }
}

# Основная точка входа
try {
    Write-Log "=== СЕЛЕКТИВНОЕ ИЗВЛЕЧЕНИЕ PDF ФАЙЛОВ ===" "Success" $LogLevel
    Write-Log "PowerShell версия: $($PSVersionTable.PSVersion)" "Verbose" $LogLevel
    Write-Log "Операционная система: $([System.Environment]::OSVersion.VersionString)" "Verbose" $LogLevel
    
    # Валидация параметров
    $SourceArchive = Resolve-Path $SourceArchive -ErrorAction Stop
    
    # Выполнение селективного извлечения
    $result = Invoke-SelectivePdfExtraction -ArchivePath $SourceArchive -OutputDirectory $DestinationDirectory -PreserveStructure $IncludeSubdirectories -Overwrite $OverwriteExisting -ShowProgressBar $ShowProgress -LogLevel $LogLevel
    
    # Вывод отчета
    Show-ExtractionReport -Result $result -LogLevel $LogLevel
    
    Write-Log "`n🎉 ОПЕРАЦИЯ ЗАВЕРШЕНА УСПЕШНО!" "Success" $LogLevel
    
    # Возврат результата для программного использования
    return $result
}
catch {
    Write-Log "`n💥 ОПЕРАЦИЯ ЗАВЕРШЕНА С ОШИБКОЙ!" "Error" $LogLevel
    Write-Log "Детали: $($_.Exception.Message)" "Error" $LogLevel
    
    if ($LogLevel -eq "Verbose") {
        Write-Log "Стек вызовов:" "Verbose" $LogLevel
        Write-Log $_.ScriptStackTrace "Verbose" $LogLevel
    }
    
    exit 1
} 