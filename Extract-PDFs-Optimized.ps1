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
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Путь к ZIP архиву")]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Leaf)) {
            throw "Файл архива не найден: $_"
        }
        if (-not ($_ -match '\.zip$')) {
            throw "Файл должен иметь расширение .zip: $_"
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

# Основная функция селективного извлечения PDF файлов
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
        
        # Создание выходной директории
        if (-not (New-DirectorySafe $OutputDirectory $LogLevel)) {
            throw "Не удалось создать выходную директорию"
        }
        
        # Открытие ZIP архива для чтения (без распаковки)
        $archiveStream = [System.IO.File]::OpenRead($ArchivePath)
        $archive = [System.IO.Compression.ZipArchive]::new($archiveStream, [System.IO.Compression.ZipArchiveMode]::Read)
        
        try {
            # Анализ содержимого архива
            $allEntries = $archive.Entries
            $result.TotalEntriesInArchive = $allEntries.Count
            $result.TotalArchiveSize = (Get-Item $ArchivePath).Length
            
            Write-Log "Всего записей в архиве: $($result.TotalEntriesInArchive)" "Info" $LogLevel
            
            # Фильтрация PDF файлов
            $pdfEntries = $allEntries | Where-Object { 
                $_.Name -match '\.pdf$' -and $_.Length -gt 0 
            }
            
            $result.PdfFilesFound = $pdfEntries.Count
            Write-Log "Обнаружено PDF файлов: $($result.PdfFilesFound)" "Success" $LogLevel
            
            if ($result.PdfFilesFound -eq 0) {
                Write-Log "В архиве не найдено PDF файлов для извлечения" "Warning" $LogLevel
                return $result
            }
            
            # Селективное извлечение PDF файлов
            $extractedCount = 0
            foreach ($entry in $pdfEntries) {
                try {
                    $extractedCount++
                    
                    # Определение пути назначения
                    $relativePath = if ($PreserveStructure) {
                        $entry.FullName
                    } else {
                        $entry.Name
                    }
                    
                    $outputPath = Join-Path $OutputDirectory $relativePath
                    $outputDir = Split-Path $outputPath -Parent
                    
                    # Проверка на перезапись
                    if ((Test-Path $outputPath) -and -not $Overwrite) {
                        Write-Log "Пропуск существующего файла: $relativePath" "Warning" $LogLevel
                        continue
                    }
                    
                    # Создание целевой директории
                    if (-not (New-DirectorySafe $outputDir $LogLevel)) {
                        throw "Не удалось создать директорию для файла"
                    }
                    
                    # Прямое извлечение файла из ZIP без временных файлов
                    $entryStream = $entry.Open()
                    try {
                        $outputFileStream = [System.IO.File]::Create($outputPath)
                        try {
                            $entryStream.CopyTo($outputFileStream)
                            $result.ExtractedPdfSize += $entry.Length
                            $result.ExtractedFiles += $relativePath
                            
                            Write-Log "Извлечен: $relativePath ($([math]::Round($entry.Length/1KB, 1)) КБ)" "Success" $LogLevel
                            
                            # Обновление прогресса
                            if ($ShowProgressBar) {
                                $percentComplete = [math]::Round(($extractedCount / $result.PdfFilesFound) * 100)
                                Write-Progress -Activity "Извлечение PDF файлов" -Status "Обработано $extractedCount из $($result.PdfFilesFound)" -PercentComplete $percentComplete
                            }
                        }
                        finally {
                            $outputFileStream.Dispose()
                        }
                    }
                    finally {
                        $entryStream.Dispose()
                    }
                    
                    $result.PdfFilesExtracted++
                }
                catch {
                    $result.ErrorsOccurred++
                    $errorMsg = "Ошибка извлечения '$($entry.FullName)': $($_.Exception.Message)"
                    $result.Errors += $errorMsg
                    Write-Log $errorMsg "Error" $LogLevel
                }
            }
            
            if ($ShowProgressBar) {
                Write-Progress -Activity "Извлечение PDF файлов" -Completed
            }
        }
        finally {
            $archive.Dispose()
            $archiveStream.Dispose()
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