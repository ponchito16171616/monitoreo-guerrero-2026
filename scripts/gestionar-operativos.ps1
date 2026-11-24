[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("List", "Validate", "SetCatalog", "Add", "Replace", "Rename", "Remove")]
    [string]$Action,

    [string]$Id,
    [string]$Name,
    [string]$SourceFile,
    [string]$PeriodStart,
    [string]$PeriodEnd,
    [switch]$ConfirmRemoval,
    [string]$RepositoryRoot
)

<#
.SYNOPSIS
Administra el manifiesto público de operativos desde una copia local del repositorio.

.EXAMPLE
.\scripts\gestionar-operativos.ps1 -Action SetCatalog -SourceFile "C:\Datos\catalogos.xlsx"

.EXAMPLE
.\scripts\gestionar-operativos.ps1 -Action Add -Id operativo-2026-01 -Name "Operativo enero" -SourceFile "C:\Datos\enero.xlsx" -PeriodStart 2026-01-01 -PeriodEnd 2026-01-31

.EXAMPLE
.\scripts\gestionar-operativos.ps1 -Action Replace -Id operativo-2026-01 -SourceFile "C:\Datos\enero-corregido.xlsx"

.EXAMPLE
.\scripts\gestionar-operativos.ps1 -Action Rename -Id operativo-2026-01 -Name "Operativo enero actualizado"

.EXAMPLE
.\scripts\gestionar-operativos.ps1 -Action Remove -Id operativo-2026-01 -ConfirmRemoval

.EXAMPLE
.\scripts\gestionar-operativos.ps1 -Action Validate

Después de revisar los cambios, publícalos desde la raíz del repositorio con git add, git commit y git push.
Remove elimina el archivo de la rama actual; Git conserva las versiones anteriores en su historial.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
    $RepositoryRoot = Split-Path -Parent $PSScriptRoot
}

$RepositoryRoot = [System.IO.Path]::GetFullPath($RepositoryRoot)
$DataRoot = Join-Path $RepositoryRoot "data"
$ManifestPath = Join-Path $DataRoot "manifest.json"

function Assert-Identifier {
    param([string]$Value)

    if ($Value -notmatch "^[a-z0-9][a-z0-9._-]{0,79}$") {
        throw "El identificador debe usar minúsculas, números, punto, guion o guion bajo; debe iniciar con letra o número y tener hasta 80 caracteres."
    }
}

function Assert-Period {
    param(
        [string]$Start,
        [string]$End
    )

    $format = "yyyy-MM-dd"
    $style = [System.Globalization.DateTimeStyles]::None
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $startDate = [datetime]::MinValue
    $endDate = [datetime]::MinValue

    if ($Start -and -not [datetime]::TryParseExact($Start, $format, $culture, $style, [ref]$startDate)) {
        throw "La fecha inicial debe tener formato AAAA-MM-DD."
    }
    if ($End -and -not [datetime]::TryParseExact($End, $format, $culture, $style, [ref]$endDate)) {
        throw "La fecha final debe tener formato AAAA-MM-DD."
    }
    if ($Start -and $End -and $endDate -lt $startDate) {
        throw "La fecha final no puede ser anterior a la fecha inicial."
    }
}

function Get-WorkbookSource {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "Indica el archivo Excel con -SourceFile."
    }

    $resolved = (Resolve-Path -LiteralPath $Path).Path
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw "No se encontró el archivo: $Path"
    }
    if ([System.IO.Path]::GetExtension($resolved) -ine ".xlsx") {
        throw "Solo se admiten libros .xlsx."
    }

    return $resolved
}

function Get-DataFilePath {
    param([string]$RelativePath)

    if ([string]::IsNullOrWhiteSpace($RelativePath) -or [System.IO.Path]::IsPathRooted($RelativePath)) {
        throw "La ruta del manifiesto debe ser relativa a la carpeta data."
    }

    $normalized = $RelativePath.Replace("/", [System.IO.Path]::DirectorySeparatorChar)
    $fullPath = [System.IO.Path]::GetFullPath((Join-Path $DataRoot $normalized))
    $dataPrefix = [System.IO.Path]::GetFullPath($DataRoot).TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar

    if (-not $fullPath.StartsWith($dataPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "La ruta del manifiesto debe permanecer dentro de data."
    }

    return $fullPath
}

function Read-Manifest {
    param([switch]$AllowMissing)

    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
        if ($AllowMissing) {
            return [pscustomobject]@{
                version = "1"
                catalog = ""
                operatives = @()
            }
        }
        throw "No existe data\manifest.json. Configura primero el catálogo con -Action SetCatalog."
    }

    try {
        $raw = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json
    } catch {
        throw "No se pudo leer data\manifest.json: $($_.Exception.Message)"
    }

    if ($null -eq $raw -or $raw -isnot [pscustomobject] -or
        $null -eq $raw.PSObject.Properties["catalog"] -or
        $null -eq $raw.PSObject.Properties["operatives"]) {
        throw "El manifiesto debe incluir las propiedades catalog y operatives."
    }

    if ($raw.operatives -is [array]) {
        $operatives = @($raw.operatives)
    } elseif ($null -eq $raw.operatives) {
        $operatives = @()
    } else {
        $operatives = @($raw.operatives)
    }

    return [pscustomobject]@{
        version = if ($raw.PSObject.Properties["version"] -and $raw.version) { [string]$raw.version } else { "1" }
        catalog = [string]$raw.catalog
        operatives = $operatives
    }
}

function Write-Manifest {
    param([pscustomobject]$Manifest)

    New-Item -ItemType Directory -Path $DataRoot -Force | Out-Null
    $document = [ordered]@{
        version = (Get-Date).ToUniversalTime().ToString("yyyyMMdd-HHmmss")
        catalog = $Manifest.catalog
        operatives = @($Manifest.operatives)
    }
    $json = ConvertTo-Json -InputObject $document -Depth 10
    $encoding = New-Object System.Text.UTF8Encoding($false)
    $temporaryPath = "$ManifestPath.tmp"

    [System.IO.File]::WriteAllText($temporaryPath, $json + [Environment]::NewLine, $encoding)
    Move-Item -LiteralPath $temporaryPath -Destination $ManifestPath -Force
}

function Copy-Workbook {
    param(
        [string]$Source,
        [string]$Destination
    )

    $directory = Split-Path -Parent $Destination
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $temporaryPath = "$Destination.new"
    try {
        Copy-Item -LiteralPath $Source -Destination $temporaryPath -Force
        Move-Item -LiteralPath $temporaryPath -Destination $Destination -Force
    } finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
    }
}

function Test-Manifest {
    param([pscustomobject]$Manifest)

    $problems = New-Object System.Collections.Generic.List[string]
    $seenIds = @{}
    $seenFiles = @{}
    $catalogPath = $null

    if ([string]::IsNullOrWhiteSpace($Manifest.catalog)) {
        $problems.Add("Falta la ruta del catálogo.")
    } else {
        try {
            $catalogPath = Get-DataFilePath $Manifest.catalog
            if (-not (Test-Path -LiteralPath $catalogPath -PathType Leaf)) {
                $problems.Add("No se encontró el catálogo: $($Manifest.catalog)")
            }
        } catch {
            $problems.Add("Ruta de catálogo inválida: $($_.Exception.Message)")
        }
    }

    foreach ($operative in $Manifest.operatives) {
        if ($null -eq $operative -or
            $null -eq $operative.PSObject.Properties["id"] -or
            $null -eq $operative.PSObject.Properties["file"]) {
            $problems.Add("Hay un operativo sin identificador o ruta de archivo.")
            continue
        }

        $operativeId = [string]$operative.id
        $operativeFile = [string]$operative.file

        if ([string]::IsNullOrWhiteSpace($operativeId)) {
            $problems.Add("Hay un operativo con identificador vacío.")
        } elseif ($seenIds.ContainsKey($operativeId)) {
            $problems.Add("El identificador '$operativeId' está duplicado.")
        } else {
            $seenIds[$operativeId] = $true
        }

        try {
            $filePath = Get-DataFilePath $operativeFile
            if ($catalogPath -and [string]::Equals($filePath, $catalogPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                $problems.Add("El archivo del operativo '$operativeId' no puede ser el mismo que el catálogo.")
            }
            if ([System.IO.Path]::GetExtension($filePath) -ine ".xlsx") {
                $problems.Add("El operativo '$operativeId' no apunta a un archivo .xlsx.")
            }
            if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
                $problems.Add("No se encontró el archivo de '$operativeId': $operativeFile")
            }

            $fileKey = $filePath.ToLowerInvariant()
            if ($seenFiles.ContainsKey($fileKey)) {
                $problems.Add("El archivo '$operativeFile' está asignado a más de un operativo.")
            } else {
                $seenFiles[$fileKey] = $true
            }
        } catch {
            $problems.Add("Ruta inválida para '$operativeId': $($_.Exception.Message)")
        }
    }

    return ,$problems
}

function Remove-PreviousFileIfUnused {
    param(
        [string]$PreviousRelativePath,
        [pscustomobject]$Manifest
    )

    if ([string]::IsNullOrWhiteSpace($PreviousRelativePath)) {
        return
    }

    $previousPath = Get-DataFilePath $PreviousRelativePath
    $isCatalog = $false
    if (-not [string]::IsNullOrWhiteSpace($Manifest.catalog)) {
        $catalogPath = Get-DataFilePath $Manifest.catalog
        $isCatalog = [string]::Equals($catalogPath, $previousPath, [System.StringComparison]::OrdinalIgnoreCase)
    }
    $isStillReferenced = $Manifest.operatives | Where-Object {
        if (-not $_.file) { return $false }
        try {
            [string]::Equals((Get-DataFilePath ([string]$_.file)), $previousPath, [System.StringComparison]::OrdinalIgnoreCase)
        } catch {
            $false
        }
    }

    if (-not $isCatalog -and -not $isStillReferenced -and
        [System.IO.Path]::GetExtension($previousPath) -ieq ".xlsx" -and
        (Test-Path -LiteralPath $previousPath -PathType Leaf)) {
        Remove-Item -LiteralPath $previousPath -Force
    }
}

switch ($Action) {
    "List" {
        $manifest = Read-Manifest
        if (-not $manifest.operatives.Count) {
            Write-Output "No hay operativos registrados."
        } else {
            $manifest.operatives |
                Select-Object id, name, file, periodStart, periodEnd |
                Format-Table -AutoSize
        }
    }

    "Validate" {
        $manifest = Read-Manifest
        $problems = Test-Manifest $manifest
        if ($problems.Count) {
            $problems | ForEach-Object { Write-Output "ERROR: $_" }
            throw "La validación encontró $($problems.Count) problema(s)."
        }
        Write-Output "Manifiesto válido: el catálogo y los $($manifest.operatives.Count) operativos referenciados existen."
    }

    "SetCatalog" {
        $source = Get-WorkbookSource $SourceFile
        $manifest = Read-Manifest -AllowMissing
        $oldCatalog = $manifest.catalog
        $manifest.catalog = "catalogos.xlsx"
        $destination = Get-DataFilePath $manifest.catalog
        Copy-Workbook -Source $source -Destination $destination
        Write-Manifest $manifest
        if ($oldCatalog -and $oldCatalog -ne $manifest.catalog) {
            Remove-PreviousFileIfUnused -PreviousRelativePath $oldCatalog -Manifest $manifest
        }
        Write-Output "Catálogo actualizado: data\catalogos.xlsx"
    }

    "Add" {
        Assert-Identifier $Id
        Assert-Period -Start $PeriodStart -End $PeriodEnd
        if ([string]::IsNullOrWhiteSpace($Name)) {
            throw "Indica un nombre visible con -Name."
        }
        $source = Get-WorkbookSource $SourceFile
        $manifest = Read-Manifest
        if ([string]::IsNullOrWhiteSpace($manifest.catalog)) {
            throw "Configura primero el catálogo con -Action SetCatalog."
        }
        if ($manifest.operatives | Where-Object { [string]$_.id -eq $Id }) {
            throw "Ya existe el identificador '$Id'. Usa -Action Replace para corregir sus datos."
        }

        $relativeFile = "operativos/$Id.xlsx"
        $destination = Get-DataFilePath $relativeFile
        if (Test-Path -LiteralPath $destination) {
            throw "Ya existe $relativeFile pero no está registrado. Revísalo antes de sobrescribir."
        }

        Copy-Workbook -Source $source -Destination $destination
        $entry = [ordered]@{
            id = $Id
            name = $Name.Trim()
            file = $relativeFile
            periodStart = $PeriodStart
            periodEnd = $PeriodEnd
            updatedAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd")
        }
        $manifest.operatives = @($manifest.operatives) + @([pscustomobject]$entry)
        Write-Manifest $manifest
        Write-Output "Operativo agregado: $Name ($Id)."
    }

    "Replace" {
        Assert-Identifier $Id
        $source = Get-WorkbookSource $SourceFile
        $manifest = Read-Manifest
        $entry = $manifest.operatives | Where-Object { [string]$_.id -eq $Id } | Select-Object -First 1
        if (-not $entry) {
            throw "No existe '$Id'. Usa -Action Add para registrar un operativo nuevo."
        }
        $existingStart = if ($entry.PSObject.Properties["periodStart"]) { [string]$entry.periodStart } else { "" }
        $existingEnd = if ($entry.PSObject.Properties["periodEnd"]) { [string]$entry.periodEnd } else { "" }
        $newStart = if ($PSBoundParameters.ContainsKey("PeriodStart")) { $PeriodStart } else { $existingStart }
        $newEnd = if ($PSBoundParameters.ContainsKey("PeriodEnd")) { $PeriodEnd } else { $existingEnd }
        Assert-Period -Start $newStart -End $newEnd

        $destination = Get-DataFilePath ([string]$entry.file)
        Copy-Workbook -Source $source -Destination $destination
        $entry | Add-Member -NotePropertyName periodStart -NotePropertyValue $newStart -Force
        $entry | Add-Member -NotePropertyName periodEnd -NotePropertyValue $newEnd -Force
        $entry | Add-Member -NotePropertyName updatedAt -NotePropertyValue ((Get-Date).ToUniversalTime().ToString("yyyy-MM-dd")) -Force
        Write-Manifest $manifest
        Write-Output "Datos reemplazados para '$($entry.name)' ($Id). Se conservó su identificador."
    }

    "Rename" {
        Assert-Identifier $Id
        if ([string]::IsNullOrWhiteSpace($Name)) {
            throw "Indica el nuevo nombre visible con -Name."
        }
        $manifest = Read-Manifest
        $entry = $manifest.operatives | Where-Object { [string]$_.id -eq $Id } | Select-Object -First 1
        if (-not $entry) {
            throw "No existe el identificador '$Id'."
        }

        $entry.name = $Name.Trim()
        Write-Manifest $manifest
        Write-Output "Nombre visible actualizado. El identificador y la ruta del Excel no cambiaron."
    }

    "Remove" {
        Assert-Identifier $Id
        if (-not $ConfirmRemoval) {
            throw "Para retirar el operativo también del repositorio actual, confirma la acción agregando -ConfirmRemoval."
        }
        $manifest = Read-Manifest
        $entry = $manifest.operatives | Where-Object { [string]$_.id -eq $Id } | Select-Object -First 1
        if (-not $entry) {
            throw "No existe el identificador '$Id'."
        }

        $previousFile = [string]$entry.file
        $manifest.operatives = @($manifest.operatives | Where-Object { [string]$_.id -ne $Id })
        Write-Manifest $manifest
        Remove-PreviousFileIfUnused -PreviousRelativePath $previousFile -Manifest $manifest
        Write-Output "Operativo '$($entry.name)' retirado de la aplicación y del estado actual del repositorio."
        Write-Output "Las versiones previas del archivo permanecen en el historial de Git."
    }
}
