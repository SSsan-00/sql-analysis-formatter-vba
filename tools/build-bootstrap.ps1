param(
    [ValidateSet('User', 'Developer')]
    [string]$Audience = 'User',
    [string]$ParserExePath = '',
    [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
if ([string]::IsNullOrWhiteSpace($ParserExePath)) {
    $ParserExePath = Join-Path $repoRoot 'dist\parser\SqlAnalysisFormatter.Parser.exe'
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $suffix = if ($Audience -eq 'Developer') { 'developer' } else { 'user' }
    $OutputPath = Join-Path $repoRoot "dist\bootstrap\SqlAnalysisFormatter.$suffix.bootstrap.ps1"
}

if (-not (Test-Path -LiteralPath $ParserExePath)) {
    & (Join-Path $PSScriptRoot 'publish-parser.ps1') | Out-Null
}

function ConvertTo-GzipBase64 {
    param([byte[]]$Bytes)

    $output = [System.IO.MemoryStream]::new()
    $gzip = [System.IO.Compression.GZipStream]::new($output, [System.IO.Compression.CompressionLevel]::Optimal, $true)
    $gzip.Write($Bytes, 0, $Bytes.Length)
    $gzip.Dispose()
    try {
        [Convert]::ToBase64String($output.ToArray())
    } finally {
        $output.Dispose()
    }
}

function Split-Base64Literal {
    param([string]$Value)

    $chunkSize = 120
    $chunks = New-Object System.Collections.Generic.List[string]
    for ($index = 0; $index -lt $Value.Length; $index += $chunkSize) {
        $length = [Math]::Min($chunkSize, $Value.Length - $index)
        $chunks.Add($Value.Substring($index, $length))
    }

    ($chunks | ForEach-Object { "            '$_'" }) -join ",`r`n"
}

function Escape-SingleQuoted {
    param([string]$Value)

    $Value.Replace("'", "''")
}

function New-Artifact {
    param(
        [string]$SourcePath,
        [string]$TargetPath
    )

    [PSCustomObject]@{
        SourcePath = $SourcePath
        TargetPath = $TargetPath
    }
}

function Test-BootstrapSourcePath {
    param([string]$Path)

    $fileName = [IO.Path]::GetFileName($Path)
    $fileName -like '*bootstrap*.ps1'
}

function Assert-NoBootstrapSourceArtifact {
    param([System.Collections.Generic.List[object]]$Artifacts)

    foreach ($artifact in $Artifacts) {
        if ((Test-BootstrapSourcePath $artifact.SourcePath) -or (Test-BootstrapSourcePath $artifact.TargetPath)) {
            throw "Bootstrap-related source must not be included: $($artifact.TargetPath)"
        }
    }
}

function Add-TreeArtifacts {
    param(
        [System.Collections.Generic.List[object]]$Artifacts,
        [string]$SourceDirectory,
        [string]$TargetDirectory
    )

    $basePath = (Resolve-Path $SourceDirectory).Path.TrimEnd('\') + '\'
    $excludedDirectories = @('\bin\', '\obj\', '\TestResults\')
    Get-ChildItem -LiteralPath $SourceDirectory -Recurse -File |
        Where-Object {
            $fullName = $_.FullName
            -not ($excludedDirectories | Where-Object { $fullName.Contains($_) })
        } |
        ForEach-Object {
            $relativePath = $_.FullName.Substring($basePath.Length)
            $Artifacts.Add((New-Artifact $_.FullName (Join-Path $TargetDirectory $relativePath)))
        }
}

function Get-UserArtifacts {
    $artifacts = [System.Collections.Generic.List[object]]::new()
    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'src\vba\SqlAnalysisFormatter.bas') 'SqlAnalysisFormatter.bas'))
    $artifacts.Add((New-Artifact $ParserExePath 'SqlAnalysisFormatter.Parser.exe'))
    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'docs\USER_GUIDE.md') 'README.md'))
    $artifacts
}

function Get-DeveloperArtifacts {
    $artifacts = [System.Collections.Generic.List[object]]::new()
    foreach ($artifact in (Get-UserArtifacts)) {
        $artifacts.Add($artifact)
    }

    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'docs\DEVELOPER_GUIDE.md') 'docs\DEVELOPER_GUIDE.md'))
    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'docs\PROVISIONAL_OUTPUT_CASES.md') 'docs\PROVISIONAL_OUTPUT_CASES.md'))
    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'SqlAnalysisFormatter.sln') 'SqlAnalysisFormatter.sln'))
    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'tests\CRUD_TEST_CASES.md') 'tests\CRUD_TEST_CASES.md'))
    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'tests\ManualOutputCases.json') 'tests\ManualOutputCases.json'))
    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'tests\OutputReportCases.json') 'tests\OutputReportCases.json'))
    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'tests\SqlAnalysisFormatter.OutputExpectations.xlsx') 'tests\SqlAnalysisFormatter.OutputExpectations.xlsx'))
    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'tools\publish-parser.ps1') 'tools\publish-parser.ps1'))
    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'tools\run-vba-tests.ps1') 'tools\run-vba-tests.ps1'))
    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'tools\run-output-golden-tests.ps1') 'tools\run-output-golden-tests.ps1'))
    $artifacts.Add((New-Artifact (Join-Path $repoRoot 'tools\Set-ManualOutputCase.ps1') 'tools\Set-ManualOutputCase.ps1'))

    Add-TreeArtifacts $artifacts (Join-Path $repoRoot 'src\vba') 'src\vba'
    Add-TreeArtifacts $artifacts (Join-Path $repoRoot 'tools\SqlAnalysisFormatter.Parser') 'tools\SqlAnalysisFormatter.Parser'
    Add-TreeArtifacts $artifacts (Join-Path $repoRoot 'tests\SqlAnalysisFormatter.Parser.Tests') 'tests\SqlAnalysisFormatter.Parser.Tests'

    $artifacts
}

$artifacts = if ($Audience -eq 'Developer') {
    Get-DeveloperArtifacts
} else {
    Get-UserArtifacts
}
Assert-NoBootstrapSourceArtifact $artifacts

$fileBlocks = New-Object System.Collections.Generic.List[string]
foreach ($artifact in $artifacts) {
    $sourcePath = Resolve-Path $artifact.SourcePath
    $targetPath = (Escape-SingleQuoted ($artifact.TargetPath -replace '\\', '/'))
    $chunks = Split-Base64Literal (ConvertTo-GzipBase64 ([IO.File]::ReadAllBytes($sourcePath)))
    $fileBlocks.Add(@"
    @{
        Path = '$targetPath'
        Chunks = @(
$chunks
        )
    }
"@)
}

$manifestText = $fileBlocks -join ",`r`n"
$bootstrap = @"
param(
    [string]`$TargetDirectory = ''
)

`$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace(`$TargetDirectory)) {
    `$TargetDirectory = Join-Path (Get-Location) 'SqlAnalysisFormatter'
}

function Expand-GzipBase64ToFile {
    param(
        [string[]]`$Chunks,
        [string]`$Path
    )

    `$parent = Split-Path -Parent `$Path
    if (-not [string]::IsNullOrWhiteSpace(`$parent)) {
        New-Item -ItemType Directory -Path `$parent -Force | Out-Null
    }

    `$compressedBytes = [Convert]::FromBase64String((`$Chunks -join ''))
    `$inputStream = [System.IO.MemoryStream]::new(`$compressedBytes)
    `$gzip = [System.IO.Compression.GZipStream]::new(`$inputStream, [System.IO.Compression.CompressionMode]::Decompress)
    `$outputStream = [System.IO.File]::Create(`$Path)
    try {
        `$gzip.CopyTo(`$outputStream)
    } finally {
        `$outputStream.Dispose()
        `$gzip.Dispose()
        `$inputStream.Dispose()
    }
}

New-Item -ItemType Directory -Path `$TargetDirectory -Force | Out-Null

`$files = @(
$manifestText
)

foreach (`$file in `$files) {
    Expand-GzipBase64ToFile `$file.Chunks (Join-Path `$TargetDirectory `$file.Path)
}

Write-Output "Created: `$TargetDirectory"
"@

New-Item -ItemType Directory -Path (Split-Path -Parent $OutputPath) -Force | Out-Null
[IO.File]::WriteAllText($OutputPath, $bootstrap, [Text.UTF8Encoding]::new($false))
Write-Output $OutputPath
