"""Experimental Microsoft Access object exporter using Access automation."""

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


POWERSHELL_SCRIPT = r'''$ErrorActionPreference = 'Stop'

$workingCopyPath = [Environment]::GetEnvironmentVariable('VBAEXPORT_ACCESS_WORKING_COPY')
$outputPath = [Environment]::GetEnvironmentVariable('VBAEXPORT_ACCESS_OUTPUT')
$access = $null

function ConvertTo-SafeBaseName([string] $name) {
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $builder = New-Object Text.StringBuilder
    foreach ($character in $name.ToCharArray()) {
        if (($invalid -contains $character) -or ([int][char] $character -lt 32)) {
            [void] $builder.Append(('_U{0:X4}_' -f [int][char] $character))
        } else {
            [void] $builder.Append($character)
        }
    }
    $safe = $builder.ToString().TrimEnd([char[]] ' .')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = 'unnamed'
    }
    if ($safe.Split('.')[0] -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$') {
        $safe = '_' + $safe
    }
    return $safe
}

function Get-UniqueBaseName([hashtable] $usedNames, [string] $name) {
    $baseName = ConvertTo-SafeBaseName $name
    $candidate = $baseName
    $number = 2
    while ($usedNames.ContainsKey($candidate.ToUpperInvariant())) {
        $candidate = "$baseName ($number)"
        $number++
    }
    $usedNames[$candidate.ToUpperInvariant()] = $true
    return $candidate
}

function Export-Objects($objects, [int] $objectType, [string] $directory,
                        [string] $extension, [string] $label) {
    $count = 0
    $usedNames = @{}
    foreach ($object in $objects) {
        $name = [string] $object.Name
        $fileName = (Get-UniqueBaseName $usedNames $name) + $extension
        try {
            $access.SaveAsText($objectType, $name, (Join-Path $directory $fileName))
            $count++
        } catch {
            Write-Warning "Failed to export ${label} '$name': $($_.Exception.Message)"
        }
    }
    return $count
}

function Export-Queries([string] $accessDirectory, [string] $sqlDirectory) {
    $rawCount = 0
    $sqlCount = 0
    $usedNames = @{}
    $database = $access.CurrentDb()
    try {
        foreach ($query in $access.CurrentData.AllQueries) {
            $name = [string] $query.Name
            if ($name.StartsWith('~sq', [StringComparison]::OrdinalIgnoreCase)) {
                Write-Warning "Skipping temporary query: $name"
                continue
            }

            $baseName = Get-UniqueBaseName $usedNames $name
            try {
                $access.SaveAsText(1, $name, (Join-Path $accessDirectory ($baseName + '.txt')))
                $rawCount++
            } catch {
                Write-Warning "Failed to export raw query '$name': $($_.Exception.Message)"
            }

            try {
                $sql = $database.QueryDefs.Item($name).SQL
                if ([string]::IsNullOrEmpty([string] $sql)) {
                    Write-Warning "Query does not expose usable SQL: $name"
                    continue
                }
                [IO.File]::WriteAllText((Join-Path $sqlDirectory ($baseName + '.sql')),
                    [string] $sql, [Text.UTF8Encoding]::new($false))
                $sqlCount++
            } catch {
                Write-Warning "Failed to export SQL for query '$name': $($_.Exception.Message)"
            }
        }
    } finally {
        if ($null -ne $database) {
            try { [void] [Runtime.InteropServices.Marshal]::FinalReleaseComObject($database) } catch {}
        }
    }
    return [PSCustomObject]@{ Raw = $rawCount; Sql = $sqlCount }
}

try {
    Write-Output "Opening temporary Access database copy: $workingCopyPath"
    $access = New-Object -ComObject Access.Application
    # Open shared rather than exclusive. Access exposes no read-only parameter here.
    $access.OpenCurrentDatabase($workingCopyPath, $false)

    $modules = Export-Objects $access.CurrentProject.AllModules 5 (Join-Path $outputPath 'Modules') '.bas' 'module'
    $queries = Export-Queries (Join-Path $outputPath 'Queries\Access') (Join-Path $outputPath 'Queries\SQL')
    $macros = Export-Objects $access.CurrentProject.AllMacros 4 (Join-Path $outputPath 'Macros') '.txt' 'macro'

    Write-Output "Export complete: $modules modules, $($queries.Raw) raw query definitions, $($queries.Sql) query SQL files, $macros macros."
    Write-Output "Output directory: $outputPath"
}
finally {
    if ($null -ne $access) {
        try { $access.CloseCurrentDatabase() } catch {}
        try { $access.Quit() } catch {}
        try { [void] [Runtime.InteropServices.Marshal]::FinalReleaseComObject($access) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''


def main():
    if len(sys.argv) != 3:
        print('Usage: export_access_objects.py <Access database path> <output directory>')
        return 1

    source = Path(sys.argv[1])
    output = Path(sys.argv[2])
    if not source.is_file():
        print(f'Error: Access database file not found: {source}')
        return 1
    if source.suffix.casefold() not in {'.accdb', '.mdb'}:
        print('Error: Access database source must have a .accdb or .mdb extension.')
        return 1

    lock_extension = '.laccdb' if source.suffix.casefold() == '.accdb' else '.ldb'
    lock_file = source.with_suffix(lock_extension)
    if lock_file.exists():
        print('Error: Access database appears to be open; close it before exporting.')
        return 1

    try:
        with tempfile.TemporaryDirectory(prefix='vbaexport-access-') as temporary_directory:
            temporary_copy = Path(temporary_directory) / source.name
            try:
                shutil.copy2(source, temporary_copy)
                if (not temporary_copy.is_file() or
                        temporary_copy.stat().st_size != source.stat().st_size):
                    raise OSError('temporary copy does not match the source file size')
            except OSError as error:
                print(f'Error: Unable to create or verify temporary Access database copy: {error}')
                return 1

            for directory in (output / 'Modules', output / 'Queries' / 'Access',
                              output / 'Queries' / 'SQL', output / 'Macros'):
                directory.mkdir(parents=True, exist_ok=True)

            environment = os.environ.copy()
            environment['VBAEXPORT_ACCESS_WORKING_COPY'] = str(temporary_copy)
            environment['VBAEXPORT_ACCESS_OUTPUT'] = str(output.resolve())
            result = subprocess.run(
                ['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', POWERSHELL_SCRIPT],
                env=environment,
                text=True,
                capture_output=True,
            )
    except OSError as error:
        print(f'Error: Unable to create temporary Access database working directory: {error}')
        return 1
    if result.stdout:
        print(result.stdout, end='')
    if result.stderr:
        print(result.stderr, end='', file=sys.stderr)
    return result.returncode


if __name__ == '__main__':
    raise SystemExit(main())
