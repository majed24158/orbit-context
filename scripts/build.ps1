param([switch]$InstallDependencies, [switch]$Clean)
$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
Push-Location -LiteralPath $projectRoot
try {
    if ($InstallDependencies) {
        python -m pip install -r requirements.txt
        if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed.' }
    }
    $buildArguments = @('-m', 'PyInstaller', '--noconfirm', '--distpath', 'dist', '--workpath', 'build')
    if ($Clean) { $buildArguments += '--clean' }
    $buildArguments += 'OrbitContext.spec'
    python @buildArguments
    if ($LASTEXITCODE -ne 0) { throw 'Executable build failed.' }
    Write-Host 'Built dist/OrbitContext.exe'
} finally { Pop-Location }
