param([switch]$WithOCR)
$ErrorActionPreference = 'Stop'
Set-Location (Split-Path $PSScriptRoot -Parent)
if ($env:OS -ne 'Windows_NT') { throw 'Build this executable on Windows x64.' }
function Invoke-Checked {
    param([string]$Program, [string[]]$Arguments)
    & $Program @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$Program failed with exit code $LASTEXITCODE" }
}
$venvName = if ($WithOCR) { '.venv-build-ocr' } else { '.venv-build-core' }
if (-not (Test-Path "$venvName/Scripts/python.exe")) {
    Invoke-Checked 'py' @('-3.11', '-m', 'venv', $venvName)
}
$python = "$venvName/Scripts/python.exe"
Invoke-Checked $python @('-m', 'pip', 'install', '-r', 'requirements-build.txt')
Invoke-Checked $python @('-m', 'pip', 'install', 'python-docx', 'openpyxl', 'python-pptx')
if ($WithOCR) { Invoke-Checked $python @('-m', 'pip', 'install', '-r', 'requirements-ocr.txt') }
Invoke-Checked $python @('-m', 'unittest', 'discover', '-s', 'tests', '-v')
Invoke-Checked $python @('scripts/smoke_gui.py')
$name = if ($WithOCR) { 'DocReplace-OCR' } else { 'DocReplace' }
$buildArgs = @('-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir', '--windowed',
    '--name', $name, '--distpath', 'dist', '--workpath', "build/$name", '--specpath', 'build',
    '--exclude-module', 'docx', '--exclude-module', 'openpyxl', '--exclude-module', 'pptx',
    '--hidden-import', 'win32com.client', '--hidden-import', 'pythoncom', '--hidden-import', 'pywintypes')
if ($WithOCR) {
    $buildArgs += @('--collect-all', 'rapidocr_onnxruntime', '--collect-all', 'onnxruntime',
        '--collect-all', 'cv2', '--collect-all', 'numpy', '--hidden-import', 'PIL.ImageTk')
} else {
    $buildArgs += @('--exclude-module', 'rapidocr_onnxruntime', '--exclude-module', 'onnxruntime',
        '--exclude-module', 'numpy', '--exclude-module', 'PIL', '--exclude-module', 'cv2')
}
$buildArgs += 'main.py'
Invoke-Checked $python $buildArgs
$exe = (Resolve-Path "dist/$name/$name.exe").Path
$report = Join-Path (Split-Path $exe) 'exe-selftest.json'
$selfTestArgs = @('--self-test', "`"$report`"")
if ($WithOCR) { $selfTestArgs += '--with-ocr' }
$process = Start-Process -FilePath $exe -ArgumentList $selfTestArgs -PassThru
if (-not $process.WaitForExit(180000)) {
    $process.Kill()
    throw 'Packaged EXE self-test timed out.'
}
if ($process.ExitCode -ne 0 -or -not (Test-Path $report)) { throw 'Packaged EXE self-test failed.' }
$result = Get-Content -Raw -Encoding UTF8 $report | ConvertFrom-Json
if (-not $result.passed) { throw 'Packaged EXE self-test reported a failure.' }
Copy-Item 'README.md' "dist/$name/使用说明.md"
Copy-Item 'docs/compatibility.md' "dist/$name/兼容范围.md"
Copy-Item 'example-rules.json' "dist/$name/示例规则.json"
& $python -m pip freeze --all | Set-Content -Encoding UTF8 "dist/$name/构建依赖版本.txt"
New-Item -ItemType Directory -Force 'release' | Out-Null
Compress-Archive -Path "dist/$name" -DestinationPath "release/$name-win-x64.zip" -Force
Write-Host "Portable ZIP: release/$name-win-x64.zip"
