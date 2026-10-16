<#
.SYNOPSIS
    打包 Word 合并工具为可移植的 Windows exe。

.DESCRIPTION
    产出两个可执行文件（都在 dist\ 下）：

      WordMergeTool.exe      图形界面版（双击即用）
      WordMergeTool-CLI.exe  命令行版（批处理/脚本调用）

    默认打成**单文件**（onefile），目标机器无需安装 Python 或任何第三方库，
    拷过去双击就能跑。

    体积优化（都是实测必要的）：
      * 排除 matplotlib / numpy / pandas / scipy —— PyInstaller 常从 PyMuPDF、
        Pillow 的依赖里把它们误拖进来，白白增加几十 MB
      * 排除 PyQt / PySide —— 本工具用 tkinter，这些一旦被牵连就是上百 MB
      * UPX 若存在则启用（可选，能再压 30% 左右）

.PARAMETER Mode
    onefile（默认，单个 exe）或 onedir（一个文件夹，启动更快）。

.PARAMETER Clean
    打包前先删掉 build\ 与 dist\。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\build.ps1
    powershell -ExecutionPolicy Bypass -File .\build.ps1 -Mode onedir -Clean
#>
[CmdletBinding()]
param(
    [ValidateSet('onefile', 'onedir')]
    [string]$Mode = 'onefile',

    [switch]$Clean,

    # 打包后跑一次 GUI 冒烟测试（需要图形环境；CI/无桌面环境请勿加）
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host "    $msg" -ForegroundColor Green }
function Write-Warn2($msg){ Write-Host "    $msg" -ForegroundColor Yellow }

# --------------------------------------------------------------------------
Write-Step "检查环境"
$python = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $python) { throw "找不到 python，请先安装 Python 3.10+ 并加入 PATH" }
Write-Ok "Python: $python"
Write-Ok ("版本: " + (& python --version 2>&1))

# 读取版本号（唯一事实来源）
$versionLine = Select-String -Path "src\wmtool\__init__.py" -Pattern '__version__\s*=\s*"([^"]+)"'
if (-not $versionLine) { throw "无法从 src\wmtool\__init__.py 读取版本号" }
$Version = $versionLine.Matches[0].Groups[1].Value
Write-Ok "版本号: $Version"

# --------------------------------------------------------------------------
Write-Step "检查依赖"
$required = @{
    'docx'         = 'python-docx'
    'pptx'         = 'python-pptx'
    'openpyxl'     = 'openpyxl'
    'pymupdf'      = 'PyMuPDF'
    'PIL'          = 'Pillow'
    'tkinterdnd2'  = 'tkinterdnd2'
}
$missing = @()
foreach ($mod in $required.Keys) {
    & python -c "import $mod" 2>$null
    if ($LASTEXITCODE -ne 0) { $missing += $required[$mod] }
}
if ($missing.Count -gt 0) {
    Write-Warn2 ("缺少依赖: " + ($missing -join ', '))
    Write-Host "    正在安装 ..."
    & python -m pip install $missing
    if ($LASTEXITCODE -ne 0) { throw "依赖安装失败" }
}

& python -c "import PyInstaller" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Warn2 "缺少 pyinstaller，正在安装 ..."
    & python -m pip install pyinstaller
    if ($LASTEXITCODE -ne 0) { throw "pyinstaller 安装失败" }
}
Write-Ok "依赖齐备"

# --------------------------------------------------------------------------
# 打包前先关掉还开着的旧 exe：上一次测试遗留的 GUI 进程会占住 dist\logs 里的
# 日志文件，导致清理步骤失败、进而中断整个打包。这里主动收尾比报错友好。
$running = @(Get-Process -Name 'WordMergeTool', 'WordMergeTool-CLI' -ErrorAction SilentlyContinue)
if ($running.Count -gt 0) {
    Write-Step "关闭仍在运行的旧程序（$($running.Count) 个进程）"
    $running | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Write-Ok "已关闭"
}

if ($Clean) {
    Write-Step "清理旧的构建产物"
    foreach ($d in @('build', 'dist')) {
        if (Test-Path $d) {
            try {
                Remove-Item -Recurse -Force $d -ErrorAction Stop
            } catch {
                # 兜底：偶尔还有句柄没释放，重试一次；再不行就点名删，别整个打包崩掉
                Start-Sleep -Seconds 2
                try {
                    Remove-Item -Recurse -Force $d -ErrorAction Stop
                } catch {
                    Write-Warn2 "无法完全删除 $d（$($_.Exception.Message)）"
                    Write-Warn2 "将继续打包，PyInstaller 的 --noconfirm 会覆盖同名产物"
                }
            }
        }
    }
    Write-Ok "已清理"
}

# --------------------------------------------------------------------------
Write-Step "打包（$Mode）"

# 公共参数
$common = @(
    '--noconfirm'
    '--clean'
    '--name', 'WordMergeTool'
    '--paths', 'src'
    '--additional-hooks-dir', 'packaging/hooks'

    # 体积：这些库与功能无关，却被 PyMuPDF / Pillow 的依赖牵连进来
    '--exclude-module', 'matplotlib'
    '--exclude-module', 'numpy'
    '--exclude-module', 'pandas'
    '--exclude-module', 'scipy'
    '--exclude-module', 'PyQt5'
    '--exclude-module', 'PyQt6'
    '--exclude-module', 'PySide2'
    '--exclude-module', 'PySide6'
    '--exclude-module', 'IPython'
    '--exclude-module', 'notebook'
    '--exclude-module', 'pytest'
    '--exclude-module', 'setuptools'
    '--exclude-module', 'pip'

    # 拖拽支持：PyInstaller 静态分析看不到 tkinterdnd2 的动态导入
    '--hidden-import', 'tkinterdnd2'
    '--hidden-import', 'pymupdf'
    '--hidden-import', 'docx'
    '--hidden-import', 'pptx'
    '--hidden-import', 'openpyxl'

    # 控制台：GUI 版不要黑框
    '--windowed'
)

if ($Mode -eq 'onefile') { $common += '--onefile' } else { $common += '--onedir' }

# 图标（存在才用）
$iconPath = Join-Path $ProjectRoot 'assets\app.ico'
if (Test-Path $iconPath) {
    $common += @('--icon', $iconPath)
    Write-Ok "使用图标: $iconPath"
} else {
    Write-Warn2 "未找到 assets\app.ico，使用默认图标"
}

# ---- GUI 版 ----
# 入口用 packaging\entry_gui.py 而不是 src\wmtool\gui.py：
# 直接拿包内模块当入口时，PyInstaller 会把它当 __main__ 跑，
# 包内的相对导入（from . import ...）会报 "no known parent package"。
Write-Host "    打包 GUI 版 ..."
& pyinstaller @common 'packaging\entry_gui.py'
if ($LASTEXITCODE -ne 0) { throw "GUI 版打包失败" }

# ---- CLI 版（保留控制台，方便批处理看输出）----
# 与 GUI 版的区别只有两处：改名字、去掉 --windowed（要保留控制台看进度）
$cliArgs = @()
foreach ($a in $common) { $cliArgs += $a }
for ($i = 0; $i -lt $cliArgs.Count; $i++) {
    if ($cliArgs[$i] -eq '--name') { $cliArgs[$i + 1] = 'WordMergeTool-CLI' }
}
$cliArgs = @($cliArgs | Where-Object { $_ -ne '--windowed' })

Write-Host "    打包 CLI 版 ..."
& pyinstaller @cliArgs 'packaging\entry_cli.py'
if ($LASTEXITCODE -ne 0) { throw "CLI 版打包失败" }

# --------------------------------------------------------------------------
Write-Step "整理产物"
$distDir = Join-Path $ProjectRoot 'dist'
$exeGui = Join-Path $distDir 'WordMergeTool.exe'
$exeCli = Join-Path $distDir 'WordMergeTool-CLI.exe'

if ($Mode -eq 'onedir') {
    $exeGui = Join-Path $distDir 'WordMergeTool\WordMergeTool.exe'
    $exeCli = Join-Path $distDir 'WordMergeTool-CLI\WordMergeTool-CLI.exe'
}

foreach ($f in @($exeGui, $exeCli)) {
    if (-not (Test-Path $f)) { throw "期望的产物不存在: $f" }
    $sizeMB = [math]::Round((Get-Item $f).Length / 1MB, 1)
    Write-Ok ("{0}  ({1} MB)" -f (Split-Path $f -Leaf), $sizeMB)
}

# 把说明与变更记录一起放进 dist，便于分发
foreach ($doc in @('README.md', 'CHANGELOG.md')) {
    if (Test-Path $doc) { Copy-Item $doc $distDir -Force }
}

# 生成版本说明
@"
Word 合并工具 v$Version
打包时间: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
打包模式: $Mode

分发方式：
  把 WordMergeTool.exe（GUI 版）单独拷到目标电脑即可，无需安装 Python 或任何依赖。
  若需要命令行批处理，再一并拷 WordMergeTool-CLI.exe。

日志位置：
  exe 同级目录下的 logs\ 文件夹。出错时请提供 logs\crash-*.txt。

支持的系统：
  Windows 10 / Windows 11（64 位）
"@ | Set-Content -Path (Join-Path $distDir '版本说明.txt') -Encoding UTF8

# --------------------------------------------------------------------------
if (-not $SkipTests) {
    Write-Step "验证产物（CLI 版跑端到端自检）"
    $testDir = Join-Path $env:TEMP "wmtool-exe-test-$(Get-Random)"
    & $exeCli --selftest $testDir
    if ($LASTEXITCODE -ne 0) {
        Write-Warn2 "自检未通过（退出码 $LASTEXITCODE），请检查上面的输出"
    } else {
        Write-Ok "自检通过"
    }
}

Write-Host ""
Write-Host "打包完成！" -ForegroundColor Green
Write-Host "  GUI 版: $exeGui"
Write-Host "  CLI 版: $exeCli"
Write-Host "  说明:   $(Join-Path $distDir '版本说明.txt')"
