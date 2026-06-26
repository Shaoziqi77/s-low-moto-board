# 一键把 s-low-moto-board 项目推送到 GitHub（main 分支）
# 用法:
#   1) 先把仓库关联到远程（只需一次），在 Git Bash 或 PowerShell 里执行:
#      git init
#      git remote add origin https://github.com/Shaoziqi77/s-low-moto-board.git
#      git branch -M main
#   2) 以后每次更新代码后，双击本脚本，或在 PowerShell 里执行:
#      .\push.ps1 "本次改动说明"
#
# 注意:
#   - 如果还没有安装 Git，请先到 https://git-scm.com/ 下载安装并重启终端
#   - 如果 commit 为空（没有任何文件改动），脚本会提示并直接退出
#   - 首次推送可能会让你在浏览器里登录 GitHub（HTTPS 方式）

param(
    [string]$Message = ""
)

$ErrorActionPreference = "Stop"

function Invoke-Git {
    param([Parameter(ValueFromRemainingArguments)][string[]]$Args)
    & git @Args
    if ($LASTEXITCODE -ne 0) {
        throw "git 命令执行失败，退出码: $LASTEXITCODE"
    }
}

# 1. 检查 Git 可用性
try {
    Invoke-Git --version | Out-Null
}
catch {
    Write-Error "未检测到 git 命令。请先安装 Git (https://git-scm.com/) 并重启终端后重试。"
    exit 1
}

# 2. 确保已关联 origin 远程
$remote = (& git remote get-url origin) 2>$null
if (-not $remote) {
    Write-Host "尚未关联远程仓库，正在关联到 https://github.com/Shaoziqi77/s-low-moto-board.git"
    Invoke-Git remote add origin https://github.com/Shaoziqi77/s-low-moto-board.git
    Invoke-Git branch -M main
}
else {
    Write-Host "当前远程仓库 origin = $remote"
}

# 3. 生成提交信息
if ([string]::IsNullOrWhiteSpace($Message)) {
    $Message = "chore: 自动提交更新（$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")）"
}

# 4. 暂存、提交（若无改动则跳过 push）
Invoke-Git add -A

$statusOutput = (& git status --porcelain) 2>$null
if ([string]::IsNullOrWhiteSpace($statusOutput)) {
    Write-Host "没有任何文件改动，无需提交。"
    exit 0
}

Invoke-Git commit -m $Message

# 5. 推送到 main（首次 -u 绑定 upstream，后续可直接 git push）
$currentBranch = (& git rev-parse --abbrev-ref HEAD) 2>$null
if ([string]::IsNullOrWhiteSpace($currentBranch)) { $currentBranch = "main" }

Write-Host "正在推送到 origin/$currentBranch ..."
try {
    & git push -u origin $currentBranch
    if ($LASTEXITCODE -ne 0) {
        # 首次遇到历史不一致时，尝试先拉取
        Write-Host "推送失败，尝试先 pull 合并远程变更后再推送..."
        & git pull --ff-only origin $currentBranch
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Pull 失败（可能存在冲突），请在可视化工具里手动处理冲突后再次 push。"
            exit 1
        }
        & git push -u origin $currentBranch
    }
}
catch {
    Write-Error "推送过程出错: $_"
    exit 1
}

Write-Host "推送完成，查看: https://github.com/Shaoziqi77/s-low-moto-board"
