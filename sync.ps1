# 把 Claude 改好的檔案推上 Apps Script 和 GitHub
# 用法：在 PowerShell 輸入  .\sync.ps1
Set-Location $PSScriptRoot
Write-Host '1/3 推上 Apps Script…' -ForegroundColor Cyan
clasp push --force
if ($LASTEXITCODE -ne 0) { Write-Host 'clasp push 失敗：如果看到 invalid_grant / invalid_rapt，先跑 clasp login 再重跑一次' -ForegroundColor Red; exit 1 }
Write-Host '2/3 存檔紀錄（git commit）…' -ForegroundColor Cyan
git add -A
git diff --cached --quiet
if ($LASTEXITCODE -ne 0) {
  $msg = if ($args.Count -gt 0) { $args[0] } else { 'Update from Claude session' }
  git commit -m $msg
}
Write-Host '3/3 推上 GitHub…' -ForegroundColor Cyan
git push
if ($LASTEXITCODE -ne 0) { Write-Host 'git push 失敗，把上面的訊息貼給 Claude' -ForegroundColor Red; exit 1 }
Write-Host '完成' -ForegroundColor Green
