# 업로드 스크립트

Set-Location -Path $PSScriptRoot
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host ""
Write-Host "  ==========================================" -ForegroundColor Cyan
Write-Host "    테노바 재고관리  GitHub 업로드" -ForegroundColor Cyan
Write-Host "  ==========================================" -ForegroundColor Cyan
Write-Host ""

# git 설치 확인
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "  [오류] Git이 설치되어 있지 않습니다." -ForegroundColor Red
    Write-Host "         https://git-scm.com 에서 설치 후 다시 실행하세요."
    Read-Host "  아무 키나 누르세요"
    exit
}

# 변경 파일 확인
$changed = git status --short 2>$null
if (-not $changed) {
    Write-Host "  변경된 파일이 없습니다. 업로드할 내용이 없어요!" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "  아무 키나 누르세요"
    exit
}

Write-Host "  [변경된 파일 목록]" -ForegroundColor White
git status --short
Write-Host ""

$msg = Read-Host "  수정 내용을 입력하세요 (예: 부품관리 버그 수정)"
if ([string]::IsNullOrWhiteSpace($msg)) {
    Write-Host ""
    Write-Host "  입력이 없어서 업로드를 취소했습니다." -ForegroundColor Yellow
    Read-Host "  아무 키나 누르세요"
    exit
}

Write-Host ""
Write-Host "  [1/3] 변경 파일 추가 중..." -ForegroundColor Gray
git add -A
if ($LASTEXITCODE -ne 0) {
    Write-Host "  [오류] git add 실패" -ForegroundColor Red
    Read-Host "  아무 키나 누르세요"
    exit
}

Write-Host "  [2/3] 커밋 중..." -ForegroundColor Gray
git commit -m $msg
if ($LASTEXITCODE -ne 0) {
    Write-Host "  [오류] git commit 실패" -ForegroundColor Red
    Read-Host "  아무 키나 누르세요"
    exit
}

Write-Host "  [3/3] GitHub에 업로드 중..." -ForegroundColor Gray
git push origin main
if ($LASTEXITCODE -ne 0) {
    Write-Host "  [오류] git push 실패" -ForegroundColor Red
    Read-Host "  아무 키나 누르세요"
    exit
}

Write-Host ""
Write-Host "  업로드 완료!" -ForegroundColor Green
Write-Host "  https://github.com/SEUNGJUNHWANG/inventory-management" -ForegroundColor DarkGray
Write-Host ""
Read-Host "  아무 키나 누르세요"
