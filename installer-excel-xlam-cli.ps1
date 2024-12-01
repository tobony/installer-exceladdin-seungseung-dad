# UTF-8 with BOM
chcp 65001 | Out-Null
$OutputEncoding = [System.Console]::InputEncoding = [System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# XLAM Add-in 설치 및 등록 스크립트

# 관리자 권한으로 실행
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator'))
{  
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments -WindowStyle Normal
    Break
}

# 경로 설정
$SourcePath = Join-Path $PSScriptRoot 'src'
$InstallPath = "$env:APPDATA\Microsoft\AddIns"

# src 폴더 존재 확인
if (-not (Test-Path $SourcePath))
{
    Write-Host '[오류] src 폴더를 찾을 수 없습니다.' -ForegroundColor Red
    Write-Host '계속하려면 아무 키나 누르세요...'
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

# src 폴더의 파일 목록 확인
$sourceFiles = Get-ChildItem -Path $SourcePath -File
if ($sourceFiles.Count -eq 0)
{
    Write-Host '[오류] src 폴더가 비어있습니다.' -ForegroundColor Red
    Write-Host '계속하려면 아무 키나 누르세요...'
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

# Excel 프로세스 확인 및 사용자 확인 후 종료
$excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProcesses)
{
    Write-Host '실행 중인 Excel을 발견했습니다. 모두 종료할까요? (Y/N)' -ForegroundColor Yellow -NoNewline
    $confirmation = Read-Host
    if ($confirmation -eq 'Y' -or $confirmation -eq 'y')
    {
        Write-Host 'Excel 종료 중...' -ForegroundColor Yellow
        $excelProcesses | Stop-Process -Force
        Start-Sleep -Seconds 2
    }
    else
    {
        Write-Host '[취소] 설치가 취소되었습니다. Excel을 먼저 종료해주세요.' -ForegroundColor Red
        Write-Host '계속하려면 아무 키나 누르세요...'
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        exit
    }
}

# 설치 디렉토리 확인 및 생성
if (-not (Test-Path $InstallPath))
{
    New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
    Write-Host "AddIn 폴더 생성됨: $InstallPath" -ForegroundColor Green
}

# 파일 복사 및 등록
$registryPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
$maxOpenKey = 0

# 기존 OPEN 키 확인
if (Test-Path $registryPath)
{
    $openKeys = Get-ItemProperty -Path $registryPath -Name 'OPEN*' -ErrorAction SilentlyContinue
    if ($openKeys)
    {
        $openKeys.PSObject.Properties | Where-Object { $_.Name -like 'OPEN*' } | ForEach-Object {
            $keyNum = [int]($_.Name -replace 'OPEN', '')
            if ($keyNum -gt $maxOpenKey)
            {
                $maxOpenKey = $keyNum
            }
        }
    }
}

# 각 파일 처리
foreach ($file in $sourceFiles)
{
    $destPath = Join-Path $InstallPath $file.Name
    try
    {
        # 파일 복사
        Copy-Item -Path $file.FullName -Destination $destPath -Force
        Write-Host "→ 파일 복사 완료: $($file.Name)" -ForegroundColor Green

        # XLAM 파일인 경우 레지스트리에 등록
        if ($file.Extension -eq '.xlam')
        {
            $maxOpenKey++
            $newKeyName = "OPEN$maxOpenKey"
            
            if (-not (Test-Path $registryPath))
            {
                New-Item -Path $registryPath -Force | Out-Null
            }
            
            New-ItemProperty -Path $registryPath -Name $newKeyName -Value $destPath -PropertyType String -Force | Out-Null
            Write-Host "→ Add-in 등록 완료: $($file.Name)" -ForegroundColor Green
        }
    }
    catch
    {
        Write-Host "[오류] 파일 처리 실패 ($($file.Name)): $_" -ForegroundColor Red
    }
}

Write-Host "`n[완료] 설치가 완료되었습니다." -ForegroundColor Green
Write-Host "Excel을 시작하면 Add-in을 사용할 수 있습니다." -ForegroundColor Green
Write-Host "`n계속하려면 아무 키나 누르세요..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')