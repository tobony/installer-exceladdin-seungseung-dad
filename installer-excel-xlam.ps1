# XLAM Add-in 설치 및 등록 스크립트

# 관리자 권한으로 실행
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {  
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments -WindowStyle Normal
    Break
}

# 설치 경로 설정
$XlamPath = Join-Path $PSScriptRoot "Addin.xlam"
$InstallPath = "$env:APPDATA\Microsoft\AddIns"

# XLAM 파일 존재 확인
if (-not (Test-Path $XlamPath)) {
    Write-Error "현재 폴더에 Addin.xlam 파일이 없습니다. 파일을 확인해주세요."
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# Excel 프로세스 확인 및 사용자 확인 후 종료
$excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProcesses) {
    $confirmation = Read-Host "실행 중인 Excel이 있습니다. 모든 Excel을 종료하시겠습니까? (Y/N)"
    if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
        Write-Host "Excel을 종료합니다..."
        $excelProcesses | Stop-Process -Force
        Start-Sleep -Seconds 2
    } else {
        Write-Host "설치가 취소되었습니다. 열려있는 Excel을 먼저 종료해주세요."
        Write-Host "Press any key to continue..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }
}

# 설치 디렉토리 확인 및 생성
if (-not (Test-Path $InstallPath)) {
    New-Item -ItemType Directory -Path $InstallPath -Force
}

# XLAM 파일 복사
$fileName = Split-Path $XlamPath -Leaf
$destPath = Join-Path $InstallPath $fileName
try {
    Copy-Item -Path $XlamPath -Destination $destPath -Force
    Write-Host "XLAM 파일이 성공적으로 복사되었습니다: $destPath"
} catch {
    Write-Error "XLAM 파일 복사 중 오류 발생: $_"
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# 레지스트리에 Add-in 등록
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
try {
    # 레지스트리 키 존재 확인
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force
    }

    # OPEN 키 확인 및 업데이트
    $openKeys = Get-ItemProperty -Path $registryPath -Name "OPEN*" -ErrorAction SilentlyContinue
    $maxOpenKey = 0
    if ($openKeys) {
        $openKeys.PSObject.Properties | Where-Object { $_.Name -like "OPEN*" } | ForEach-Object {
            $keyNum = [int]($_.Name -replace "OPEN", "")
            if ($keyNum -gt $maxOpenKey) {
                $maxOpenKey = $keyNum
            }
        }
    }
    
    # 새로운 OPEN 키 생성
    $newKeyName = "OPEN" + ($maxOpenKey + 1)
    New-ItemProperty -Path $registryPath -Name $newKeyName -Value $destPath -PropertyType String -Force
    Write-Host "Add-in이 레지스트리에 성공적으로 등록되었습니다."

} catch {
    Write-Error "레지스트리 등록 중 오류 발생: $_"
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

Write-Host "설치가 완료되었습니다. Excel을 실행하면 Add-in을 사용할 수 있습니다."
Write-Host "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
