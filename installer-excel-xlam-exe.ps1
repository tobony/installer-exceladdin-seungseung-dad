# GUI 모드 설정
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# PowerShell script encoding setting
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 > $null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# Get the executable path
$exePath = [System.IO.Path]::GetDirectoryName([System.Windows.Forms.Application]::ExecutablePath)
if ([string]::IsNullOrEmpty($exePath)) {
    # Fallback for PS1 script execution
    $exePath = $PSScriptRoot
    if ([string]::IsNullOrEmpty($exePath)) {
        # Final fallback to current directory
        $exePath = (Get-Location).Path
    }
}

# Source paths definition at the start
$SourcePath = Join-Path $exePath "src"
$InstallPath = "$env:APPDATA\Microsoft\AddIns"

# Write to log function
function Write-Log {
    param([string]$Message)
    $logTextBox.AppendText("$Message`r`n")
    $logTextBox.Select($logTextBox.Text.Length, 0)
    $logTextBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Excel Add-in Installer'
$form.Size = New-Object System.Drawing.Size(600,400)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Log textbox
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10,120)
$logTextBox.Size = New-Object System.Drawing.Size(565,200)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = 'Vertical'
$logTextBox.ReadOnly = $true
$form.Controls.Add($logTextBox)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,330)
$progressBar.Size = New-Object System.Drawing.Size(565,20)
$form.Controls.Add($progressBar)

# Install button
$installButton = New-Object System.Windows.Forms.Button
$installButton.Location = New-Object System.Drawing.Point(200,70)
$installButton.Size = New-Object System.Drawing.Size(200,30)
$installButton.Text = 'Install Excel Add-in'
$form.Controls.Add($installButton)

# Description label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(565,40)
$label.Text = 'This tool will install Excel Add-in from the src folder and register it in Excel.'
$form.Controls.Add($label)

# Form load event handler
$form.Add_Shown({
    # 폼이 완전히 로드된 후 Excel 프로세스 체크
    $installButton.Enabled = $false
    Write-Log "Excel 프로세스 확인 중..."
    
    $excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
    if ($excelProcesses) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Excel이 실행 중입니다. 설치를 진행하려면 Excel을 종료해야 합니다.`n`n계속 진행하시겠습니까?",
            "Excel 실행 확인",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        
        if ($result -eq [System.Windows.Forms.DialogResult]::No) {
            Write-Log "사용자가 Excel 종료를 취소했습니다."
            [System.Windows.Forms.MessageBox]::Show(
                "Excel을 종료한 후 설치 프로그램을 다시 실행해주세요.",
                "설치 취소",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
            $form.Close()
            return
        }

        Write-Log "Excel 프로세스 종료 중..."
        foreach ($proc in $excelProcesses) {
            $proc.CloseMainWindow() | Out-Null
            Start-Sleep -Milliseconds 500
            if (-not $proc.HasExited) {
                $proc.Kill()
            }
        }

        Start-Sleep -Seconds 2
        $remainingExcel = Get-Process excel -ErrorAction SilentlyContinue
        if ($remainingExcel) {
            Write-Log "Excel을 완전히 종료하지 못했습니다."
            [System.Windows.Forms.MessageBox]::Show(
                "Excel을 완전히 종료하지 못했습니다.`n모든 Excel 작업을 저장하고 직접 종료한 후 다시 시도해주세요.",
                "Excel 종료 실패",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            $form.Close()
            return
        }
        Write-Log "Excel 프로세스가 종료되었습니다."
    } else {
        Write-Log "Excel이 실행중이지 않습니다."
    }
    
    $installButton.Enabled = $true
})

# Installation function
function Install-AddIn {
    $installButton.Enabled = $false
    $progressBar.Value = 0
    
    try {
        Write-Log "설치를 시작합니다..."
        $progressBar.Value = 10
        
        # Check source folder
        if (-not (Test-Path $SourcePath)) {
            Write-Log "[오류] src 폴더를 찾을 수 없습니다: $SourcePath"
            [System.Windows.Forms.MessageBox]::Show(
                "src 폴더를 찾을 수 없습니다.`n설치 파일 구조를 확인해주세요.`n경로: $SourcePath",
                "오류",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        $progressBar.Value = 30
        # Check source files
        $sourceFiles = Get-ChildItem -Path $SourcePath -File
        if ($sourceFiles.Count -eq 0) {
            Write-Log "[오류] src 폴더가 비어있습니다."
            [System.Windows.Forms.MessageBox]::Show(
                "src 폴더에 설치할 파일이 없습니다.",
                "오류",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        $progressBar.Value = 40
        # Create install directory if not exists
        if (-not (Test-Path $InstallPath)) {
            New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
            Write-Log "AddIn 폴더 생성됨: $InstallPath"
        }
        
        $progressBar.Value = 50
        # Registry path for Excel add-ins
        $registryPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
        $maxOpenKey = 0
        
        # Check existing OPEN keys
        if (Test-Path $registryPath) {
            $openKeys = Get-ItemProperty -Path $registryPath -Name 'OPEN*' -ErrorAction SilentlyContinue
            if ($openKeys) {
                $openKeys.PSObject.Properties | Where-Object { $_.Name -like 'OPEN*' } | ForEach-Object {
                    $keyNum = [int]($_.Name -replace 'OPEN', '')
                    if ($keyNum -gt $maxOpenKey) {
                        $maxOpenKey = $keyNum
                    }
                }
            }
        }
        
        $progressBar.Value = 60
        # Process each file
        $fileCount = $sourceFiles.Count
        $current = 0
        
        foreach ($file in $sourceFiles) {
            $current++
            $progressValue = 60 + (30 * ($current / $fileCount))
            $progressBar.Value = [math]::Min(90, $progressValue)
            
            $destPath = Join-Path $InstallPath $file.Name
            try {
                # Copy file
                Copy-Item -Path $file.FullName -Destination $destPath -Force
                Write-Log "→ 파일 복사 완료: $($file.Name)"
                
                # Register if it's an XLAM file
                if ($file.Extension -eq '.xlam') {
                    $maxOpenKey++
                    $newKeyName = "OPEN$maxOpenKey"
                    
                    if (-not (Test-Path $registryPath)) {
                        New-Item -Path $registryPath -Force | Out-Null
                    }
                    
                    New-ItemProperty -Path $registryPath -Name $newKeyName -Value $destPath -PropertyType String -Force | Out-Null
                    Write-Log "→ Add-in 등록 완료: $($file.Name)"
                }
            }
            catch {
                Write-Log "[오류] 파일 처리 실패 ($($file.Name)): $_"
                [System.Windows.Forms.MessageBox]::Show(
                    "파일 처리 중 오류가 발생했습니다: $($file.Name)`n$_",
                    "오류",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        
        $progressBar.Value = 100
        Write-Log "`n[완료] 설치가 완료되었습니다."
        Write-Log "Excel을 시작하면 Add-in을 사용할 수 있습니다."
        
        [System.Windows.Forms.MessageBox]::Show(
            "설치가 완료되었습니다.`nExcel을 시작하면 Add-in을 사용할 수 있습니다.",
            "설치 완료",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        Write-Log "[오류] 설치 중 오류가 발생했습니다: $_"
        [System.Windows.Forms.MessageBox]::Show(
            "설치 중 오류가 발생했습니다:`n$_",
            "오류",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $installButton.Enabled = $true
        $progressBar.Value = 0
    }
}

# Install button click event
$installButton.Add_Click({ Install-AddIn })

# Check and request admin privileges
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "이 프로그램은 관리자 권한이 필요합니다. 관리자 권한으로 다시 시작하시겠습니까?",
        "관리자 권한 필요",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processStartInfo.FileName = "powershell.exe"
        $processStartInfo.Arguments = "-File `"$($myinvocation.mycommand.definition)`""
        $processStartInfo.Verb = "runas"
        [System.Diagnostics.Process]::Start($processStartInfo)
    }
    exit
}

[System.Windows.Forms.Application]::Run($form)