# 승승아빠-Excel-AddIn-Installer
- src 폴더의 파일을 `APPDATA\Microsoft\AddIns`에 복사합니다.
- `xlam`파일은 레지스트리에 등록합니다. (별도 addin 추가과정을 대신합니다.)
- 엑셀을 실행 중이라면 재실행이 필요하여 이를 확인합니다.

<br><br>


# How to Use
- 설치할 파일을 `src`하위폴더에 넣습니다.
- exe 파일을 실행해서 절차에 따라 진행합니다.

<br>

<img width="700" alt="image" src="https://github.com/user-attachments/assets/a31d3300-3cc8-41ca-9e2f-46d531719214">

<br><br>

<img width="700" alt="image" src="https://github.com/user-attachments/assets/7cf00fdf-e599-450b-8276-beae23287e93">


<br><br>

## make .exe file
# 1. PS2EXE 모듈 설치
Install-Module -Name ps2exe -Scope CurrentUser -Force


# 2. 스크립트를 EXE로 변환
Invoke-ps2exe .\installer-excel-xlam-exe.ps1 .\승승아빠-Excel-AddIn-Installer.exe -noConsole -RequireAdmin
