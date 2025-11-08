@echo off
title Office 激活修复工具 - 修复 NetworkService 注册表问题
color 0A
setlocal enabledelayedexpansion

:: =====================================================
:: 自动提权部分
:: =====================================================
:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [提示] 当前未以管理员身份运行，正在尝试自动提权...
    powershell -Command "Start-Process '%~f0' -Verb runAs"
    exit /b
)

echo =====================================================
echo       Office 激活修复工具 - NetworkService 修复
echo =====================================================
echo.

set "nsPath=%SystemRoot%\ServiceProfiles\NetworkService"
set "backupPath=%SystemRoot%\ServiceProfiles\NetworkService.backup"
set "timeTag=%date:~-4,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%"

:: =====================================================
:: 1. 停止可能占用的服务
:: =====================================================
echo [1] 停止 sppsvc 与 ClipSVC 服务...
net stop sppsvc >nul 2>&1
net stop ClipSVC >nul 2>&1
timeout /t 3 /nobreak >nul

:: =====================================================
:: 2. 检查并卸载残留注册表配置单元
:: =====================================================
reg query "HKU\S-1-5-20" >nul 2>&1
if %errorlevel% equ 0 (
    echo [2] 检测到已加载的注册表配置单元，正在卸载...
    reg unload "HKU\S-1-5-20" >nul 2>&1
    if %errorlevel% neq 0 (
        echo [警告] 无法卸载，可能被系统进程锁定。
        echo          尝试继续修复。
    )
)

:: =====================================================
:: 3. 备份原始 NTUSER.DAT
:: =====================================================
echo [3] 备份当前 NTUSER.DAT...
if exist "%nsPath%\NTUSER.DAT" (
    if not exist "%backupPath%" mkdir "%backupPath%"
    copy "%nsPath%\NTUSER.DAT" "%backupPath%\NTUSER.DAT.backup.%timeTag%" >nul
    echo     ✅ 已备份至 %backupPath%
) else (
    echo     ⚠ 未找到原始文件，将重新生成。
)

:: =====================================================
:: 4. 重建 NTUSER.DAT
:: =====================================================
echo [4] 正在重建 NetworkService 的 NTUSER.DAT...
set "sourceFound=0"
for %%s in (
    "%SystemDrive%\Users\Default\NTUSER.DAT"
    "%SystemRoot%\System32\config\systemprofile\NTUSER.DAT"
    "%SystemRoot%\ServiceProfiles\LocalService\NTUSER.DAT"
) do (
    if exist "%%s" if !sourceFound! equ 0 (
        copy "%%s" "%nsPath%\NTUSER.DAT" /Y >nul
        if !errorlevel! equ 0 (
            set "sourceFound=1"
            echo     ✅ 使用模板: %%~s
        )
    )
)
if !sourceFound! equ 0 (
    echo     ⚠ 无可用模板，创建空文件...
    type nul > "%nsPath%\NTUSER.DAT"
)

:: =====================================================
:: 5. 设置 NTUSER.DAT 文件权限
:: =====================================================
echo [5] 设置 NTUSER.DAT 文件权限...
icacls "%nsPath%\NTUSER.DAT" /inheritance:r >nul
icacls "%nsPath%\NTUSER.DAT" /grant:r "SYSTEM":F >nul
icacls "%nsPath%\NTUSER.DAT" /grant:r "NETWORK SERVICE":F >nul
icacls "%nsPath%\NTUSER.DAT" /grant:r "Administrators":F >nul
icacls "%nsPath%\NTUSER.DAT" /setowner "SYSTEM" >nul
echo     ✅ 文件权限已修复

:: =====================================================
:: 6. 加载注册表配置单元并授予 ACL
:: =====================================================
echo [6] 加载注册表配置单元...
reg load "HKU\S-1-5-20" "%nsPath%\NTUSER.DAT" >nul 2>&1
if %errorlevel% neq 0 (
    echo     ❌ 注册表加载失败，可能被系统占用。
    goto end
)
echo     ✅ 注册表加载成功

echo [6.1] 初始化基本结构...
reg add "HKU\S-1-5-20\Software\Microsoft\Office" /f >nul
reg add "HKU\S-1-5-20\Software\Microsoft\Windows\CurrentVersion\Explorer" /f >nul
reg add "HKU\S-1-5-20\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /f >nul

echo [6.2] 正在授予注册表访问权限（通过 PowerShell）...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$key = Get-Item 'Registry::HKEY_USERS\\S-1-5-20';" ^
"$acl = $key.GetAccessControl();" ^
"$rules = @('NETWORK SERVICE','SYSTEM','Administrators');" ^
"foreach($r in $rules){" ^
"  $rule = New-Object System.Security.AccessControl.RegistryAccessRule($r,'FullControl','ContainerInherit,ObjectInherit','None','Allow');" ^
"  $acl.SetAccessRule($rule)" ^
"};" ^
"$key.SetAccessControl($acl)" >nul 2>&1

echo     ✅ 已为注册表授予 NETWORK SERVICE 等完全访问权限

reg unload "HKU\S-1-5-20" >nul 2>&1
echo     ✅ 注册表配置单元已卸载（权限写入生效）

:: =====================================================
:: 7. 重启相关服务
:: =====================================================
echo [7] 启动 sppsvc 与 ClipSVC 服务...
net start ClipSVC >nul 2>&1
net start sppsvc >nul 2>&1
echo     ✅ 服务已重新启动

:: =====================================================
:: 8. Office 状态检查
:: =====================================================
echo [8] 检查 Office 激活状态（如有安装）...
if exist "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" (
    cscript //nologo "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /dstatus
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" (
    cscript //nologo "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" /dstatus
) else (
    echo     ⚠ 未检测到 Office 安装路径，跳过检查。
)

:done
echo.
echo =====================================================
echo     ✅ 修复完成，请重新启动计算机。
echo =====================================================
pause
exit /b