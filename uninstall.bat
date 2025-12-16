@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo.
echo ========================================
echo        TextCaptureTool 卸载程序
echo ========================================
echo.

REM 检查是否以管理员权限运行
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 警告：建议以管理员权限运行此脚本以获得最佳卸载效果
    echo 右键点击脚本，选择"以管理员身份运行"
    echo.
)

echo 正在卸载 TextCaptureTool...
echo.

REM 删除桌面快捷方式
set "desktop=%USERPROFILE%\Desktop"
set "shortcut_name=TextCaptureTool.lnk"
set "desktop_shortcut=%desktop%\%shortcut_name%"

if exist "%desktop_shortcut%" (
    echo 删除桌面快捷方式: %desktop_shortcut%
    del /f /q "%desktop_shortcut%"
    if !errorlevel! equ 0 (
        echo ✓ 桌面快捷方式已删除
    ) else (
        echo ✗ 删除桌面快捷方式失败，可能需要管理员权限
    )
) else (
    echo 桌面快捷方式不存在，跳过删除
)

REM 删除开始菜单快捷方式
set "start_menu=%APPDATA%\Microsoft\Windows\Start Menu\Programs"
set "start_menu_shortcut=%start_menu%\%shortcut_name%"

if exist "%start_menu_shortcut%" (
    echo 删除开始菜单快捷方式: %start_menu_shortcut%
    del /f /q "%start_menu_shortcut%"
    if !errorlevel! equ 0 (
        echo ✓ 开始菜单快捷方式已删除
    ) else (
        echo ✗ 删除开始菜单快捷方式失败，可能需要管理员权限
    )
) else (
    echo 开始菜单快捷方式不存在，跳过删除
)

REM 询问是否删除配置文件
set /p "delete_config=是否删除配置文件？(y/N): "
if /i "!delete_config!"=="y" (
    set "config_dir=%APPDATA%\TextCaptureTool"
    if exist "!config_dir!" (
        echo 删除配置文件目录: !config_dir!
        rmdir /s /q "!config_dir!"
        if !errorlevel! equ 0 (
            echo ✓ 配置文件已删除
        ) else (
            echo ✗ 删除配置文件失败，可能需要手动删除
        )
    ) else (
        echo 配置文件目录不存在，跳过删除
    )
) else (
    echo 保留配置文件
)

echo.
echo ========================================
echo           卸载完成！
echo ========================================
echo.
echo 卸载完成！以下内容已被删除：
echo ✓ 桌面快捷方式
echo ✓ 开始菜单快捷方式
echo.
echo 注意：
echo - Python环境和依赖包不会被删除
echo - 程序文件仍保留在原位置
echo - 如需完全卸载，请手动删除程序目录
echo.
echo 感谢使用 TextCaptureTool！
echo.

pause