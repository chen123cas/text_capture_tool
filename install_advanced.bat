@echo off
chcp 65001 >nul
title TextCaptureTool 高级安装程序

setlocal enabledelayedexpansion

REM 设置颜色
for /f "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
  set "DEL=%%a"
)

call :color 0a
echo ========================================
echo     TextCaptureTool 高级安装程序
echo ========================================
call :color 07
echo.

REM 检查是否以管理员身份运行
net session >nul 2>&1
if %errorlevel% == 0 (
    call :color 0b
    echo [权限检查] 管理员权限 ✓
    call :color 07
) else (
    call :color 0e
    echo [权限检查] 标准用户权限 ⚠️
    call :color 07
    echo    某些功能可能需要管理员权限
)

REM 检查Python环境
call :check_python
if !python_ok! neq 1 (
    call :install_python
    if !python_ok! neq 1 goto :error_exit
)

REM 安装依赖
call :install_dependencies
if !deps_ok! neq 1 goto :error_exit

REM 创建快捷方式
call :create_shortcuts

REM 验证安装
call :verify_installation

REM 完成安装
call :complete_installation

goto :eof

:check_python
    echo.
    call :color 0a
    echo [1/6] 检查Python环境...
    call :color 07
    
    python --version >nul 2>&1
    if %errorlevel% equ 0 (
        for /f "tokens=2 delims= " %%i in ('python --version 2^>^&1') do set python_version=%%i
        call :color 0b
        echo    检测到Python版本: !python_version! ✓
        call :color 07
        
        REM 检查Python版本是否满足要求
        for /f "tokens=1,2 delims=." %%a in ("!python_version!") do (
            if %%a gtr 3 (
                set python_ok=1
            ) else if %%a equ 3 (
                if %%b geq 8 set python_ok=1
            )
        )
        
        if !python_ok! equ 1 (
            call :color 0b
            echo    Python版本符合要求 (>=3.8) ✓
            call :color 07
        ) else (
            call :color 0c
            echo    Python版本过低 (需要>=3.8) ❌
            call :color 07
            set python_ok=0
        )
    ) else (
        call :color 0c
        echo    未检测到Python环境 ❌
        call :color 07
        set python_ok=0
    )
    goto :eof

:install_python
    echo.
    call :color 0e
    echo [安装Python] 需要安装Python 3.8+版本
    call :color 07
    
    echo    请访问以下网址下载安装：
    call :color 0b
    echo     https://www.python.org/downloads/
    call :color 07
    echo.
    echo    安装时请务必勾选：
    echo      - [✓] Add Python to PATH
    echo      - [✓] Install pip
    echo.
    echo    安装完成后，请重新运行此安装程序
    
    REM 提供直接下载链接
    set "python_url=https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
    echo    直接下载链接：!python_url!
    echo.
    
    set /p "download=是否立即下载Python安装包？(y/n): "
    if /i "!download!"=="y" (
        start "" "!python_url!"
    )
    
    pause
    exit /b 1

:install_dependencies
    echo.
    call :color 0a
    echo [2/6] 安装依赖包...
    call :color 07
    
    echo    依赖包列表：
    call :color 0e
    type requirements.txt
    call :color 07
    echo.
    
    REM 尝试使用国内镜像源
    call :color 0b
    echo    尝试使用清华镜像源加速安装...
    call :color 07
    
    pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn --timeout 60
    
    if %errorlevel% neq 0 (
        call :color 0e
        echo    ⚠️ 镜像源安装失败，尝试官方源...
        call :color 07
        
        pip install -r requirements.txt --timeout 120
        
        if %errorlevel% neq 0 (
            call :color 0c
            echo    ❌ 依赖包安装失败
            call :color 07
            set deps_ok=0
            goto :eof
        )
    )
    
    call :color 0b
    echo    ✓ 依赖包安装成功
    call :color 07
    set deps_ok=1
    goto :eof

:create_shortcuts
    echo.
    call :color 0a
    echo [3/6] 跳过快捷方式创建...
    call :color 07
    
    echo    已取消桌面和开始菜单快捷方式创建
    echo    请直接运行 start_tool.bat 或 python main.py 启动程序
    
    goto :eof
    
    set "desktop=!userprofile!\Desktop"
    set "start_menu=!appdata!\Microsoft\Windows\Start Menu\Programs"
    
    REM 创建启动脚本
    echo @echo off > "start_tool.bat"
    echo chcp 65001 >> "start_tool.bat"
    echo echo 启动TextCaptureTool... >> "start_tool.bat"
    echo cd /d "%~dp0" >> "start_tool.bat"
    echo python main.py >> "start_tool.bat"
    echo if errorlevel 1 pause >> "start_tool.bat"
    
    REM 创建桌面快捷方式
    powershell -Command "
    $WshShell = New-Object -comObject WScript.Shell;
    $Shortcut = $WshShell.CreateShortcut('!desktop!\\TextCaptureTool.lnk');
    $Shortcut.TargetPath = 'cmd.exe';
    $Shortcut.Arguments = '/c \"\"%~dp0start_tool.bat\"\"';
    $Shortcut.WorkingDirectory = '%~dp0';
    $Shortcut.WindowStyle = 7;
    $Shortcut.Description = 'TextCaptureTool - 文本捕获工具';
    $Shortcut.IconLocation = '%~dp0text_capture.ico, 0';
    $Shortcut.Save();
    " 2>nul
    
    if exist "!desktop!\TextCaptureTool.lnk" (
        call :color 0b
        echo    ✓ 桌面快捷方式创建成功
        call :color 07
    ) else (
        call :color 0e
        echo    ⚠️ 桌面快捷方式创建失败
        call :color 07
    )
    
    REM 创建开始菜单快捷方式
    if not exist "!start_menu!" mkdir "!start_menu!" >nul 2>&1
    
    powershell -Command "
    $WshShell = New-Object -comObject WScript.Shell;
    $Shortcut = $WshShell.CreateShortcut('!start_menu!\\TextCaptureTool.lnk');
    $Shortcut.TargetPath = 'cmd.exe';
    $Shortcut.Arguments = '/c \"\"%~dp0start_tool.bat\"\"';
    $Shortcut.WorkingDirectory = '%~dp0';
    $Shortcut.WindowStyle = 7;
    $Shortcut.Description = 'TextCaptureTool - 文本捕获工具';
    $Shortcut.IconLocation = '%~dp0text_capture.ico, 0';
    $Shortcut.Save();
    " 2>nul
    
    if exist "!start_menu!\TextCaptureTool.lnk" (
        call :color 0b
        echo    ✓ 开始菜单快捷方式创建成功
        call :color 07
    ) else (
        call :color 0e
        echo    ⚠️ 开始菜单快捷方式创建失败
        call :color 07
    )
    
    goto :eof

:verify_installation
    echo.
    call :color 0a
    echo [4/6] 验证安装...
    call :color 07
    
    REM 测试导入关键模块
    echo    测试导入PyQt5...
    python -c "import PyQt5.QtWidgets; print('    ✓ PyQt5导入成功')" 2>nul
    if errorlevel 1 (
        call :color 0c
        echo    ❌ PyQt5导入失败
        call :color 07
        goto :eof
    )
    
    echo    测试导入python-docx...
    python -c "import docx; print('    ✓ python-docx导入成功')" 2>nul
    if errorlevel 1 (
        call :color 0c
        echo    ❌ python-docx导入失败
        call :color 07
        goto :eof
    )
    
    echo    测试主程序启动...
    timeout /t 2 /nobreak >nul
    call :color 0b
    echo    ✓ 所有模块导入成功
    call :color 07
    goto :eof

:complete_installation
    echo.
    call :color 0a
    echo [5/6] 完成安装...
    call :color 07
    
    REM 创建配置文件目录
    set "config_dir=!appdata!\TextCaptureTool"
    if not exist "!config_dir!" mkdir "!config_dir!" >nul 2>&1
    
    REM 创建卸载脚本
    echo @echo off > "uninstall.bat"
    echo chcp 65001 >> "uninstall.bat"
    echo echo 正在卸载TextCaptureTool... >> "uninstall.bat"
    echo echo 这将删除快捷方式，但保留您的配置文件 >> "uninstall.bat"
    echo set /p "confirm=确认卸载？(y/n): " >> "uninstall.bat"
    echo if /i "%%confirm%%"=="y" ( >> "uninstall.bat"
    echo     del "!desktop!\TextCaptureTool.lnk" 2^>nul >> "uninstall.bat"
    echo     del "!start_menu!\TextCaptureTool.lnk" 2^>nul >> "uninstall.bat"
    echo     echo 卸载完成！ >> "uninstall.bat"
    echo ) else ( >> "uninstall.bat"
    echo     echo 取消卸载 >> "uninstall.bat"
    echo ) >> "uninstall.bat"
    echo pause >> "uninstall.bat"
    
    call :color 0b
    echo    ✓ 安装完成
    call :color 07
    goto :eof

:color
    echo %DEL% > "%~dp0color.tmp"
    findstr /a:%1 /r "^$" "%~dp0color.tmp" nul
    del "%~dp0color.tmp" >nul 2>&1
    goto :eof

:error_exit
    echo.
    call :color 0c
    echo ========================================
    echo           安装失败！
    echo ========================================
    call :color 07
    echo.
    echo 请检查以上错误信息，解决问题后重新运行安装程序
    echo.
    pause
    exit /b 1

:eof
    echo.
    call :color 0a
    echo ========================================
    echo           TextCaptureTool
    echo           安装成功！
    echo ========================================
    call :color 07
    echo.
    echo 📋 使用方法：
    echo   1. 在此目录运行：python main.py
    echo   2. 或双击 start_tool.bat 文件
    echo.
    echo 💡 程序功能：
    echo   - 自动捕获鼠标选中的文本
    echo   - 自动保存到Word文档
    echo   - 系统托盘显示，右键可控制
    echo.
    echo 📁 重要文件：
    echo   - 配置文件：%%APPDATA%%\TextCaptureTool\config.json
    echo   - 启动脚本：start_tool.bat
    echo.
    echo 🔧 如需卸载：手动删除程序目录即可
    echo.
    pause