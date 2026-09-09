; -*- coding: utf-8 -*-
; installer.nsi - 价格牌处理工具 v1.3 安装包脚本
;
; 用法：先运行 build_exe.py 生成 dist\价格牌处理工具.exe，
; 再运行 makensis.exe installer.nsi 即可生成安装包。
;
; 特性：
; - Modern UI 2（与腾讯会议/微信同款的现代向导外观）
; - 许可协议页 → 安装路径页 → 安装进度页 → 完成页
; - 开始菜单 + 桌面快捷方式
; - 控制面板卸载项（UninstallString / DisplayIcon / Publisher / DisplayVersion）
; - 卸载时自动清理注册表 + 快捷方式 + 安装目录

Unicode true
SetCompressor /SOLID lzma
SetDatablockOptimize on

!include "MUI2.nsh"
!include "LogicLib.nsh"

; 应用信息
Name "价格牌处理工具"
OutFile "dist\价格牌处理工具_安装程序_v1.3.exe"
InstallDir "$PROGRAMFILES64\价格牌处理工具"
InstallDirRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\价格牌处理工具" "InstallLocation"
RequestExecutionLevel highest

; 安装包自带图标与顶图（与 EXE 同款徽章）
!define MUI_ICON "assets\logo.ico"
!define MUI_UNICON "assets\logo.ico"

; Modern UI 2 设置
!define MUI_ABORTWARNING
!define MUI_WELCOMEPAGE_TITLE "欢迎使用 价格牌处理工具"
!define MUI_WELCOMEPAGE_TEXT "本安装程序将引导您完成 价格牌处理工具 v1.3 的安装。$\r$\n$\r$\n建议在安装前关闭其他应用程序，以避免与本安装程序冲突。$\r$\n$\r$\n点击下一步继续。"
!define MUI_FINISHPAGE_TITLE "安装完成"
!define MUI_FINISHPAGE_TEXT "价格牌处理工具 已成功安装到您的电脑。$\r$\n$\r$\n点击完成关闭安装程序。"
!define MUI_FINISHPAGE_RUN "$INSTDIR\价格牌处理工具.exe"
!define MUI_FINISHPAGE_RUN_TEXT "立即运行 价格牌处理工具"
!define MUI_FINISHPAGE_SHORTCUT "$SMPROGRAMS\价格牌处理工具\价格牌处理工具.lnk"
!define MUI_FINISHPAGE_SHORTCUT_TITLE "创建桌面快捷方式"
!define MUI_FINISHPAGE_SHORTCUT_DESTINATION "$DESKTOP"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

!insertmacro MUI_LANGUAGE "SimpChinese"

; 预估安装大小（用于安装进度）
Section -SET_INSTALL_SIZE
    SectionIn RO
SectionEnd

; ---- 安装段 ----
Section "主程序（必装）" SecMain
    SectionIn RO
    SetOutPath "$INSTDIR"
    File "dist\价格牌处理工具.exe"
    ; 写入控制面板"程序和功能"的卸载信息
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\价格牌处理工具" "DisplayName" "价格牌处理工具"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\价格牌处理工具" "DisplayVersion" "1.3"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\价格牌处理工具" "Publisher" "Internal"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\价格牌处理工具" "InstallLocation" "$INSTDIR"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\价格牌处理工具" "UninstallString" "$\"$INSTDIR\Uninstall.exe$\""
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\价格牌处理工具" "QuietUninstallString" "$\"$INSTDIR\Uninstall.exe$\" /S"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\价格牌处理工具" "DisplayIcon" "$\"$INSTDIR\价格牌处理工具.exe$\",0"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\价格牌处理工具" "NoModify" 1
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\价格牌处理工具" "NoRepair" 1
    ; 生成卸载器
    WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

Section "开始菜单快捷方式" SecStartMenu
    CreateDirectory "$SMPROGRAMS\价格牌处理工具"
    CreateShortcut "$SMPROGRAMS\价格牌处理工具\价格牌处理工具.lnk" "$INSTDIR\价格牌处理工具.exe" "" "$INSTDIR\价格牌处理工具.exe" 0
    CreateShortcut "$SMPROGRAMS\价格牌处理工具\卸载 价格牌处理工具.lnk" "$INSTDIR\Uninstall.exe" "" "$INSTDIR\Uninstall.exe" 0
SectionEnd

Section "桌面快捷方式" SecDesktop
    CreateShortcut "$DESKTOP\价格牌处理工具.lnk" "$INSTDIR\价格牌处理工具.exe" "" "$INSTDIR\价格牌处理工具.exe" 0
SectionEnd

; 段说明（左侧列表）
LangString DESC_SecMain ${LANG_SIMPCHINESE} "价格牌处理工具 主程序。"
LangString DESC_SecStartMenu ${LANG_SIMPCHINESE} "在开始菜单中创建快捷方式（推荐）。"
LangString DESC_SecDesktop ${LANG_SIMPCHINESE} "在桌面创建快捷方式。"

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
    !insertmacro MUI_DESCRIPTION_TEXT ${SecMain} $(DESC_SecMain)
    !insertmacro MUI_DESCRIPTION_TEXT ${SecStartMenu} $(DESC_SecStartMenu)
    !insertmacro MUI_DESCRIPTION_TEXT ${SecDesktop} $(DESC_SecDesktop)
!insertmacro MUI_FUNCTION_DESCRIPTION_END

; ---- 卸载段 ----
Section "Uninstall"
    ; 清理注册表
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\价格牌处理工具"
    ; 删除快捷方式
    Delete "$SMPROGRAMS\价格牌处理工具\价格牌处理工具.lnk"
    Delete "$SMPROGRAMS\价格牌处理工具\卸载 价格牌处理工具.lnk"
    RMDir "$SMPROGRAMS\价格牌处理工具"
    Delete "$DESKTOP\价格牌处理工具.lnk"
    ; 删除文件
    Delete "$INSTDIR\价格牌处理工具.exe"
    Delete "$INSTDIR\Uninstall.exe"
    RMDir "$INSTDIR"
    ; 弹出卸载完成提示
    MessageBox MB_OK|MB_ICONINFORMATION "价格牌处理工具 已成功从您的电脑中移除。"
SectionEnd