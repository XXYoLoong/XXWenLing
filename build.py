import PyInstaller.__main__
import os
import sys
import shutil

def create_installer():
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 首先使用PyInstaller打包程序
    PyInstaller.__main__.run([
        'ui_main.py',  # 主程序文件
        '--name=文档合并工具',  # 生成的exe名称
        '--onefile',  # 打包成单个exe文件
        '--noconsole',  # 不显示控制台窗口
        '--clean',  # 清理临时文件
        '--add-data=README.md;.',  # 添加说明文档
        '--add-data=assets;assets',  # 添加资源文件
    ])
    
    # 创建安装程序脚本
    iss_content = """
#define MyAppName "文档合并工具"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "倪家诚"
#define MyAppExeName "文档合并工具.exe"
#define MyAppCopyright "© 倪家诚 2025/5/29"

[Setup]
AppId={{A1B2C3D4-E5F6-4A5B-8C7D-9E0F1A2B3C4D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppCopyright={#MyAppCopyright}
DefaultDirName={autopf}\\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer
OutputBaseFilename=文档合并工具安装程序
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\\*"; DestDir: "{app}\\assets"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\\{#MyAppName}"; Filename: "{app}\\{#MyAppExeName}"
Name: "{autodesktop}\\{#MyAppName}"; Filename: "{app}\\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
"""
    
    # 写入安装程序脚本
    with open("installer.iss", "w", encoding="utf-8") as f:
        f.write(iss_content)
    
    print("请按照以下步骤完成安装程序打包：")
    print("1. 下载并安装 Inno Setup: https://jrsoftware.org/isdl.php")
    print("2. 安装完成后，右键点击 installer.iss 文件")
    print("3. 选择 'Compile' 选项")
    print("4. 等待编译完成，安装程序将生成在 installer 目录中")

if __name__ == "__main__":
    create_installer() 