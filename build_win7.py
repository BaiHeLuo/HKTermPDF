"""
Win7 兼容打包脚本
自动下载 Python 3.8 → 创建虚拟环境 → 安装依赖 → 用 PyInstaller 打包
生成可在 Windows 7 上运行的 exe

使用方法: python build_win7.py
"""

import os
import sys
import subprocess
import urllib.request
import shutil

# ===== 配置 =====
PYTHON38_VERSION = "3.8.10"
PYTHON38_INSTALLER_URL = f"https://www.python.org/ftp/python/{PYTHON38_VERSION}/python-{PYTHON38_VERSION}-amd64.exe"

# 项目目录
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON38_DIR = os.path.join(PROJECT_DIR, "_python38")
VENV_DIR = os.path.join(PROJECT_DIR, "_venv38")
BUILD_DIR = os.path.join(PROJECT_DIR, "build")
DIST_DIR = os.path.join(PROJECT_DIR, "dist")

# 依赖（兼容 Python 3.8 的版本）
DEPENDENCIES = [
    "flask==2.3.3",
    "openpyxl==3.1.5",
    "PyMuPDF==1.23.8",
    "pyinstaller==6.6.0",
    "Werkzeug==2.3.8",
    "Jinja2==3.1.4",
    "MarkupSafe==2.1.5",
    "itsdangerous==2.1.2",
    "click==8.1.7",
    "blinker==1.7.0",
]


def run(cmd, cwd=None, check=True):
    """运行命令并打印"""
    print(f"\n>>> {cmd}")
    result = subprocess.run(cmd, shell=True, cwd=cwd, check=check)
    return result


def download_python38():
    """下载 Python 3.8 安装包"""
    installer = os.path.join(PROJECT_DIR, f"_python-{PYTHON38_VERSION}-amd64.exe")
    if os.path.exists(installer):
        print(f"  安装包已存在: {installer}")
        return installer

    print(f"  正在下载 Python {PYTHON38_VERSION}...")
    urllib.request.urlretrieve(PYTHON38_INSTALLER_URL, installer)
    print(f"  下载完成: {installer}")
    return installer


def install_python38():
    """安装 Python 3.8 到项目目录"""
    python38_exe = os.path.join(PYTHON38_DIR, "python.exe")
    if os.path.exists(python38_exe):
        print(f"  Python 3.8 已安装: {PYTHON38_DIR}")
        return python38_exe

    print("  正在安装 Python 3.8...")
    installer = download_python38()

    # 静默安装到指定目录
    cmd = (
        f'"{installer}" /quiet '
        f'TargetDir="{PYTHON38_DIR}" '
        f'InstallAllUsers=0 '
        f'PrependPath=0 '
        f'Include_test=0 '
        f'Include_launcher=0'
    )
    run(cmd, check=False)

    if not os.path.exists(python38_exe):
        print(f"  错误: 安装失败，未找到 {python38_exe}")
        sys.exit(1)

    print(f"  Python 3.8 安装完成: {PYTHON38_DIR}")
    return python38_exe


def create_venv(python38_exe):
    """创建虚拟环境"""
    venv_python = os.path.join(VENV_DIR, "Scripts", "python.exe")
    if os.path.exists(venv_python):
        print(f"  虚拟环境已存在: {VENV_DIR}")
        return venv_python

    print("  正在创建虚拟环境...")
    run(f'"{python38_exe}" -m venv "{VENV_DIR}"')
    print(f"  虚拟环境创建完成: {VENV_DIR}")
    return venv_python


def install_deps(venv_python):
    """安装依赖"""
    venv_pip = os.path.join(VENV_DIR, "Scripts", "pip.exe")
    print("  正在升级 pip...")
    run(f'"{venv_python}" -m pip install --upgrade pip')
    print("  正在安装依赖...")
    deps_str = " ".join(f'"{d}"' for d in DEPENDENCIES)
    run(f'"{venv_pip}" install {deps_str}')
    print("  依赖安装完成")


def build_exe(venv_python):
    """用 PyInstaller 打包"""
    print("  正在打包（可能需要几分钟）...")

    # 清理旧的构建文件
    for d in [BUILD_DIR, DIST_DIR]:
        if os.path.exists(d):
            shutil.rmtree(d)

    spec_file = os.path.join(PROJECT_DIR, "app.spec")
    venv_pyinstaller = os.path.join(VENV_DIR, "Scripts", "pyinstaller.exe")

    run(f'"{venv_pyinstaller}" "{spec_file}" --noconfirm --distpath "{DIST_DIR}" --workpath "{BUILD_DIR}"')

    exe_path = os.path.join(DIST_DIR, "勞工通知書生成器.exe")
    if os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        print(f"\n  ✓ 打包成功！")
        print(f"  ✓ 文件: {exe_path}")
        print(f"  ✓ 大小: {size_mb:.1f} MB")
        return exe_path
    else:
        print("  ✗ 打包失败，未找到 exe 文件")
        sys.exit(1)


def main():
    print("=" * 55)
    print("  Win7 兼容打包工具")
    print("  将使用 Python 3.8 生成兼容 Windows 7 的 exe")
    print("=" * 55)

    # 步骤 1: 安装 Python 3.8
    print("\n[1/4] 安装 Python 3.8...")
    python38_exe = install_python38()

    # 步骤 2: 创建虚拟环境
    print("\n[2/4] 创建 Python 3.8 虚拟环境...")
    venv_python = create_venv(python38_exe)

    # 步骤 3: 安装依赖
    print("\n[3/4] 安装项目依赖...")
    install_deps(venv_python)

    # 步骤 4: 打包
    print("\n[4/4] 使用 PyInstaller 打包...")
    exe_path = build_exe(venv_python)

    print("\n" + "=" * 55)
    print("  全部完成！")
    print(f"  Win7 兼容的 exe 位于: {exe_path}")
    print("  将此 exe 和「使用说明.txt」一起发给用户即可。")
    print("=" * 55)


if __name__ == "__main__":
    main()
