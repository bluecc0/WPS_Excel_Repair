#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WPS Excel修复工具 - EXE打包脚本
"""

import os
import subprocess
import sys
import shutil

def build_exe():
    """打包成EXE文件"""
    print("=== WPS Excel修复工具 EXE打包 ===")
    
    # 检查PyInstaller是否已安装
    try:
        import PyInstaller
        print(f"检测到PyInstaller版本: {PyInstaller.__version__}")
    except ImportError:
        print("未检测到PyInstaller，正在安装...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        print("PyInstaller安装完成")
    
    # 清理旧的构建文件
    print("\n清理旧文件...")
    for folder in ["build", "dist", "__pycache__"]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
            print(f"删除 {folder}")
    
    # 检查必要文件
    required_files = ["wps_repair_standalone.py", "assset/icon.ico"]
    for file in required_files:
        if not os.path.exists(file):
            print(f"缺少必要文件: {file}")
            return False
        print(f"检测到: {file}")
    
    # PyInstaller命令参数
    cmd = [
        "pyinstaller",
        "--onefile",                    # 打包成单个EXE文件
        "--windowed",                   # 无控制台窗口
        "--icon=assset/icon.ico",       # 设置图标
        "--name=WPS_Excel_Repair_Tool", # EXE文件名
        "--add-data=assset/icon.ico;assset",  # 包含图标文件
        "--distpath=.",                 # 输出到当前目录
        "wps_repair_standalone.py"      # 主程序文件
    ]
    
    print(f"\n开始打包...")
    print(f"执行命令: {' '.join(cmd)}")
    
    try:
        # 执行打包命令
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        
        print("打包成功!")
        
        # 检查生成的EXE文件
        exe_path = "WPS_Excel_Repair_Tool.exe"
        if os.path.exists(exe_path):
            size = os.path.getsize(exe_path) / (1024 * 1024)  # MB
            print(f"生成文件: {exe_path} ({size:.1f} MB)")
            
            # 清理构建文件
            print("\n清理构建文件...")
            for folder in ["build"]:
                if os.path.exists(folder):
                    shutil.rmtree(folder)
                    print(f"删除 {folder}")
            
            # 删除spec文件
            spec_file = "WPS_Excel_Repair_Tool.spec"
            if os.path.exists(spec_file):
                os.remove(spec_file)
                print(f"删除 {spec_file}")
            
            print(f"\n打包完成!")
            print(f"EXE文件位置: {os.path.abspath(exe_path)}")
            print(f"使用方法: 将.xlsx文件拖拽到EXE图标上即可修复")
            return True
        else:
            print("EXE文件生成失败")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"打包失败: {e}")
        if e.stdout:
            print(f"输出: {e.stdout}")
        if e.stderr:
            print(f"错误: {e.stderr}")
        return False

if __name__ == "__main__":
    success = build_exe()
    if not success:
        sys.exit(1)
    
    print("\n按任意键退出...")
    input()