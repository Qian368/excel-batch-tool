# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec文件配置
# DOS终端执行命令：PyInstaller excel_tool.spec

# 导入os模块用于路径操作
# import os

# 加密配置
block_cipher = None  # 设置加密方式，None表示不加密

# UPX配置 - 用于压缩可执行文件
upx_dir = r'D:\Software Data\upx'  # UPX工具路径，使用原始字符串避免转义问题
upx = True  # 是否启用UPX压缩，True表示启用

# 设置输出目录 - 使用绝对路径指向上级目录的release文件夹
# 注意：PyInstaller在处理相对路径时可能会有问题，建议使用绝对路径
# 由于在PyInstaller执行spec文件时__file__变量不存在，因此使用当前工作目录
# current_dir = os.getcwd()  # 获取当前工作目录
# 确保release目录存在
# release_dir = os.path.abspath(os.path.join(current_dir, '..', 'release'))
# if not os.path.exists(release_dir):
#    os.makedirs(release_dir)
# dist_path = release_dir  # 编译输出目录

# 设置图标路径 - 使用绝对路径
# icon_path = os.path.abspath(os.path.join(current_dir, '.', 'ui', 'image', 'icon.ico'))

# 主程序分析配置
a = Analysis(
    ['__main__.py'],  # 主程序入口文件                                					自定义配置
    pathex=[],  # 额外Python路径
    binaries=[],  # 需要包含的二进制文件
    datas=[(r'.\ui\image\icon.ico', 'ui/image')],  # 需要包含的非Python文件，显式添加应用里面的图标资源   自定义配置
    # distpath='..\release',  # 指定输出目录为release目录								自定义配置
    hiddenimports=[],  # 需要显式导入的隐藏模块
    hookspath=[],  # 自定义hook路径
    hooksconfig={},  # hook配置
    runtime_hooks=[],  # 运行时hook
    excludes=[],  # 需要排除的模块
    win_no_prefer_redirects=False,  # Windows重定向设置
    win_private_assemblies=False,  # Windows私有程序集
    cipher=block_cipher,  # 加密方式
    noarchive=False,  # 是否创建归档
    upx=upx,  # 是否使用UPX压缩
    upx_exclude=[],  # UPX排除列表
    name='Excel批量操作工具'  # 应用名称                               					自定义配置
)

# PYZ输出配置 - 创建Python字节码归档
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# EXE可执行文件配置
exe = EXE(
    pyz,  # Python字节码归档
    a.scripts,  # 脚本文件
    [],  # 附加选项
    exclude_binaries=True,  # 排除二进制文件，避免重复打包
    name='Excel批量操作工具',  # 输出文件名
    # distpath='..\release',  # 指定输出目录为release目录
    debug=False,  # 是否启用调试
    bootloader_ignore_signals=False,  # 是否忽略信号
    strip=False,  # 是否去除调试信息
    upx=upx,  # 是否使用UPX压缩                                
    upx_exclude=[],  # UPX排除列表                               			可以自定义配置
    runtime_tmpdir=None,  # 临时目录
    console=False,  # 是否显示控制台窗口                                	可以自定义配置
    disable_windowed_traceback=False,  # 是否禁用窗口化traceback
    argv_emulation=False,  # 是否模拟argv
    target_arch=None,  # 目标架构
    codesign_identity=None,  # 代码签名标识
    entitlements_file=None,  # 授权文件
    uac_admin=False,  # 是否请求管理员权限                               	可以自定义配置
    icon=r'..\resources\image\icon.ico'  # 应用图标路径                						自定义配置
)

# 收集依赖文件 - 用于单文件夹分发模式
coll = COLLECT(
    exe,  # 可执行文件
    a.binaries,  # 二进制文件
    a.zipfiles,  # zip文件
    a.datas,  # 数据文件
    # distpath='..\release',  # 指定输出目录为release目录
    strip=False,  # 是否去除调试信息和符号表
    upx=True,  # 使用UPX压缩依赖文件
    upx_exclude=[],  # UPX排除列表
    name='upx_output',  # 输出目录名称                                
    debug=False,  # 是否启用调试模式
    bootloader_ignore_signals=False,  # 是否忽略引导加载程序信号
    console=False,  # 是否显示控制台窗口
    disable_windowed_traceback=False,  # 是否禁用窗口化的回溯信息
    argv_emulation=False,  # 是否模拟命令行参数
    target_arch=None,  # 目标架构，None表示使用当前系统架构
    codesign_identity=None,  # 代码签名标识，用于macOS
    entitlements_file=None,  # 授权文件路径，用于macOS	
    # icon=r'..\resources\image\icon.ico' # 应用图标路径,这里收集图标过去程序没有意义          	自定义配置
)