# -*- mode: python ; coding: utf-8 -*-

"""
PyInstaller 打包配置文件
用途：将Python项目打包成单个EXE文件，方便在没有Python环境的电脑上运行

使用命令：
    pyinstaller build.spec

打包原理：
    1. 分析(Analysis): 扫描所有Python文件和依赖，收集需要打包的内容
    2. 压缩(PYZ): 将Python源码编译为.pyc并打包
    3. 执行文件(EXE): 生成最终的.exe可执行文件
"""

# ============================================================
# 1. 分析阶段 - 收集所有需要打包的文件和依赖
# ============================================================
block_cipher = None
a = Analysis(
    # 入口文件：程序从这里开始执行
    ['main.py'],
    
    # 搜索路径：额外添加的模块搜索路径（空表示使用默认）
    pathex=[],
    
    # 二进制文件：需要打包的动态链接库等（空表示自动收集）
    binaries=[],
    
    # 数据文件：需要打包的非Python文件（配置目录）
    # 将 config 目录下的所有文件打包，运行时解压到 ./config
    datas=[('config', 'config')],
    
    # 隐藏导入：PyInstaller无法自动检测到的模块
    # win32com、pythoncom、pywintypes、win32file 是通过动态方式导入的，
    # 需要手动声明，否则打包后会出现 "ModuleNotFoundError"
    hiddenimports=[
        'win32com',
        'pythoncom',
        'pywintypes',
        'win32file',
    ],
    
    # 钩子路径：自定义打包钩子（空表示使用默认）
    hookspath=[],
    
    # 钩子配置：额外的配置选项（空表示使用默认）
    hooksconfig={},
    
    # 运行时钩子：在运行时执行的脚本（空表示不需要）
    runtime_hooks=[],
    
    # 排除模块：不需要打包的模块（空表示全部打包）
    excludes=[],
    
    # Windows特定选项：不优先使用重定向（保持默认）
    win_no_prefer_redirects=False,
    
    # Windows特定选项：不将程序作为私有程序集（保持默认）
    win_private_assemblies=False,
    
    # 加密密钥：用于加密Python字节码（空表示不加密）
    cipher=block_cipher,
    
    # 是否压缩：将多个文件打包到一个文件中（False表示不压缩）
    noarchive=False,
)

# ============================================================
# 2. 压缩阶段 - 将Python源码编译并打包
# ============================================================
pyd = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ============================================================
# 3. 生成EXE阶段 - 创建可执行文件
# ============================================================
exe = EXE(
    pyd,                        # 压缩后的Python字节码
    a.scripts,                  # 脚本文件
    a.binaries,                 # 二进制文件
    a.zipfiles,                 # ZIP文件
    a.datas,                    # 数据文件
    [],                         # 额外文件
    name='Excel批量生成工具',    # 生成的EXE文件名（不含.exe后缀）
    debug=False,                # 是否开启调试模式（False表示不输出调试信息）
    bootloader_ignore_signals=False,  # 是否忽略信号（保持默认）
    strip=False,                # 是否去除调试符号（False表示保留）
    upx=True,                   # 是否使用UPX压缩（True表示压缩减小体积）
    upx_exclude=[],             # UPX排除列表（空表示全部压缩）
    runtime_tmpdir=None,        # 运行时临时目录（空表示使用系统默认）
    console=False,              # 是否显示控制台窗口（False表示只显示GUI窗口）
    disable_windowed_traceback=False,  # 是否禁用窗口化异常信息（保持默认）
    argv_emulation=False,       # 是否模拟命令行参数（保持默认）
    target_arch=None,           # 目标架构（空表示自动检测）
    codesign_identity=None,     # 代码签名身份（空表示不签名）
    entitlements_file=None,     # 权限文件（空表示使用默认）
    #icon='icon.ico',            # 程序图标文件（需要自行准备icon.ico）
)