📋 项目简介
本工具用于根据 Excel 模板批量生成品质检验报告（首件/过程/成品），自动填充日期、随机数、检验员等信息，并设置文件时间属性。支持自定义生产工期、非工作日、内容修改规则，配置可持久化保存。

适用场景
制造业品质部门批量生成检验报告

需要按生产工期自动生成对应日期的检验表格

模板格式复杂（含合并单元格、上下标、斜线等），需保留原样

需要模拟真实文件时间（创建/修改/访问时间）

技术栈
组件	技术
GUI 框架	PyQt5
Excel 操作	win32com (调用本地 Excel 应用程序)
日志系统	Python logging
配置存储	JSON
打包工具	PyInstaller
Python 版本	3.13.x
✨ 主要功能
1. 三种模板类型支持
首件：生成工期第一天 + 每段工作日的第一天

过程：生成工期内所有工作日（排除非工作日）

成品：生成工期最后一天 + 每段工作日的最后一天

2. 灵活的日期生成规则
支持连续范围（如 2026-06-11 ~ 2026-06-18）

支持多段范围（可添加多组日期段）

支持单天生成

自动识别并跳过非工作日

日期范围按产品+模板类型分别保存

3. 可视化配置界面
模板拖入/浏览加载

日历控件选择日期范围

非工作日点击切换（高亮显示）

内容规则增删改查

4. 非工作日管理
日历点击切换，高亮显示所有已选日期

支持导入/导出 JSON 配置

配置自动保存/读取

一键清空所有

5. 内容修改规则
支持对 Excel 单元格进行多种修改：

值类型	说明	示例
随机数	生成指定范围的随机数	C9-L9: 92.31-92.69
日期替换	替换单元格中的日期	D3: 2026-06-11
文本中的日期	替换文本中的所有日期	A28: "检验日期：2026-06-11"
夜班检验员	根据日期自动替换	A18: "夜班检验员：颜大丰/张志宇"
支持指定工作表

支持行范围、列范围、单个单元格、离散单元格

规则按产品+模板类型分别保存

支持规则导入/导出

6. 批量生成与进度监控
多线程生成（界面不卡顿）

实时进度条

预估剩余时间

支持中止取消

断点续做（跳过已存在文件）

7. 文件时间自动设置
时间属性	设置规则
创建时间	日期当天 8:30~9:00 随机
修改时间	创建时间 + 30~60 分钟随机
访问时间	修改时间 + 0~2 分钟随机
8. 配置持久化
非工作日配置：config/non_workdays.json

产品规则配置：config/products/{产品名}/{模板类型}.json

日志配置：config/log_config.json

支持配置导入/导出

9. 完善的日志系统
四级日志：DEBUG / INFO / WARNING / ERROR

日志级别可通过配置文件控制

按天滚动，自动清理旧日志

DEBUG 级别记录每个单元格的修改详情

📁 项目目录结构
text
excel_batch_generator/
│
├── main.py                      # 程序入口
├── requirements.txt             # Python 依赖清单
├── build.spec                   # PyInstaller 打包配置
│
├── config/                      # 配置文件目录（程序运行时自动生成）
│   ├── log_config.json          # 日志级别配置
│   ├── non_workdays.json        # 非工作日配置
│   └── products/                # 产品配置目录
│       └── {产品名}/            # 按产品名分组
│           ├── 首件.json        # 首件模板的规则配置
│           ├── 过程.json        # 过程模板的规则配置
│           └── 成品.json        # 成品模板的规则配置
│
├── ui/                          # 界面模块
│   ├── __init__.py
│   ├── main_window.py           # 主窗口（模板加载、日期范围、规则列表、生成控制）
│   ├── non_workdays_dialog.py   # 非工作日配置对话框（日历选择）
│   ├── rule_dialog.py           # 规则配置对话框（添加/编辑规则）
│   └── preview_dialog.py        # 预览对话框（显示将要生成的日期列表）
│
├── core/                        # 核心业务逻辑
│   ├── __init__.py
│   ├── config_manager.py        # 配置管理器（读写 JSON 配置文件）
│   ├── date_calculator.py       # 日期计算器（根据规则计算目标日期）
│   ├── excel_generator.py       # Excel 生成器（win32com 操作 Excel）
│   ├── generate_thread.py       # 生成线程（多线程执行，支持取消）
│   └── log_manager.py           # 日志管理器（单例模式，统一日志输出）
│
├── models/                      # 数据模型
│   ├── __init__.py
│   └── config_models.py         # 配置数据类（Rule / ProductConfig / NonWorkdaysConfig）
│
├── utils/                       # 工具函数
│   ├── __init__.py
│   └── helpers.py               # 日期解析、格式化、文本提取等工具函数
│
└── logs/                        # 日志文件目录（程序运行时自动生成）
    └── generate_2026-08-10.log  # 按天滚动的日志文件
各文件功能说明
文件/目录	功能说明
main.py	程序入口，初始化 QApplication，显示主窗口
ui/main_window.py	主窗口，所有用户交互的入口
ui/non_workdays_dialog.py	非工作日配置弹出窗口
ui/rule_dialog.py	内容规则添加/编辑弹出窗口
ui/preview_dialog.py	日期预览弹出窗口
core/config_manager.py	读写 JSON 配置，管理产品/非工作日配置
core/date_calculator.py	根据模板类型和规则计算目标日期
core/excel_generator.py	win32com 操作 Excel，应用规则，保存文件
core/generate_thread.py	独立线程执行生成任务，支持取消
core/log_manager.py	统一日志管理，支持多级别输出
models/config_models.py	数据类定义，序列化/反序列化
utils/helpers.py	日期格式解析、文本处理等工具函数
🚀 快速开始
环境要求
依赖	说明
Windows 操作系统	需要安装 Microsoft Excel（win32com 依赖）
Python 3.13.x	推荐使用 3.13 版本
pip	Python 包管理器
安装步骤
克隆或下载项目代码

安装依赖

bash
pip install -r requirements.txt
运行程序

bash
python main.py
首次运行
程序启动后，会自动创建 config/ 和 logs/ 目录

点击「浏览...」或拖入 Excel 模板文件

程序自动识别产品名称和模板类型

配置生产工期、非工作日、内容规则

点击「预览」查看将要生成的日期

点击「开始生成」选择输出目录，开始批量生成

⚙️ 配置文件说明
1. 日志配置 config/log_config.json
json
{
  "level": "INFO",          // DEBUG | INFO | WARNING | ERROR
  "max_files": 30,          // 保留最近30个日志文件
  "console_output": true,   // 同时输出到控制台
  "file_output": true,      // 写入日志文件
  "log_dir": "logs"         // 日志文件目录
}
日志级别说明：

级别	记录内容
DEBUG	最详细：每个单元格修改、规则应用、文件时间、完整堆栈
INFO	常规：程序启动/完成、文件生成成功/失败、配置加载
WARNING	警告：文件已存在跳过、规则异常、配置缺失
ERROR	错误：生成失败、模板加载失败、配置文件损坏
2. 非工作日配置 config/non_workdays.json
json
{
  "non_workdays": [
    "2026-06-07",
    "2026-06-14",
    "2026-06-19",
    "2026-06-20",
    "2026-06-28"
  ]
}
3. 产品规则配置 config/products/{产品名}/{模板类型}.json
json
{
  "product_name": "23036_N24四方管",
  "template_type": "首件",
  "date_ranges": [
    ["2026-06-11", "2026-06-18"],
    ["2026-07-02", "2026-08-11"]
  ],
  "rules": [
    {
      "id": "rule_001",
      "target_type": "range",
      "target": "C9-L9",
      "value_type": "random",
      "sheet_name": "首件检验记录",
      "min_val": 45.73,
      "max_val": 45.99,
      "decimals": 2,
      "enabled": true
    },
    {
      "id": "rule_002",
      "target_type": "cell",
      "target": "D3",
      "value_type": "date",
      "sheet_name": "首件检验记录",
      "enabled": true
    }
  ]
}
📖 使用指南
1. 加载模板
点击「浏览...」按钮选择 Excel 模板文件

或将模板文件直接拖入窗口

程序自动识别：

产品名称：从文件名提取（如 23036_N24四方管）

模板类型：关键词识别（首件/过程/成品）

2. 配置生产工期
选择起始日期和结束日期

点击「添加」加入日期范围列表

支持多段工期（点击「添加」逐条加入）

点击「清空所有」一键清空

日期范围会自动保存到产品配置中

3. 配置非工作日
点击「配置非工作日」打开日历对话框

点击日期切换非工作日标记（高亮显示）

支持跨月选择

点击「导入」/「导出」备份配置

点击「确定」自动保存

4. 配置内容修改规则
点击「添加规则」打开规则配置对话框

目标类型：行范围 / 列范围 / 单个单元格 / 离散单元格

目标位置：如 C9-L9、D6-D15、F9,K9

工作表：选择规则应用到哪个工作表（默认所有）

值类型：

随机数：设置最小/最大值和小数位数

日期替换：替换单元格中的日期

文本中的日期：替换文本中所有日期

夜班检验员：根据日期替换夜班检验员名字

支持启用/禁用规则

支持导入/导出规则配置

5. 预览与生成
点击「预览」查看将要生成的日期列表

可导出预览列表为文本文件

点击「开始生成」选择输出目录

程序开始批量生成，显示进度和剩余时间

点击「中止」可安全取消生成

生成日志自动保存到 logs/ 目录

🛠️ 打包为 EXE
使用 PyInstaller 打包
安装 PyInstaller

bash
pip install pyinstaller
执行打包命令

bash
pyinstaller build.spec
输出位置
打包后的 EXE 文件位于 dist/Excel批量生成工具.exe

打包注意事项
确保 build.spec 文件在项目根目录

打包前确认所有依赖已安装

首次打包可能需要较长时间（需打包 Python 环境）

build.spec 文件
spec
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('config', 'config')],
    hiddenimports=[
        'win32com',
        'pythoncom',
        'pywintypes',
        'win32file',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyd = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyd,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Excel批量生成工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
❓ 常见问题 (FAQ)
Q1: 生成的 Excel 文件日期变成数字串（如 46183）？
A: 这是因为 Excel 将日期字符串识别为日期序列号。程序已通过两种方式解决：

纯日期单元格：使用日期序列号 + NumberFormat 控制显示

文本中的日期：先设置单元格格式为文本，再写入字符串

如仍有问题，请确认规则类型选择正确（date vs text_with_date）。

Q2: 日期显示比预期少一天？
A: 这是因为 win32com 在传递 datetime 对象时会经过时区转换。程序已改为使用 Excel 日期序列号（浮点数）赋值，绕过了时区转换问题。请更新到最新版本的 excel_generator.py。

Q3: 夜班检验员没有替换？
A: 请确认：

规则类型选择了 night_shift

目标单元格包含「夜班检验员：」文本

模板类型为「过程」（夜班检验员仅在过程表中使用）

日期在 1-15 日替换为「颜大丰」，16-31 日替换为「张志宇」

Q4: 如何查看详细的运行日志？
A:

打开 config/log_config.json

将 "level" 改为 "DEBUG"

重新运行程序

查看 logs/generate_YYYY-MM-DD.log 文件

Q5: 配置了非工作日但生成的日期没有跳过？
A: 请确认：

非工作日已点击「确定」保存（不是关闭窗口）

非工作日日期格式为 YYYY-MM-DD

在日期计算前已重新加载配置（切换模板类型或重新加载模板时会自动刷新）

Q6: 如何备份/迁移配置？
A: 复制整个 config/ 目录即可。包含所有非工作日配置、产品规则配置和日志配置。

Q7: 生成的 Excel 格式不对或内容丢失？
A:

本工具使用 win32com（调用本地 Excel），理论上应完美保留格式

如遇到问题，请确认本地已安装 Microsoft Excel

检查规则是否误修改了不应修改的单元格

查看 DEBUG 日志确认规则应用详情

Q8: 程序启动时报 libpng warning？
A: 这是 Qt 加载 PNG 图标时产生的警告，不影响程序功能，可以忽略。已在 main.py 中屏蔽该警告。

Q9: 如何只生成特定某几天？
A: 在「生产工期」区域，将起始日期和结束日期都设置为同一天，点击「添加」即可。添加多组单天可实现任意日期的组合。

Q10: 生成过程中可以关闭程序吗？
A: 不建议在生成过程中强制关闭程序。请点击「中止」按钮安全停止生成，程序会清理资源后再退出。