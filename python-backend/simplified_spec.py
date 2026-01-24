# -*- mode: python ; coding: utf-8 -*-

import os

# 定义分析配置
a = Analysis(
    ['cli.py'],
    pathex=[os.path.abspath('.')],
    hiddenimports=[
        'rules',
        'core.engine',
        'core.context',
        'utils.port_manager',
        'core.config_loader',
        'core.config_repository',
        'core.yaml_config_repository',
        'ruamel.yaml',
        'ruamel.yaml.scalarstring',
        'ruamel.yaml.representer',
        'ruamel.yaml.constructor',
        'ruamel.yaml.nodes',
        'ruamel.yaml.parser',
        'ruamel.yaml.scanner',
        'ruamel.yaml.resolver',
        'ruamel.yaml.dumper',
        'ruamel.yaml.loader',
        'pyyaml',
        'yaml',
    ],
    datas=[
        # 配置文件
        ('config/presets.yaml', 'config'),
        # 规则文件 - 递归包含整个rules目录
        ('rules', 'rules'),
    ],
    binaries=[],
    runtime_hooks=[],
    hookspath=[],
    hooksconfig={},
    module_aliases={},
    noarchive=False,
    debug=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='word_format_fixer_backend',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# 使用COLLECT创建onedir模式的输出目录
after_build = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='word_format_fixer_backend',
)
