# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['selenium_gui.py'],
             pathex=['C:\\Users\\jjen\\PycharmProjects\\excel_data\\GUI\\Web'],
             binaries=[('C:\\Users\\jjen\\PycharmProjects\\excel_data\\GUI\\Web\\chromedriver.exe', '.')],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='selenium_gui_one_cache_timer',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          manifest=None)
