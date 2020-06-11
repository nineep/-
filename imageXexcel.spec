# -*- mode: python ; coding: utf-8 -*-
import site
import os
import shutil


block_cipher = None

tkfilebrowser_path = os.path.join(site.getsitepackages()[1], 'tkfilebrowser')

a = Analysis(['imageXexcel.py'],
             pathex=[],
             binaries=[],
             datas=[(tkfilebrowser_path, 'tkfilebrowser')],
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
          [('imageXexcel.png', 'assets/imageXexcel.png', 'DATA')],
          name='imageXexcel',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False, ico='assets/imageXexcel.ico')

shutil.copyfile('config.ini', 'dist/config.ini')
shutil.copyfile('assets/附件2.标准勘察表--基站名.xlsx', 'dist/附件2.标准勘察表--基站名.xlsx')
shutil.copyfile('config-template.ini', 'dist/config-template.ini')
if os.path.exists('imageXexcel'):
    shutil.rmtree('imageXexcel')
    os.rename('dist', 'imageXexcel')
shutil.make_archive('imageXexcel', 'zip', '.')