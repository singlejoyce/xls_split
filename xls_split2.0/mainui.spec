# -*- mode: python -*-

block_cipher = None


a = Analysis(['mainui.py'],
             pathex=['mythread.py', 'D:\\python_prj\\xls_split2.0'],
             binaries=[],
             datas=[],
             hiddenimports=['mythread'],
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
          name='mainui',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False , icon='excel.ico')
