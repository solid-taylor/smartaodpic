# -*- mode: python ; coding: utf-8 -*-

block_cipher = pyi_crypto.PyiBlockCipher(key='6e28644f2fe3301')


a = Analysis(['P:\\Temp\\Tapio_Develop\\extract.py'],
             pathex=['P:\\Temp\\Tapio_Develop\\venv\\Scripts'],
             binaries=[],
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
          name='SmartAOD',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir='P:\\Temp\\Tapio_Develop\\pyinst_temp',
          console=True , icon='P:\\Temp\\Tapio_Develop\\.local_misc\\icon8\\app.ico')
