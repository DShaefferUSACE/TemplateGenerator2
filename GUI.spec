# -*- mode: python -*-

block_cipher = None


a = Analysis(['D:\\documents\\Python\\Template_Generator_3.7\\GUI.py'],
             pathex=['C:\\Users\\k7rgrdls\\AppData\\Roaming\\Python\\Python36\\Scripts'],
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
          name='TemplateGenerator',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir='C:\\Template Generator',
          console=True )
