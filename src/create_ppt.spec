# -*- mode: python -*-

block_cipher = None


a = Analysis(['create_ppt.py'],
             pathex=['C:\\Users\\gaodl\\PycharmProjects\\excel2ppt\\src'],
             binaries=[],
             datas=[("D:\\resource.txt", ".\\resource")],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='create_ppt',
          debug=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='create_ppt')
