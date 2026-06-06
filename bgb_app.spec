# PyInstaller spec — BGB Notice Generator
# Produces a proper macOS .app bundle (onedir mode)
# Build: pyinstaller bgb_app.spec --noconfirm --clean

block_cipher = None

a = Analysis(
    ['app_gui.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'docx',
        'docx.shared',
        'docx.enum.text',
        'docx.enum.table',
        'docx.oxml',
        'docx.oxml.ns',
        'dateutil',
        'dateutil.parser',
        'dateutil.relativedelta',
        'excel_parser',
        'letter_generator',
        'amount_to_words',
        'utils',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'pandas', 'PIL', 'scipy'],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,   # onedir mode
    name='BGB_Notice_Generator',
    debug=False,
    strip=False,
    upx=True,
    console=False,           # no terminal window
    target_arch=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='BGB_Notice_Generator',
)

# macOS .app bundle
app = BUNDLE(
    coll,
    name='BGB Notice Generator.app',
    icon=None,
    bundle_identifier='in.gov.mes.bgb-notice-generator',
    info_plist={
        'NSHighResolutionCapable': True,
        'CFBundleShortVersionString': '1.0.0',
        'CFBundleName': 'BGB Notice Generator',
    },
)
