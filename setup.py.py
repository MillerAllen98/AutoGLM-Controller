from setuptools import setup

APP = ['autoglm_combat_platform.py']
DATA_FILES = [('', ['logo.png'])]
OPTIONS = {
    'argv_emulation': False,
    'iconfile': 'logo.png',
    'packages': ['tkmacosx', 'tkinter', 'threading', 'subprocess', 'datetime'],
    'plist': {
        'CFBundleName': '中央信息作战平台',
        'CFBundleDisplayName': '中央信息作战平台',
        'CFBundleIdentifier': 'com.combat.platform',
        'LSMinimumSystemVersion': '11.0',
        'NSHighResolutionCapable': True,
    }
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)