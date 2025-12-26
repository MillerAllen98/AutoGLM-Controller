# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['autoglm_IDE.py'],
    pathex=[],
    binaries=[],
    datas=[('phone_agent', 'phone_agent')],
    hiddenimports=['phone_agent', 'phone_agent.agent', 'phone_agent.agent_ios', 'phone_agent.model', 'phone_agent.device_factory'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AutoGLM_Controller',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AutoGLM_Controller',
)
app = BUNDLE(
    coll,
    name='AutoGLM_Controller.app',
    icon=None,
    bundle_identifier=None,
)
