# -*- mode: python ; coding: utf-8 -*-
import os, sys
from PyInstaller.utils.hooks import (
    collect_data_files, collect_submodules, collect_dynamic_libs
)
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT

block_cipher  = None
project_root  = os.getcwd()
spec_dir      = os.path.join(project_root, 'source')
script_path   = os.path.join(spec_dir, 'OCTec.py')

# OCRmyPDF
ocrmypdf_data   = collect_data_files('ocrmypdf')
ocrmypdf_hidden = collect_submodules('ocrmypdf')

# OpenCV / NumPy
cv2_data     = collect_data_files('cv2')
cv2_bins     = collect_dynamic_libs('cv2')
numpy_hidden = collect_submodules('numpy')

# PikePDF (qpdf*.dll)
pikepdf_hidden = collect_submodules('pikepdf')
pikepdf_bins   = collect_dynamic_libs('pikepdf')

# PyMuPDF (fitz/mupdf*.dll)
pymupdf_bins = collect_dynamic_libs('fitz')

# Endesive e deps
endesive_hidden = (
    collect_submodules('endesive')
    + collect_submodules('oscrypto')
    + collect_submodules('asn1crypto')
)

# Watchdog (garante backends/observers no frozen)
watchdog_hidden = collect_submodules('watchdog')

# Certificados HTTPS/OCSP etc.
certifi_data = collect_data_files('certifi')

# PySide6 (plugins e afins)
pyside6_data = collect_data_files('PySide6')

# (Opcional) PIL: descomente se faltar plugins de imagem em alguma máquina
pil_data = collect_data_files('PIL')
# ------------------------------------------------------

resources_tree = [(os.path.join(project_root, 'resources'), 'resources')]
datas    = ocrmypdf_data + cv2_data + pyside6_data + certifi_data + resources_tree  # + pil_data (se usar)
binaries = cv2_bins + pikepdf_bins + pymupdf_bins
runtime_hooks = [os.path.join(spec_dir, 'add_to_path.py')]

hidden = (
    ocrmypdf_hidden + numpy_hidden + pikepdf_hidden +
resources_tree = [(os.path.join(project_root, 'resources'), 'resources')]
datas    = ocrmypdf_data + cv2_data + pyside6_data + certifi_data + resources_tree  # + pil_data (se usar)
binaries = cv2_bins + pikepdf_bins + pymupdf_bins
runtime_hooks = [os.path.join(spec_dir, 'add_to_path.py')]
 endesive_hidden + watchdog_hidden
)

a = Analisys(
    [script_path],
    pathex=[project_root],
    binaries=binaries,
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    runtime_hooks=runtime_hooks,
    excludes=[],
    cipher=block_cipher
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    exclude_binaries=True,
    name='OCTec',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon=os.path.join(project_root, 'resources', 'app_icon.ico')
)

coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas,
    strip=False, upx=False, upx_exclude=[], name='OCTec'
)
