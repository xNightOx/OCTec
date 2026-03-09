# source/add_to_path.py

import os, sys

# dist_root = pasta onde está OCTecOCR.exe após a instalação
dist_root = os.path.dirname(sys.executable)

# pasta _internal fica ao lado do executável
internal = os.path.join(dist_root, '_internal')

# estas são as pastas que têm os .exe que você vai chamar
paths = [
    # Caso você copie o ocrmypdf.exe direto em _internal
    internal,

    # Tesseract (contém tesseract.exe, tesseract-lt.exe etc)
    os.path.join(internal, 'resources', 'Tesseract-OCR'),

    # Poppler → subpasta bin
    os.path.join(internal, 'resources', 'poppler', 'bin'),

    # Ghostscript → subpasta bin
    os.path.join(internal, 'resources', 'gs10.04.0', 'bin'),
]

# Insere cada caminho no início do PATH do processo
orig = os.environ.get('PATH', '')
for p in paths:
    if os.path.isdir(p) and p not in orig:
        orig = p + os.pathsep + orig

os.environ['PATH'] = orig
