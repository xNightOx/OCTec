import os
import sys
import unicodedata
import logging
from logging import handlers
import threading
import time
import queue
import json
import shutil
import re
import asyncio
import subprocess
import signal
import hashlib
import platform
import zipfile
from datetime import datetime
from pathlib import Path

import os
import uuid
import time
import shutil

def _is_unc_path(path: Path) -> bool:
    try:
        pstr = str(path)
        return pstr.startswith("\\\\") or pstr.startswith(r"\\?\\UNC\\")
    except Exception:
        return False

def _same_drive(a: Path, b: Path) -> bool:
    try:
        return a.drive.lower() == b.drive.lower()
    except Exception:
        return False

def _ensure_parent(dir_path: Path) -> None:
    try:
        dir_path.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass

def atomic_copy_replace(src: Path, dst: Path, *, retries: int = 10, base_sleep: float = 0.25) -> None:
    if not src.exists():
        raise FileNotFoundError(f"Fonte ausente para atomic_copy_replace: {src}")
    _ensure_parent(dst.parent)

    tmp = dst.with_name(f".octec_tmp_{uuid.uuid4().hex}_{dst.name}.part")
    try:
        if tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass

        with open(src, "rb") as fin, open(tmp, "wb") as fout:
            import shutil as _sh
            _sh.copyfileobj(fin, fout, length=2**20)
            try:
                fout.flush()
                os.fsync(fout.fileno())
            except Exception:
                pass

        last_err = None
        for attempt in range(1, retries + 1):
            try:
                if dst.exists():
                    try:
                        sz = dst.stat().st_size
                    except Exception:
                        sz = -1
                    if sz == 0:
                        try:
                            dst.unlink()
                        except Exception:
                            pass
                os.replace(str(tmp), str(dst))
                try:
                    if dst.stat().st_size <= 0:
                        raise OSError("Destino ficou com 0 bytes após replace.")
                except Exception as ve:
                    last_err = ve
                    raise ve
                return
            except PermissionError as e:
                last_err = e
                time.sleep(base_sleep * attempt)
                continue
            except OSError as e:
                last_err = e
                time.sleep(base_sleep * attempt)
                continue
        raise last_err if last_err else OSError("Falha desconhecida no atomic_copy_replace.")
    finally:
        try:
            if tmp.exists():
                tmp.unlink()
        except Exception:
            pass

def robust_move(src, dst) -> None:
    """Move robusto que aceita str ou Path e lida com C:\\ -> \\servidor."""
    src = Path(src)
    dst = Path(dst)
    try:
        if not _is_unc_path(dst) and _same_drive(src, dst):
            _ensure_parent(dst.parent)
            try:
                os.replace(str(src), str(dst))
                return
            except OSError:
                pass
        atomic_copy_replace(src, dst)
        try:
            if src.exists():
                src.unlink()
        except Exception:
            pass
    except Exception:
        raise
from io import BytesIO
import tempfile
import itertools
import math
import uuid
import random
from concurrent.futures import ProcessPoolExecutor

# Importações de biblioteca externas
try:
    import fitz  # PyMuPDF
    from endesive.pdf import cms
    from cryptography.hazmat.primitives.serialization import pkcs12
    from PIL import Image, ImageFilter, ExifTags
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
        QAbstractItemView, QProgressBar, QLabel, QLineEdit, QDialog,
        QMessageBox, QGroupBox, QGridLayout, QTabWidget, QCheckBox,
        QComboBox, QFileDialog, QInputDialog, QMenu, QSystemTrayIcon,
        QTextEdit, QSplitter, QGraphicsView, QGraphicsScene, QGraphicsPixmapItem,
        QGraphicsRectItem
    )
    from PySide6.QtGui import (
        QIcon, QDesktopServices, QAction, QFont, QCloseEvent, QPalette, QColor,
        QIntValidator, QPixmap, QPen, QBrush, QWheelEvent, QTransform, QPainter
    )
    from PySide6.QtCore import Qt, QTimer, Signal, QSize, QUrl, QRectF, QProcess
    from PySide6.QtWidgets import QStyle, QSizePolicy, QToolBar  # Added QSizePolicy
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    import ocrmypdf
    from ocrmypdf import ocr, api as ocr_api, exceptions as ocr_exceptions

    # --- Patch para ocultar qualquer janela criada por ocrmypdf (Ghostscript/Tesseract) ---
    try:
        import ocrmypdf.subprocess as _ocrsp

        _orig_run = _ocrsp.run
        _orig_run_poll = _ocrsp.run_polling_stderr


        def _ocr_run_no_console(*args, **kwargs):
            if os.name == 'nt':
                kwargs['creationflags'] = (kwargs.get('creationflags', 0) | CREATE_NO_WINDOW)
                si = kwargs.get('startupinfo') or subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                try:
                    si.wShowWindow = subprocess.SW_HIDE
                except Exception:
                    pass
                kwargs['startupinfo'] = si
            return _orig_run(*args, **kwargs)


        def _ocr_run_poll_no_console(*args, **kwargs):
            if os.name == 'nt':
                kwargs['creationflags'] = (kwargs.get('creationflags', 0) | CREATE_NO_WINDOW)
                si = kwargs.get('startupinfo') or subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                try:
                    si.wShowWindow = subprocess.SW_HIDE
                except Exception:
                    pass
                kwargs['startupinfo'] = si
            return _orig_run_poll(*args, **kwargs)


        _ocrsp.run = _ocr_run_no_console
        _ocrsp.run_polling_stderr = _ocr_run_poll_no_console
    except Exception as _e:
        logging.debug(f"Falha ao aplicar patch no módulo ocrmypdf.subprocess: {_e}")

    from dataclasses import dataclass
    import inspect
    import cv2  # Import OpenCV
    import numpy as np  # Import NumPy for image conversion
    import base64
    import hmac
    from datetime import date, timedelta, datetime, timezone
    from dateutil.relativedelta import relativedelta

except ImportError as e:
    print(
        f"ERRO CRÍTICO: Falha ao importar componentes essenciais. Detalhes: {e}")
    print(
        "Por favor, verifique a instalação e tente novamente.")
    sys.exit(1)

# Importações adicionais globais para compatibilidade multiplataforma
try:
    import winreg  # type: ignore[attr-defined]
except ImportError:
    winreg = None  # type: ignore[assignment]

try:
    from platformdirs import user_config_dir, user_log_dir
except ImportError:
    # Se platformdirs não estiver instalado, usamos um fallback simples para não quebrar
    def user_config_dir(appname: str) -> str:
        return str(Path.home() / ".config" / appname)

    def user_log_dir(appname: str) -> str:
        return str(Path.home() / ".config" / appname)


if os.name == 'nt':
    CREATE_NO_WINDOW = 0x08000000
    _OldPopen = subprocess.Popen


    def _popen_no_console(*args, **kwargs):
        # Garante que NENHUMA janela de console seja aberta por processos filhos (ex.: Ghostscript)
        # Aplica CREATE_NO_WINDOW sem remover flags previamente definidas
        kwargs['creationflags'] = (kwargs.get('creationflags', 0) | CREATE_NO_WINDOW)
        # Configura STARTUPINFO com SW_HIDE para ocultar qualquer janela
        si = kwargs.get('startupinfo') or subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        try:
            si.wShowWindow = subprocess.SW_HIDE
        except Exception:
            pass
        kwargs['startupinfo'] = si
        return _OldPopen(*args, **kwargs)


    subprocess.Popen = _popen_no_console
    # Também envolve subprocess.run para capturar chamadas indiretas
    _OldRun = subprocess.run


    def _run_no_console(*args, **kwargs):
        kwargs['creationflags'] = (kwargs.get('creationflags', 0) | CREATE_NO_WINDOW)
        si = kwargs.get('startupinfo') or subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        try:
            si.wShowWindow = subprocess.SW_HIDE
        except Exception:
            pass
        kwargs['startupinfo'] = si
        return _OldRun(*args, **kwargs)


    subprocess.run = _run_no_console

# --- Configurações Globais ---
APP_NAME = "OCTec"
APP_VERSION = "1.0"  # Versão com OCR Zonal e Zoom
ENGINE_CHECK_INTERVAL = 1000
FILE_WATCHER_INTERVAL = 2

# --- NOVAS CONSTANTES PARA LICENCIAMENTO ---
SECRET_KEY = base64.b64decode("ovN9vSSD8MgEptVg6rvR62nrtT/aQPrvo03LBRFvyi8=")

# --- Novos Constantes para Limites ---
MAX_FILES_IN_GUI = 20
MAX_LOGS_PER_FILE = 50

# Constantes de layout (72 pt = 1 polegada)
BASE_FIELD_WIDTH_PTS = 92
BASE_FIELD_HEIGHT_PTS = 46
MARGIN_RIGHT_PTS = 26
MARGIN_BOTTOM_PTS = 26
OCRMYPDF_SCALING_FACTOR = 4.17

try:
    if getattr(sys, 'frozen', False):
        # Executando empacotado pelo PyInstaller
        CURRENT_SCRIPT_PATH = Path(sys.executable).parent
        BASE_DIR = CURRENT_SCRIPT_PATH
        if (BASE_DIR / "_internal" / "resources").exists():
            RESOURCES_DIR = BASE_DIR / "_internal" / "resources"
        elif (BASE_DIR / "resources").exists():
            RESOURCES_DIR = BASE_DIR / "resources"
        else:
            RESOURCES_DIR = BASE_DIR
            logging.warning(f"AVISO: Pasta resources não encontrada! Usando {RESOURCES_DIR}")
    else:
        # Executando no IDE (não empacotado)
        CURRENT_SCRIPT_PATH = Path(__file__).resolve().parent
        BASE_DIR = CURRENT_SCRIPT_PATH.parent
        if (BASE_DIR / "resources").exists():
            RESOURCES_DIR = BASE_DIR / "resources"
        else:
            RESOURCES_DIR = CURRENT_SCRIPT_PATH / "resources"
            logging.warning(f"AVISO: Pasta resources não encontrada na raiz, usando {RESOURCES_DIR}")
except Exception as e:
    BASE_DIR = Path.cwd()
    RESOURCES_DIR = BASE_DIR / "resources"
    logging.warning(
        f"AVISO: Não foi possível determinar o BASE_DIR de forma precisa. Usando diretório atual: {BASE_DIR}. Detalhe: {e}")

BASE_CONFIG_ROOT = Path("C:\\")
CONFIG_DIR = BASE_CONFIG_ROOT / APP_NAME
RULES_INPUT_BASE_DIR = CONFIG_DIR / "REGRAS"
CONFIG_FILE = CONFIG_DIR / "config.json"
CREDENTIALS_FILE = CONFIG_DIR / "credentials.json"
LOG_FILE = CONFIG_DIR / "octec.log"
LICENSE_FILE = CONFIG_DIR / "license.txt"

try:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    RULES_INPUT_BASE_DIR.mkdir(parents=True, exist_ok=True)
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
except OSError as e:
    logging.critical(
        f"ERRO CRÍTICO: Não foi possível criar diretórios de configuração em {CONFIG_DIR}. Verifique as permissões. Detalhes: {e}")
    sys.exit(2)


class SafeRotatingFileHandler(handlers.RotatingFileHandler):
    """Handler que evita cascata de erros quando arquivo de log está bloqueado."""
    def doRollover(self):
        try:
            super().doRollover()
        except (PermissionError, OSError):
            pass  # Silenciosamente ignora erro de rotação


def setup_logging():
    log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(threadName)s] - %(funcName)s - %(message)s')
    file_handler = SafeRotatingFileHandler(
        LOG_FILE, maxBytes=10 * 1024 * 1024, backupCount=5, encoding='utf-8'
    )
    file_handler.setLevel(logging.ERROR)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_formatter)
    console_handler.setLevel(logging.INFO)

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)

    logging.getLogger('PySide6').setLevel(logging.WARNING)
    logging.info(f'Logging inicializado. Arquivo de log: {LOG_FILE}')


def get_executable_for_autostart() -> Path:
    """
    Retorna o caminho que deve ser usado no Exec do .desktop.
    - Em binário PyInstaller: sys.executable
    - Em modo desenvolvimento: python + caminho do script
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    return Path(__file__).resolve()


def ensure_autostart_linux():
    """
    Garante que exista um arquivo ~/.config/autostart/octec.desktop
    para iniciar o OCTec em modo tray no login do usuário (Linux).
    Não faz nada em outros sistemas.
    """
    if not sys.platform.startswith("linux"):
        return

    autostart_dir = Path.home() / ".config" / "autostart"
    try:
        autostart_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        logging.error(f"Não foi possível criar diretório de autostart {autostart_dir}: {e}")
        return

    desktop_file = autostart_dir / "octec.desktop"
    if desktop_file.exists():
        logging.info(f"Autostart já configurado em {desktop_file}")
        return

    exe_path = get_executable_for_autostart()
    exec_cmd = f'"{exe_path}" --tray'

    icon_path = RESOURCES_DIR / "app_icon.png"
    if not icon_path.exists():
        possible_ico = RESOURCES_DIR / "app_icon.ico"
        icon_path = possible_ico if possible_ico.exists() else icon_path

    desktop_content = f"""[Desktop Entry]
Type=Application
Version=1.0
Name=OCTec
Comment=Monitor de OCR e processamento de documentos
Exec={exec_cmd}
Icon={icon_path}
Terminal=false
X-GNOME-Autostart-enabled=true
"""

    try:
        desktop_file.write_text(desktop_content, encoding="utf-8")
        logging.info(f"Autostart criado em {desktop_file} com Exec={exec_cmd}")
    except Exception as e:
        logging.error(f"Falha ao gravar arquivo de autostart {desktop_file}: {e}")




setup_logging()

# --- Verificar componentes de processamento ---
try:
    logging.info(
        f"Componente de processamento de documentos importado com sucesso (Versão: {getattr(ocrmypdf, '__version__', 'N/A')}).")
except ImportError:
    logging.critical(
        "ERRO CRÍTICO: O componente de processamento de documentos não foi encontrado. Funcionalidades avançadas de PDF estarão desabilitadas.")
    logging.critical("Por favor, verifique a instalação.")
    sys.exit(1)

# --- Variáveis de Estado do Motor Global ---


# --- Otimização de Concorrência e OCR / Deduplicação ---
try:
    _cpu_count = os.cpu_count() or 1
except Exception:
    _cpu_count = 1
# Aumenta o número máximo de processamentos simultâneos de OCR para
# aproveitar todos os núcleos disponíveis. O valor anterior limitava
# drasticamente a concorrência (por exemplo, em máquinas com 8 cores
# apenas 2 jobs podiam executar ao mesmo tempo). Ao usar o total de
# _cpu_count, permitimos que vários arquivos sejam processados em
# paralelo, cada um utilizando apenas uma fração da CPU quando ocrmypdf
# está configurado com jobs=1.
OCR_MAX_CONCURRENT = max(1, _cpu_count)
OCR_SEMAPHORE = threading.Semaphore(OCR_MAX_CONCURRENT)
_OCR_ACTIVE = 0
_OCR_ACTIVE_LOCK = threading.Lock()

# Fila de prioridade e controle de duplicação/estado
_file_queue_counter = itertools.count()
queued_files_lock = threading.Lock()
queued_files_set = set()  # {(rule_id, norm_path)}
processing_files_set = set()  # {(rule_id, norm_path)}

# Executor global para processamento em paralelo (ProcessPool). Será inicializado
# durante initialize_engine(). Cada tarefa de processamento de arquivo será
# enviada para este executor, permitindo que várias conversões sejam
# processadas ao mesmo tempo em processos separados.
process_executor = None


# --- Função utilitária para escanear pastas de entrada e enfileirar arquivos existentes ---
def scan_and_enqueue_existing_files() -> None:
    """
    Procura arquivos nos diretórios de entrada de cada regra e os coloca na fila de
    processamento. Isso é útil ao iniciar ou retomar o motor para garantir que
    arquivos já presentes sejam processados.
    """
    accepted_exts = {".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"}
    for rid, cfg in list(rules_data.items()):
        # Ignora regras pausadas
        if cfg.get("paused", False):
            continue
        input_dir = Path(cfg.get("input", ""))
        if not input_dir.exists():
            continue
        try:
            for file in input_dir.iterdir():
                if file.is_file() and file.suffix.lower() in accepted_exts:
                    enqueue_file(rid, file)
        except Exception as e_scan:
            logging.error(f"Falha ao escanear diretório de entrada {input_dir} para regra '{rid}': {e_scan}")


def _norm_path_key(path_str: str) -> str:
    # Robust against None/sentinels and weird inputs
    if not path_str:
        return ""
    try:
        return str(Path(path_str).resolve()).lower()
    except Exception:
        try:
            return str(path_str).replace("\\", "/").lower()
        except Exception:
            return ""


# --- Função para processamento em subprocesso ---
# Esta função será executada dentro de um processo separado via ProcessPoolExecutor.
# Recebe o id da regra, o caminho do arquivo como string e um dicionário contendo
# as opções da regra (rule_options). Devolve um tuplo (sucesso, caminho_saida)
# para ser tratado na thread principal. Exceptions são capturadas e retornadas
# como falha.

def process_file_task(params: tuple[str, str, dict]) -> tuple[bool, str | None]:
    rule_id, file_path_str, rule_options = params
    path = Path(file_path_str)

    try:
        # Decide qual função usar com base na extensão
        if path.suffix.lower() == ".pdf":
            success, out_path = process_pdf_file(rule_id, path, rule_options)
        elif path.suffix.lower() in [".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"]:
            success, out_path = process_image_file(rule_id, path, rule_options)
        else:
            # Tipo de arquivo não suportado nesta função; retorna falha
            return False, None
        # Converte Path para string para facilitar comunicação entre processos
        return success, str(out_path) if out_path else None
    except Exception:
        # Qualquer erro inesperado é tratado como falha
        return False, None


# Diretório temporário local rápido (definição explícita, sem "unresolved reference")
try:
    CONFIG_DIR
except NameError:
    # fallback padrão se CONFIG_DIR não existir no projeto
    CONFIG_DIR = Path("C:/OCTec")
FAST_TMP_DIR = (Path(CONFIG_DIR) / "tmp")
try:
    FAST_TMP_DIR.mkdir(parents=True, exist_ok=True)
except Exception as _e:
    logging.warning(f"Falha ao criar FAST_TMP_DIR {FAST_TMP_DIR}: {_e}")

engine_running = threading.Event()
rules_data = {}
file_queue = queue.PriorityQueue()
files_status = {}
files_status_lock = threading.Lock()
worker_threads = []
monitor_threads = []
rule_event_handlers = {}


# --- Caminhos de Ferramentas Externas ---

def compute_adaptive_dpi(page_rect, target_megapixels: float = 8.0, min_dpi: int = 150, max_dpi: int = 300) -> int:
    """Calcula um DPI adaptativo visando ~target_megapixels por página."""
    try:
        inch_w = max(1e-3, page_rect.width / 72.0)
        inch_h = max(1e-3, page_rect.height / 72.0)
        dpi = int(((target_megapixels * 1_000_000) / (inch_w * inch_h)) ** 0.5)
        return max(min_dpi, min(max_dpi, dpi))
    except Exception:
        return 200


def compute_priority_for_file(path: Path) -> int:
    """Menor tamanho = maior prioridade (0 primeiro)."""
    try:
        size = path.stat().st_size
        size = max(0, min(size, 1_000_000_000))
        return int(size)
    except Exception:
        return 0


def enqueue_file(rule_id: str, file_path: Path) -> None:
    """Evita duplicatas e enfileira com prioridade."""
    key = (rule_id, _norm_path_key(str(file_path)))
    with queued_files_lock:
        if key in queued_files_set or key in processing_files_set:
            return
        queued_files_set.add(key)
    prio = compute_priority_for_file(file_path)
    cnt = next(_file_queue_counter)
    file_queue.put((prio, cnt, rule_id, str(file_path)))
    update_file_status(rule_id, file_path.name, "Queued")

def dequeue_file(timeout: float = 1.0):
    prio, cnt, rule_id, file_path_str = file_queue.get(timeout=timeout)
    # Sentinel item to stop workers gracefully: (.., None, None)
    if rule_id is None and file_path_str is None:
        return None, None
    key = (rule_id, _norm_path_key(file_path_str))
    with queued_files_lock:
        queued_files_set.discard(key)
        processing_files_set.add(key)
    return rule_id, file_path_str

    key = (rule_id, _norm_path_key(file_path_str))
    with queued_files_lock:
        queued_files_set.discard(key)
        processing_files_set.add(key)
    return rule_id, file_path_str


def _clear_processing_state(rule_id: str, path: Path) -> None:
    """
    Remove any tracking state for the given file.  When a file finishes
    processing (successfully or not), we need to ensure it can be
    re‑enqueued if it shows up again.  We therefore discard its key
    from both the processing and queued sets.  This prevents stale
    entries from blocking future processing of the same path.
    """
    key = (rule_id, _norm_path_key(str(path)))
    with queued_files_lock:
        processing_files_set.discard(key)
        # Also remove from the queued set in case a stale entry lingered
        queued_files_set.discard(key)


if os.name == 'nt':
    EXTERNAL_TOOLS = {
        "TESSERACT_PATH": str(RESOURCES_DIR / "Tesseract-OCR" / "tesseract.exe"),
        "POPPLER_PATH": str(RESOURCES_DIR / "poppler" / "bin"),
        "GHOSTSCRIPT_PATH": str(RESOURCES_DIR / "gs10.04.0" / "bin" / "gswin64c.exe"),
    }
else:
    # Em Linux (e outros), assumimos que as ferramentas estão instaladas no sistema
    EXTERNAL_TOOLS = {
        "TESSERACT_PATH": "tesseract",
        "POPPLER_PATH": "",  # usaremos pdftoppm/pdfinfo do PATH, quando necessário
        "GHOSTSCRIPT_PATH": "gs",
    }


def add_external_tools_to_path():
    """
    No Windows, adiciona os diretórios internos ao PATH.
    No Linux, apenas verifica se as ferramentas básicas existem no PATH
    e registra avisos se estiverem ausentes.
    """
    if os.name == 'nt':
        tesseract_dir = str(Path(EXTERNAL_TOOLS["TESSERACT_PATH"]).parent)
        if tesseract_dir and Path(tesseract_dir).exists() and tesseract_dir not in os.environ.get("PATH", ""):
            os.environ["PATH"] = tesseract_dir + os.pathsep + os.environ["PATH"]
            logging.info(f"Adicionado diretório de ferramenta de reconhecimento de texto ao PATH: {tesseract_dir}")

        poppler_bin_dir = EXTERNAL_TOOLS["POPPLER_PATH"]
        if poppler_bin_dir and Path(poppler_bin_dir).exists() and poppler_bin_dir not in os.environ.get("PATH", ""):
            os.environ["PATH"] = poppler_bin_dir + os.pathsep + os.environ["PATH"]
            logging.info(f"Adicionado diretório de ferramenta de manipulação de PDF ao PATH: {poppler_bin_dir}")

        gs_exe_path = Path(EXTERNAL_TOOLS["GHOSTSCRIPT_PATH"])
        if gs_exe_path.exists():
            gs_dir = str(gs_exe_path.parent)
            if gs_dir not in os.environ.get("PATH", ""):
                os.environ["PATH"] = gs_dir + os.pathsep + os.environ["PATH"]
                logging.info(f"Ferramenta de renderização de PDF encontrada: {gs_exe_path}")
            else:
                logging.info(f"Ferramenta de renderização de PDF encontrada: {gs_exe_path}")
        else:
            logging.warning(
                f"Ferramenta de renderização de PDF não encontrada em {gs_exe_path}. Funcionalidades de PDF podem ser limitadas."
            )
    else:
        import shutil as _shutil
        for tool in ("tesseract", "gs", "pdftoppm"):
            if _shutil.which(tool) is None:
                logging.warning(f"Ferramenta externa '{tool}' não encontrada no PATH. Algumas funções podem não funcionar.")
            else:
                logging.info(f"Ferramenta externa '{tool}' encontrada no PATH.")


# Define explicitamente o executável do Ghostscript para o OCRmyPDF, garantindo uso do gswin64c.exe
try:
    if os.name == 'nt':
        gs_exe = EXTERNAL_TOOLS.get('GHOSTSCRIPT_PATH')
        if gs_exe and os.path.exists(gs_exe):
            os.environ['OCRMYPDF_GS'] = gs_exe
except Exception:
    pass

add_external_tools_to_path()


# --- Funções Utilitárias ---
def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            logging.error(f"Erro ao carregar {CONFIG_FILE}: {e}. Retornando configuração vazia.")
    return {}


def save_config(config_data: dict) -> None:
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4)
    except Exception as e:
        logging.error(f"Erro ao salvar {CONFIG_FILE}: {e}")


_net_connections_status = {}


def get_net_credentials() -> dict:
    if CREDENTIALS_FILE.exists():
        try:
            with open(CREDENTIALS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            logging.error(f"Erro ao carregar {CREDENTIALS_FILE}: {e}. Retornando credenciais vazias.")
    return {}


def save_net_credentials(creds_data: dict) -> None:
    try:
        with open(CREDENTIALS_FILE, 'w', encoding='utf-8') as f:
            json.dump(creds_data, f, indent=4)
    except IOError as e:
        logging.error(f"Erro ao salvar {CREDENTIALS_FILE}: {e}")


def add_or_update_net_credential(server_path: str, username: str, password: str) -> None:
    creds = get_net_credentials()
    normalized_server_path_key = Path(server_path).as_posix().lower()
    creds[normalized_server_path_key] = {"username": username, "password": password}
    save_net_credentials(creds)


def remove_net_credential(server_path: str) -> None:
    creds = get_net_credentials()
    normalized_server_path_key = Path(server_path).as_posix().lower()
    if normalized_server_path_key in creds:
        del creds[normalized_server_path_key]
        save_net_credentials(creds)
        disconnect_network_share(server_path)


def find_credential_for_unc(unc_path: str) -> dict | None:
    creds = get_net_credentials()
    norm_unc_path_to_find = Path(unc_path).as_posix().lower()

    best_match = None
    longest_match_len = 0
    for stored_share_key, cred_info in creds.items():
        if norm_unc_path_to_find.startswith(stored_share_key):
            if len(stored_share_key) > longest_match_len:
                longest_match_len = len(stored_share_key)
                best_match = cred_info

    if best_match:
        logging.info(
            f"Credencial encontrada para '{unc_path}' através da correspondência com a chave '{list(creds.keys())[list(creds.values()).index(best_match)]}'.")
        return best_match

    logging.error(f"Nenhuma credencial encontrada para o caminho UNC '{unc_path}'.")
    return None


def disconnect_network_share(share_path_to_disconnect: str):
    if not str(share_path_to_disconnect).startswith("\\\\"):
        return

    logging.info(f"Tentando garantir a desconexão de '{share_path_to_disconnect}' para evitar conflitos...")
    try:
        cmd = ["net", "use", str(share_path_to_disconnect), "/delete", "/y"]
        creation_flags = 0
        if sys.platform == "win32" and getattr(sys, 'frozen', False):
            creation_flags = subprocess.CREATE_NO_WINDOW
        result = subprocess.run(cmd, capture_output=True, check=False, text=True, encoding='latin-1', timeout=15,
                                creationflags=creation_flags)
        if result.returncode == 0:
            logging.info(f"Desconexão de '{share_path_to_disconnect}' realizada com sucesso.")
        elif "não foi encontrada nenhuma conexão de rede" in result.stderr.lower() or "network connection could not be found" in result.stderr.lower():
            logging.info(f"Nenhuma conexão preexistente encontrada para '{share_path_to_disconnect}', o que é bom.")
        else:
            logging.warning(
                f"Comando para desconectar '{share_path_to_disconnect}' retornou uma mensagem inesperada. stdout: {result.stdout.strip()}, stderr: '{result.stderr.strip()}'")
        share_key = Path(share_path_to_disconnect).as_posix().lower()
        _net_connections_status.pop(share_key, None)
    except subprocess.TimeoutExpired:
        logging.error(f"Timeout ao tentar desconectar de '{share_path_to_disconnect}'.")
    except Exception as e:
        logging.error(f"Erro inesperado ao desconectar de '{share_path_to_disconnect}': {e}")


def ensure_persistent_connection(unc_path: str) -> bool:
    unc_path_obj = Path(unc_path)
    if not unc_path_obj.is_absolute() or not unc_path_obj.drive:
        return True

    share_root = unc_path_obj.drive
    if not share_root:
        logging.error(f"Não foi possível extrair a raiz do compartilhamento do caminho UNC '{unc_path}'.")
        return False

    share_root_str = str(Path(share_root))
    # Se já houver acesso ao UNC, valide e retorne
    try:
        if Path(unc_path).exists():
            return True
    except Exception:
        pass
    # (otimizado) não desconectar preventivamente

    cred_info = find_credential_for_unc(share_root_str)
    if not cred_info:
        logging.error(f"Nenhuma credencial encontrada para '{share_root_str}'. Não é possível conectar.")
        return False

    username = cred_info.get("username")
    password = cred_info.get("password")

    if not username or password is None:
        logging.error(f"Credenciais incompletas para '{share_root_str}' (usuário ou senha faltando).")
        return False

    logging.info(f"Tentando estabelecer nova conexão de rede com '{share_root_str}' usando o usuário '{username}'...")
    try:
        cmd = ["net", "use", share_root_str]
        if password:
            cmd.append(password)
        cmd.extend([f"/user:{username}", "/persistent:no"])

        creation_flags = 0
        if sys.platform == "win32" and getattr(sys, 'frozen', False):
            creation_flags = subprocess.CREATE_NO_WINDOW

        result = subprocess.run(cmd, capture_output=True, text=True, encoding='latin-1', timeout=30,
                                creationflags=creation_flags)

        if result.returncode == 0:
            logging.info(f"Conectado com sucesso a '{share_root_str}'.")
            if Path(unc_path).exists():
                logging.info(f"Caminho completo '{unc_path}' verificado e acessível.")
                _net_connections_status[Path(share_root_str).as_posix().lower()] = True
                return True
            else:
                logging.error(
                    f"Conectado a '{share_root_str}', mas o caminho completo '{unc_path}' não foi encontrado.")
                return False
        else:
            logging.error(
                f"Falha ao executar 'net use' para '{share_root_str}'. Código: {result.returncode}, stdout: '{result.stdout.strip()}', stderr: '{result.stderr.strip()}'")
            _net_connections_status[Path(share_root_str).as_posix().lower()] = False
            return False
    except subprocess.TimeoutExpired:
        logging.error(f"Timeout ao tentar conectar a '{share_root_str}'.")
        return False
    except Exception as e:
        logging.error(f"Erro inesperado na execução do 'net use' para '{share_root_str}': {e}")
        return False


# --- Auto-Resume de Regras pausadas por caminhos ausentes (input/output) ---
_auto_resume_thread = None
_auto_resume_stop_event = threading.Event()


def _path_accessible_for_resume(path_str: str) -> bool:
    """Verifica se um caminho está acessível.
    - Para caminhos UNC, tenta (re)conectar via ensure_persistent_connection().
    - Para locais, apenas verifica existência.
    Nunca cria pastas aqui para evitar efeitos colaterais indesejados.
    """
    if not path_str:
        return True
    try:
        p = Path(path_str)
        # UNC?
        if str(p).startswith('\\'):
            if not ensure_persistent_connection(path_str):
                return False
        # Por fim, checa existência
        return p.exists()
    except Exception:
        return False


def _auto_resume_worker(poll_interval: float = 15.0) -> None:
    thread_name = threading.current_thread().name
    logging.info(f"{thread_name}: Monitor de auto‑retomada iniciado.")
    while True:
        # Permite finalizar rapidamente
        if _auto_resume_stop_event.is_set():
            break
        any_resumed = False
        try:
            # Percorre cópia estável para evitar RuntimeError em alterações simultâneas
            for rid, cfg in list(rules_data.items()):
                try:
                    # Apenas regras pausadas especificamente por caminhos ausentes
                    if not cfg.get('paused', False) or not cfg.get('missing_paths', False):
                        continue
                    inp = cfg.get('input', '') or ''
                    outp = cfg.get('output', '') or ''
                    inp_ok = _path_accessible_for_resume(inp)
                    out_ok = _path_accessible_for_resume(outp)
                    if inp_ok and out_ok:
                        # Retoma automaticamente
                        logging.info(
                            f"Regra '{rid}' detectada com pastas acessíveis novamente. Retomando automaticamente.")
                        cfg['paused'] = False
                        # Remove o marcador de falta de caminhos para não ficar re‑checando agressivamente
                        cfg.pop('missing_paths', None)
                        rules_data[rid] = cfg
                        save_config(rules_data)
                        try:
                            start_watching_thread_for_rule(rid)
                        except Exception as e_sw:
                            logging.error(f"Falha ao reiniciar observador para regra '{rid}': {e_sw}")
                        any_resumed = True
                except Exception as e_rule:
                    logging.error(f"Erro no auto‑retomador ao analisar regra '{rid}': {e_rule}")
        except Exception as e_loop:
            logging.error(f"Erro no laço do monitor de auto‑retomada: {e_loop}")
        # Se alguma regra foi retomada, executa uma varredura para enfileirar arquivos já existentes
        if any_resumed:
            try:
                scan_and_enqueue_existing_files()
            except Exception as e_scan:
                logging.error(f"Falha ao escanear arquivos após auto‑retomada: {e_scan}")
        # Aguarda próximo ciclo
        _auto_resume_stop_event.wait(poll_interval)
    logging.info(f"{thread_name}: Monitor de auto‑retomada finalizado.")


def cleanup_all_managed_persistent_connections():
    logging.info("Limpando todas as conexões de rede gerenciadas...")
    active_shares = list(_net_connections_status.keys())
    for share_path in active_shares:
        if _net_connections_status.get(share_path):
            disconnect_network_share(share_path)
    _net_connections_status.clear()
    logging.info("Limpeza de conexões de rede concluída.")


def debug_log(rule_id: str, file_name: str, message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}"
    with files_status_lock:
        log_list = files_status.setdefault(rule_id, {}).setdefault(file_name, {}).setdefault("logs", [])
        log_list.append(log_entry)
        if len(log_list) > MAX_LOGS_PER_FILE:
            files_status[rule_id][file_name]["logs"] = log_list[-MAX_LOGS_PER_FILE:]
    logging.debug(f"[{rule_id}][{file_name}] {message}")


def update_file_status(rule_id: str, file_name: str, status: str, progress: int = None, size: int = None) -> None:
    with files_status_lock:
        if rule_id not in files_status: files_status[rule_id] = {}
        if file_name not in files_status[rule_id]:
            files_status[rule_id][file_name] = {
                "progress": 0,
                "status": "Queued",
                "logs": [],
                "size": 0,
                "entry_time": time.time()
            }
        files_status[rule_id][file_name]["status"] = status
        if progress is not None: files_status[rule_id][file_name]["progress"] = progress
        if size is not None: files_status[rule_id][file_name]["size"] = size
    logging.info(
        f"[{rule_id}][{file_name}] Status: {status}, Progresso: {progress if progress is not None else files_status.get(rule_id, {}).get(file_name, {}).get('progress', 'N/A')}%")


# --- Funções de Licenciamento ---
def base36encode(n: int) -> str:
    """Converte um número inteiro para uma string Base36."""
    chars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    s = ''
    while n > 0:
        n, i = divmod(n, 36)
        s = chars[i] + s
    return s or '0'


def base36decode(s: str) -> int:
    """Converte string Base36 em inteiro."""
    return int(s, 36)


def calc_hmac(payload: str) -> str:
    """Trunca HMAC‑SHA256(payload) a 5 bytes e retorna Base36."""
    digest = hmac.new(SECRET_KEY, payload.encode(), hashlib.sha256).digest()[:5]
    return base36encode(int.from_bytes(digest, 'big')).rjust(5, '0')


def _read_license_from_file() -> str | None:
    """Lê a chave de licença de um arquivo de texto (qualquer SO)."""
    try:
        return LICENSE_FILE.read_text(encoding="utf-8").strip()
    except FileNotFoundError:
        return None
    except Exception as e:
        logging.error(f"Erro ao ler arquivo de licença {LICENSE_FILE}: {e}")
        return None


def _write_license_to_file(key: str) -> None:
    """Grava a chave de licença em um arquivo de texto."""
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        LICENSE_FILE.write_text(key.strip(), encoding="utf-8")
    except Exception as e:
        logging.error(f"Erro ao gravar arquivo de licença {LICENSE_FILE}: {e}")


def _delete_license_file() -> None:
    try:
        LICENSE_FILE.unlink(missing_ok=True)
    except Exception as e:
        logging.error(f"Erro ao remover arquivo de licença {LICENSE_FILE}: {e}")


def _read_license_from_registry() -> str | None:
    """Lê a chave de licença do Registro do Windows, se disponível."""
    if winreg is None:
        return None
    try:
        acc64 = getattr(winreg, 'KEY_READ', 0) | getattr(winreg, 'KEY_WOW64_64KEY', 0)
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\OCTec", 0, acc64) as h:
            k, _ = winreg.QueryValueEx(h, "LicenseKey")
            return k
    except OSError:
        try:
            acc32 = getattr(winreg, 'KEY_READ', 0) | getattr(winreg, 'KEY_WOW64_32KEY', 0)
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\WOW6432Node\OCTec", 0, acc32) as h:
                k, _ = winreg.QueryValueEx(h, "LicenseKey")
                logging.warning("LicenseKey lida de WOW6432Node (32-bit hive). Padronize para 64-bit.")
                return k
        except OSError:
            return None


def read_license_key():
    """Lê a chave de licença.
    - No Windows: tenta Registro; se falhar, tenta arquivo.
    - Em outros SO: lê diretamente do arquivo.
    """
    if sys.platform.startswith("win"):
        key = _read_license_from_registry()
        if key:
            return key
        return _read_license_from_file()
    else:
        return _read_license_from_file()


def save_license_key(key: str):
    """Salva a chave de licença.
    - No Windows: grava no Registro e no arquivo.
    - Em outros SO: grava apenas no arquivo.
    """
    _write_license_to_file(key)
    if sys.platform.startswith("win") and winreg is not None:
        try:
            reg = winreg.CreateKeyEx(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\OCTec",
                0,
                getattr(winreg, 'KEY_WRITE', 0) | getattr(winreg, 'KEY_WOW64_64KEY', 0),
            )
            winreg.SetValueEx(reg, "LicenseKey", 0, winreg.REG_SZ, key)
            winreg.CloseKey(reg)
        except Exception as e:
            logging.error(f"Erro ao salvar licença no Registro: {e}")
            raise  # Re-raise para que o chamador possa lidar com isso


def delete_license_key_from_registry():
    """Remove a licença do Registro do Windows e do arquivo local (se existir)."""
    _delete_license_file()
    if winreg is None:
        return
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\OCTec",
            0,
            getattr(winreg, 'KEY_SET_VALUE', 0) | getattr(winreg, 'KEY_WOW64_64KEY', 0),
        )
        winreg.DeleteValue(key, "LicenseKey")
        winreg.CloseKey(key)
        logging.info("License key successfully deleted from registry.")
    except FileNotFoundError:
        logging.info("License key not found in registry, no deletion needed.")
    except Exception as e:
        logging.error(f"Error deleting license key from registry: {e}")


def is_license_valid(key: str) -> bool:
    """
    Valida diferentes tipos de chaves de licença.
    TEST:  TEST-<ts8>-<hmac5>          → ts em segundos
    TRIAL: TRIAL-<days5>-<rand5>-<hmac5> → days desde 1970-01-01
    PROD:  PROD-<days5>-<rand5>-<hmac5>  → days desde 1970-01-01
    """
    if not key or '-' not in key:
        logging.warning(f"Chave inválida (formato): '{key}'")
        return False

    parts = key.split('-')
    prefix = parts[0].upper()

    # TEST (5 Minutos)
    if prefix == "TEST" and len(parts) == 3:
        data_code, hmac_code = parts[1], parts[2]
        try:
            exp_ts = base36decode(data_code)
        except ValueError:
            logging.warning(f"TEST inválida (data_code): '{data_code}'")
            return False
        if calc_hmac(data_code) != hmac_code:
            logging.warning("TEST inválida (HMAC)")
            return False
        if int(datetime.now(timezone.utc).timestamp()) > exp_ts:
            logging.warning("TEST expirada")
            return False
        return True

    # PROD (1-10 Anos)
    if prefix == "PROD" and len(parts) == 4:
        data_code, rand_code, hmac_code = parts[1], parts[2], parts[3]
        payload = f"{data_code}-{rand_code}"
        if calc_hmac(payload) != hmac_code:
            logging.warning("PROD inválida (HMAC)")
            return False
        try:
            days = base36decode(data_code)
        except ValueError:
            logging.warning(f"PROD inválida (data_code): '{data_code}'")
            return False
        exp_date = date(1970, 1, 1) + timedelta(days=days)
        if datetime.now(timezone.utc).date() > exp_date:
            logging.warning(f"PROD expirada em {exp_date.isoformat()}")
            return False
        return True

    # TRIAL (15 Dias—timestamp-based)
    if prefix == "TRIAL" and len(parts) == 3:
        ts_code, hmac_code = parts[1], parts[2]
        # Verifica HMAC
        if calc_hmac(ts_code) != hmac_code:
            logging.warning("TRIAL inválida (HMAC)")
            return False
        # Decodifica timestamp em segundos UTC
        try:
            exp_ts = base36decode(ts_code)
        except ValueError:
            logging.warning(f"TRIAL inválida (ts_code): '{ts_code}'")
            return False
        # Expira após exatamente 15×24h
        if int(datetime.now(timezone.utc).timestamp()) > exp_ts:
            logging.warning("TRIAL expirada")
            return False
        return True

    logging.warning(f"Prefixo desconhecido: '{prefix}'")
    return False


# --- recuperação de info para exibir na GUI ---
def get_license_info(key: str):
    """
    Retorna (dias_restantes: int, data_expiracao: date, tipo: str)
    ou None se inválido.
    """
    if not key or '-' not in key:
        return None

    parts = key.split('-')
    prefix = parts[0].upper()

    # TEST
    if prefix == "TEST" and len(parts) == 3:
        try:
            exp_ts = base36decode(parts[1])
            exp_dt_utc = datetime.fromtimestamp(exp_ts, timezone.utc)  # Full datetime object for precise comparison
            exp_date = exp_dt_utc.date()  # Date part for "dias" calculation
        except Exception:
            return None

        dias = (exp_date - datetime.now(timezone.utc).date()).days
        tipo = "Teste"

        if dias < 0:
            dias = 0
            tipo += " (Expirada)"
        elif dias == 0:  # License expires today
            current_dt_utc = datetime.now(timezone.utc)
            if current_dt_utc > exp_dt_utc:
                tipo += " (Expirada)"
            else:
                remaining_seconds = (exp_dt_utc - current_dt_utc).total_seconds()
                if remaining_seconds > 0:
                    remaining_minutes = int(remaining_seconds / 60)
                    remaining_hours = int(remaining_minutes / 60)
                    if remaining_hours > 0:
                        tipo += f" (Expira em {remaining_hours}h {remaining_minutes % 60}m)"
                    elif remaining_minutes > 0:
                        tipo += f" (Expira em {remaining_minutes}m)"
                    else:
                        tipo += f" (Expira em {int(remaining_seconds)}s)"
                else:
                    tipo += " (Expirada)"  # Should not happen if current_dt_utc <= exp_dt_utc, but for safety

        return dias, exp_dt_utc.date(), tipo  # Return exp_dt_utc.date() for consistency

    # PROD
    if prefix == "PROD" and len(parts) == 4:
        try:
            days = base36decode(parts[1])
            exp_date = date(1970, 1, 1) + timedelta(days=days)
        except Exception:
            return None
        dias = (exp_date - datetime.now(timezone.utc).date()).days
        tipo = "Produção"
        if dias < 0:
            dias = 0
            tipo += " (Expirada)"
        return dias, exp_date, tipo

    # TRIAL (15 Dias—timestamp-based)
    if prefix == "TRIAL" and len(parts) == 3:
        try:
            exp_ts = base36decode(parts[1])
            exp_dt = datetime.fromtimestamp(exp_ts, timezone.utc)
            exp_date = exp_dt.date()
        except Exception:
            return None
        dias = (exp_date - datetime.now(timezone.utc).date()).days
        tipo = "Trial"
        if dias < 0:
            dias = 0
            tipo += " (Expirada)"
        return dias, exp_date, tipo

    return None


# --- Funções de Processamento de Arquivo ---
def convert_pdf_to_text(input_pdf_path: Path, output_txt_path: Path, rule_options: dict) -> bool:
    rule_id = rule_options.get("rule_id", "Desconhecido")
    base_name = input_pdf_path.name
    debug_log(rule_id, base_name, f"Convertendo documento para Texto: {input_pdf_path}")
    try:
        doc = fitz.open(input_pdf_path)
        text_content = ""
        for page_num in range(doc.page_count):
            page = doc[page_num]
            text_content += page.get_text() + "\n"
        doc.close()
        with open(output_txt_path, "w", encoding="utf-8") as f:
            f.write(text_content)
        debug_log(rule_id, base_name, f"Documento convertido para TXT com sucesso: {output_txt_path}")
        return True
    except Exception as e:
        debug_log(rule_id, base_name, f"Erro ao converter documento para TXT: {e}")
        logging.error(f"Erro ao converter documento para TXT '{input_pdf_path}': {e}")
        return False


def _create_docx_from_text_pages(pages_text, output_docx_path: Path) -> bool:
    """
    Cria um DOCX simples (apenas texto) sem dependências externas.
    pages_text: lista de strings (um item por página). Cada string pode conter múltiplas linhas.
    """
    try:
        def _xml_escape(t: str) -> str:
            return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

        # Constrói o corpo do documento com quebras de linha como novos parágrafos e quebra de página entre páginas.
        body_parts = []
        for p_idx, page in enumerate(pages_text):
            lines = page.replace("\r\n", "\n").replace("\r", "\n").split("\n")
            for line in lines:
                # usa xml:space="preserve" para manter espaços
                body_parts.append(f'<w:p><w:r><w:t xml:space="preserve">{_xml_escape(line)}</w:t></w:r></w:p>')
            if p_idx < len(pages_text) - 1:
                body_parts.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        document_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:body>' + "".join(body_parts) + '<w:sectPr/></w:body></w:document>'
        )
        content_types_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '</Types>'
        )
        rels_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>'
        )
        with zipfile.ZipFile(str(output_docx_path), "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", content_types_xml)
            z.writestr("_rels/.rels", rels_xml)
            z.writestr("word/document.xml", document_xml)
        return True
    except Exception as e:
        logging.error(f"Falha ao criar DOCX '{output_docx_path}': {e}")
        return False


# ==== Export Helpers: XLSX / PPTX (moved near DOCX for clarity) ====
def _create_xlsx_from_text_pages(pages_text, output_xlsx_path: Path) -> bool:
    try:
        def _xml_escape(t: str) -> str:
            return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        rows_xml = []
        for idx, page in enumerate(pages_text, start=1):
            text = _xml_escape(page.replace("\r\n", "\n").replace("\r", "\n"))
            a = f'<c r="A{idx}" t="n"><v>{idx}</v></c>'
            b = f'<c r="B{idx}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'
            rows_xml.append(f'<row r="{idx}">{a}{b}</row>')
        dim_ref = f'A1:B{len(pages_text) if pages_text else 1}'
        worksheet_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="{dim_ref}"/>'
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'<sheetData>{"".join(rows_xml) if rows_xml else ""}</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            '</worksheet>'
        )
        workbook_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Pages" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>'
        )
        workbook_rels = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/>'
            '</Relationships>'
        )
        content_types = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '</Types>'
        )
        root_rels = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>'
            '</Relationships>'
        )
        with zipfile.ZipFile(str(output_xlsx_path), "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", content_types)
            z.writestr("_rels/.rels", root_rels)
            z.writestr("xl/workbook.xml", workbook_xml)
            z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
            z.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
        return True
    except Exception as e:
        logging.error(f"Falha ao criar XLSX '{output_xlsx_path}': {e}")
        return False


def _create_pptx_from_text_pages(pages_text, output_pptx_path: Path) -> bool:
    try:
        def _xml_escape(t: str) -> str:
            return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        slide_xmls = []
        for i, page in enumerate(pages_text, start=1):
            lines = page.replace("\r\n", "\n").replace("\r", "\n").split("\n")
            paras = [
                        f'<a:p><a:r><a:t xml:space="preserve">{_xml_escape(ln)}</a:t></a:r><a:endParaRPr lang="pt-BR"/></a:p>'
                        for ln in lines] or ['<a:p/>']
            slide_xmls.append(
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<p:cSld><p:spTree>'
                '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                '<p:grpSpPr><a:xfrm/></p:grpSpPr>'
                '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Texto"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>'
                '<p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="8229600" cy="5943600"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
                '<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>' + "".join(paras) + '</p:txBody>'
                                                                                      '</p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
            )
        sld_ids = []
        rels = []
        for i in range(len(slide_xmls)):
            rid = f"rId{i + 1}";
            sid = 256 + i
            sld_ids.append(f'<p:sldId id="{sid}" r:id="{rid}"/>')
            rels.append(
                f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i + 1}.xml"/>')
        presentation_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<p:sldIdLst>{"".join(sld_ids)}</p:sldIdLst>'
            '<p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>'
            '<p:notesSz cx="6858000" cy="9144000"/>'
            '</p:presentation>'
        )
        presentation_rels = (
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + "".join(rels) + '</Relationships>'
        )
        content_types = (
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
                + "".join(
            f'<Override PartName="/ppt/slides/slide{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
            for i in range(len(slide_xmls))) +
                '</Types>'
        )
        root_rels = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
            '</Relationships>'
        )
        with zipfile.ZipFile(str(output_pptx_path), "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", content_types)
            z.writestr("_rels/.rels", root_rels)
            z.writestr("ppt/presentation.xml", presentation_xml)
            z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels)
            for i, xml in enumerate(slide_xmls, start=1):
                z.writestr(f"ppt/slides/slide{i}.xml", xml)
        return True
    except Exception as e:
        logging.error(f"Falha ao criar PPTX '{output_pptx_path}': {e}")
        return False


def convert_pdf_to_rtf(input_pdf_path: Path, output_rtf_path: Path, rule_options: dict) -> bool:
    rule_id = rule_options.get("rule_id", "Desconhecido")
    base_name = input_pdf_path.name
    debug_log(rule_id, base_name, f"Convertendo documento para RTF: {input_pdf_path}")
    try:
        doc = fitz.open(input_pdf_path)
        text_content = ""
        for page_num in range(doc.page_count):
            page = doc[page_num]
            text_content += page.get_text() + "\n"
        doc.close()
        rtf_header = r"{\rtf1\ansi\deff0\nouicompat{\fonttbl{\f0\fnil\fcharset0 Calibri;}}" \
                     r"{\pard\sa200\sl276\slmult1\f0\fs22 "
        rtf_footer = r"}\par}"
        escaped_text = text_content.replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}").replace("\n",
                                                                                                          "\\par\n")
        with open(output_rtf_path, "w",
                  encoding="utf-8") as f:
            f.write(rtf_header + escaped_text + rtf_footer)
        debug_log(rule_id, base_name, f"Documento convertido para RTF com sucesso: {output_rtf_path}")
        return True
    except Exception as e:
        debug_log(rule_id, base_name, f"Erro ao converter documento para RTF: {e}")
        logging.error(f"Erro ao converter documento para RTF '{input_pdf_path}': {e}")
        return False


def convert_pdf_to_jpeg_pages(input_pdf_path: Path, output_dir_path: Path, base_output_name: str, rule_options: dict) -> \
        list[Path]:
    rule_id = rule_options.get("rule_id", "Desconhecido")
    input_base_name = input_pdf_path.name
    debug_log(rule_id, input_base_name, f"Convertendo documento para páginas JPEG: {input_pdf_path}")
    output_jpeg_paths = []
    try:
        doc = fitz.open(input_pdf_path)
        for page_num in range(doc.page_count):
            page = doc[page_num]
            pix = page.get_pixmap()
            jpeg_path = output_dir_path / f"{base_output_name}_page{page_num + 1}.jpeg"
            pix.save(str(jpeg_path))
            output_jpeg_paths.append(jpeg_path)
            debug_log(rule_id, input_base_name, f"Página {page_num + 1} salva como {jpeg_path.name}")
        doc.close()
        return output_jpeg_paths
    except Exception as e:
        debug_log(rule_id, input_base_name, f"Erro ao converter página de documento para JPEG: {e}")
        logging.error(f"Erro ao converter páginas de '{input_pdf_path}' para JPEG: {e}")
        return []


def convert_to_pdfa_ocrmypdf(
        input_path: Path,
        output_pdfa_path: Path,
        rule_options: dict
) -> bool:
    """
    Converte um arquivo PDF ou imagem para PDF pesquisável ou PDF/A,
    utilizando funcionalidades avançadas de rotação e alinhamento.
    """
    if rule_options.get("show_debug_messages", False):
        ocr_api.configure_logging(ocr_api.Verbosity.debug)
    else:
        ocr_api.configure_logging(ocr_api.Verbosity.default)

    rule_id = rule_options.get("rule_id", "Desconhecido")
    base_name = input_path.name

    debug_log(rule_id, base_name, f"Iniciando processamento de documento: {input_path} -> {output_pdfa_path}")
    try:
        ocr_kwargs = {
            "input_file": input_path,
            "output_file": output_pdfa_path,
            "jobs": os.cpu_count() or 1,
            "keep_temporary_files": False,
            "progress_bar": False,
        }

        if rule_options.get("orientation"):
            ocr_kwargs["rotate_pages"] = True
            ocr_kwargs["rotate_pages_threshold"] = 2.0
            ocr_kwargs["skip_text"] = True
            if not ocr_kwargs.get("redo_ocr", False):
                ocr_kwargs["deskew"] = True
            debug_log(rule_id, base_name, "Orientação e alinhamento ativados.")
        else:
            ocr_kwargs["force_ocr"] = True
            debug_log(rule_id, base_name,
                      "Forçando a recriação do reconhecimento de texto pois a orientação não está ativa.")

        lang = rule_options.get("language")
        if lang:
            ocr_kwargs["language"] = lang
            debug_log(rule_id, base_name, f"Idioma de reconhecimento de texto: {lang}")

        fmt = rule_options.get("output_format", "").lower()
        if fmt == "pdf/a":
            ocr_kwargs["output_type"] = "pdfa"
            debug_log(rule_id, base_name, "Formato de saída: PDF/A.")
        else:
            ocr_kwargs["output_type"] = "pdf"
            debug_log(rule_id, base_name, "Formato de saída: PDF Pesquisável.")

        dpi = rule_options.get("image_dpi_for_ocrmypdf")
        if dpi:
            ocr_kwargs["image_dpi"] = dpi
            debug_log(rule_id, base_name, f"DPI da imagem forçada para: {dpi}")

        ocr_kwargs["optimize"] = 0
        debug_log(rule_id, base_name, f"Nível de otimização: {ocr_kwargs['optimize']}")

        ocr_kwargs["tesseract_timeout"] = rule_options.get("ocr_timeout", 600)

        debug_log(rule_id, base_name,
                  f"Chamando função de processamento de documento com os seguintes argumentos: {ocr_kwargs}")
        # Permite múltiplos processamentos em paralelo sem reduzir
        # drasticamente o número de tarefas em execução. Cada job de ocrmypdf
        # usará apenas um núcleo (jobs=1), permitindo que vários arquivos
        # sejam processados simultaneamente. O semáforo limita a
        # concorrência ao número de núcleos disponíveis.
        OCR_SEMAPHORE.acquire()
        try:
            # Incrementa contador interno de OCR (mantido para compatibilidade)
            with _OCR_ACTIVE_LOCK:
                global _OCR_ACTIVE
                _OCR_ACTIVE += 1
            # Usa um único job de ocrmypdf para que cada processo consuma
            # somente um núcleo, permitindo que outras tarefas compartilhem a CPU.
            jobs = rule_options.get("ocr_jobs", 1)
            ocr_kwargs["jobs"] = jobs
            ocr(**ocr_kwargs)
        finally:
            with _OCR_ACTIVE_LOCK:
                _OCR_ACTIVE = max(0, _OCR_ACTIVE - 1)
            OCR_SEMAPHORE.release()

        debug_log(rule_id, base_name,
                  f"Processamento de documento concluído com sucesso. Arquivo de saída: {output_pdfa_path}")
        return True

    except ocr_exceptions.BadArgsError as e:
        logging.error(
            f"[{rule_id}] Erro de argumentos incompatíveis no processamento de documento: {e}. Arquivo: {input_path}")
        debug_log(rule_id, base_name, f"Erro de argumentos incompatíveis no processamento de documento: {e}")
        return False
    except ocr_exceptions.OutputFileAccessError as e:
        logging.error(
            f"[{rule_id}] Erro de acesso ao arquivo de saída: {e}. Certifique-se de que o diretório temporário e de saída são graváveis. Arquivo: {output_pdfa_path}")
        debug_log(rule_id, base_name, f"Erro de acesso temporário: {e}")
        return False
    except ocr_exceptions.EncryptedPdfError:
        logging.error(f"[{rule_id}] PDF de entrada está criptografado e não pôde ser processado: {input_path}")
        debug_log(rule_id, base_name, "Erro: PDF de entrada criptografado.")
        return False
    except ocr_exceptions.MissingDependencyError as e:
        logging.critical(f"[{rule_id}] Dependência crítica de processamento de documento está faltando: {e}")
        debug_log(rule_id, base_name, f"Erro: Dependência faltando: {e}")
        return False
    except ocr_exceptions.InputFileError as e:
        logging.warning(f"[{rule_id}] Erro no arquivo de entrada (pode estar corrompido): {e}")
        debug_log(rule_id, base_name, f"Erro no arquivo de entrada: {e}")
        return False
    except ocr_exceptions.SubprocessOutputError as e:
        stderr = e.stderr.decode(errors="ignore") if e.stderr else "N/A"
        logging.error(f"[{rule_id}] Erro em um subprocesso do processamento de documento. Detalhes: {stderr}")
        debug_log(rule_id, base_name, f"Erro de subprocesso: {stderr}")
        return False
    except Exception as e:
        logging.exception(f"[{rule_id}] Erro inesperado no processamento de documento: {e.__class__.__name__}: {e}")
        debug_log(rule_id, base_name, f"Erro inesperado no processamento de documento: {e.__class__.__name__}: {e}")
        return False


def _apply_image_manipulations(img: Image.Image, rule_options: dict, rule_id: str, original_file_name: str,
                               page_info_for_log: str = "", perform_exif_rotation: bool = True) -> tuple[
    Image.Image, bool]:
    current_img = img
    image_modified = False

    log_prefix = f"{page_info_for_log} " if page_info_for_log else ""

    if rule_options.get("orientation") and perform_exif_rotation:
        debug_log(rule_id, original_file_name, f"{log_prefix}Aplicando auto-orientação via dados de imagem.")
        try:
            exif = current_img._getexif()
            if exif:
                orientation_tag = 274
                if orientation_tag in exif:
                    orientation = exif[orientation_tag]
                    if orientation == 3:
                        current_img = current_img.rotate(180, expand=True, resample=Image.LANCZOS)
                        image_modified = True
                    elif orientation == 6:
                        current_img = current_img.rotate(270, expand=True, resample=Image.LANCZOS)
                        image_modified = True
                    elif orientation == 8:
                        current_img = current_img.rotate(90, expand=True, resample=Image.LANCZOS)
                        image_modified = True
                    if image_modified:
                        debug_log(rule_id, original_file_name,
                                  f"{log_prefix}Orientação de imagem aplicada (código: {orientation})")
        except (AttributeError, KeyError, IndexError, TypeError):
            debug_log(rule_id, original_file_name,
                      f"{log_prefix}Nenhum dado de orientação válido encontrado ou erro ao processar.")
        except Exception as e_exif:
            debug_log(rule_id, original_file_name,
                      f"{log_prefix}Erro ao processar dados de imagem para orientação: {e_exif}")

    if rule_options.get("noise_removal"):
        debug_log(rule_id, original_file_name, f"{log_prefix}Aplicando redução de ruído.")
        try:
            arr = np.array(current_img)
            if current_img.mode == "L":
                dst = cv2.fastNlMeansDenoising(arr, None, h=10, templateWindowSize=7, searchWindowSize=21)
            else:
                dst = cv2.fastNlMeansDenoisingColored(arr, None, h=10, hColor=10, templateWindowSize=7,
                                                      searchWindowSize=21)
            current_img = Image.fromarray(dst)
            image_modified = True
            debug_log(rule_id, original_file_name, f"{log_prefix}Redução de ruído aplicada com sucesso.")
        except Exception as e_noise:
            debug_log(rule_id, original_file_name,
                      f"{log_prefix}Erro ao aplicar redução de ruído: {e_noise}")

    if rule_options.get("binarization"):
        debug_log(rule_id, original_file_name, f"{log_prefix}Aplicando binarização.")
        try:
            if current_img.mode != 'L':
                current_img = current_img.convert('L')
            threshold = int(rule_options.get("binarization_threshold", 128))
            current_img = current_img.point(lambda p: 255 if p > threshold else 0, mode='1')
            image_modified = True
        except Exception as e_bin:
            debug_log(rule_id, original_file_name, f"{log_prefix}Erro ao aplicar binarização: {e_bin}")

    return current_img, image_modified


# --- LÓGICA DE OCR ZONAL E RENOMEAÇÃO ---

def do_ocr_on_region(image: Image.Image, rect: tuple, lang: str, psm: int) -> str:
    """Recorta uma imagem e executa o reconhecimento de texto na região especificada."""
    try:
        # The box is a 4-tuple defining the left, upper, right, and lower pixel coordinate.
        box = (int(rect[0]), int(rect[1]), int(rect[0] + rect[2]), int(rect[1] + rect[3]))
        cropped_img = image.crop(box)

        # Use a temporary file to pass to the OCR tool
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            cropped_img.save(tmp.name)
            tmp_path_str = tmp.name

        output_base = os.path.splitext(tmp_path_str)[0]
        cmd = [
            EXTERNAL_TOOLS["TESSERACT_PATH"],
            tmp_path_str,
            output_base,  # OCR tool appends .txt to this
            "-l", lang,
            "--psm", str(psm)
        ]

        subprocess.run(cmd, capture_output=True, check=True, text=True, encoding='utf-8', timeout=60,
                       creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))

        # Read the result from the output file
        with open(output_base + ".txt", "r", encoding="utf-8") as f:
            text = f.read().strip().replace("\n", " ").replace("\r", "")

        # Sanitize text for filename
        return re.sub(r'[\\/*?:"<>|]', "", text)
    except Exception as e:
        logging.error(f"Erro durante reconhecimento de texto zonal em uma região: {e}")
        return "OCR_ERROR"


def get_image_for_zonal_ocr(input_path: Path, temp_dir: Path) -> Path | None:
    """Converte a primeira página de um documento ou uma imagem para um PNG temporário para reconhecimento de texto."""
    try:
        temp_image_path = temp_dir / f"zonal_ocr_source_{input_path.stem}.png"
        if input_path.suffix.lower() == ".pdf":
            doc = fitz.open(input_path)
            if doc.page_count > 0:
                page = doc[0]
                pix = page.get_pixmap(dpi=300)  # Higher DPI for better OCR
                pix.save(str(temp_image_path))
            doc.close()
        else:  # It's an image file
            img = Image.open(input_path)
            img.save(temp_image_path, "PNG")
            img.close()

        return temp_image_path if temp_image_path.exists() else None
    except Exception as e:
        logging.error(f"Falha ao criar imagem para reconhecimento de texto zonal a partir de '{input_path.name}': {e}")
        return None


def handle_zonal_ocr_renaming(rule_id: str, rule_options: dict, original_input_path: Path,
                              processed_local_path: Path) -> Path:
    """Verifica se há um template de reconhecimento de texto zonal e renomeia o arquivo processado de acordo."""
    template = rule_options.get("zonal_ocr_template")
    if not template or not template.get("zones"):
        return processed_local_path  # No template, no change

    debug_log(rule_id, original_input_path.name,
              "Template de reconhecimento de texto zonal encontrado. Iniciando renomeação.")

    temp_dir = processed_local_path.parent
    source_image_for_ocr = get_image_for_zonal_ocr(original_input_path, temp_dir)

    if not source_image_for_ocr:
        debug_log(rule_id, original_input_path.name,
                  "Falha ao gerar imagem para reconhecimento de texto zonal. Renomeação cancelada.")
        return processed_local_path

    try:
        pil_image = Image.open(source_image_for_ocr)
        extracted_parts = []
        # Sort zones by order to ensure correct filename construction
        for zone in sorted(template["zones"], key=lambda z: z.get('order', 0)):
            rect_coords = zone['rect']  # [x, y, w, h]
            text = do_ocr_on_region(pil_image, rect_coords, template.get("lang", "por"), template.get("psm", 7))
            extracted_parts.append(text)
        pil_image.close()

        delimiter = template.get("delimiter", "_")
        new_base_name = delimiter.join(part for part in extracted_parts if part)

        if template.get("prefix"):
            new_base_name = f"{template['prefix']}{delimiter}{new_base_name}"
        if template.get("suffix"):
            new_base_name = f"{new_base_name}{delimiter}{template['suffix']}"
        if template.get("keep_previous_indexing"):
            # Uses original stem, not the processed one (which might be temporary)
            new_base_name = f"{original_input_path.stem}{delimiter}{new_base_name}"

        if not new_base_name:
            debug_log(rule_id, original_input_path.name,
                      "Reconhecimento de texto zonal não extraiu texto. Mantendo nome original.")
            return processed_local_path

        new_path = processed_local_path.with_name(new_base_name + processed_local_path.suffix)
        os.rename(processed_local_path, new_path)
        debug_log(rule_id, original_input_path.name, f"Arquivo renomeado para: {new_path.name}")
        return new_path

    except Exception as e:
        debug_log(rule_id, original_input_path.name,
                  f"Erro durante o processo de renomeação com reconhecimento de texto zonal: {e}")
        logging.exception(f"Erro na renomeação zonal para {original_input_path.name}")
        return processed_local_path
    finally:
        if source_image_for_ocr and source_image_for_ocr.exists():
            try:
                os.unlink(source_image_for_ocr)
            except OSError:
                pass


def process_image_file(rule_id: str, original_input_path: Path, rule_options: dict) -> tuple[bool, Path | None]:
    file_name = original_input_path.name
    base_name_no_ext = original_input_path.stem
    local_processing_temp_dir = FAST_TMP_DIR / f".octec_proc_{uuid.uuid4().hex}_{base_name_no_ext}"

    try:
        local_processing_temp_dir.mkdir(parents=True, exist_ok=True)
        debug_log(rule_id, file_name, f"Diretório temporário criado com permissões 0o777: {local_processing_temp_dir}")
    except OSError as e:
        debug_log(rule_id, file_name, f"Falha ao criar ou definir permissões para o diretório temporário: {e}")
        logging.error(f"Erro ao criar/definir permissões para dir temporário '{local_processing_temp_dir}': {e}")
        return False, None

    try:
        pil_img = Image.open(original_input_path)
        pil_img.load()
        image_dpi = 300
        if 'dpi' in pil_img.info and pil_img.info['dpi'] is not None:
            if isinstance(pil_img.info['dpi'], tuple) and len(pil_img.info['dpi']) > 0:
                image_dpi = int(pil_img.info['dpi'][0])
                debug_log(rule_id, file_name, f"DPI detectado na imagem: {image_dpi}")
            else:
                debug_log(rule_id, file_name,
                          f"DPI da imagem não formatado como esperado: {pil_img.info['dpi']}. Usando padrão {image_dpi}.")
        else:
            debug_log(rule_id, file_name, f"DPI não encontrado na imagem. Usando padrão: {image_dpi}.")

        rule_options["image_dpi_for_ocrmypdf"] = image_dpi

        processed_pil_img, image_was_modified = _apply_image_manipulations(pil_img, rule_options, rule_id, file_name,
                                                                           perform_exif_rotation=True)

        input_for_ocr_or_conversion = original_input_path
        if image_was_modified:
            temp_processed_img_path = local_processing_temp_dir / f"{base_name_no_ext}_octec_processed{original_input_path.suffix}"
            save_params = {}
            if processed_pil_img.mode == '1' and rule_options.get("output_format", "").lower() != "pdf searchable":
                pass
            elif rule_options.get("output_format", "").lower() == "jpeg":
                save_params['format'] = 'JPEG'
                save_params['quality'] = rule_options.get("jpeg_quality", 85)
                if processed_pil_img.mode == 'RGBA' or processed_pil_img.mode == 'P':
                    processed_pil_img = processed_pil_img.convert('RGB')
            processed_pil_img.save(temp_processed_img_path, **save_params)
            input_for_ocr_or_conversion = temp_processed_img_path
            debug_log(rule_id, file_name, f"Imagem processada e salva temporariamente em: {temp_processed_img_path}")

        pil_img.close()
        if image_was_modified: processed_pil_img.close()

        output_dir = Path(rule_options["output"])
        output_dir.mkdir(parents=True, exist_ok=True)
        output_format = rule_options.get("output_format", "pdf searchable").lower()
        sanitized_output_base = base_name_no_ext  # Preserve original filename as requested
        final_output_local_path = None

        if output_format == "checksum":
            debug_log(rule_id, file_name, "Calculando checksum SHA-256 do arquivo original.")
            hasher = hashlib.sha256()
            with open(original_input_path, 'rb') as afile:
                buf = afile.read(65536)
                while len(buf) > 0:
                    hasher.update(buf)
                    buf = afile.read(65536)
            checksum_val = hasher.hexdigest()
            final_output_network_path = output_dir / f"{sanitized_output_base}.sha256.txt"
            with open(final_output_network_path, "w") as f:
                f.write(f"{checksum_val} *{original_input_path.name}\n")
            return True, final_output_network_path

        tesseract_timeout = rule_options.get("tesseract_timeout", 120)
        creation_flags = 0
        if sys.platform == "win32" and getattr(sys, 'frozen', False):
            creation_flags = subprocess.CREATE_NO_WINDOW

        if output_format == "pdf searchable" or output_format == "pdf/a":
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.pdf"
            if not convert_to_pdfa_ocrmypdf(input_for_ocr_or_conversion, final_output_local_path, rule_options):
                debug_log(rule_id, file_name, f"Falha no processamento de documento para '{output_format}'. Abortando.")
                return False, None

        elif output_format == "pdf":
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.pdf"
            try:
                img_doc = fitz.open()
                img_to_insert = Image.open(input_for_ocr_or_conversion)
                img_bytes = BytesIO()
                img_to_insert.save(img_bytes, format=img_to_insert.format or 'PNG')
                img_bytes.seek(0)
                rect = fitz.Rect(0, 0, img_to_insert.width, img_to_insert.height)
                pdf_page = img_doc.new_page(width=rect.width, height=rect.height)
                pdf_page.insert_image(rect, stream=img_bytes.read())
                img_to_insert.close()
                if rule_options.get("compact"):
                    img_doc.save(str(final_output_local_path), garbage=4, deflate=True, pretty=True)
                else:
                    img_doc.save(str(final_output_local_path), garbage=0, deflate=False, pretty=False)
                img_doc.close()
                debug_log(rule_id, file_name,
                          f"Imagem convertida para PDF (somente imagem) localmente: {final_output_local_path}")
            except Exception as e_pdf_img:
                debug_log(rule_id, file_name, f"Erro ao criar PDF (somente imagem) localmente: {e_pdf_img}")
                return False, None

        elif output_format == "jpeg":
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.jpeg"
            img_opened = Image.open(input_for_ocr_or_conversion)
            if img_opened.mode == 'RGBA' or img_opened.mode == 'P':
                img_opened = img_opened.convert('RGB')
            img_opened.save(final_output_local_path, "JPEG", quality=rule_options.get("jpeg_quality", 85))
            img_opened.close()
            debug_log(rule_id, file_name, f"Imagem convertida para JPEG localmente: {final_output_local_path}")

        elif output_format == "tiff":
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.tiff"
            img_opened = Image.open(input_for_ocr_or_conversion)
            tiff_compression = rule_options.get("tiff_compression", "tiff_lzw")
            img_opened.save(final_output_local_path, "TIFF", compression=tiff_compression)
            img_opened.close()
            debug_log(rule_id, file_name,
                      f"Imagem convertida para TIFF ({tiff_compression}) localmente: {final_output_local_path}")

        elif output_format == "text":
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.txt"
            cmd = [EXTERNAL_TOOLS["TESSERACT_PATH"], str(input_for_ocr_or_conversion),
                   str(final_output_local_path.with_suffix('')),
                   "-l", rule_options.get("language", "eng")]
            if rule_options.get("tesseract_psm"): cmd.extend(["--psm", str(rule_options.get("tesseract_psm"))])
            if rule_options.get("tesseract_oem"): cmd.extend(["--oem", str(rule_options.get("tesseract_oem"))])
            debug_log(rule_id, file_name, f"Executando reconhecimento de texto (texto) localmente: {cmd}")
            subprocess.run(cmd, capture_output=True, check=True, text=True, encoding='utf-8', timeout=tesseract_timeout,
                           creationflags=creation_flags)

        elif output_format == "docx":
            # OCR para TXT e empacota em DOCX simples (somente texto)
            tmp_txt_base = (local_processing_temp_dir / f"{sanitized_output_base}_ocr_tmp").as_posix()
            cmd = [EXTERNAL_TOOLS["TESSERACT_PATH"], str(input_for_ocr_or_conversion), tmp_txt_base,
                   "-l", rule_options.get("language", "eng")]
            if rule_options.get("tesseract_psm"): cmd.extend(["--psm", str(rule_options.get("tesseract_psm"))])
            if rule_options.get("tesseract_oem"): cmd.extend(["--oem", str(rule_options.get("tesseract_oem"))])
            debug_log(rule_id, file_name, f"Executando reconhecimento de texto (DOCX via TXT) localmente: {cmd}")
            subprocess.run(cmd, capture_output=True, check=True, text=True, encoding='utf-8', timeout=tesseract_timeout,
                           creationflags=creation_flags)
            txt_path = Path(tmp_txt_base + ".txt")
            text_data = txt_path.read_text(encoding="utf-8", errors="ignore") if txt_path.exists() else ""
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.docx"
            if not _create_docx_from_text_pages([text_data], final_output_local_path):
                debug_log(rule_id, file_name, "Falha ao criar DOCX a partir de TXT.")
                return False, None

        elif output_format == "rtf":
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.rtf"
            cmd = [EXTERNAL_TOOLS["TESSERACT_PATH"], str(input_for_ocr_or_conversion),
                   str(final_output_local_path.with_suffix('')),
                   "-l", rule_options.get("language", "eng"), "rtf"]
            if rule_options.get("tesseract_psm"): cmd.extend(["--psm", str(rule_options.get("tesseract_psm"))])
            if rule_options.get("tesseract_oem"): cmd.extend(["--oem", str(rule_options.get("tesseract_oem"))])
            debug_log(rule_id, file_name, f"Executando reconhecimento de texto (RTF) localmente: {cmd}")
            subprocess.run(cmd, capture_output=True, check=True, text=True, encoding='utf-8', timeout=tesseract_timeout,
                           creationflags=creation_flags)

        else:
            debug_log(rule_id, file_name, f"Formato de saída não suportado para imagem: {output_format}")
            return False, None

        if final_output_local_path and final_output_local_path.exists():
            # --- PONTO DE INTEGRAÇÃO DO OCR ZONAL ---
            final_output_local_path = handle_zonal_ocr_renaming(rule_id, rule_options, original_input_path,
                                                                final_output_local_path)
            # --- FIM DA INTEGRAÇÃO ---

            # --- Assinatura PDF (se aplicável) ---
            try:
                final_output_local_path = maybe_sign_pdf(
                    rule_id, file_name, rule_options,
                    final_output_local_path,
                    local_processing_temp_dir
                )
                # Verifica assinatura pelo conteúdo
                if _pdf_has_signature(final_output_local_path):
                    debug_log(rule_id, file_name,
                              f"Arquivo assinado gerado (verificado por ByteRange): {final_output_local_path.name}")
                else:
                    debug_log(rule_id, file_name, "Assinatura NÃO detectada no PDF resultante.")
            except Exception as _e_sign:
                debug_log(rule_id, file_name, f"Assinatura ignorada por erro: {_e_sign}")

            final_output_network_path = output_dir / final_output_local_path.name
            debug_log(rule_id, file_name,
                      f"Movendo arquivo processado de {final_output_local_path} para rede: {final_output_network_path}")
            robust_move(str(final_output_local_path), str(final_output_network_path))
            debug_log(rule_id, file_name,
                      f"Arquivo de imagem processado com sucesso. Saída na rede: {final_output_network_path}")
            return True, final_output_network_path
        else:
            debug_log(rule_id, file_name,
                      f"Processamento de imagem concluído, mas arquivo de saída local esperado ({final_output_local_path}) não foi encontrado.")
            return False, None

    except FileNotFoundError as e:
        debug_log(rule_id, file_name,
                  f"Erro: Ferramenta de reconhecimento de texto ou externa não encontrada: {e}. Verifique a instalação e o PATH.")
        logging.error(f"Ferramenta de reconhecimento de texto ou externa não encontrada para '{file_name}': {e}")
        return False, None
    except subprocess.TimeoutExpired:
        debug_log(rule_id, file_name,
                  f"Timeout ({tesseract_timeout}s) durante processamento de imagem com ferramenta de reconhecimento de texto.")
        logging.error(f"Reconhecimento de texto timeout para '{file_name}'.")
        return False, None
    except subprocess.CalledProcessError as e:
        debug_log(rule_id, file_name,
                  f"Erro durante processamento de imagem (subprocesso de reconhecimento de texto): {e.stderr}")
        logging.error(f"Subprocesso de reconhecimento de texto falhou para '{file_name}': {e}")
        return False, None
    except Image.DecompressionBombError as e:
        debug_log(rule_id, file_name,
                  f"Erro de DecompressionBomb ao processar imagem (potencialmente muito grande ou maliciosa): {e}")
        logging.error(f"DecompressionBombError para '{file_name}': {e}")
        return False, None
    except Exception as e:
        debug_log(rule_id, file_name, f"Erro inesperado durante processamento de imagem: {e.__class__.__name__}: {e}")
        logging.exception(f"Erro inesperado durante processamento de imagem para '{file_name}'")
        return False, None
    finally:
        if local_processing_temp_dir and local_processing_temp_dir.exists():
            try:
                shutil.rmtree(local_processing_temp_dir)
                debug_log(rule_id, file_name, f"Diretório temporário local removido: {local_processing_temp_dir}")
            except Exception as e_del:
                debug_log(rule_id, file_name,
                          f"Falha ao remover diretório temporário local {local_processing_temp_dir}: {e_del}")


def process_pdf_file(rule_id: str, original_input_path: Path, rule_options: dict) -> tuple[bool, Path | None]:
    file_name = original_input_path.name
    base_name_no_ext = original_input_path.stem
    output_dir = Path(rule_options["output"])
    output_dir.mkdir(parents=True, exist_ok=True)
    sanitized_output_base = base_name_no_ext  # Preserve original filename as requested
    final_output_local_path = None
    local_processing_temp_dir = FAST_TMP_DIR / f".octec_proc_{uuid.uuid4().hex}_{base_name_no_ext}"

    try:
        local_processing_temp_dir.mkdir(parents=True, exist_ok=True)
        debug_log(rule_id, file_name, f"Diretório temporário criado com permissões 0o777: {local_processing_temp_dir}")
        working_input_path = local_processing_temp_dir / f"{base_name_no_ext}_{uuid.uuid4().hex}{original_input_path.suffix.lower()}"
        shutil.copy2(str(original_input_path), str(working_input_path))
        debug_log(rule_id, file_name, f"Cópia de trabalho criada: {working_input_path}")
    except OSError as e:
        debug_log(rule_id, file_name, f"Falha ao criar ou definir permissões para o diretório temporário: {e}")
        logging.error(f"Erro ao criar/definir permissões para dir temporário '{local_processing_temp_dir}': {e}")
        return False, None

    output_format = rule_options.get("output_format", "pdf searchable").lower()

    try:
        if output_format == "checksum":
            debug_log(rule_id, file_name, "Calculando checksum SHA-256 do PDF original.")
            hasher = hashlib.sha256()
            with open(working_input_path, 'rb') as afile:
                buf = afile.read(65536)
                while len(buf) > 0:
                    hasher.update(buf)
                    buf = afile.read(65536)
            checksum_val = hasher.hexdigest()
            final_output_network_path = output_dir / f"{sanitized_output_base}.sha256.txt"
            with open(final_output_network_path, "w") as f:
                f.write(f"{checksum_val} *{original_input_path.name}\n")
            return True, final_output_network_path

        manipulations_that_require_raster = ["framing", "noise_removal", "binarization"]
        needs_rasterization_for_manipulation = any(rule_options.get(opt) for opt in manipulations_that_require_raster)
        needs_rasterization_for_output = output_format in ["jpeg", "tiff", "text", "rtf", "docx"]

        input_for_final_step = working_input_path

        if not needs_rasterization_for_manipulation and not needs_rasterization_for_output and output_format in [
            "pdf searchable", "pdf/a"]:
            debug_log(rule_id, file_name,
                      "Otimização: Pulando rasterização. O processador de documentos irá processar o PDF original diretamente.")
            update_file_status(rule_id, file_name, "Processando documento", progress=10)
            # --- [FIX] Carimbar antes de converter para PDF/A quando assinatura está habilitada ---
            if output_format == "pdf/a" and rule_options.get("sign_pdf"):
                display_name = (rule_options.get("sign_display_name") or "").strip()
                if display_name:
                    try:
                        pre_stamped = local_processing_temp_dir / f"{sanitized_output_base}_prestamp.pdf"
                        stamp_all_pages(str(working_input_path), str(pre_stamped),
                                        f"Assinado digitalmente por: {display_name}")
                        working_input_path = pre_stamped
                        debug_log(rule_id, file_name, "Carimbo aplicado antes da conversão para PDF/A.")
                    except Exception as e:
                        debug_log(rule_id, file_name, f"Falha ao carimbar antes do PDF/A: {e}")
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.pdf"
            if not convert_to_pdfa_ocrmypdf(working_input_path, final_output_local_path, rule_options):
                debug_log(rule_id, file_name, "Falha no processamento direto de documento.")
                return False, None
        else:
            debug_log(rule_id, file_name,
                      "Documento requer rasterização para manipulação de imagem ou formato de saída específico.")
            temp_raster_dir_for_pdf = local_processing_temp_dir / "raster_pages"
            temp_raster_dir_for_pdf.mkdir(parents=True, exist_ok=True)
            pdf_doc = fitz.open(working_input_path)
            processed_image_paths_for_this_pdf = []

            for page_num in range(pdf_doc.page_count):
                update_file_status(rule_id, file_name, f"Rasterizando pág {page_num + 1}/{pdf_doc.page_count}...",
                                   progress=int(10 + (page_num / pdf_doc.page_count) * 40))
                page = pdf_doc[page_num]
                raster_dpi = compute_adaptive_dpi(page.rect)
                pix = page.get_pixmap(dpi=raster_dpi)
                temp_page_img_path = temp_raster_dir_for_pdf / f"{sanitized_output_base}_page{page_num + 1}.png"
                pix.save(str(temp_page_img_path))

                if needs_rasterization_for_manipulation:
                    page_img_pil = Image.open(temp_page_img_path)
                    page_img_pil.load()
                    processed_page_pil, page_was_modified = _apply_image_manipulations(
                        page_img_pil, rule_options, rule_id, file_name,
                        page_info_for_log=f"Pág {page_num + 1}:",
                        perform_exif_rotation=False
                    )
                    page_img_pil.close()
                    if page_was_modified:
                        processed_page_pil.save(temp_page_img_path)
                        debug_log(rule_id, file_name, f"Página rasterizada {page_num + 1} manipulada e salva.")
                    if page_was_modified: processed_page_pil.close()
                processed_image_paths_for_this_pdf.append(temp_page_img_path)
            pdf_doc.close()

            if not processed_image_paths_for_this_pdf:
                debug_log(rule_id, file_name, "Nenhuma página foi rasterizada do documento. Abortando.")
                return False, None

            if output_format in ["pdf searchable", "pdf", "pdf/a"]:
                input_for_final_step = local_processing_temp_dir / f"{sanitized_output_base}_combined_raster.pdf"
                combined_doc = fitz.open()
                for i, img_path in enumerate(processed_image_paths_for_this_pdf):
                    update_file_status(rule_id, file_name, "Combinando págs...",
                                       progress=int(50 + (i / len(processed_image_paths_for_this_pdf)) * 20))
                    img_pil = Image.open(img_path)
                    rect = fitz.Rect(0, 0, img_pil.width, img_pil.height)
                    img_pil.close()
                    pdf_page = combined_doc.new_page(width=rect.width, height=rect.height)
                    pdf_page.insert_image(rect, filename=str(img_path))

                if rule_options.get("compact"):
                    combined_doc.save(str(input_for_final_step), garbage=4, deflate=True, pretty=True)
                else:
                    combined_doc.save(str(input_for_final_step), garbage=0, deflate=False, pretty=False)
                combined_doc.close()
                debug_log(rule_id, file_name, f"PDF temporário combinado criado em: {input_for_final_step}")

                if output_format in ["pdf searchable", "pdf/a"]:
                    final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.pdf"
                    if not convert_to_pdfa_ocrmypdf(input_for_final_step, final_output_local_path, rule_options):
                        return False, None
                else:
                    final_output_local_path = input_for_final_step

            else:
                # --- Handle non-PDF outputs from rasterized pages ---
                try:
                    if output_format in ["jpeg", "jpg"]:
                        # Save each page as JPEG; if multiple pages, create a ZIP, else a single file
                        jpeg_paths = []
                        for i, png_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"Gerando JPEG pág {i}/{len(processed_image_paths_for_this_pdf)}...",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            with Image.open(png_path) as im:
                                rgb = im.convert("RGB")
                                out_path = local_processing_temp_dir / f"{sanitized_output_base}_page{i}.jpeg"
                                rgb.save(out_path, format="JPEG", quality=rule_options.get("jpeg_quality", 85))
                                jpeg_paths.append(out_path)
                        if len(jpeg_paths) == 1:
                            final_output_local_path = jpeg_paths[0]
                        else:
                            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}_pages_jpeg.zip"
                            with zipfile.ZipFile(final_output_local_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                                for p in jpeg_paths:
                                    zf.write(p, arcname=p.name)
                    elif output_format in ["tiff", "tif"]:
                        # Create a (multi‑page) TIFF if more than one page; otherwise single-page TIFF
                        tiff_compression = rule_options.get("tiff_compression",
                                                            "tiff_deflate")  # 'tiff_deflate','tiff_lzw','group4'
                        pil_mode = "1" if rule_options.get("tiff_bilevel", False) else "RGB"
                        pil_images = []
                        for i, png_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"Gerando TIFF pág {i}/{len(processed_image_paths_for_this_pdf)}...",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            with Image.open(png_path) as im:
                                conv = im.convert(pil_mode)
                                pil_images.append(conv.copy())
                        if len(pil_images) == 1:
                            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.tiff"
                            pil_images[0].save(final_output_local_path, format="TIFF", compression=tiff_compression)
                        else:
                            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.tiff"
                            first, *rest = pil_images
                            first.save(final_output_local_path, save_all=True, append_images=rest, format="TIFF",
                                       compression=tiff_compression)
                        # Close copies
                        for im in pil_images:
                            im.close()
                    elif output_format in ["txt", "text"]:
                        # OCR each raster page to text using Tesseract
                        lang = rule_options.get("ocr_language", "por+eng")
                        tesseract_exe = EXTERNAL_TOOLS["TESSERACT_PATH"]
                        tesseract_timeout = rule_options.get("tesseract_timeout", 180)
                        creation_flags = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
                        texts = []
                        for i, img_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"OCR TXT pág {i}/{len(processed_image_paths_for_this_pdf)}...",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            out_base = (local_processing_temp_dir / f"{sanitized_output_base}_page{i}_ocr").as_posix()
                            cmd = [tesseract_exe, str(img_path), out_base, "-l", lang]
                            try:
                                subprocess.run(cmd, check=True, timeout=tesseract_timeout, creationflags=creation_flags)
                                page_txt = Path(out_base + ".txt").read_text(encoding="utf-8", errors="ignore")
                                texts.append(page_txt)
                            except subprocess.TimeoutExpired:
                                debug_log(rule_id, file_name, f"Tesseract timeout na página {i}.")
                                return False, None
                            except subprocess.CalledProcessError as e:
                                debug_log(rule_id, file_name, f"Erro Tesseract na página {i}: {e}")
                                return False, None
                        final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.txt"
                        final_output_local_path.write_text("\n".join(texts), encoding="utf-8")

                    elif output_format in ["docx"]:
                        # OCR cada página rasterizada para TXT com Tesseract e empacota como DOCX simples
                        lang = rule_options.get("ocr_language", "por+eng")
                        tesseract_exe = EXTERNAL_TOOLS["TESSERACT_PATH"]
                        tesseract_timeout = rule_options.get("tesseract_timeout", 180)
                        creation_flags = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
                        page_texts = []
                        for i, img_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"OCR DOCX pág {i}/{len(processed_image_paths_for_this_pdf)}...",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            out_base = (local_processing_temp_dir / f"{sanitized_output_base}_page{i}_ocr").as_posix()
                            cmd = [tesseract_exe, str(img_path), out_base, "-l", lang]
                            try:
                                subprocess.run(cmd, check=True, timeout=tesseract_timeout, creationflags=creation_flags)
                                page_txt = Path(out_base + ".txt").read_text(encoding="utf-8", errors="ignore")
                                page_texts.append(page_txt)
                            except subprocess.TimeoutExpired:
                                debug_log(rule_id, file_name, f"Tesseract timeout na página {i}.")
                                return False, None
                            except subprocess.CalledProcessError as e:
                                debug_log(rule_id, file_name, f"Erro Tesseract na página {i}: {e}")
                                return False, None
                        final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.docx"
                        if not _create_docx_from_text_pages(page_texts, final_output_local_path):
                            debug_log(rule_id, file_name, "Falha ao criar DOCX a partir do OCR.")
                            return False, None
                    elif output_format in ["xlsx"]:
                        lang = rule_options.get("ocr_language", "por+eng")
                        tesseract_exe = EXTERNAL_TOOLS["TESSERACT_PATH"]
                        tesseract_timeout = rule_options.get("tesseract_timeout", 180)
                        creation_flags = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
                        page_texts = []
                        for i, img_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"OCR XLSX pág {i}/{len(processed_image_paths_for_this_pdf)}.",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            out_base = (local_processing_temp_dir / f"{sanitized_output_base}_page{i}_ocr").as_posix()
                            cmd = [tesseract_exe, str(img_path), out_base, "-l", lang]
                            try:
                                subprocess.run(cmd, check=True, timeout=tesseract_timeout, creationflags=creation_flags)
                                page_txt = Path(out_base + ".txt").read_text(encoding="utf-8", errors="ignore")
                                page_texts.append(page_txt)
                            except subprocess.TimeoutExpired:
                                debug_log(rule_id, file_name, f"Tesseract timeout na página {i}.")
                                return False, None
                            except subprocess.CalledProcessError as e:
                                debug_log(rule_id, file_name, f"Erro Tesseract na página {i}: {e}")
                                return False, None
                        final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.xlsx"
                        if not _create_xlsx_from_text_pages(page_texts, final_output_local_path):
                            debug_log(rule_id, file_name, "Falha ao criar XLSX a partir do OCR.")
                            return False, None

                    elif output_format in ["pptx"]:
                        lang = rule_options.get("ocr_language", "por+eng")
                        tesseract_exe = EXTERNAL_TOOLS["TESSERACT_PATH"]
                        tesseract_timeout = rule_options.get("tesseract_timeout", 180)
                        creation_flags = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
                        page_texts = []
                        for i, img_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"OCR PPTX pág {i}/{len(processed_image_paths_for_this_pdf)}.",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            out_base = (local_processing_temp_dir / f"{sanitized_output_base}_page{i}_ocr").as_posix()
                            cmd = [tesseract_exe, str(img_path), out_base, "-l", lang]
                            try:
                                subprocess.run(cmd, check=True, timeout=tesseract_timeout, creationflags=creation_flags)
                                page_txt = Path(out_base + ".txt").read_text(encoding="utf-8", errors="ignore")
                                page_texts.append(page_txt)
                            except subprocess.TimeoutExpired:
                                debug_log(rule_id, file_name, f"Tesseract timeout na página {i}.")
                                return False, None
                            except subprocess.CalledProcessError as e:
                                debug_log(rule_id, file_name, f"Erro Tesseract na página {i}: {e}")
                                return False, None
                        final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.pptx"
                        if not _create_pptx_from_text_pages(page_texts, final_output_local_path):
                            debug_log(rule_id, file_name, "Falha ao criar PPTX a partir do OCR.")
                            return False, None

                    elif output_format in ["rtf"]:
                        # OCR cada página rasterizada para TXT com Tesseract e empacota como RTF único
                        lang = rule_options.get("ocr_language", "por+eng")
                        tesseract_exe = EXTERNAL_TOOLS["TESSERACT_PATH"]
                        tesseract_timeout = rule_options.get("tesseract_timeout", 180)
                        creation_flags = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
                        texts = []
                        for i, img_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"OCR RTF pág {i}/{len(processed_image_paths_for_this_pdf)}.",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            out_base = (local_processing_temp_dir / f"{sanitized_output_base}_page{i}_ocr").as_posix()
                            cmd = [tesseract_exe, str(img_path), out_base, "-l", lang]
                            try:
                                subprocess.run(cmd, check=True, timeout=tesseract_timeout, creationflags=creation_flags)
                                page_txt = Path(out_base + ".txt").read_text(encoding="utf-8", errors="ignore")
                                texts.append(page_txt)
                            except subprocess.TimeoutExpired:
                                debug_log(rule_id, file_name, f"Tesseract timeout na página {i}.")
                                return False, None
                            except subprocess.CalledProcessError as e:
                                debug_log(rule_id, file_name, f"Erro Tesseract na página {i}: {e}")
                                return False, None

                        # Empacota todo o texto em um único RTF simples
                        def _escape_rtf(s: str) -> str:
                            # Escapa barras, chaves e converte quebras de linha em \par
                            return s.replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}").replace("\r",
                                                                                                           "").replace(
                                "\n", "\\par\n")

                        # Constrói RTF com fonte Calibri 11pt
                        rtf_header = r"{\rtf1\ansi\deff0\nouicompat{\fonttbl{\f0\fnil\fcharset0 Calibri;}}{\pard\sa200\sl276\slmult1\f0\fs22 "
                        rtf_footer = r"}\par}"

                        escaped_text = _escape_rtf("\n".join(texts))

                        # Constrói RTF com Unicode usando \uN? e \uc1 para compatibilidade ampla
                        def _to_rtf_unicode(s: str) -> str:
                            # Normaliza quebras de linha
                            s = s.replace("\r\n", "\n").replace("\r", "\n")
                            out_parts = []
                            for ch in s:
                                code = ord(ch)
                                if ch == "\\":
                                    out_parts.append(r"\\\\")
                                elif ch == "{":
                                    out_parts.append(r"\{")
                                elif ch == "}":
                                    out_parts.append(r"\}")
                                elif ch == "\n":
                                    out_parts.append(r"\par" + "\n")
                                else:
                                    # ASCII imprimível seguro
                                    if 32 <= code <= 126:
                                        out_parts.append(ch)
                                    else:
                                        # RTF usa signed 16-bit para \uN
                                        if code > 0x7FFF:
                                            code_signed = code - 0x10000
                                        else:
                                            code_signed = code
                                        out_parts.append(fr"\u{code_signed}?")
                            return "".join(out_parts)

                        rtf_header = (
                            r"{\rtf1\ansi\ansicpg1252\deff0\uc1"
                            r"{\fonttbl{\f0\fnil\fcharset0 Calibri;}}"
                            r"{\pard\sa200\sl276\slmult1\f0\fs22 "
                        )
                        rtf_footer = r"}\par}"

                        escaped_text = _to_rtf_unicode("\n".join(texts))
                        final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.rtf"
                        with open(final_output_local_path, "w", encoding="utf-8") as _rtf_out:
                            _rtf_out.write(rtf_header + escaped_text + rtf_footer)

                    else:
                        debug_log(rule_id, file_name, f"Lógica de rasterização não cobriu o formato: {output_format}")
                        return False, None
                except Exception as e_nonpdf:
                    debug_log(rule_id, file_name,
                              f"Falha ao processar formato {output_format} a partir de páginas rasterizadas: {e_nonpdf}")
                    return False, None

        update_file_status(rule_id, file_name, "Concluindo", progress=95)

        if final_output_local_path and final_output_local_path.exists():
            # --- PONTO DE INTEGRAÇÃO DO OCR ZONAL ---
            final_output_local_path = handle_zonal_ocr_renaming(rule_id, rule_options, original_input_path,
                                                                final_output_local_path)
            # --- FIM DA INTEGRAÇÃO ---

            # --- Assinatura PDF (se aplicável) ---
            try:
                debug_log(rule_id, file_name, "Iniciando etapa de assinatura (PDF - entrada é PDF).")
                new_signed_path = maybe_sign_pdf(
                    rule_id, file_name, rule_options,
                    final_output_local_path,
                    local_processing_temp_dir
                )
                if _pdf_has_signature(new_signed_path if new_signed_path else final_output_local_path):
                    debug_log(rule_id, file_name,
                              f"Arquivo assinado gerado (verificado por ByteRange): {(new_signed_path if new_signed_path else final_output_local_path).name}")
                    final_output_local_path = new_signed_path if new_signed_path else final_output_local_path
                else:
                    debug_log(rule_id, file_name, "Assinatura NÃO detectada no PDF resultante.")
                    final_output_local_path = new_signed_path
            except Exception as _e_sign:
                debug_log(rule_id, file_name, f"Erro ao executar assinatura: {_e_sign}")
            final_output_network_path = output_dir / final_output_local_path.name
            debug_log(rule_id, file_name,
                      f"Movendo arquivo processado de {final_output_local_path} para rede: {final_output_network_path}")
            robust_move(str(final_output_local_path), str(final_output_network_path))
            update_file_status(rule_id, file_name, "Concluído", progress=100)
            debug_log(rule_id, file_name,
                      f"Arquivo PDF processado com sucesso. Saída na rede: {final_output_network_path}")
            return True, final_output_network_path
        else:
            debug_log(rule_id, file_name,
                      f"Processamento de PDF concluído, mas arquivo de saída local esperado ({final_output_local_path}) não foi encontrado.")
            update_file_status(rule_id, file_name, "Falha: Arquivo de saída ausente", progress=0)
            return False, None

    except Exception as e:
        debug_log(rule_id, file_name, f"Erro crítico durante processamento de PDF: {e.__class__.__name__}: {e}")
        logging.exception(f"Erro crítico durante processamento de PDF para '{file_name}'")
        update_file_status(rule_id, file_name, "Falha Crítica", progress=0)
        return False, None
    finally:
        if local_processing_temp_dir and local_processing_temp_dir.exists():
            try:
                shutil.rmtree(local_processing_temp_dir)
                debug_log(rule_id, file_name, f"Diretório temporário local removido: {local_processing_temp_dir}")
            except Exception as e_rm:
                debug_log(rule_id, file_name,
                          f"Falha ao remover diretório temporário local {local_processing_temp_dir}: {e_rm}")


# --- Engine Management ---
observers = {}


def process_image_file(rule_id: str, original_input_path: Path, rule_options: dict) -> tuple[bool, Path | None]:
    file_name = original_input_path.name
    base_name_no_ext = original_input_path.stem
    local_processing_temp_dir = FAST_TMP_DIR / f".octec_proc_{uuid.uuid4().hex}_{base_name_no_ext}"

    try:
        local_processing_temp_dir.mkdir(parents=True, exist_ok=True)
        debug_log(rule_id, file_name, f"Diretório temporário criado com permissões 0o777: {local_processing_temp_dir}")
        working_input_path = local_processing_temp_dir / f"{base_name_no_ext}_{uuid.uuid4().hex}{original_input_path.suffix.lower()}"
        shutil.copy2(str(original_input_path), str(working_input_path))
        debug_log(rule_id, file_name, f"Cópia de trabalho criada: {working_input_path}")
    except OSError as e:
        debug_log(rule_id, file_name, f"Falha ao criar ou definir permissões para o diretório temporário: {e}")
        logging.error(f"Erro ao criar/definir permissões para dir temporário '{local_processing_temp_dir}': {e}")
        return False, None

    try:
        pil_img = Image.open(working_input_path)
        pil_img.load()
        image_dpi = 300
        if 'dpi' in pil_img.info and pil_img.info['dpi'] is not None:
            if isinstance(pil_img.info['dpi'], tuple) and len(pil_img.info['dpi']) > 0:
                image_dpi = int(pil_img.info['dpi'][0])
                debug_log(rule_id, file_name, f"DPI detectado na imagem: {image_dpi}")
            else:
                debug_log(rule_id, file_name,
                          f"DPI da imagem não formatado como esperado: {pil_img.info['dpi']}. Usando padrão {image_dpi}.")
        else:
            debug_log(rule_id, file_name, f"DPI não encontrado na imagem. Usando padrão: {image_dpi}.")

        rule_options["image_dpi_for_ocrmypdf"] = image_dpi

        processed_pil_img, image_was_modified = _apply_image_manipulations(pil_img, rule_options, rule_id, file_name,
                                                                           perform_exif_rotation=True)

        input_for_ocr_or_conversion = original_input_path
        if image_was_modified:
            temp_processed_img_path = local_processing_temp_dir / f"{base_name_no_ext}_octec_processed{original_input_path.suffix}"
            save_params = {}
            if processed_pil_img.mode == '1' and rule_options.get("output_format", "").lower() != "pdf searchable":
                pass
            elif rule_options.get("output_format", "").lower() == "jpeg":
                save_params['format'] = 'JPEG'
                save_params['quality'] = rule_options.get("jpeg_quality", 85)
                if processed_pil_img.mode == 'RGBA' or processed_pil_img.mode == 'P':
                    processed_pil_img = processed_pil_img.convert('RGB')
            processed_pil_img.save(temp_processed_img_path, **save_params)
            input_for_ocr_or_conversion = temp_processed_img_path
            debug_log(rule_id, file_name, f"Imagem processada e salva temporariamente em: {temp_processed_img_path}")

        pil_img.close()
        if image_was_modified: processed_pil_img.close()

        output_dir = Path(rule_options["output"])
        output_dir.mkdir(parents=True, exist_ok=True)
        output_format = rule_options.get("output_format", "pdf searchable").lower()
        sanitized_output_base = base_name_no_ext  # Preserve original filename as requested
        final_output_local_path = None

        if output_format == "checksum":
            debug_log(rule_id, file_name, "Calculando checksum SHA-256 do arquivo original.")
            hasher = hashlib.sha256()
            with open(original_input_path, 'rb') as afile:
                buf = afile.read(65536)
                while len(buf) > 0:
                    hasher.update(buf)
                    buf = afile.read(65536)
            checksum_val = hasher.hexdigest()
            final_output_network_path = output_dir / f"{sanitized_output_base}.sha256.txt"
            with open(final_output_network_path, "w") as f:
                f.write(f"{checksum_val} *{original_input_path.name}\n")
            return True, final_output_network_path

        tesseract_timeout = rule_options.get("tesseract_timeout", 120)
        creation_flags = 0
        if sys.platform == "win32" and getattr(sys, 'frozen', False):
            creation_flags = subprocess.CREATE_NO_WINDOW

        if output_format == "pdf searchable" or output_format == "pdf/a":
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.pdf"
            if not convert_to_pdfa_ocrmypdf(input_for_ocr_or_conversion, final_output_local_path, rule_options):
                debug_log(rule_id, file_name, f"Falha no processamento de documento para '{output_format}'. Abortando.")
                return False, None

        elif output_format == "pdf":
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.pdf"
            try:
                img_doc = fitz.open()
                img_to_insert = Image.open(input_for_ocr_or_conversion)
                img_bytes = BytesIO()
                img_to_insert.save(img_bytes, format=img_to_insert.format or 'PNG')
                img_bytes.seek(0)
                rect = fitz.Rect(0, 0, img_to_insert.width, img_to_insert.height)
                pdf_page = img_doc.new_page(width=rect.width, height=rect.height)
                pdf_page.insert_image(rect, stream=img_bytes.read())
                img_to_insert.close()
                if rule_options.get("compact"):
                    img_doc.save(str(final_output_local_path), garbage=4, deflate=True, pretty=True)
                else:
                    img_doc.save(str(final_output_local_path), garbage=0, deflate=False, pretty=False)
                img_doc.close()
                debug_log(rule_id, file_name,
                          f"Imagem convertida para PDF (somente imagem) localmente: {final_output_local_path}")
            except Exception as e_pdf_img:
                debug_log(rule_id, file_name, f"Erro ao criar PDF (somente imagem) localmente: {e_pdf_img}")
                return False, None

        elif output_format == "jpeg":
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.jpeg"
            img_opened = Image.open(input_for_ocr_or_conversion)
            if img_opened.mode == 'RGBA' or img_opened.mode == 'P':
                img_opened = img_opened.convert('RGB')
            img_opened.save(final_output_local_path, "JPEG", quality=rule_options.get("jpeg_quality", 85))
            img_opened.close()
            debug_log(rule_id, file_name, f"Imagem convertida para JPEG localmente: {final_output_local_path}")

        elif output_format == "tiff":
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.tiff"
            img_opened = Image.open(input_for_ocr_or_conversion)
            tiff_compression = rule_options.get("tiff_compression", "tiff_lzw")
            img_opened.save(final_output_local_path, "TIFF", compression=tiff_compression)
            img_opened.close()
            debug_log(rule_id, file_name,
                      f"Imagem convertida para TIFF ({tiff_compression}) localmente: {final_output_local_path}")

        elif output_format == "text":
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.txt"
            cmd = [EXTERNAL_TOOLS["TESSERACT_PATH"], str(input_for_ocr_or_conversion),
                   str(final_output_local_path.with_suffix('')),
                   "-l", rule_options.get("language", "eng")]
            if rule_options.get("tesseract_psm"): cmd.extend(["--psm", str(rule_options.get("tesseract_psm"))])
            if rule_options.get("tesseract_oem"): cmd.extend(["--oem", str(rule_options.get("tesseract_oem"))])
            debug_log(rule_id, file_name, f"Executando reconhecimento de texto (texto) localmente: {cmd}")
            subprocess.run(cmd, capture_output=True, check=True, text=True, encoding='utf-8', timeout=tesseract_timeout,
                           creationflags=creation_flags)

        elif output_format == "docx":
            # OCR para TXT e empacota em DOCX simples (somente texto)
            tmp_txt_base = (local_processing_temp_dir / f"{sanitized_output_base}_ocr_tmp").as_posix()
            cmd = [EXTERNAL_TOOLS["TESSERACT_PATH"], str(input_for_ocr_or_conversion), tmp_txt_base,
                   "-l", rule_options.get("language", "eng")]
            if rule_options.get("tesseract_psm"): cmd.extend(["--psm", str(rule_options.get("tesseract_psm"))])
            if rule_options.get("tesseract_oem"): cmd.extend(["--oem", str(rule_options.get("tesseract_oem"))])
            debug_log(rule_id, file_name, f"Executando reconhecimento de texto (DOCX via TXT) localmente: {cmd}")
            subprocess.run(cmd, capture_output=True, check=True, text=True, encoding='utf-8', timeout=tesseract_timeout,
                           creationflags=creation_flags)
            txt_path = Path(tmp_txt_base + ".txt")
            text_data = txt_path.read_text(encoding="utf-8", errors="ignore") if txt_path.exists() else ""
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.docx"
            if not _create_docx_from_text_pages([text_data], final_output_local_path):
                debug_log(rule_id, file_name, "Falha ao criar DOCX a partir de TXT.")
                return False, None

        elif output_format == "rtf":
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.rtf"
            cmd = [EXTERNAL_TOOLS["TESSERACT_PATH"], str(input_for_ocr_or_conversion),
                   str(final_output_local_path.with_suffix('')),
                   "-l", rule_options.get("language", "eng"), "rtf"]
            if rule_options.get("tesseract_psm"): cmd.extend(["--psm", str(rule_options.get("tesseract_psm"))])
            if rule_options.get("tesseract_oem"): cmd.extend(["--oem", str(rule_options.get("tesseract_oem"))])
            debug_log(rule_id, file_name, f"Executando reconhecimento de texto (RTF) localmente: {cmd}")
            subprocess.run(cmd, capture_output=True, check=True, text=True, encoding='utf-8', timeout=tesseract_timeout,
                           creationflags=creation_flags)

        else:
            debug_log(rule_id, file_name, f"Formato de saída não suportado para imagem: {output_format}")
            return False, None

        if final_output_local_path and final_output_local_path.exists():
            # --- PONTO DE INTEGRAÇÃO DO OCR ZONAL ---
            final_output_local_path = handle_zonal_ocr_renaming(rule_id, rule_options, original_input_path,
                                                                final_output_local_path)
            # --- FIM DA INTEGRAÇÃO ---

            # --- Assinatura PDF (se aplicável) ---
            try:
                debug_log(rule_id, file_name, "Iniciando etapa de assinatura (PDF - entrada é PDF).")
                new_signed_path = maybe_sign_pdf(
                    rule_id, file_name, rule_options,
                    final_output_local_path,
                    local_processing_temp_dir
                )
                if _pdf_has_signature(new_signed_path if new_signed_path else final_output_local_path):
                    debug_log(rule_id, file_name,
                              f"Arquivo assinado gerado (verificado por ByteRange): {(new_signed_path if new_signed_path else final_output_local_path).name}")
                    final_output_local_path = new_signed_path if new_signed_path else final_output_local_path
                else:
                    debug_log(rule_id, file_name, "Assinatura NÃO detectada no PDF resultante.")
                    final_output_local_path = new_signed_path
            except Exception as _e_sign:
                debug_log(rule_id, file_name, f"Erro ao executar assinatura: {_e_sign}")
            final_output_network_path = output_dir / final_output_local_path.name
            debug_log(rule_id, file_name,
                      f"Movendo arquivo processado de {final_output_local_path} para rede: {final_output_network_path}")
            robust_move(str(final_output_local_path), str(final_output_network_path))
            debug_log(rule_id, file_name,
                      f"Arquivo de imagem processado com sucesso. Saída na rede: {final_output_network_path}")
            return True, final_output_network_path
        else:
            debug_log(rule_id, file_name,
                      f"Processamento de imagem concluído, mas arquivo de saída local esperado ({final_output_local_path}) não foi encontrado.")
            return False, None

    except FileNotFoundError as e:
        debug_log(rule_id, file_name,
                  f"Erro: Ferramenta de reconhecimento de texto ou externa não encontrada: {e}. Verifique a instalação e o PATH.")
        logging.error(f"Ferramenta de reconhecimento de texto ou externa não encontrada para '{file_name}': {e}")
        return False, None
    except subprocess.TimeoutExpired:
        debug_log(rule_id, file_name,
                  f"Timeout ({tesseract_timeout}s) durante processamento de imagem com ferramenta de reconhecimento de texto.")
        logging.error(f"Reconhecimento de texto timeout para '{file_name}'.")
        return False, None
    except subprocess.CalledProcessError as e:
        debug_log(rule_id, file_name,
                  f"Erro durante processamento de imagem (subprocesso de reconhecimento de texto): {e.stderr}")
        logging.error(f"Subprocesso de reconhecimento de texto falhou para '{file_name}': {e}")
        return False, None
    except Image.DecompressionBombError as e:
        debug_log(rule_id, file_name,
                  f"Erro de DecompressionBomb ao processar imagem (potencialmente muito grande ou maliciosa): {e}")
        logging.error(f"DecompressionBombError para '{file_name}': {e}")
        return False, None
    except Exception as e:
        debug_log(rule_id, file_name, f"Erro inesperado durante processamento de imagem: {e.__class__.__name__}: {e}")
        logging.exception(f"Erro inesperado durante processamento de imagem para '{file_name}'")
        return False, None
    finally:
        if local_processing_temp_dir and local_processing_temp_dir.exists():
            try:
                shutil.rmtree(local_processing_temp_dir)
                debug_log(rule_id, file_name, f"Diretório temporário local removido: {local_processing_temp_dir}")
            except Exception as e_del:
                debug_log(rule_id, file_name,
                          f"Falha ao remover diretório temporário local {local_processing_temp_dir}: {e_del}")


def process_pdf_file(rule_id: str, original_input_path: Path, rule_options: dict) -> tuple[bool, Path | None]:
    file_name = original_input_path.name
    base_name_no_ext = original_input_path.stem
    output_dir = Path(rule_options["output"])
    output_dir.mkdir(parents=True, exist_ok=True)
    sanitized_output_base = base_name_no_ext  # Preserve original filename as requested
    final_output_local_path = None
    local_processing_temp_dir = FAST_TMP_DIR / f".octec_proc_{uuid.uuid4().hex}_{base_name_no_ext}"

    try:
        local_processing_temp_dir.mkdir(parents=True, exist_ok=True)
        debug_log(rule_id, file_name, f"Diretório temporário criado com permissões 0o777: {local_processing_temp_dir}")
    except OSError as e:
        debug_log(rule_id, file_name, f"Falha ao criar ou definir permissões para o diretório temporário: {e}")
        logging.error(f"Erro ao criar/definir permissões para dir temporário '{local_processing_temp_dir}': {e}")
        return False, None

    output_format = rule_options.get("output_format", "pdf searchable").lower()

    try:
        if output_format == "checksum":
            debug_log(rule_id, file_name, "Calculando checksum SHA-256 do PDF original.")
            hasher = hashlib.sha256()
            with open(original_input_path, 'rb') as afile:
                buf = afile.read(65536)
                while len(buf) > 0:
                    hasher.update(buf)
                    buf = afile.read(65536)
            checksum_val = hasher.hexdigest()
            final_output_network_path = output_dir / f"{sanitized_output_base}.sha256.txt"
            with open(final_output_network_path, "w") as f:
                f.write(f"{checksum_val} *{original_input_path.name}\n")
            return True, final_output_network_path

        manipulations_that_require_raster = ["framing", "noise_removal", "binarization"]
        needs_rasterization_for_manipulation = any(rule_options.get(opt) for opt in manipulations_that_require_raster)
        needs_rasterization_for_output = output_format in ["jpeg", "tiff", "text", "rtf", "docx"]

        input_for_final_step = original_input_path

        if not needs_rasterization_for_manipulation and not needs_rasterization_for_output and output_format in [
            "pdf searchable", "pdf/a"]:
            debug_log(rule_id, file_name,
                      "Otimização: Pulando rasterização. O processador de documentos irá processar o PDF original diretamente.")
            update_file_status(rule_id, file_name, "Processando documento", progress=10)
            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.pdf"
            if not convert_to_pdfa_ocrmypdf(original_input_path, final_output_local_path, rule_options):
                debug_log(rule_id, file_name, "Falha no processamento direto de documento.")
                return False, None
        else:
            debug_log(rule_id, file_name,
                      "Documento requer rasterização para manipulação de imagem ou formato de saída específico.")
            temp_raster_dir_for_pdf = local_processing_temp_dir / "raster_pages"
            temp_raster_dir_for_pdf.mkdir(parents=True, exist_ok=True)
            pdf_doc = fitz.open(original_input_path)
            processed_image_paths_for_this_pdf = []

            for page_num in range(pdf_doc.page_count):
                update_file_status(rule_id, file_name, f"Rasterizando pág {page_num + 1}/{pdf_doc.page_count}...",
                                   progress=int(10 + (page_num / pdf_doc.page_count) * 40))
                page = pdf_doc[page_num]
                raster_dpi = compute_adaptive_dpi(page.rect)
                pix = page.get_pixmap(dpi=raster_dpi)
                temp_page_img_path = temp_raster_dir_for_pdf / f"{sanitized_output_base}_page{page_num + 1}.png"
                pix.save(str(temp_page_img_path))

                if needs_rasterization_for_manipulation:
                    page_img_pil = Image.open(temp_page_img_path)
                    page_img_pil.load()
                    processed_page_pil, page_was_modified = _apply_image_manipulations(
                        page_img_pil, rule_options, rule_id, file_name,
                        page_info_for_log=f"Pág {page_num + 1}:",
                        perform_exif_rotation=False
                    )
                    page_img_pil.close()
                    if page_was_modified:
                        processed_page_pil.save(temp_page_img_path)
                        debug_log(rule_id, file_name, f"Página rasterizada {page_num + 1} manipulada e salva.")
                    if page_was_modified: processed_page_pil.close()
                processed_image_paths_for_this_pdf.append(temp_page_img_path)
            pdf_doc.close()

            if not processed_image_paths_for_this_pdf:
                debug_log(rule_id, file_name, "Nenhuma página foi rasterizada do documento. Abortando.")
                return False, None

            if output_format in ["pdf searchable", "pdf", "pdf/a"]:
                input_for_final_step = local_processing_temp_dir / f"{sanitized_output_base}_combined_raster.pdf"
                combined_doc = fitz.open()
                for i, img_path in enumerate(processed_image_paths_for_this_pdf):
                    update_file_status(rule_id, file_name, "Combinando págs...",
                                       progress=int(50 + (i / len(processed_image_paths_for_this_pdf)) * 20))
                    img_pil = Image.open(img_path)
                    rect = fitz.Rect(0, 0, img_pil.width, img_pil.height)
                    img_pil.close()
                    pdf_page = combined_doc.new_page(width=rect.width, height=rect.height)
                    pdf_page.insert_image(rect, filename=str(img_path))

                if rule_options.get("compact"):
                    combined_doc.save(str(input_for_final_step), garbage=4, deflate=True, pretty=True)
                else:
                    combined_doc.save(str(input_for_final_step), garbage=0, deflate=False, pretty=False)
                combined_doc.close()
                debug_log(rule_id, file_name, f"PDF temporário combinado criado em: {input_for_final_step}")

                if output_format in ["pdf searchable", "pdf/a"]:
                    final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.pdf"
                    if not convert_to_pdfa_ocrmypdf(input_for_final_step, final_output_local_path, rule_options):
                        return False, None
                else:
                    final_output_local_path = input_for_final_step

            else:
                # --- Handle non-PDF outputs from rasterized pages ---
                try:
                    if output_format in ["jpeg", "jpg"]:
                        # Save each page as JPEG; if multiple pages, create a ZIP, else a single file
                        jpeg_paths = []
                        for i, png_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"Gerando JPEG pág {i}/{len(processed_image_paths_for_this_pdf)}...",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            with Image.open(png_path) as im:
                                rgb = im.convert("RGB")
                                out_path = local_processing_temp_dir / f"{sanitized_output_base}_page{i}.jpeg"
                                rgb.save(out_path, format="JPEG", quality=rule_options.get("jpeg_quality", 85))
                                jpeg_paths.append(out_path)
                        if len(jpeg_paths) == 1:
                            final_output_local_path = jpeg_paths[0]
                        else:
                            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}_pages_jpeg.zip"
                            with zipfile.ZipFile(final_output_local_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                                for p in jpeg_paths:
                                    zf.write(p, arcname=p.name)
                    elif output_format in ["tiff", "tif"]:
                        # Create a (multi‑page) TIFF if more than one page; otherwise single-page TIFF
                        tiff_compression = rule_options.get("tiff_compression",
                                                            "tiff_deflate")  # 'tiff_deflate','tiff_lzw','group4'
                        pil_mode = "1" if rule_options.get("tiff_bilevel", False) else "RGB"
                        pil_images = []
                        for i, png_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"Gerando TIFF pág {i}/{len(processed_image_paths_for_this_pdf)}...",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            with Image.open(png_path) as im:
                                conv = im.convert(pil_mode)
                                pil_images.append(conv.copy())
                        if len(pil_images) == 1:
                            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.tiff"
                            pil_images[0].save(final_output_local_path, format="TIFF", compression=tiff_compression)
                        else:
                            final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.tiff"
                            first, *rest = pil_images
                            first.save(final_output_local_path, save_all=True, append_images=rest, format="TIFF",
                                       compression=tiff_compression)
                        # Close copies
                        for im in pil_images:
                            im.close()
                    elif output_format in ["txt", "text"]:
                        # OCR each raster page to text using Tesseract
                        lang = rule_options.get("ocr_language", "por+eng")
                        tesseract_exe = EXTERNAL_TOOLS["TESSERACT_PATH"]
                        tesseract_timeout = rule_options.get("tesseract_timeout", 180)
                        creation_flags = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
                        texts = []
                        for i, img_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"OCR TXT pág {i}/{len(processed_image_paths_for_this_pdf)}...",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            out_base = (local_processing_temp_dir / f"{sanitized_output_base}_page{i}_ocr").as_posix()
                            cmd = [tesseract_exe, str(img_path), out_base, "-l", lang]
                            try:
                                subprocess.run(cmd, check=True, timeout=tesseract_timeout, creationflags=creation_flags)
                                page_txt = Path(out_base + ".txt").read_text(encoding="utf-8", errors="ignore")
                                texts.append(page_txt)
                            except subprocess.TimeoutExpired:
                                debug_log(rule_id, file_name, f"Tesseract timeout na página {i}.")
                                return False, None
                            except subprocess.CalledProcessError as e:
                                debug_log(rule_id, file_name, f"Erro Tesseract na página {i}: {e}")
                                return False, None
                        final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.txt"
                        final_output_local_path.write_text("\n".join(texts), encoding="utf-8")

                    elif output_format in ["docx"]:
                        # OCR cada página rasterizada para TXT com Tesseract e empacota como DOCX simples
                        lang = rule_options.get("ocr_language", "por+eng")
                        tesseract_exe = EXTERNAL_TOOLS["TESSERACT_PATH"]
                        tesseract_timeout = rule_options.get("tesseract_timeout", 180)
                        creation_flags = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
                        page_texts = []
                        for i, img_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"OCR DOCX pág {i}/{len(processed_image_paths_for_this_pdf)}...",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            out_base = (local_processing_temp_dir / f"{sanitized_output_base}_page{i}_ocr").as_posix()
                            cmd = [tesseract_exe, str(img_path), out_base, "-l", lang]
                            try:
                                subprocess.run(cmd, check=True, timeout=tesseract_timeout, creationflags=creation_flags)
                                page_txt = Path(out_base + ".txt").read_text(encoding="utf-8", errors="ignore")
                                page_texts.append(page_txt)
                            except subprocess.TimeoutExpired:
                                debug_log(rule_id, file_name, f"Tesseract timeout na página {i}.")
                                return False, None
                            except subprocess.CalledProcessError as e:
                                debug_log(rule_id, file_name, f"Erro Tesseract na página {i}: {e}")
                                return False, None
                        final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.docx"
                        if not _create_docx_from_text_pages(page_texts, final_output_local_path):
                            debug_log(rule_id, file_name, "Falha ao criar DOCX a partir do OCR.")
                            return False, None
                    elif output_format in ["xlsx"]:
                        lang = rule_options.get("ocr_language", "por+eng")
                        tesseract_exe = EXTERNAL_TOOLS["TESSERACT_PATH"]
                        tesseract_timeout = rule_options.get("tesseract_timeout", 180)
                        creation_flags = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
                        page_texts = []
                        for i, img_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"OCR XLSX pág {i}/{len(processed_image_paths_for_this_pdf)}.",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            out_base = (local_processing_temp_dir / f"{sanitized_output_base}_page{i}_ocr").as_posix()
                            cmd = [tesseract_exe, str(img_path), out_base, "-l", lang]
                            try:
                                subprocess.run(cmd, check=True, timeout=tesseract_timeout, creationflags=creation_flags)
                                page_txt = Path(out_base + ".txt").read_text(encoding="utf-8", errors="ignore")
                                page_texts.append(page_txt)
                            except subprocess.TimeoutExpired:
                                debug_log(rule_id, file_name, f"Tesseract timeout na página {i}.")
                                return False, None
                            except subprocess.CalledProcessError as e:
                                debug_log(rule_id, file_name, f"Erro Tesseract na página {i}: {e}")
                                return False, None
                        final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.xlsx"
                        if not _create_xlsx_from_text_pages(page_texts, final_output_local_path):
                            debug_log(rule_id, file_name, "Falha ao criar XLSX a partir do OCR.")
                            return False, None

                    elif output_format in ["pptx"]:
                        lang = rule_options.get("ocr_language", "por+eng")
                        tesseract_exe = EXTERNAL_TOOLS["TESSERACT_PATH"]
                        tesseract_timeout = rule_options.get("tesseract_timeout", 180)
                        creation_flags = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
                        page_texts = []
                        for i, img_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"OCR PPTX pág {i}/{len(processed_image_paths_for_this_pdf)}.",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            out_base = (local_processing_temp_dir / f"{sanitized_output_base}_page{i}_ocr").as_posix()
                            cmd = [tesseract_exe, str(img_path), out_base, "-l", lang]
                            try:
                                subprocess.run(cmd, check=True, timeout=tesseract_timeout, creationflags=creation_flags)
                                page_txt = Path(out_base + ".txt").read_text(encoding="utf-8", errors="ignore")
                                page_texts.append(page_txt)
                            except subprocess.TimeoutExpired:
                                debug_log(rule_id, file_name, f"Tesseract timeout na página {i}.")
                                return False, None
                            except subprocess.CalledProcessError as e:
                                debug_log(rule_id, file_name, f"Erro Tesseract na página {i}: {e}")
                                return False, None
                        final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.pptx"
                        if not _create_pptx_from_text_pages(page_texts, final_output_local_path):
                            debug_log(rule_id, file_name, "Falha ao criar PPTX a partir do OCR.")
                            return False, None

                    elif output_format in ["rtf"]:
                        # OCR cada página rasterizada para TXT com Tesseract e empacota como RTF único
                        lang = rule_options.get("ocr_language", "por+eng")
                        tesseract_exe = EXTERNAL_TOOLS["TESSERACT_PATH"]
                        tesseract_timeout = rule_options.get("tesseract_timeout", 180)
                        creation_flags = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
                        texts = []
                        for i, img_path in enumerate(processed_image_paths_for_this_pdf, start=1):
                            update_file_status(rule_id, file_name,
                                               f"OCR RTF pág {i}/{len(processed_image_paths_for_this_pdf)}.",
                                               progress=int(70 + (i / len(processed_image_paths_for_this_pdf)) * 10))
                            out_base = (local_processing_temp_dir / f"{sanitized_output_base}_page{i}_ocr").as_posix()
                            cmd = [tesseract_exe, str(img_path), out_base, "-l", lang]
                            try:
                                subprocess.run(cmd, check=True, timeout=tesseract_timeout, creationflags=creation_flags)
                                page_txt = Path(out_base + ".txt").read_text(encoding="utf-8", errors="ignore")
                                texts.append(page_txt)
                            except subprocess.TimeoutExpired:
                                debug_log(rule_id, file_name, f"Tesseract timeout na página {i}.")
                                return False, None
                            except subprocess.CalledProcessError as e:
                                debug_log(rule_id, file_name, f"Erro Tesseract na página {i}: {e}")
                                return False, None

                        # Empacota todo o texto em um único RTF simples
                        def _escape_rtf(s: str) -> str:
                            # Escapa barras, chaves e converte quebras de linha em \par
                            return s.replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}").replace("\r",
                                                                                                           "").replace(
                                "\n", "\\par\n")

                        # Constrói RTF com fonte Calibri 11pt
                        rtf_header = r"{\rtf1\ansi\deff0\nouicompat{\fonttbl{\f0\fnil\fcharset0 Calibri;}}{\pard\sa200\sl276\slmult1\f0\fs22 "
                        rtf_footer = r"}\par}"

                        escaped_text = _escape_rtf("\n".join(texts))

                        # Constrói RTF com Unicode usando \uN? e \uc1 para compatibilidade ampla
                        def _to_rtf_unicode(s: str) -> str:
                            # Normaliza quebras de linha
                            s = s.replace("\r\n", "\n").replace("\r", "\n")
                            out_parts = []
                            for ch in s:
                                code = ord(ch)
                                if ch == "\\":
                                    out_parts.append(r"\\\\")
                                elif ch == "{":
                                    out_parts.append(r"\{")
                                elif ch == "}":
                                    out_parts.append(r"\}")
                                elif ch == "\n":
                                    out_parts.append(r"\par" + "\n")
                                else:
                                    # ASCII imprimível seguro
                                    if 32 <= code <= 126:
                                        out_parts.append(ch)
                                    else:
                                        # RTF usa signed 16-bit para \uN
                                        if code > 0x7FFF:
                                            code_signed = code - 0x10000
                                        else:
                                            code_signed = code
                                        out_parts.append(fr"\u{code_signed}?")
                            return "".join(out_parts)

                        rtf_header = (
                            r"{\rtf1\ansi\ansicpg1252\deff0\uc1"
                            r"{\fonttbl{\f0\fnil\fcharset0 Calibri;}}"
                            r"{\pard\sa200\sl276\slmult1\f0\fs22 "
                        )
                        rtf_footer = r"}\par}"

                        escaped_text = _to_rtf_unicode("\n".join(texts))
                        final_output_local_path = local_processing_temp_dir / f"{sanitized_output_base}.rtf"
                        with open(final_output_local_path, "w", encoding="utf-8") as _rtf_out:
                            _rtf_out.write(rtf_header + escaped_text + rtf_footer)

                    else:
                        debug_log(rule_id, file_name, f"Lógica de rasterização não cobriu o formato: {output_format}")
                        return False, None
                except Exception as e_nonpdf:
                    debug_log(rule_id, file_name,
                              f"Falha ao processar formato {output_format} a partir de páginas rasterizadas: {e_nonpdf}")
                    return False, None

        update_file_status(rule_id, file_name, "Concluindo", progress=95)

        if final_output_local_path and final_output_local_path.exists():
            # --- PONTO DE INTEGRAÇÃO DO OCR ZONAL ---
            final_output_local_path = handle_zonal_ocr_renaming(rule_id, rule_options, original_input_path,
                                                                final_output_local_path)
            # --- FIM DA INTEGRAÇÃO ---

            # --- Assinatura PDF (se aplicável) ---
            try:
                debug_log(rule_id, file_name, "Iniciando etapa de assinatura (PDF - entrada é PDF).")
                new_signed_path = maybe_sign_pdf(
                    rule_id, file_name, rule_options,
                    final_output_local_path,
                    local_processing_temp_dir
                )
                if _pdf_has_signature(new_signed_path if new_signed_path else final_output_local_path):
                    debug_log(rule_id, file_name,
                              f"Arquivo assinado gerado (verificado por ByteRange): {(new_signed_path if new_signed_path else final_output_local_path).name}")
                    final_output_local_path = new_signed_path if new_signed_path else final_output_local_path
                else:
                    debug_log(rule_id, file_name, "Assinatura NÃO detectada no PDF resultante.")
                    final_output_local_path = new_signed_path
            except Exception as _e_sign:
                debug_log(rule_id, file_name, f"Erro ao executar assinatura: {_e_sign}")
            final_output_network_path = output_dir / final_output_local_path.name
            debug_log(rule_id, file_name,
                      f"Movendo arquivo processado de {final_output_local_path} para rede: {final_output_network_path}")
            robust_move(str(final_output_local_path), str(final_output_network_path))
            update_file_status(rule_id, file_name, "Concluído", progress=100)
            debug_log(rule_id, file_name,
                      f"Arquivo PDF processado com sucesso. Saída na rede: {final_output_network_path}")
            return True, final_output_network_path
        else:
            debug_log(rule_id, file_name,
                      f"Processamento de PDF concluído, mas arquivo de saída local esperado ({final_output_local_path}) não foi encontrado.")
            update_file_status(rule_id, file_name, "Falha: Arquivo de saída ausente", progress=0)
            return False, None

    except Exception as e:
        debug_log(rule_id, file_name, f"Erro crítico durante processamento de PDF: {e.__class__.__name__}: {e}")
        logging.exception(f"Erro crítico durante processamento de PDF para '{file_name}'")
        update_file_status(rule_id, file_name, "Falha Crítica", progress=0)
        return False, None
    finally:
        if local_processing_temp_dir and local_processing_temp_dir.exists():
            try:
                shutil.rmtree(local_processing_temp_dir)
                debug_log(rule_id, file_name, f"Diretório temporário local removido: {local_processing_temp_dir}")
            except Exception as e_rm:
                debug_log(rule_id, file_name,
                          f"Falha ao remover diretório temporário local {local_processing_temp_dir}: {e_rm}")


# --- Engine Management ---
observers = {}


def is_file_stable(path: Path, wait_time_seconds: int = 3, check_interval_seconds: float = 0.5) -> bool:
    previous_size = -1
    elapsed_time = 0.0
    debug_log("SYSTEM", path.name, f"Verificando estabilidade do arquivo: {path}")
    if not path.exists():
        debug_log("SYSTEM", path.name, f"Arquivo não existe durante a verificação de estabilidade: {path}")
        return False
    time.sleep(check_interval_seconds)
    while elapsed_time < wait_time_seconds:
        try:
            current_size = path.stat().st_size
        except FileNotFoundError:
            debug_log("SYSTEM", path.name,
                      f"Arquivo '{path.name}' não encontrado durante a verificação de estabilidade, pode ter sido movido ou excluído.")
            return False
        except Exception as e:
            debug_log("SYSTEM", path.name,
                      f"Erro ao verificar tamanho do arquivo para '{path.name}' durante a estabilidade: {e}")
            time.sleep(check_interval_seconds)
            continue
        if current_size == previous_size:
            elapsed_time += check_interval_seconds
            debug_log("SYSTEM", path.name,
                      f"Tamanho estável: {current_size} bytes. Tempo decorrido: {elapsed_time:.1f}s/{wait_time_seconds}s")
        else:
            elapsed_time = 0.0
            previous_size = current_size
            debug_log("SYSTEM", path.name,
                      f"Tamanho alterado: {current_size} bytes. Reiniciando contagem de estabilidade.")
        time.sleep(check_interval_seconds)
    debug_log("SYSTEM", path.name, f"Arquivo estável após {wait_time_seconds} segundos: {path}")
    return True


# === Single-flight lock helpers (per-input file) ===
LOCK_EXT = ".octec.lock"
LOCK_TTL_SECONDS = 2 * 60 * 60  # 2h to guard against stale locks


def _lock_path_for(path: Path) -> Path:
    try:
        return path.with_suffix(path.suffix + LOCK_EXT)
    except Exception:
        # Fallback: append extension as text
        return Path(str(path) + LOCK_EXT)


def acquire_file_lock(path: Path, ttl_seconds: int = LOCK_TTL_SECONDS) -> bool:
    """Try to obtain an exclusive lock for 'path'. If a stale lock exists (older than ttl), reclaim it.
    Returns True if lock is held by this process now, False otherwise."""
    import os, time
    lp = _lock_path_for(path)
    now = time.time()
    try:
        fd = os.open(str(lp), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.close(fd)
        try:
            lp.write_text(str(int(now)))
        except Exception:
            pass
        return True
    except FileExistsError:
        # Check for stale lock
        try:
            mtime = lp.stat().st_mtime
        except FileNotFoundError:
            # raced with remover; try once more
            try:
                fd = os.open(str(lp), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.close(fd)
                return True
            except Exception:
                return False
        except Exception:
            return False
        if now - mtime > ttl_seconds:
            try:
                lp.unlink()
            except Exception:
                return False
            # retry once
            try:
                fd = os.open(str(lp), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.close(fd)
                try:
                    lp.write_text(str(int(now)))
                except Exception:
                    pass
                return True
            except Exception:
                return False
        return False
    except Exception:
        return False


def release_file_lock(path: Path) -> None:
    try:
        _lock_path_for(path).unlink()
    except FileNotFoundError:
        pass
    except Exception:
        pass


# === Deletion tolerant to races/locks ===
def safe_unlink(path: Path, retries: int = 3, delay: float = 0.3) -> bool:
    import time
    for _ in range(max(1, retries)):
        try:
            # Python 3.8+: missing_ok
            path.unlink(missing_ok=True)  # type: ignore
            return True
        except TypeError:
            # For Python <3.8 environments
            try:
                path.unlink()
                return True
            except FileNotFoundError:
                return True
        except PermissionError:
            time.sleep(delay)
            continue
        except FileNotFoundError:
            return True
        except Exception:
            return False
    return False


class RuleEventHandler(FileSystemEventHandler):
    def __init__(self, rule_id: str, file_extensions: list) -> None:
        super().__init__()
        self.rule_id = rule_id
        self.file_extensions = [ext.lower() for ext in file_extensions]
        self.last_modified_times = {}

    def _is_valid_file(self, event_path: str) -> bool:
        path = Path(event_path)
        if path.is_dir():
            return False
        if path.parent.name.startswith(".octec_proc_temp_"):
            logging.debug(f"[{self.rule_id}] Ignorando arquivo em diretório temporário de processamento: {path.name}")
            return False
        return path.suffix.lower() in self.file_extensions

    def on_created(self, event):
        if not engine_running.is_set() or rules_data.get(self.rule_id, {}).get("paused", False):
            return
        if event.is_directory:
            return
        file_path = Path(event.src_path)
        if self._is_valid_file(str(file_path)):
            logging.info(f"[{self.rule_id}] Novo arquivo detectado: {file_path}")
            enqueue_file(self.rule_id, file_path)
        else:
            logging.debug(f"[{self.rule_id}] Arquivo criado não é válido para processamento: {file_path.name}")

    def on_modified(self, event):
        if not engine_running.is_set() or rules_data.get(self.rule_id, {}).get("paused", False):
            return
        if event.is_directory:
            return
        file_path = Path(event.src_path)
        if self._is_valid_file(str(file_path)):
            # Debounce simples por mtime e tamanho, evita enxurrada de enfileiramentos
            try:
                stat = file_path.stat()
                key = (file_path,)
                prev = self.last_modified_times.get(key)
                curr = (stat.st_mtime, stat.st_size)
                if prev == curr:
                    return
                self.last_modified_times[key] = curr
            except Exception:
                pass
            logging.info(f"[{self.rule_id}] Arquivo modificado detectado: {file_path}")
            enqueue_file(self.rule_id, file_path)
        else:
            logging.debug(f"[{self.rule_id}] Arquivo modificado não é válido para processamento: {file_path.name}")

    def clear_entry(self, file_path: Path) -> None:
        key = (file_path,)
        if key in self.last_modified_times:
            del self.last_modified_times[key]
            logging.debug(f"[{self.rule_id}] Removido '{file_path.name}' do rastreamento de RuleEventHandler.")


class RulesRootWatcher(FileSystemEventHandler):
    """Observa RULES_INPUT_BASE_DIR e remove a regra quando a pasta de entrada é apagada manualmente."""

    def on_deleted(self, event):
        try:
            if event.is_directory:
                deleted_path = Path(event.src_path)
                # Verifica se corresponde a alguma regra
                for rid, cfg in list(rules_data.items()):
                    if Path(cfg.get('input', '')) == deleted_path:
                        logging.info(f"Pasta de entrada removida no Explorer: {deleted_path}. Removendo regra '{rid}'.")
                        stop_watching_thread_for_rule(rid)
                        with files_status_lock:
                            files_status.pop(rid, None)
                        rules_data.pop(rid, None)
                        save_config(rules_data)
                        break
        except Exception as e:
            logging.error(f"Erro no RulesRootWatcher.on_deleted: {e}")


def initialize_engine() -> None:
    logging.info("Inicializando Motor OCTec...")
    engine_running.set()
    # Inicia o executador de processos para paralelizar o processamento de arquivos.
    global process_executor
    try:
        detected_cpus = os.cpu_count() or 1
    except Exception:
        detected_cpus = 1
    # Cria o executor com tantos processos quanto núcleos disponíveis
    process_executor = ProcessPoolExecutor(max_workers=max(1, detected_cpus))
    # Define o número de threads trabalhadoras com base no número de CPUs disponíveis.
    # Ao escalar os workers proporcionalmente aos núcleos, permitimos que múltiplos
    # arquivos sejam processados em paralelo sem sobrecarregar o sistema. Um mínimo
    # de 2 workers é mantido para garantir algum paralelismo em máquinas com apenas
    # um núcleo lógico.
    num_workers = max(2, detected_cpus)
    for i in range(num_workers):
        worker_thread = threading.Thread(target=file_worker, name=f"FileWorker-{i + 1}")
        worker_thread.daemon = True
        worker_threads.append(worker_thread)
        worker_thread.start()
    logging.info(f"Iniciando {num_workers} threads trabalhadoras de arquivo...")
    global rules_data
    rules_data = load_config()
    logging.info(f"Iniciando threads de monitoramento para {len(rules_data)} regras...")
    for rule_id, rule_config in rules_data.items():
        if not rule_config.get("paused", False):
            start_watching_thread_for_rule(rule_id)
        else:
            logging.info(f"Regra '{rule_id}' está pausada, observador não iniciado.")
    # Inicia observador da raiz para detectar deleções de pastas de entrada
    global _root_rules_observer
    _root_rules_observer = Observer()
    _root_rules_observer.schedule(RulesRootWatcher(), str(RULES_INPUT_BASE_DIR), recursive=False)
    _root_rules_observer.daemon = True
    _root_rules_observer.start()
    # Após iniciar observadores para regras e raiz, faça uma varredura inicial
    # para enfileirar quaisquer arquivos que já existam nos diretórios de entrada.
    try:
        scan_and_enqueue_existing_files()
    except Exception as e_scan:
        logging.error(f"Falha ao escanear arquivos existentes na inicialização: {e_scan}")

    # Inicia o monitor de auto-retomada de regras pausadas por caminhos ausentes
    try:
        global _auto_resume_thread, _auto_resume_stop_event
        _auto_resume_stop_event.clear()
        _auto_resume_thread = threading.Thread(target=_auto_resume_worker, name="AutoResume", daemon=True)
        _auto_resume_thread.start()
        logging.info("Monitor de auto-retomada iniciado.")
    except Exception as e:
        logging.warning(f"Falha ao iniciar monitor de auto-retomada: {e}")


logging.info("Motor OCTec inicializado.")


def shutdown_engine() -> None:
    logging.info("Encerrando Motor OCTec...")

    # Finaliza o monitor de auto-retomada
    try:
        global _auto_resume_thread, _auto_resume_stop_event
        if ' _auto_resume_stop_event' in globals():
            pass  # só para evitar linter
        _auto_resume_stop_event.set()
        thr = globals().get('_auto_resume_thread')
        if thr:
            thr.join(timeout=5)
            if thr.is_alive():
                logging.warning("Thread 'AutoResume' não parou graciosamente.")
    except Exception as e:
        logging.warning(f"Falha ao encerrar monitor de auto-retomada: {e}")
    engine_running.clear()
    # Para o observador da raiz
    global _root_rules_observer
    try:
        if '_root_rules_observer' in globals() and _root_rules_observer:
            _root_rules_observer.stop()
            _root_rules_observer.join(timeout=5)
    except Exception as e:
        logging.warning(f"Falha ao parar observador raiz: {e}")
    for rule_id, observer in observers.items():
        logging.info(f"Parando observador para regra {rule_id}...")
        observer.stop()
    for rule_id, observer in observers.items():
        observer.join(timeout=5)
    # Encerra o executor de processos
        # (Removido) Checagem redundante de 'observer' fora do loop de encerramento.
        try:
            # Encerra sem bloquear. Em versões de Python anteriores a 3.11, o argumento
            # cancel_futures não está disponível. Por isso, chamamos shutdown
            # simplesmente com wait=False.
            process_executor.shutdown(wait=False)
        except Exception as e:
            logging.warning(f"Falha ao encerrar executor de processos: {e}")
        if observer.is_alive():
            logging.warning(f"Observador para regra {rule_id} não parou graciosamente.")
    observers.clear()
    rule_event_handlers.clear()
    for _ in worker_threads:
        file_queue.put((float('inf'), next(_file_queue_counter), None, None))
    for worker_thread in worker_threads:
        worker_thread.join(timeout=10)
        if worker_thread.is_alive():
            logging.warning(f"Thread '{worker_thread.name}' não parou graciosamente.")
    worker_threads.clear()
    while not file_queue.empty():
        try:
            file_queue.get_nowait()
            file_queue.task_done()
        except queue.Empty:
            pass
    cleanup_all_managed_persistent_connections()
    logging.info("Motor OCTec encerrado.")


def start_watching_thread_for_rule(rule_id: str) -> None:
    if rule_id not in rules_data:
        logging.warning(f"Não é possível iniciar o observador para regra {rule_id}: regra não encontrada.")
        return
    rule_config = rules_data[rule_id]
    input_path = Path(rule_config["input"])

    # Se a pasta de entrada for UNC, garante conexão antes de observar
    try:
        if str(input_path).startswith('\\\\'):
            if not ensure_persistent_connection(str(input_path)):
                logging.error(
                    f"Não foi possível conectar à pasta de entrada UNC '{input_path}' para a regra '{rule_id}'. Observador não iniciado.")
                return
    except Exception as e:
        logging.error(f"Erro ao tentar conectar na pasta de entrada UNC '{input_path}': {e}")
    if rule_id in observers and observers[rule_id].is_alive():
        logging.info(f"Observador para regra '{rule_id}' já está em execução, ignorando início duplicado.")
        return
    if not input_path.exists():
        try:
            input_path.mkdir(parents=True, exist_ok=True)
            logging.info(f"Pasta de entrada para regra '{rule_id}' criada: {input_path}")
        except Exception as e:
            logging.error(f"Não foi possível criar pasta de entrada para regra '{rule_id}' em {input_path}: {e}")
            return
    accepted_extensions = [".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"]
    event_handler = RuleEventHandler(rule_id, accepted_extensions)
    observer = Observer()
    observer.schedule(event_handler, str(input_path), recursive=False)
    observer.daemon = True
    observer.start()
    observers[rule_id] = observer
    rule_event_handlers[rule_id] = event_handler
    logging.info(f"Observador iniciado para regra {rule_id}.")


def stop_watching_thread_for_rule(rule_id: str) -> None:
    if rule_id in observers:
        logging.info(f"Parando observador para regra {rule_id}...")
        observer = observers.pop(rule_id)
        observer.stop()
        observer.join(timeout=5)
        if observer.is_alive():
            logging.warning(f"Observador para regra {rule_id} não parou graciosamente após 5s.")
    else:
        logging.info(f"Nenhum observador ativo para a regra {rule_id} para parar.")
    if rule_id in rule_event_handlers:
        del rule_event_handlers[rule_id]


# --- Threads de Trabalho ---
def file_worker():
    thread_name = threading.current_thread().name
    logging.info(f"{thread_name}: Trabalhador de arquivo iniciado.")
    while True:
        rule_id = None
        file_path_str = None
        re_enqueue = False
        try:
            rule_id, file_path_str = dequeue_file(timeout=1)
            logging.debug(
                f"{thread_name}: Item retirado da fila. Rule ID: '{rule_id}', File: '{file_path_str}'")
            if rule_id is None and file_path_str is None:
                file_queue.task_done()
                break
            original_input_path = Path(file_path_str)
            # Single-flight: try to acquire per-file lock; skip if someone else is processing
            try:
                if not acquire_file_lock(original_input_path):
                    continue
            except Exception:
                # Conservador: se não conseguimos criar/ver lock, seguimos sem travar
                pass
            file_name = original_input_path.name
            logging.info(f"{thread_name}: Processando arquivo '{file_name}' para regra '{rule_id}'...")

            if not original_input_path.exists():
                debug_log(rule_id, file_name, "Arquivo de entrada não existe. Pulando processamento.")
                update_file_status(rule_id, file_name, "Falha: Arquivo Inexistente", progress=0)
                file_queue.task_done()
                continue

            if not is_file_stable(original_input_path):
                debug_log(rule_id, file_name, "Arquivo de entrada não está estável, reenfileirando.")
                update_file_status(rule_id, file_name, "Reenfileirado: Arquivo Instável", progress=0)
                _clear_processing_state(rule_id, original_input_path)
                enqueue_file(rule_id, Path(file_path_str))
                re_enqueue = True
                time.sleep(5)
                continue

            update_file_status(rule_id, file_name, "Processando", progress=0,
                               size=original_input_path.stat().st_size if original_input_path.exists() else 0)
            debug_log(rule_id, file_name, "Iniciando processamento do arquivo.")
            success = False
            output_path = None
            if rule_id not in rules_data:
                debug_log(rule_id, file_name, f"ID da Regra {rule_id} não encontrado. Pulando arquivo.")
                update_file_status(rule_id, file_name, "Falha: Regra Ausente", progress=0)
            else:
                rule_config = rules_data[rule_id]
                output_dir_str = rule_config["output"]
                if output_dir_str.startswith("\\\\"):
                    debug_log(rule_id, file_name, f"Verificando conexão de rede para pasta de saída: {output_dir_str}.")
                    if not ensure_persistent_connection(output_dir_str):
                        debug_log(rule_id, file_name,
                                  f"Falha ao estabelecer conexão de rede para {output_dir_str}. Reenfileirando.")
                        update_file_status(rule_id, file_name, "Falha: Rede", progress=0)
                        _clear_processing_state(rule_id, original_input_path)
                        enqueue_file(rule_id, Path(file_path_str))
                        re_enqueue = True
                        time.sleep(15)
                        continue
                # Processa o arquivo em um processo separado para permitir paralelismo real
                file_extension = original_input_path.suffix.lower()
                if file_extension in [".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"]:
                    # Certifica-se de que o executor de processos está inicializado
                    if process_executor is None:
                        # Se por algum motivo não foi inicializado, processa de forma síncrona
                        if file_extension == ".pdf":
                            success, output_path = process_pdf_file(rule_id, original_input_path, rule_config)
                        else:
                            success, output_path = process_image_file(rule_id, original_input_path, rule_config)
                    else:
                        try:
                            # Passa uma cópia das opções da regra para evitar efeitos colaterais entre processos
                            rule_opts = rule_config.copy()
                            # Agenda o processamento em um subprocesso
                            future = process_executor.submit(process_file_task, (rule_id, file_path_str, rule_opts))
                            # Indica progresso intermediário enquanto aguarda o término do processo
                            update_file_status(rule_id, file_name, "Processando", progress=50)
                            res_success, res_output_str = future.result()
                            success = res_success
                            output_path = Path(res_output_str) if res_output_str else None
                        except Exception as exec_err:
                            success = False
                            output_path = None
                            debug_log(rule_id, file_name, f"Erro ao enviar tarefa para executor: {exec_err}")
                else:
                    debug_log(rule_id, file_name,
                              f"Tipo de arquivo não suportado: {file_extension}. Arquivo será ignorado, não removido.")
                    logging.info(
                        f"{thread_name}: Tipo de arquivo não suportado {file_extension} para {file_name}. Ignorando.")
                    update_file_status(rule_id, file_name, "Ignorado: Tipo Não Suportado", progress=100)
                    success = False

            if success:
                # Remove o estado de processamento para permitir que o mesmo arquivo seja reprocessado futuramente
                _clear_processing_state(rule_id, original_input_path)
                # Tenta descobrir o tamanho real do arquivo de saída (útil para a GUI)
                out_size = None
                if output_path:
                    try:
                        out_size = Path(output_path).stat().st_size
                    except Exception:
                        out_size = None
                if out_size is not None:
                    update_file_status(rule_id, file_name, "Concluído", progress=100, size=out_size)
                else:
                    update_file_status(rule_id, file_name, "Concluído", progress=100)
                debug_log(rule_id, file_name, f"Processamento do arquivo concluído com sucesso. Saída: {output_path}")
                try:
                    # Só apagar o arquivo de entrada se nenhuma outra regra tiver ele na fila ou em processamento
                    key_norm = _norm_path_key(str(original_input_path))
                    should_delete = False
                    with queued_files_lock:
                        other_processing = any(k[1] == key_norm and k[0] != rule_id for k in processing_files_set)
                        other_queued = any(k[1] == key_norm and k[0] != rule_id for k in queued_files_set)
                        if not other_processing and not other_queued:
                            should_delete = True
                    if should_delete:
                        safe_unlink(original_input_path)

                        debug_log(rule_id, file_name, f"Arquivo de entrada original {original_input_path} removido.")
                    else:
                        debug_log(rule_id, file_name,
                                  f"Entrada mantida: outra regra ainda tem pendência sobre {original_input_path}.")
                except Exception as e_del_orig:
                    debug_log(rule_id, file_name,
                              f"Erro ao remover arquivo de entrada original {original_input_path}: {e_del_orig}")
                # Independente de deletar ou não, limpa rastreio para permitir novo enfileiramento
                if rule_id in rule_event_handlers:
                    rule_event_handlers[rule_id].clear_entry(original_input_path)
            elif not re_enqueue:
                # Falha definitiva (não reenqueue). Marca status, limpa estado e libera rastreio
                with files_status_lock:
                    current_status = files_status.get(rule_id, {}).get(file_name, {}).get("status", "Falha")
                if "Falha" not in current_status and "Ignorado" not in current_status:
                    update_file_status(rule_id, file_name, "Falha Processamento", progress=0)
                    debug_log(rule_id, file_name, "Falha no processamento do arquivo.")
                _clear_processing_state(rule_id, original_input_path)
                if rule_id in rule_event_handlers:
                    rule_event_handlers[rule_id].clear_entry(original_input_path)
        except queue.Empty:
            if not engine_running.is_set() and file_queue.empty(): break
            continue
        except Exception as e:
            current_item_details = f"Regra: {rule_id if rule_id else 'N/A'}, Arquivo: {file_path_str if file_path_str else 'N/A'}"
            logging.critical(f"{thread_name}: Erro crítico no loop do trabalhador ({current_item_details}): {e}",
                             exc_info=True)
            if rule_id and file_path_str:
                update_file_status(rule_id, Path(file_path_str).name, "Falha Crítica Interna", progress=0)
            re_enqueue = False
        finally:
            # Release file lock if held
            try:
                if file_path_str:
                    release_file_lock(Path(file_path_str))
            except Exception:
                pass
            if not re_enqueue and rule_id is not None:
                file_queue.task_done()
    logging.info(f"{thread_name}: Trabalhador de arquivo parado.")


class FileDetailWidget(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.current_rule_id: str | None = None
        self.file_data_cache = {}

        layout = QVBoxLayout(self)
        self.file_table = QTableWidget()
        self.file_table.setColumnCount(4)
        self.file_table.setHorizontalHeaderLabels(["Arquivo", "Progresso", "Status", "Tamanho (MB)"])
        self.file_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.file_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.file_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.file_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.file_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.file_table.itemSelectionChanged.connect(self.update_logs_for_selected)
        self.file_table.setSortingEnabled(True)
        layout.addWidget(self.file_table)

        self.tab_widget = QTabWidget()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFontFamily("Consolas")
        self.log_text.setLineWrapMode(QTextEdit.WidgetWidth)
        self.tab_widget.addTab(self.log_text, "Logs do Arquivo")
        layout.addWidget(self.tab_widget)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.refresh_files_display)
        self.timer.start(1500)

    def populate(self, rule_id: str | None) -> None:
        self.current_rule_id = rule_id
        self.file_table.setSortingEnabled(False)
        self.file_table.clearContents()
        self.file_table.setRowCount(0)
        self.log_text.clear()
        self.file_data_cache = {}
        if rule_id:
            self.refresh_files_display()
        self.file_table.setSortingEnabled(True)

    def refresh_files_display(self) -> None:
        if not self.current_rule_id:
            self.file_table.setRowCount(0)
            self.log_text.clear()
            return

        rule_file_data = {}
        with files_status_lock:
            if self.current_rule_id not in files_status:
                self.file_table.setRowCount(0)
                self.log_text.clear()
                return

            rule_files_dict = files_status[self.current_rule_id]
            if len(rule_files_dict) > MAX_FILES_IN_GUI:
                sorted_keys = sorted(rule_files_dict, key=lambda k: rule_files_dict[k].get('entry_time', 0))
                for key_to_remove in sorted_keys[:-MAX_FILES_IN_GUI]:
                    del rule_files_dict[key_to_remove]
            rule_file_data = rule_files_dict.copy()

        self.file_table.setSortingEnabled(False)
        self.file_table.setUpdatesEnabled(False)

        selected_file_name = None
        current_selection_rows = self.file_table.selectionModel().selectedRows()
        if current_selection_rows:
            try:
                selected_file_name = self.file_table.item(current_selection_rows[0].row(), 0).text()
            except AttributeError:
                pass

        self.file_table.clearContents()
        self.file_table.setRowCount(0)
        self.file_data_cache = {}

        new_file_names_sorted = sorted(rule_file_data.keys())
        self.file_table.setRowCount(len(new_file_names_sorted))

        for row_idx, fname in enumerate(new_file_names_sorted):
            info = rule_file_data[fname]

            item_fname = QTableWidgetItem(fname)
            self.file_table.setItem(row_idx, 0, item_fname)

            bar = QProgressBar()
            bar.setAlignment(Qt.AlignCenter)
            bar.setValue(info.get("progress", 0))
            bar.setFormat("%p%")
            self.file_table.setCellWidget(row_idx, 1, bar)

            item_st = QTableWidgetItem(info.get("status", "Desconhecido"))
            item_st.setTextAlignment(Qt.AlignCenter)
            self.file_table.setItem(row_idx, 2, item_st)

            size_bytes = info.get("size", 0)
            size_mb = size_bytes / (1024 * 1024) if size_bytes else 0
            item_sz = QTableWidgetItem(f"{size_mb:.2f}")
            item_sz.setTextAlignment(Qt.AlignCenter)
            self.file_table.setItem(row_idx, 3, item_sz)

            self.file_data_cache[fname] = info.copy()

        if selected_file_name and selected_file_name in new_file_names_sorted:
            row_to_select = new_file_names_sorted.index(selected_file_name)
            self.file_table.selectRow(row_to_select)
        elif self.file_table.rowCount() > 0:
            self.file_table.selectRow(0)

        self.file_table.setSortingEnabled(True)
        self.file_table.setUpdatesEnabled(True)
        self.update_logs_for_selected()

    def update_logs_for_selected(self) -> None:
        if not self.current_rule_id: self.log_text.clear(); return
        selected_items = self.file_table.selectedItems()
        if not selected_items: self.log_text.clear(); return

        row = selected_items[0].row()
        fname_item = self.file_table.item(row, 0)
        if not fname_item: self.log_text.clear(); return
        fname = fname_item.text()

        with files_status_lock:
            info = files_status.get(self.current_rule_id, {}).get(fname, {})
            logs_to_display = info.get("logs", []).copy()

        if not logs_to_display: self.log_text.clear(); return

        self.log_text.setPlainText("\n".join(logs_to_display))
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())


class NetworkCredentialDialog(QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Gerenciar Credenciais de Rede")
        self.setMinimumWidth(500)
        self.layout = QVBoxLayout(self)

        self.cred_table = QTableWidget()
        self.cred_table.setColumnCount(2)
        self.cred_table.setHorizontalHeaderLabels(["Caminho do Servidor/Compartilhamento (UNC)", "Usuário"])
        self.cred_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.cred_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.cred_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.cred_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.layout.addWidget(self.cred_table)

        form_group = QGroupBox("Adicionar/Editar Credencial")
        form_layout = QGridLayout(form_group)
        form_layout.addWidget(QLabel("Caminho Servidor (UNC):"), 0, 0)
        self.edt_server_path = QLineEdit()
        self.edt_server_path.setPlaceholderText(r"Ex: \\servidor\compartilhamento")
        form_layout.addWidget(self.edt_server_path, 0, 1)
        form_layout.addWidget(QLabel("Usuário:"), 1, 0)
        self.edt_user = QLineEdit()
        self.edt_user.setPlaceholderText("Ex: DOMINIO\\usuario ou usuario")
        form_layout.addWidget(self.edt_user, 1, 1)
        form_layout.addWidget(QLabel("Senha:"), 2, 0)
        self.edt_pass = QLineEdit()
        self.edt_pass.setEchoMode(QLineEdit.Password)
        form_layout.addWidget(self.edt_pass, 2, 1)
        self.layout.addWidget(form_group)

        btn_box = QHBoxLayout()
        self.btn_add = QPushButton(self.style().standardIcon(QStyle.SP_DialogSaveButton), " Salvar Credencial")
        self.btn_add.clicked.connect(self.save_credential_entry)
        self.btn_remove = QPushButton(self.style().standardIcon(QStyle.SP_TrashIcon), " Remover Selecionada")
        self.btn_remove.clicked.connect(self.remove_credential_entry)
        btn_box.addStretch()
        btn_box.addWidget(self.btn_add)
        btn_box.addWidget(self.btn_remove)
        self.layout.addLayout(btn_box)

        self.load_credentials_to_table()
        self.cred_table.itemClicked.connect(self.populate_form_from_table)

    def load_credentials_to_table(self) -> None:
        self.cred_table.setSortingEnabled(False)
        self.cred_table.setRowCount(0)
        creds = get_net_credentials()
        self.cred_table.setRowCount(len(creds))
        for row, (server_path_key, data) in enumerate(sorted(creds.items())):
            display_path = str(Path(server_path_key))
            self.cred_table.setItem(row, 0, QTableWidgetItem(display_path))
            self.cred_table.setItem(row, 1, QTableWidgetItem(data.get("username", "")))
        self.cred_table.resizeColumnsToContents()
        self.cred_table.setSortingEnabled(True)

    def populate_form_from_table(self, item: QTableWidgetItem) -> None:
        row = item.row()
        path_item = self.cred_table.item(row, 0)
        user_item = self.cred_table.item(row, 1)
        if path_item: self.edt_server_path.setText(path_item.text())
        if user_item: self.edt_user.setText(user_item.text())
        self.edt_pass.clear()
        self.edt_pass.setPlaceholderText("Deixe em branco para manter a senha existente")

    def save_credential_entry(self) -> None:
        server_path_raw = self.edt_server_path.text().strip()
        username = self.edt_user.text().strip()
        password = self.edt_pass.text()
        if not server_path_raw or not server_path_raw.startswith("\\\\"):
            QMessageBox.warning(self, "Erro de Validação",
                                "Caminho do Servidor UNC é obrigatório e deve iniciar com \\\\.")
            return
        if not username:
            QMessageBox.warning(self, "Erro de Validação", "Usuário é obrigatório.")
            return
        try:
            p = Path(server_path_raw)
            unc_drive = p.drive
            if not unc_drive or len(unc_drive.split('\\')) < 4:
                QMessageBox.warning(self, "Erro de Validação",
                                    r"Formato do caminho UNC inválido. Use \\servidor\compartilhamento.")
                return
            unc_drive_str = str(Path(unc_drive))
            normalized_server_path_key = Path(unc_drive).as_posix().lower()
        except Exception as e:
            logging.error(f"Erro inesperado ao validar o caminho UNC '{server_path_raw}': {e}")
            QMessageBox.warning(self, "Erro de Validação", r"Formato do caminho UNC inválido ou erro inesperado.")
            return

        current_creds = get_net_credentials()
        is_new_entry = normalized_server_path_key not in current_creds
        if is_new_entry and not password:
            QMessageBox.warning(self, "Erro de Validação", "Senha é obrigatória para uma nova credencial.")
            return

        password_to_save = password
        if not password and not is_new_entry:
            password_to_save = current_creds[normalized_server_path_key].get("password")
        add_or_update_net_credential(unc_drive_str, username, password_to_save)
        QMessageBox.information(self, "Sucesso", f"Credencial para '{unc_drive_str}' salva/atualizada.")
        self.load_credentials_to_table()
        self.edt_server_path.clear();
        self.edt_user.clear();
        self.edt_pass.clear()
        self.edt_pass.setPlaceholderText("")

    def remove_credential_entry(self) -> None:
        selected_rows = self.cred_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "Remover Credencial", "Nenhuma credencial selecionada.")
            return
        path_to_remove_item = self.cred_table.item(selected_rows[0].row(), 0)
        if not path_to_remove_item: return
        path_to_remove = path_to_remove_item.text()
        if QMessageBox.question(self, "Confirmar Remoção",
                                f"Tem certeza que deseja remover a credencial para '{path_to_remove}'?",
                                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.Yes:
            remove_net_credential(path_to_remove)
            self.load_credentials_to_table()
            self.edt_server_path.clear();
            self.edt_user.clear();
            self.edt_pass.clear()
            QMessageBox.information(self, "Sucesso", f"Credencial para '{path_to_remove}' removida.")


class RuleDialog(QDialog):
    def __init__(self, parent=None, rule_id: str | None = None, existing_data: dict | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Editar Regra [{rule_id}]" if rule_id else "Nova Regra")
        self.setMinimumWidth(700)
        self.rule_id_original = rule_id
        self.data = existing_data.copy() if existing_data else {}

        layout = QVBoxLayout(self)
        main_tab_widget = QTabWidget()

        # --- Aba Geral ---
        tab_general = QWidget()
        general_layout = QGridLayout(tab_general)
        general_layout.addWidget(
            QLabel("<b>ID da Regra:</b><br><small>(Nome único para a regra; usado para pasta de entrada)</small>"), 0,
            0, Qt.AlignTop)
        self.edt_rule_id = QLineEdit()
        self.edt_rule_id.setToolTip("Este será o nome da pasta criada dentro de 'REGRAS'. Ex: 'Contratos_SCAN'")
        general_layout.addWidget(self.edt_rule_id, 0, 1)
        self.lbl_input_dir_display = QLabel()
        self.lbl_input_dir_display.setTextInteractionFlags(Qt.TextSelectableByMouse)
        general_layout.addWidget(QLabel("Pasta de Entrada Real:"), 1, 0)
        general_layout.addWidget(self.lbl_input_dir_display, 1, 1)
        self.edt_rule_id.textChanged.connect(self._update_input_dir_display)
        # Campo opcional de Pasta de Entrada (customizada). Se vazio, usa o padrão baseado no ID.
        general_layout.addWidget(QLabel(
            "<b>Pasta de Entrada (Opcional):</b><br><small>Deixe em branco para usar a pasta padrão baseada no ID. Aceita caminhos locais ou UNC.</small>"),
            6, 0, Qt.AlignTop)
        self.edt_input_dir = QLineEdit()
        btn_sel_in = QPushButton(self.style().standardIcon(QStyle.SP_DirOpenIcon), " Selecionar Pasta")
        btn_sel_in.clicked.connect(self.select_input_dir)
        btn_test_in = QPushButton("Testar Pasta de Entrada")
        btn_test_in.clicked.connect(self.test_input_path)
        input_dir_layout = QHBoxLayout()
        input_dir_layout.addWidget(self.edt_input_dir)
        input_dir_layout.addWidget(btn_sel_in)
        general_layout.addLayout(input_dir_layout, 6, 1)
        general_layout.addWidget(btn_test_in, 7, 1, Qt.AlignRight)

        general_layout.addWidget(
            QLabel("<b>Pasta de Destino (Output):</b><br><small>(Caminho local ou de rede UNC)</small>"), 2, 0,
            Qt.AlignTop)
        self.edt_output_dir = QLineEdit()
        btn_sel_out = QPushButton(self.style().standardIcon(QStyle.SP_DirOpenIcon), " Selecionar Pasta")
        btn_sel_out.clicked.connect(self.select_output_dir)
        btn_test_out = QPushButton("Testar Caminho de Destino")
        btn_test_out.clicked.connect(self.test_output_path)
        output_dir_layout = QHBoxLayout()
        output_dir_layout.addWidget(self.edt_output_dir);
        output_dir_layout.addWidget(btn_sel_out)
        general_layout.addLayout(output_dir_layout, 2, 1)
        general_layout.addWidget(btn_test_out, 3, 1, Qt.AlignRight)
        main_tab_widget.addTab(tab_general, " Geral ")

        # --- Aba Processamento de Imagem ---
        tab_processing = QWidget()
        processing_layout = QGridLayout(tab_processing)
        self.chk_orientation = QCheckBox("Auto Orientação e Alinhamento (analisa conteúdo da página)")
        self.chk_orientation.setToolTip(
            "Detecta e corrige a rotação (90/180/270 graus)\ne endireita pequenas inclinações. Essencial para scans.")
        processing_layout.addWidget(self.chk_orientation, 0, 0, 1, 2)

        self.chk_framing = QCheckBox("Auto Framing (adicionar borda branca)")
        processing_layout.addWidget(self.chk_framing, 1, 0)

        self.chk_noise = QCheckBox("Redução de Ruído")
        self.chk_noise.setToolTip(
            "Aplica algoritmo de redução de ruído para melhorar a qualidade da imagem.")
        processing_layout.addWidget(self.chk_noise, 2, 0, 1, 2)

        self.chk_binar = QCheckBox("Binarização (converter para P&B, limiar 128)")
        processing_layout.addWidget(self.chk_binar, 3, 0, 1, 2)
        main_tab_widget.addTab(tab_processing, " Processamento de Imagem ")

        # --- Aba Formato de Saída e OCR ---
        tab_output = QWidget()
        output_layout = QGridLayout(tab_output)
        output_layout.addWidget(QLabel("Formato de Saída Final:"), 0, 0)
        self.cmb_output_format = QComboBox()
        self.cmb_output_format.addItems([
            "PDF Pesquisável (com reconhecimento de texto)", "PDF (Somente Imagem)",
            "PDF/A (com reconhecimento de texto)",
            "JPEG (Página única ou Zip de páginas para PDF)", "TIFF (Página única ou Zip de páginas para PDF)",
            "Texto (TXT)", "RTF (Texto Formatado)", "DOCX (Texto Editável)",

            "XLSX (Planilha: 1 linha por página)", "PPTX (1 slide por página)",
            "Checksum (SHA-256 do original)"
        ])
        self.cmb_output_format.currentIndexChanged.connect(self._update_options_for_format)
        output_layout.addWidget(self.cmb_output_format, 0, 1)

        self.lbl_lang = QLabel("Idioma para reconhecimento de texto:")
        output_layout.addWidget(self.lbl_lang, 1, 0)
        self.cmb_lang = QComboBox()
        self.cmb_lang.addItems(
            ["eng (Inglês)", "por (Português)", "spa (Espanhol)", "fra (Francês)", "deu (Alemão)", "ita (Italiano)",
             "osd (Detecção Auto Script/Orient.)"])
        output_layout.addWidget(self.cmb_lang, 1, 1)

        self.chk_compact = QCheckBox("Compactar Saída (se aplicável, ex: PDF deflate, otimizar imagens)")
        output_layout.addWidget(self.chk_compact, 2, 0, 1, 2)

        self.group_jpeg_opts = QGroupBox("Opções JPEG")
        jpeg_layout = QGridLayout(self.group_jpeg_opts)
        jpeg_layout.addWidget(QLabel("Qualidade JPEG (1-95):"), 0, 0)
        self.spn_jpeg_quality = QLineEdit("85")
        self.spn_jpeg_quality.setValidator(QIntValidator(1, 95, self))
        jpeg_layout.addWidget(self.spn_jpeg_quality, 0, 1)
        self.group_jpeg_opts.setVisible(False)
        output_layout.addWidget(self.group_jpeg_opts, 4, 0, 1, 2)

        self.group_tiff_opts = QGroupBox("Opções TIFF")
        tiff_layout = QGridLayout(self.group_tiff_opts)
        tiff_layout.addWidget(QLabel("Compressão TIFF:"), 0, 0)
        self.cmb_tiff_compression = QComboBox()
        self.cmb_tiff_compression.addItems(["tiff_lzw", "tiff_adobe_deflate", "group4", "jpeg"])
        tiff_layout.addWidget(self.cmb_tiff_compression, 0, 1)
        self.group_tiff_opts.setVisible(False)
        output_layout.addWidget(self.group_tiff_opts, 5, 0, 1, 2)
        main_tab_widget.addTab(tab_output, " Formato de Saída e Reconhecimento de Texto ")

        # --- Aba Assinatura PDF ---
        tab_pdfsign = QWidget()
        pdfsign_layout = QGridLayout(tab_pdfsign)

        self.chk_sign_pdf = QCheckBox("Assinar PDFs desta regra")
        self.lbl_cert_info = QLabel("<small>Selecione um certificado (Cert:\\CurrentUser\\My)</small>")
        self.cmb_certificates = QComboBox()
        self.btn_reload_certs = QPushButton("Recarregar certificados")

        pdfsign_layout.addWidget(self.chk_sign_pdf, 0, 0, 1, 2)
        pdfsign_layout.addWidget(self.lbl_cert_info, 1, 0, 1, 2)
        pdfsign_layout.addWidget(self.cmb_certificates, 2, 0)
        pdfsign_layout.addWidget(self.btn_reload_certs, 2, 1)

        main_tab_widget.addTab(tab_pdfsign, " Assinatura PDF ")
        self._idx_tab_pdfsign = main_tab_widget.indexOf(tab_pdfsign)

        self._certs_cache = []  # cache de CertificateInfo

        def _load_certs_into_combo():
            self.cmb_certificates.clear()
            try:
                certs = list_currentuser_my_certs()
                self._certs_cache = certs
                if not certs:
                    self.cmb_certificates.addItem("Nenhum certificado com chave privada foi encontrado", "")
                else:
                    for c in certs:
                        display = c.friendly_name or extract_cn_from_subject(c.subject) or c.thumbprint
                        self.cmb_certificates.addItem(f"{display}   —   {c.thumbprint}", c.thumbprint)
            except Exception as e:
                self.cmb_certificates.addItem(f"Erro ao listar certificados: {e}", "")

        _load_certs_into_combo()
        self.btn_reload_certs.clicked.connect(_load_certs_into_combo)

        def _update_pdfsign_tab_visibility():
            try:
                txt = self.cmb_output_format.currentText()
            except Exception:
                txt = ""
            is_pdf = "PDF" in (txt.upper() if isinstance(txt, str) else "")
            try:
                main_tab_widget.setTabVisible(self._idx_tab_pdfsign, bool(is_pdf))
            except Exception:
                # PySide6 < 6.7 não tem setTabVisible; fallback: mostrar sempre
                pass

        _update_pdfsign_tab_visibility()
        try:
            self.cmb_output_format.currentTextChanged.connect(lambda _t: _update_pdfsign_tab_visibility())
        except Exception:
            pass

        layout.addWidget(main_tab_widget)

        btn_save = QPushButton(self.style().standardIcon(QStyle.SP_DialogSaveButton), " Salvar Regra")
        btn_save.clicked.connect(self.save_and_close)
        btn_cancel = QPushButton(self.style().standardIcon(QStyle.SP_DialogCancelButton), " Cancelar")
        btn_cancel.clicked.connect(self.reject)
        button_box = QHBoxLayout();
        button_box.addStretch();
        button_box.addWidget(btn_save);
        button_box.addWidget(btn_cancel)
        layout.addLayout(button_box)

        self.load_data()
        self._update_input_dir_display(self.edt_rule_id.text())
        self._update_options_for_format()

    def _update_input_dir_display(self, text: str):
        rule_id_str = text.strip()
        custom = (self.edt_input_dir.text().strip() if hasattr(self, "edt_input_dir") else "")
        if custom:
            self.lbl_input_dir_display.setText(f"<i>Usando pasta personalizada: {Path(custom)}</i>")
        elif rule_id_str:
            self.lbl_input_dir_display.setText(f"<i>Padrão: {RULES_INPUT_BASE_DIR / rule_id_str}</i>")
        else:
            self.lbl_input_dir_display.setText("<i>(Será definido com base no ID da Regra)</i>")

    def _update_options_for_format(self):
        selected_format_text = self.cmb_output_format.currentText().lower()
        is_ocr_format = (
                "pdf pesquisável" in selected_format_text or "texto (txt)" in selected_format_text or "rtf (texto" in selected_format_text or "pdf/a" in selected_format_text or "docx" in selected_format_text or "xlsx" in selected_format_text or "pptx" in selected_format_text)
        self.lbl_lang.setVisible(is_ocr_format)
        self.cmb_lang.setVisible(is_ocr_format)
        self.group_jpeg_opts.setVisible("jpeg" in selected_format_text)
        self.group_tiff_opts.setVisible("tiff" in selected_format_text)

    def load_data(self) -> None:
        if self.data:
            self.edt_rule_id.setText(self.rule_id_original or "")
            self.edt_output_dir.setText(self.data.get("output", ""))
            # Preenche pasta de entrada custom se diferir do padrão baseado no ID
            try:
                _rid = self.rule_id_original or self.edt_rule_id.text().strip()
                default_in = str(RULES_INPUT_BASE_DIR / (_rid or "")) if _rid else ""
                stored_in = self.data.get("input", "")
                if stored_in and default_in and Path(stored_in) != Path(default_in):
                    self.edt_input_dir.setText(stored_in)
            except Exception:
                pass
            self.chk_orientation.setChecked(self.data.get("orientation", False))
            self.chk_framing.setChecked(self.data.get("framing", False))
            self.chk_noise.setChecked(self.data.get("noise_removal", False))
            self.chk_binar.setChecked(self.data.get("binarization", False))
            self.chk_compact.setChecked(self.data.get("compact", True))

            fmt_map_load = {"pdf searchable": "PDF Pesquisável (com reconhecimento de texto)",
                            "pdf": "PDF (Somente Imagem)", "pdf/a": "PDF/A (com reconhecimento de texto)",
                            "jpeg": "JPEG (Página única ou Zip de páginas para PDF)",
                            "tiff": "TIFF (Página única ou Zip de páginas para PDF)",
                            "text": "Texto (TXT)",
                            "rtf": "RTF (Texto Formatado)",
                            "checksum": "Checksum (SHA-256 do original)",
                            "docx": "DOCX (Texto Editável)",
                            "xlsx": "XLSX (Planilha: 1 linha por página)",
                            "pptx": "PPTX (1 slide por página)"}
            stored_fmt_key = self.data.get("output_format", "pdf searchable")
            found_display_text = next(
                (display_val for key, display_val in fmt_map_load.items() if key == stored_fmt_key), None)
            if found_display_text:
                self.cmb_output_format.setCurrentText(found_display_text)
            else:
                self.cmb_output_format.setCurrentText(fmt_map_load["pdf searchable"])
                logging.info(
                    f"Formato de saída desconhecido '{stored_fmt_key}' na regra '{self.rule_id_original}'. Usando padrão.")
            lang_map_load = {"eng": "eng (Inglês)", "por": "por (Português)", "spa": "spa (Espanhol)",
                             "fra": "fra (Francês)", "deu": "deu (Alemão)", "ita": "ita (Italiano)",
                             "osd": "osd (Detecção Auto Script/Orient.)"}
            self.cmb_lang.setCurrentText(lang_map_load.get(self.data.get("language", "por"), "por (Português)"))
            self.spn_jpeg_quality.setText(str(self.data.get("jpeg_quality", 85)))
            self.cmb_tiff_compression.setCurrentText(self.data.get("tiff_compression", "tiff_lzw"))

            # --- Assinatura PDF (restaura) ---
            self.chk_sign_pdf.setChecked(bool(self.data.get("sign_pdf", False)))
            saved_tp = (self.data.get("sign_thumbprint") or "").strip().upper()
            if saved_tp:
                for i in range(self.cmb_certificates.count()):
                    data = self.cmb_certificates.itemData(i) or ""
                    if (data or "").upper() == saved_tp:
                        self.cmb_certificates.setCurrentIndex(i)
                        break

        else:
            self.cmb_lang.setCurrentText("por (Português)")
            self.chk_orientation.setChecked(True)
            self.chk_compact.setChecked(True)
            self.chk_noise.setChecked(False)

        self._update_input_dir_display(self.edt_rule_id.text())
        self._update_options_for_format()

    def select_input_dir(self) -> None:

        path_str = QFileDialog.getExistingDirectory(self, "Selecione a Pasta de Entrada",

                                                    self.edt_input_dir.text() or str(Path.home()))

        if path_str:
            self.edt_input_dir.setText(str(Path(path_str)))

            # Atualiza visualização

            self._update_input_dir_display(self.edt_rule_id.text())

    def test_input_path(self) -> None:

        # Usa o customizado se fornecido, caso contrário, o padrão baseado no ID

        rule_id_str = self.edt_rule_id.text().strip()

        candidate = self.edt_input_dir.text().strip() or (
            str(RULES_INPUT_BASE_DIR / rule_id_str) if rule_id_str else "")

        if not candidate:
            QMessageBox.warning(self, "Teste de Caminho", "Informe o ID da regra ou selecione uma pasta de entrada.");

            return

        input_path = Path(candidate)

        is_unc = candidate.startswith('\\\\')

        if is_unc:

            if not ensure_persistent_connection(candidate):
                QMessageBox.critical(self, "Teste de Caminho de Rede",

                                     f"Falha ao conectar/autenticar no compartilhamento '{candidate}'. Verifique as credenciais e a acessibilidade.")

                return

        try:

            if not input_path.exists():

                if QMessageBox.question(self, "Pasta Inexistente",

                                        f"A pasta de entrada '{input_path}' não existe. Deseja criá-la?",

                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.Yes:

                    input_path.mkdir(parents=True, exist_ok=True)

                    QMessageBox.information(self, "Pasta Criada", f"Pasta '{input_path}' criada com sucesso.")


                else:

                    return

            _ = list(input_path.iterdir())  # teste de leitura simples

            QMessageBox.information(self, "Teste de Caminho", f"Caminho de entrada '{input_path}' é acessível.")


        except PermissionError:

            QMessageBox.critical(self, "Teste de Caminho", f"Sem permissão para acessar '{input_path}'.")


        except Exception as e:

            QMessageBox.critical(self, "Teste de Caminho", f"Erro ao testar caminho '{input_path}':\n{e}")

    def select_output_dir(self) -> None:
        path_str = QFileDialog.getExistingDirectory(self, "Selecione a Pasta de Destino",
                                                    self.edt_output_dir.text() or str(Path.home()))
        if path_str: self.edt_output_dir.setText(str(Path(path_str)))

    def test_output_path(self) -> None:
        path_str = self.edt_output_dir.text().strip()
        if not path_str: QMessageBox.warning(self, "Teste de Caminho",
                                             "Nenhum caminho de destino especificado."); return
        output_path = Path(path_str)
        is_unc = path_str.startswith("\\\\")
        if is_unc:
            if not ensure_persistent_connection(path_str):
                QMessageBox.critical(self, "Teste de Caminho de Rede",
                                     f"Falha ao conectar/autenticar no compartilhamento de '{path_str}'.\nVerifique as credenciais de rede configuradas e a acessibilidade do servidor.")
                return
        try:
            if not output_path.exists():
                if QMessageBox.question(self, "Pasta Inexistente",
                                        f"A pasta '{output_path}' não existe. Deseja criá-la?",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.Yes:
                    output_path.mkdir(parents=True, exist_ok=True)
                    QMessageBox.information(self, "Pasta Criada", f"Pasta '{output_path}' criada com sucesso.")
                else:
                    return
            test_file = output_path / f"octec_write_test_{os.getpid()}.tmp"
            with open(test_file, "w") as f:
                f.write("OCTec test write")
            test_file.unlink()
            QMessageBox.information(self, "Teste de Caminho", f"Caminho '{output_path}' é acessível e gravável.")
        except PermissionError:
            QMessageBox.critical(self, "Teste de Caminho",
                                 f"Sem permissão para escrever em '{output_path}'.\nPara caminhos de rede, verifique as credenciais e permissões do compartilhamento.")
        except Exception as e:
            QMessageBox.critical(self, "Teste de Caminho", f"Erro ao testar caminho '{output_path}':\n{e}")

    def save_and_close(self) -> None:
        new_rule_id_str = self.edt_rule_id.text().strip()
        out_dir_str = str(Path(self.edt_output_dir.text().strip()))
        if not new_rule_id_str: QMessageBox.warning(self, "Erro", "O 'ID da Regra' é obrigatório."); return
        if not out_dir_str: QMessageBox.warning(self, "Erro", "A 'Pasta de Destino' é obrigatória."); return
        if not re.match(r"^[a-zA-Z0-9_-]+$", new_rule_id_str):
            QMessageBox.warning(self, "Erro", "ID da Regra deve conter apenas letras, números, '_' ou '-'.");
            return
        if self.rule_id_original and self.rule_id_original != new_rule_id_str and new_rule_id_str in rules_data:
            QMessageBox.warning(self, "Erro", f"O novo ID da Regra '{new_rule_id_str}' já existe.");
            return
        elif not self.rule_id_original and new_rule_id_str in rules_data:
            QMessageBox.warning(self, "Erro", f"O ID da Regra '{new_rule_id_str}' já existe.");
            return

        fmt_display_to_key = {
            "PDF Pesquisável (com reconhecimento de texto)": "pdf searchable",
            "PDF (Somente Imagem)": "pdf",
            "PDF/A (com reconhecimento de texto)": "pdf/a",
            "JPEG (Página única ou Zip de páginas para PDF)": "jpeg",
            "TIFF (Página única ou Zip de páginas para PDF)": "tiff",
            "Texto (TXT)": "text",
            "RTF (Texto Formatado)": "rtf",
            "Checksum (SHA-256 do original)": "checksum",
            "DOCX (Texto Editável)": "docx"
            ,
            "XLSX (Planilha: 1 linha por página)": "xlsx",
            "PPTX (1 slide por página)": "pptx"
        }
        output_format_code = fmt_display_to_key.get(self.cmb_output_format.currentText(), "pdf searchable")
        logging.debug(f"RuleDialog: Saving rule '{new_rule_id_str}' with output_format_code: '{output_format_code}'")

        lang_map_save = {
            "eng (Inglês)": "eng",
            "por (Português)": "por",
            "spa (Espanhol)": "spa",
            "fra (Francês)": "fra",
            "deu (Alemão)": "deu",
            "ita (Italiano)": "ita",
            "osd (Detecção Auto Script/Orient.)": "osd"
        }
        lang_code = lang_map_save.get(self.cmb_lang.currentText(), "por")

        self.data = {
            "input": (self.edt_input_dir.text().strip() or str(RULES_INPUT_BASE_DIR / new_rule_id_str)),
            "output": out_dir_str,
            "orientation": self.chk_orientation.isChecked(),
            "framing": self.chk_framing.isChecked(),
            "noise_removal": self.chk_noise.isChecked(),
            "binarization": self.chk_binar.isChecked(),
            "compact": self.chk_compact.isChecked(),
            "output_format": output_format_code,
            "language": lang_code,
            "jpeg_quality": int(self.spn_jpeg_quality.text() or "85"),
            "tiff_compression": self.cmb_tiff_compression.currentText(),
            "paused": self.data.get("paused", False),
            # --- Assinatura PDF ---
            "sign_pdf": self.chk_sign_pdf.isChecked(),
            "sign_thumbprint": (self.cmb_certificates.currentData() or "").strip(),
            "sign_display_name": (self.cmb_certificates.currentText().split("   —   ")[0].strip()
                                  if self.cmb_certificates.currentData() else ""),

        }
        # Preserve zonal_ocr_template if it exists and rule is being edited
        if self.rule_id_original and 'zonal_ocr_template' in rules_data.get(self.rule_id_original, {}):
            self.data['zonal_ocr_template'] = rules_data[self.rule_id_original]['zonal_ocr_template']
        self.accept()

    def get_rule_data(self) -> tuple[str, dict]:
        return self.edt_rule_id.text().strip(), self.data


# --- NOVAS CLASSES PARA OCR ZONAL ---

class ZoomableGraphicsView(QGraphicsView):
    """ View customizada para exibir a imagem e desenhar retângulos de seleção, com funcionalidade de zoom. """

    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        # Optional: for smoother rendering
        self.setRenderHint(QPainter.Antialiasing)
        self.setRenderHint(QPainter.SmoothPixmapTransform)
        self.setOptimizationFlag(QGraphicsView.DontSavePainterState)
        self.setViewportUpdateMode(QGraphicsView.FullViewportUpdate)
        # Anchor the transformation (zoom) under the mouse cursor
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorViewCenter)  # Keep view centered on resize
        self._zoom_factor = 1.0
        self._zoom_step = 1.15  # Adjust for desired zoom speed

    def wheelEvent(self, event: QWheelEvent):
        # Only zoom if no modifier keys are pressed, to avoid conflicts with other interactions
        if event.modifiers() == Qt.NoModifier:
            if event.angleDelta().y() > 0:  # Scroll up (zoom in)
                self.zoom_in()
            else:  # Scroll down (zoom out)
                self.zoom_out()
        else:
            super().wheelEvent(event)  # Pass other wheel events to the base class

    def zoom_in(self):
        self.scale(self._zoom_step, self._zoom_step)
        self._zoom_factor *= self._zoom_step

    def zoom_out(self):
        self.scale(1.0 / self._zoom_step, 1.0 / self._zoom_step)
        self._zoom_factor /= self._zoom_step

    def reset_zoom(self):
        self.setTransform(QTransform())  # Resets all transformations
        self._zoom_factor = 1.0
        # After resetting transform, ensure the pixmap fits the view again
        if self.scene() and self.scene().items():
            # Find the pixmap item (assuming it's the first item added or identifiable)
            pixmap_item = next((item for item in self.scene().items() if isinstance(item, QGraphicsPixmapItem)), None)
            if pixmap_item:
                self.fitInView(pixmap_item, Qt.KeepAspectRatio)


class ZonalOCRImageView(ZoomableGraphicsView):  # Inherit from ZoomableGraphicsView
    """ View customizada para exibir a imagem e desenhar retângulos de seleção. """
    zone_drawn = Signal(QRectF)

    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self.rubber_band = None
        self.origin = None
        self.setDragMode(QGraphicsView.RubberBandDrag)  # Enable rubber band selection by default

    def mousePressEvent(self, event):
        # Only start rubber band if left button is pressed and no modifier keys
        if event.button() == Qt.LeftButton and event.modifiers() == Qt.NoModifier and self.scene().items():
            self.origin = self.mapToScene(event.pos())
            self.rubber_band = QGraphicsRectItem(QRectF(self.origin, self.origin))
            self.rubber_band.setPen(QPen(QColor(255, 0, 0, 200), 2, Qt.DashLine))
            self.scene().addItem(self.rubber_band)
            # Do not call super().mousePressEvent(event) here if you want to prevent default drag mode
            # from interfering with rubber band drawing.
        else:
            super().mousePressEvent(event)  # Allow pan/zoom if modifiers are pressed or other buttons

    def mouseMoveEvent(self, event):
        if self.rubber_band:
            end_point = self.mapToScene(event.pos())
            self.rubber_band.setRect(QRectF(self.origin, end_point).normalized())
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.rubber_band:
            self.scene().removeItem(self.rubber_band)
            final_rect = self.rubber_band.rect()
            self.rubber_band = None
            if final_rect.width() > 5 and final_rect.height() > 5:  # Ensure a meaningful rectangle was drawn
                self.zone_drawn.emit(final_rect)
        else:  # Ensure base class method is called for other mouse releases (e.g., pan)
            super().mouseReleaseEvent(event)


class ZonalOCRDialog(QDialog):
    """ Diálogo principal para configuração do reconhecimento de texto zonal. """

    def __init__(self, parent=None, rules_data=None):
        super().__init__(parent)
        self.setWindowTitle("Configurar Template de Reconhecimento de Texto Zonal")
        self.setMinimumSize(900, 700)
        self.rules_data = rules_data
        self.current_image_path = None
        self.pixmap_item = None
        self.zones = []  # Lista de {'name': str, 'rect': QRectF, 'item': QGraphicsRectItem, 'order': int}

        # Layout principal com splitter
        main_layout = QHBoxLayout(self)
        splitter = QSplitter(Qt.Horizontal)

        # Painel Esquerdo: Imagem
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        self.scene = QGraphicsScene(self)
        self.view = ZonalOCRImageView(self.scene, self)  # Use the updated ZonalOCRImageView
        self.view.zone_drawn.connect(self.add_new_zone)
        self.view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # Ensure it expands
        left_layout.addWidget(self.view)

        # Zoom Controls below the image view
        zoom_buttons_layout = QHBoxLayout()
        self.btn_zoom_in = QPushButton("+ Zoom")
        self.btn_zoom_out = QPushButton("- Zoom")
        self.btn_reset_zoom = QPushButton("Reset Zoom")

        self.btn_zoom_in.clicked.connect(self.view.zoom_in)
        self.btn_zoom_out.clicked.connect(self.view.zoom_out)
        self.btn_reset_zoom.clicked.connect(self.view.reset_zoom)

        zoom_buttons_layout.addStretch()
        zoom_buttons_layout.addWidget(self.btn_zoom_in)
        zoom_buttons_layout.addWidget(self.btn_zoom_out)
        zoom_buttons_layout.addWidget(self.btn_reset_zoom)
        zoom_buttons_layout.addStretch()

        left_layout.addLayout(zoom_buttons_layout)

        # Painel Direito: Controles
        right_widget = QWidget()
        right_layout = QGridLayout(right_widget)

        # Controles
        btn_load_image = QPushButton(self.style().standardIcon(QStyle.SP_DirOpenIcon), " Carregar Imagem de Modelo")
        btn_load_image.clicked.connect(self.load_image)

        self.zones_table = QTableWidget()
        self.zones_table.setColumnCount(3)
        self.zones_table.setHorizontalHeaderLabels(["Ordem", "Nome da Zona", "Ação"])
        self.zones_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.zones_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.zones_table.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Make table read-only
        self.zones_table.setSortingEnabled(True)  # Enable sorting

        # Connect item click to highlight zone on image
        self.zones_table.itemSelectionChanged.connect(self.highlight_selected_zone)

        btn_preview = QPushButton("Preview Reconhecimento de Texto da Zona")
        btn_preview.clicked.connect(self.preview_zone_ocr)

        # Grupo de Configuração
        config_group = QGroupBox("Configuração do Nome do Arquivo")
        config_layout = QGridLayout(config_group)
        config_layout.addWidget(QLabel("Prefixo:"), 0, 0)
        self.edt_prefix = QLineEdit()
        config_layout.addWidget(self.edt_prefix, 0, 1)
        config_layout.addWidget(QLabel("Sufixo:"), 1, 0)
        self.edt_suffix = QLineEdit()
        config_layout.addWidget(self.edt_suffix, 1, 1)
        config_layout.addWidget(QLabel("Delimitador:"), 2, 0)
        self.cmb_delimiter = QComboBox()
        self.cmb_delimiter.addItems(["_", "-"])
        config_layout.addWidget(self.cmb_delimiter, 2, 1)
        self.chk_keep_indexing = QCheckBox("Manter Indexação Prévia")
        config_layout.addWidget(self.chk_keep_indexing, 3, 0, 1, 2)

        # Grupo de Configuração OCR
        ocr_config_group = QGroupBox("Configurações de Reconhecimento de Texto")
        ocr_config_layout = QGridLayout(ocr_config_group)
        ocr_config_layout.addWidget(QLabel("Idioma Reconhecimento de Texto:"), 0, 0)
        self.cmb_ocr_lang = QComboBox()
        self.cmb_ocr_lang.addItems(["por", "eng", "spa", "fra", "deu", "ita", "osd"])
        ocr_config_layout.addWidget(self.cmb_ocr_lang, 0, 1)
        ocr_config_layout.addWidget(QLabel("PSM (Page Segmentation Mode):"), 1, 0)
        self.spn_ocr_psm = QLineEdit("7")  # Default PSM for single text line
        self.spn_ocr_psm.setValidator(QIntValidator(0, 13, self))
        ocr_config_layout.addWidget(self.spn_ocr_psm, 1, 1)

        # Grupo de Salvamento
        save_group = QGroupBox("Salvar Template")
        save_layout = QGridLayout(save_group)
        save_layout.addWidget(QLabel("Atribuir à Regra:"), 0, 0)
        self.cmb_rules = QComboBox()
        self.cmb_rules.addItems(sorted(self.rules_data.keys()))
        self.cmb_rules.currentIndexChanged.connect(self.load_template_for_rule)
        save_layout.addWidget(self.cmb_rules, 0, 1)
        btn_save = QPushButton(self.style().standardIcon(QStyle.SP_DialogSaveButton), " Salvar Template")
        btn_save.clicked.connect(self.save_template)

        # Adicionando widgets ao layout direito
        right_layout.addWidget(btn_load_image, 0, 0, 1, 2)
        right_layout.addWidget(self.zones_table, 1, 0, 1, 2)
        right_layout.addWidget(btn_preview, 2, 0, 1, 2)
        right_layout.addWidget(config_group, 3, 0, 1, 2)
        right_layout.addWidget(ocr_config_group, 4, 0, 1, 2)  # Add OCR config group
        right_layout.addWidget(save_group, 5, 0, 1, 2)
        save_layout.addWidget(btn_save, 1, 0, 1, 2)
        right_layout.setRowStretch(1, 1)  # Tabela de zonas expande

        # Montagem final
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([600, 300])  # Initial sizes for the splitter
        main_layout.addWidget(splitter)

        if self.cmb_rules.count() > 0:
            self.load_template_for_rule(0)

    def load_image(self):
        path, _ = QFileDialog.getOpenFileName(self, "Selecionar Imagem de Modelo", "",
                                              "Imagens (*.png *.jpg *.jpeg *.bmp *.tif *.tiff);;PDF (*.pdf)")
        if path:
            self.current_image_path = Path(path)
            # Create a temporary PNG for display regardless of input type
            with tempfile.TemporaryDirectory() as tmpdir:
                display_image_path = get_image_for_zonal_ocr(self.current_image_path, Path(tmpdir))
                if display_image_path:
                    pixmap = QPixmap(str(display_image_path))
                    self.scene.clear()  # Limpa zonas antigas da cena
                    self.pixmap_item = self.scene.addPixmap(pixmap)
                    self.view.fitInView(self.pixmap_item, Qt.KeepAspectRatio)  # Fit the image initially
                    self.view.reset_zoom()  # Reset zoom after loading new image
                    # Recarrega as zonas do template atual na nova imagem
                    self.load_template_for_rule(self.cmb_rules.currentIndex())
                else:
                    QMessageBox.critical(self, "Erro", "Não foi possível carregar a imagem do arquivo selecionado.")

    def add_new_zone(self, rect: QRectF):
        name, ok = QInputDialog.getText(self, "Nova Zona", "Nome para a nova zona de reconhecimento de texto:")
        if ok and name:
            zone_item = self.scene.addRect(rect, QPen(QColor(0, 255, 0, 200), 2))
            order = len(self.zones)
            self.zones.append({'name': name, 'rect': rect, 'item': zone_item, 'order': order})
            self.update_zones_table()

    def update_zones_table(self):
        self.zones_table.setSortingEnabled(False)  # Disable sorting during update
        self.zones_table.setRowCount(0)

        # Sort zones by order before populating table
        sorted_zones = sorted(self.zones, key=lambda z: z['order'])
        self.zones_table.setRowCount(len(sorted_zones))

        for i, zone in enumerate(sorted_zones):
            order_item = QTableWidgetItem(str(zone['order']))
            order_item.setFlags(order_item.flags() & ~Qt.ItemIsEditable)  # Make order column read-only
            name_item = QTableWidgetItem(zone['name'])
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)  # Make name column read-only

            self.zones_table.setItem(i, 0, order_item)
            self.zones_table.setItem(i, 1, name_item)

            btn_del = QPushButton("Remover")
            btn_del.clicked.connect(lambda checked, z=zone: self.remove_zone(z))
            self.zones_table.setCellWidget(i, 2, btn_del)
        self.zones_table.setSortingEnabled(True)  # Re-enable sorting

    def remove_zone(self, zone_to_remove):
        self.scene.removeItem(zone_to_remove['item'])
        self.zones.remove(zone_to_remove)
        # Re-assign orders to maintain sequential numbering
        for i, zone in enumerate(sorted(self.zones, key=lambda z: z['order'])):
            zone['order'] = i
        self.update_zones_table()

    def highlight_selected_zone(self):
        # Clear previous highlights
        for item in self.scene.items():
            if isinstance(item, QGraphicsRectItem) and item != self.pixmap_item:
                item.setPen(QPen(QColor(0, 255, 0, 200), 2))  # Reset to green

        selected_items = self.zones_table.selectedItems()
        if selected_items and self.pixmap_item:
            row = selected_items[0].row()
            zone_name = self.zones_table.item(row, 1).text()
            zone_data = next((z for z in self.zones if z['name'] == zone_name), None)
            if zone_data and zone_data['item'].scene() == self.scene:
                zone_data['item'].setPen(QPen(QColor(255, 255, 0, 255), 3))  # Highlight in yellow
                self.view.ensureVisible(zone_data['item'])  # Scroll to make it visible

    def preview_zone_ocr(self):
        selected = self.zones_table.selectedItems()
        if not selected or not self.current_image_path:
            QMessageBox.warning(self, "Preview",
                                "Selecione uma zona na tabela e carregue uma imagem de modelo primeiro.")
            return

        row = selected[0].row()
        zone_name = self.zones_table.item(row, 1).text()
        zone_data = next((z for z in self.zones if z['name'] == zone_name), None)

        if zone_data:
            try:
                # Ensure the image is reloaded to get fresh data, especially if it was a PDF
                with tempfile.TemporaryDirectory() as tmpdir:
                    source_image_for_ocr = get_image_for_zonal_ocr(self.current_image_path, Path(tmpdir))
                    if not source_image_for_ocr:
                        QMessageBox.critical(self, "Erro de Preview",
                                             "Não foi possível preparar a imagem para reconhecimento de texto.")
                        return

                    pil_image = Image.open(source_image_for_ocr)

                    text = do_ocr_on_region(pil_image,
                                            (zone_data['rect'].x(), zone_data['rect'].y(), zone_data['rect'].width(),
                                             zone_data['rect'].height()),
                                            self.cmb_ocr_lang.currentText(),
                                            int(self.spn_ocr_psm.text()))
                    pil_image.close()
                    QMessageBox.information(self, f"Preview Reconhecimento de Texto: {zone_name}",
                                            f"Texto extraído:\n\n{text}")
            except Exception as e:
                QMessageBox.critical(self, "Erro de Preview",
                                     f"Não foi possível realizar o reconhecimento de texto: {e}")

    def load_template_for_rule(self, index):
        # Limpa o estado atual
        self.edt_prefix.clear()
        self.edt_suffix.clear()
        self.chk_keep_indexing.setChecked(False)
        self.cmb_delimiter.setCurrentIndex(0)
        self.cmb_ocr_lang.setCurrentIndex(0)  # Reset language
        self.spn_ocr_psm.setText("7")  # Reset PSM

        for z in self.zones:
            if z['item'].scene() == self.scene:
                self.scene.removeItem(z['item'])
        self.zones.clear()
        self.update_zones_table()

        if index < 0: return
        rule_id = self.cmb_rules.itemText(index)
        rule_conf = self.rules_data.get(rule_id, {})
        template = rule_conf.get("zonal_ocr_template")

        if template:
            self.edt_prefix.setText(template.get("prefix", ""))
            self.edt_suffix.setText(template.get("suffix", ""))
            self.chk_keep_indexing.setChecked(template.get("keep_previous_indexing", False))
            self.cmb_delimiter.setCurrentText(template.get("delimiter", "_"))
            self.cmb_ocr_lang.setCurrentText(template.get("lang", "por"))
            self.spn_ocr_psm.setText(str(template.get("psm", 7)))

            if self.pixmap_item:  # Only draw zones if an image is loaded
                for zone_data in template.get("zones", []):
                    rect_list = zone_data['rect']
                    rect = QRectF(rect_list[0], rect_list[1], rect_list[2], rect_list[3])
                    zone_item = self.scene.addRect(rect, QPen(QColor(0, 255, 0, 200), 2))
                    self.zones.append({
                        'name': zone_data['name'],
                        'rect': rect,
                        'item': zone_item,
                        'order': zone_data.get('order', 99)
                    })
            self.update_zones_table()

    def save_template(self):
        rule_id = self.cmb_rules.currentText()
        if not rule_id:
            QMessageBox.warning(self, "Salvar", "Nenhuma regra selecionada para salvar o template.")
            return

        template_data = {
            "prefix": self.edt_prefix.text(),
            "suffix": self.edt_suffix.text(),
            "delimiter": self.cmb_delimiter.currentText(),
            "keep_previous_indexing": self.chk_keep_indexing.isChecked(),
            "lang": self.cmb_ocr_lang.currentText(),
            "psm": int(self.spn_ocr_psm.text()),
            "zones": [
                {
                    "name": z['name'],
                    "rect": [z['rect'].x(), z['rect'].y(), z['rect'].width(), z['rect'].height()],
                    "order": z['order']
                } for z in self.zones
            ]
        }

        self.rules_data[rule_id]['zonal_ocr_template'] = template_data
        save_config(self.rules_data)
        QMessageBox.information(self, "Sucesso",
                                f"Template de Reconhecimento de Texto Zonal salvo para a regra '{rule_id}'.")
        self.close()


class MainWindow(QMainWindow):
    request_gui_update = Signal()

    def __init__(self) -> None:
        super().__init__()
        # Inicializa o estado da licença como desabilitado por padrão
        self.engine_enabled = False
        self.license_key = None  # Inicializa a chave de licença

        self.setWindowTitle(f"{APP_NAME} - Processamento de Documentos e Gerenciador de Arquivos - v{APP_VERSION}")
        icon_path = RESOURCES_DIR / "app_icon.ico"  # Corrected icon path
        self.setWindowIcon(QIcon(str(icon_path)) if icon_path.exists() else self.load_icon_with_fallback(None,
                                                                                                         "application-x-executable",
                                                                                                         QStyle.SP_ComputerIcon))
        self.set_dark_theme()
        self.center_window()
        self.detail_widget = FileDetailWidget()
        self.request_gui_update.connect(self.update_rules_table_display)
        self.request_gui_update.connect(self.detail_widget.refresh_files_display)
        self.init_tray_icon()
        self.rules_table = QTableWidget()
        self.rules_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.rules_table.customContextMenuRequested.connect(self.on_rules_table_context_menu)
        self.rules_table.setColumnCount(6)
        self.rules_table.setHorizontalHeaderLabels(
            ["Status", "ID da Regra", "Destino", "Fila", "Formato Saída", "Progresso Médio"])
        self.rules_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.rules_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.rules_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.rules_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Interactive)
        self.rules_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.rules_table.setSelectionMode(QAbstractItemView.SingleSelection)
        # Timer para coalescer atualizações da tabela de regras.  Em vez de
        # reconstruir a tabela em cada chamada, agendamos uma atualização para
        # daqui a alguns milissegundos.  Isso reduz o impacto de múltiplas
        # chamadas consecutivas e torna a interface mais leve.
        self._rules_update_timer = QTimer(self)
        self._rules_update_timer.setSingleShot(True)
        self._rules_update_timer.timeout.connect(self._perform_rules_table_update)
        self.rules_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.rules_table.itemSelectionChanged.connect(self.on_rule_selection_changed)
        self.build_toolbar()
        self.build_menu_bar()
        self.splitter = QSplitter(Qt.Vertical)
        self.splitter.addWidget(self.rules_table)
        self.splitter.addWidget(self.detail_widget)
        self.setCentralWidget(self.splitter)
        self.status_bar = self.statusBar()
        self.status_bar.showMessage(f"{APP_NAME} Pronto. Verificando licença...", 5000)
        self.gui_update_timer = QTimer(self)
        self.gui_update_timer.timeout.connect(self.request_gui_update.emit)
        self.gui_update_timer.start(ENGINE_CHECK_INTERVAL + 500)

        self.update_rules_table_display()
        self.is_quitting_app = False
        QTimer.singleShot(100, lambda: self.splitter.setSizes([self.height() // 2, self.height() // 2]))
        try:
            test_perm_file = CONFIG_DIR / f".octec_perm_test_{os.getpid()}"
            with open(test_perm_file, "w") as f:
                f.write("ok")
            test_perm_file.unlink()
        except Exception as e_perm:
            logging.error(f"Falha no teste de permissão de escrita em {CONFIG_DIR}: {e_perm}")
            QMessageBox.critical(self, "Erro de Permissão",
                                 f"Não foi possível escrever no diretório de configuração ({CONFIG_DIR}).\nVerifique as permissões ou execute como administrador.\nO {APP_NAME} pode não funcionar corretamente.")

        # Inicia a verificação da licença após a GUI estar pronta
        QTimer.singleShot(0, self._check_license_on_startup)

        # Cria um timer que a cada 5 segundos verifica a licença
        self._license_poll_timer = QTimer(self)
        self._license_poll_timer.timeout.connect(self._poll_license)
        self._license_poll_timer.start(5000)  # 5000 ms = 5 segundos

    def create_dark_palette(self) -> QPalette:
        palette = QPalette();
        palette.setColor(QPalette.Window, QColor(53, 53, 53));
        palette.setColor(QPalette.WindowText, Qt.white);
        palette.setColor(QPalette.Base, QColor(25, 25, 25));
        palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53));
        palette.setColor(QPalette.ToolTipBase, Qt.white);
        palette.setColor(QPalette.ToolTipText, Qt.white);
        palette.setColor(QPalette.Text, Qt.white);
        palette.setColor(QPalette.Button, QColor(53, 53, 53));
        palette.setColor(QPalette.ButtonText, Qt.white);
        palette.setColor(QPalette.BrightText, Qt.red);
        palette.setColor(QPalette.Link, QColor(42, 130, 218));
        palette.setColor(QPalette.Highlight, QColor(42, 130, 218));
        palette.setColor(QPalette.HighlightedText, Qt.black);
        palette.setColor(QPalette.Disabled, QPalette.Text, QColor(127, 127, 127));
        palette.setColor(QPalette.Disabled, QPalette.ButtonText, QColor(127, 127, 127));
        return palette

    def set_dark_theme(self) -> None:
        QApplication.instance().setPalette(self.create_dark_palette())
        self.setStyleSheet(
            "QToolTip { color: #ffffff; background-color: #2a82da; border: 1px solid white; } QTableWidget { gridline-color: #777777; } QHeaderView::section { background-color: #444; border: 1px solid #555; padding: 4px; } QProgressBar { border: 1px solid #555; border-radius: 5px; text-align: center; color: white; background-color: #333; } QProgressBar::chunk { background-color: #05B8CC; border-radius: 5px; } QSplitter::handle { background-color: #444; border: 1px solid #555; } QSplitter::handle:hover { background-color: #555; } QMainWindow { background-color: #353535; } QDialog { background-color: #353535; } QGroupBox { border:1px solid gray; border-radius: 5px; margin-top: 1ex; font-weight: bold; color: white;} QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding:0 3px; color: #ccc;} QLabel { color: white; } QLineEdit { background-color: #252525; color: white; border: 1px solid #555; border-radius: 3px; padding: 3px;} QPushButton { background-color: #535353; color: white; border: 1px solid #666; border-radius: 3px; padding: 5px;} QPushButton:hover { background-color: #636363; } QPushButton:pressed { background-color: #434343; } QComboBox { background-color: #252525; color: white; border: 1px solid #555; border-radius: 3px; padding: 3px;} QComboBox::drop-down { border: none; background-color: #535353;} QComboBox QAbstractItemView { background-color: #252525; color: white; selection-background-color: #4282DA; } QCheckBox { color: white; } QTextEdit { background-color: #202020; color: #f0f0f0; border: 1px solid #555;}")

    def center_window(self) -> None:
        try:
            screen_geo = QApplication.primaryScreen().availableGeometry()
            desired_width, desired_height = int(screen_geo.width() * 0.8), int(screen_geo.height() * 0.8)
            min_width, min_height = 800, 600
            max_width, max_height = max(min_width, int(screen_geo.width() * 0.90)), max(min_height,
                                                                                        int(screen_geo.height() * 0.85))
            final_width, final_height = max(min_width, min(desired_width, max_width)), max(min_height,
                                                                                           min(desired_height,
                                                                                               max_height))
            self.resize(final_width, final_height)
            self.setGeometry(QStyle.alignedRect(Qt.LeftToRight, Qt.AlignCenter, self.size(), screen_geo))
        except Exception as e:
            logging.warning(f"Não foi possível centralizar e redimensionar a janela adequadamente: {e}")
            self.resize(1024, 768)
            frame_geometry = self.frameGeometry()
            screen_center = QApplication.primaryScreen().availableGeometry().center()
            frame_geometry.moveCenter(screen_center)
            self.move(frame_geometry.topLeft())

    def load_icon_with_fallback(self, filename: str | None, theme_name: str,
                                fallback_pixmap: QStyle.StandardPixmap) -> QIcon:
        if filename:
            path = RESOURCES_DIR / filename
            if path.exists(): return QIcon(str(path))
        theme_icon = QIcon.fromTheme(theme_name)
        if not theme_icon.isNull(): return theme_icon
        return self.style().standardIcon(fallback_pixmap)

    def build_toolbar(self) -> None:
        toolbar = self.addToolBar("Ações Principais")
        toolbar.setMovable(False)
        icon_size = self.style().pixelMetric(QStyle.PM_ToolBarIconSize)
        toolbar.setIconSize(QSize(icon_size, icon_size))

        def _add_tb_action(icon_file, theme_name, fallback, text, slot, tooltip=""):
            action = QAction(self.load_icon_with_fallback(icon_file, theme_name, fallback), text, self)
            action.triggered.connect(slot)
            if tooltip: action.setToolTip(tooltip)
            toolbar.addAction(action)
            return action

        _add_tb_action("key_icon.png", "preferences-system-network", QStyle.SP_ComputerIcon, "Credenciais",
                       self.open_credentials_dialog, "Gerenciar credenciais de rede")
        toolbar.addSeparator()
        _add_tb_action(None, "document-new", QStyle.SP_FileIcon, "Nova", self.create_new_rule, "Criar nova regra")
        _add_tb_action("edit_icon.png", "document-edit", QStyle.SP_FileDialogDetailedView, "Editar",
                       self.edit_selected_rule, "Editar regra selecionada")
        _add_tb_action(None, "edit-delete", QStyle.SP_TrashIcon, "Remover", self.remove_selected_rule,
                       "Remover regra selecionada")
        toolbar.addSeparator()
        self.pause_engine_action_tb = _add_tb_action(None, "", QStyle.SP_MediaStop, "Pausar Motor",
                                                     self.toggle_engine_pause_state, "Pausar/Resumir motor global")
        self.pause_rule_action_tb = _add_tb_action(None, "", QStyle.SP_MediaPause, "Pausar Regra",
                                                   self.toggle_selected_rule_pause_state,
                                                   "Pausar/Resumir regra selecionada")
        self.pause_rule_action_tb.setEnabled(False)
        self.update_pause_actions_visuals()

        # --- BOTÃO OCR ZONAL ADICIONADO AQUI ---
        toolbar.addSeparator()
        _add_tb_action("target_icon.png", "zoom-fit-best", QStyle.SP_FileDialogDetailedView,
                       "Reconhecimento de Texto Zonal",
                       self.open_zonal_ocr_dialog, "Configurar reconhecimento de texto zonal para uma regra")

        # --- Split/Merge PDF ---
        toolbar.addSeparator()
        _add_tb_action("splitmerge_icon.png", "", QStyle.SP_FileDialogListView, "Split/Merge PDF",
                       self.open_split_merge_tool, "Abrir utilitário para dividir/mesclar PDFs")

        # --- License Query Button (pushed to the right) ---
        toolbar.addSeparator()
        # Adiciona um spacer para empurrar o botão de licença para a direita
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        toolbar.addWidget(spacer)

        # Usar o ícone de chave do tema; senão, usa um padrão Qt para o botão de licença
        license_button_icon = QIcon.fromTheme("emblem-keys")
        if license_button_icon.isNull():
            license_button_icon = self.style().standardIcon(QStyle.SP_MessageBoxInformation)

        self.btn_query_license = QPushButton(license_button_icon, "Consultar Licença")
        self.btn_query_license.clicked.connect(self._show_license_dialog)  # Conecta ao novo método
        toolbar.addWidget(self.btn_query_license)

    def build_menu_bar(self):
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("&Arquivo")
        cred_action = file_menu.addAction(
            self.load_icon_with_fallback("key_icon.png", "preferences-system-network", QStyle.SP_ComputerIcon),
            "&Credenciais de Rede...")
        cred_action.triggered.connect(self.open_credentials_dialog)
        file_menu.addSeparator()

        # --- Ação "Ativar Licença" (agora chama o mesmo diálogo) ---
        activate_action = file_menu.addAction("Ativar Licença")
        activate_action.triggered.connect(self._show_license_dialog)
        file_menu.addSeparator()
        # --- Fim da Ação "Ativar Licença" ---

        exit_action = file_menu.addAction(
            self.load_icon_with_fallback(None, "application-exit", QStyle.SP_DialogCloseButton), "&Sair")
        exit_action.triggered.connect(self.confirm_quit_application)

        rules_menu = menu_bar.addMenu("&Regras")
        new_rule_action = rules_menu.addAction(self.load_icon_with_fallback(None, "document-new", QStyle.SP_FileIcon),
                                               "&Nova Regra...")
        new_rule_action.triggered.connect(self.create_new_rule)
        edit_rule_action = rules_menu.addAction(
            self.load_icon_with_fallback("edit_icon.png", "document-edit", QStyle.SP_FileDialogDetailedView),
            "&Editar Regra Selecionada...")
        edit_rule_action.triggered.connect(self.edit_selected_rule)
        remove_rule_action = rules_menu.addAction(
            self.load_icon_with_fallback(None, "edit-delete", QStyle.SP_TrashIcon), "&Remover Regra Selecionada")
        remove_rule_action.triggered.connect(self.remove_selected_rule)
        rules_menu.addSeparator()
        self.menu_pause_rule_action = rules_menu.addAction("Pausar/Resumir Regra")
        self.menu_pause_rule_action.triggered.connect(self.toggle_selected_rule_pause_state)
        self.menu_pause_rule_action.setEnabled(False)
        engine_menu = menu_bar.addMenu("&Motor")
        self.menu_pause_engine_action = engine_menu.addAction("Pausar/Resumir Motor Global")
        self.menu_pause_engine_action.triggered.connect(self.toggle_engine_pause_state)
        help_menu = menu_bar.addMenu("&Ajuda")
        about_action = help_menu.addAction(f"Sobre {APP_NAME}...")
        about_action.triggered.connect(self.show_about_dialog)

    def init_tray_icon(self) -> None:
        self.tray_icon = QSystemTrayIcon(self)
        icon_path = RESOURCES_DIR / "app_icon.ico"
        tray_qicon = QIcon(str(icon_path)) if icon_path.exists() else self.load_icon_with_fallback(None,
                                                                                                   "application-x-executable",
                                                                                                   QStyle.SP_ComputerIcon)
        self.tray_icon.setIcon(tray_qicon)
        tray_menu = QMenu(self)
        show_action = tray_menu.addAction(f"Abrir {APP_NAME}")
        show_action.triggered.connect(self.show_window_from_tray)
        tray_menu.addSeparator()
        quit_action = tray_menu.addAction(f"Sair do {APP_NAME}")
        quit_action.triggered.connect(self.confirm_quit_application)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_icon_activated)
        self.tray_icon.setToolTip(f"{APP_NAME} - Processamento de Documentos")
        self.tray_icon.show()

    def update_pause_actions_visuals(self):
        is_engine_on = engine_running.is_set() and self.engine_enabled  # Only consider engine "on" if licensed
        engine_text = "Pausar Motor" if is_engine_on else "Resumir Motor"
        engine_qicon = self.load_icon_with_fallback(None, "process-stop" if is_engine_on else "view-refresh",
                                                    QStyle.SP_BrowserStop if is_engine_on else QStyle.SP_BrowserReload)
        self.pause_engine_action_tb.setIcon(engine_qicon)
        self.pause_engine_action_tb.setText(engine_text)
        self.pause_engine_action_tb.setEnabled(self.engine_enabled)  # Disable if license is invalid
        if hasattr(self, 'menu_pause_engine_action'):
            self.menu_pause_engine_action.setIcon(engine_qicon);
            self.menu_pause_engine_action.setText(engine_text)
            self.menu_pause_engine_action.setEnabled(self.engine_enabled)  # Disable if license is invalid

        selected_rule_id = self.get_selected_rule_id()
        rule_text = "Pausar Regra"
        rule_fallback_default, rule_theme_default = QStyle.SP_MediaPause, "media-playback-pause"
        self.pause_rule_action_tb.setEnabled(False)
        if hasattr(self, 'menu_pause_rule_action'): self.menu_pause_rule_action.setEnabled(False)

        if selected_rule_id and selected_rule_id in rules_data and self.engine_enabled:  # Also check license
            is_rule_paused = rules_data[selected_rule_id].get("paused", False)
            rule_theme_default = "media-playback-start" if is_rule_paused else "media-playback-pause"
            rule_fallback_default = QStyle.SP_MediaPlay if is_rule_paused else QStyle.SP_MediaPause
            rule_text = "Resumir Regra" if is_rule_paused else "Pausar Regra"
            self.pause_rule_action_tb.setEnabled(True)
            if hasattr(self, 'menu_pause_rule_action'): self.menu_pause_rule_action.setEnabled(True)

        rule_qicon = self.load_icon_with_fallback(None, rule_theme_default, rule_fallback_default)
        self.pause_rule_action_tb.setIcon(rule_qicon);
        self.pause_rule_action_tb.setText(rule_text)
        if hasattr(self, 'menu_pause_rule_action'):
            self.menu_pause_rule_action.setIcon(rule_qicon);
            self.menu_pause_rule_action.setText(rule_text)

    def get_selected_rule_id(self) -> str | None:
        selected_rows = self.rules_table.selectionModel().selectedRows()
        if selected_rows:
            id_item = self.rules_table.item(selected_rows[0].row(), 1)
            if id_item: return id_item.text()
        return None

    def on_rule_selection_changed(self) -> None:
        selected_id = self.get_selected_rule_id()
        self.detail_widget.populate(selected_id)
        self.update_pause_actions_visuals()

    def on_rules_table_context_menu(self, pos) -> None:
        selected_item = self.rules_table.itemAt(pos)
        if not selected_item: return
        rule_id = self.rules_table.item(selected_item.row(), 1).text()
        if not rule_id or rule_id not in rules_data: return
        menu = QMenu(self)
        act_edit = menu.addAction(
            self.load_icon_with_fallback("edit_icon.png", "document-edit", QStyle.SP_FileDialogDetailedView),
            "Editar Regra...")
        act_delete = menu.addAction(self.load_icon_with_fallback(None, "edit-delete", QStyle.SP_TrashIcon),
                                    "Remover Regra")
        menu.addSeparator()
        is_paused = rules_data[rule_id].get("paused", False)
        pause_icon = self.load_icon_with_fallback(None, "media-playback-start" if is_paused else "media-playback-pause",
                                                  QStyle.SP_MediaPlay if is_paused else QStyle.SP_MediaPause)
        act_pause_resume = menu.addAction(pause_icon, "Resumir Monitoramento" if is_paused else "Pausar Monitoramento")
        menu.addSeparator()
        act_open_input = menu.addAction(self.load_icon_with_fallback(None, "folder-open", QStyle.SP_DirOpenIcon),
                                        "Abrir Pasta de Entrada")
        act_open_output = menu.addAction(self.load_icon_with_fallback(None, "folder-remote", QStyle.SP_DriveNetIcon),
                                         "Abrir Pasta de Destino")
        action = menu.exec(self.rules_table.mapToGlobal(pos))
        if action == act_edit:
            self.edit_selected_rule(rule_id_to_edit=rule_id)
        elif action == act_delete:
            self.remove_selected_rule(rule_id_to_remove=rule_id)
        elif action == act_pause_resume:
            self.toggle_selected_rule_pause_state(rule_id_to_toggle=rule_id)
        elif action == act_open_input:
            QDesktopServices.openUrl(QUrl.fromLocalFile(rules_data[rule_id]["input"]))
        elif action == act_open_output:
            output_path_str = rules_data[rule_id]["output"]
            if output_path_str.startswith("\\\\"): ensure_persistent_connection(output_path_str)
            if not QDesktopServices.openUrl(QUrl.fromLocalFile(output_path_str)):
                QMessageBox.warning(self, "Abrir Pasta",
                                    f"Não foi possível abrir '{output_path_str}'. Verifique se o caminho é válido e acessível.")

    def update_rules_table_display(self) -> None:
        """
        Solicita uma atualização da tabela de regras.  A lógica de reconstrução
        da tabela é feita de forma atrasada pelo temporizador para suavizar
        múltiplas chamadas consecutivas e evitar sobrecarga na thread da GUI.
        """
        # Programamos uma execução do actualizador se ainda não houver uma
        # programada.  O atraso de 200 ms permite que várias actualizações sejam
        # combinadas numa única operação.
        if not self._rules_update_timer.isActive():
            self._rules_update_timer.start(200)

    def _perform_rules_table_update(self) -> None:
        """
        Executa a reconstrução da tabela de regras.  Esta função contém a
        implementação original de update_rules_table_display().  É chamada
        pelo temporizador de coalescência para reduzir a frequência de
        actualizações pesadas da GUI.
        """
        self.rules_table.setUpdatesEnabled(False)
        self.rules_table.setSortingEnabled(False)
        current_selection_id = self.get_selected_rule_id()
        # Remove regras cujas pastas de entrada ou saída foram apagadas manualmente
        removed_rids = []
        changed = False
        # Controle de rate limiting para logs de erro
        current_time = time.time()

        for rid, cfg in list(rules_data.items()):
            try:
                # Evita "pingar" caminhos de regras pausadas para não sobrecarregar rede/USB
                if cfg.get("paused", False):
                    continue
                inp = Path(cfg.get("input", ""))
                outp = Path(cfg.get("output", ""))
                missing = False
                # Verifica existência com tratamento de erro silencioso para rede
                try:
                    if inp and not inp.exists():
                        missing = True
                except (PermissionError, OSError):
                    pass  # Ignora erros de rede/permissão silenciosamente
                try:
                    if outp and not outp.exists():
                        missing = True
                except (PermissionError, OSError):
                    pass  # Ignora erros de rede/permissão silenciosamente
                if missing:
                    logging.warning(f"Caminho ausente para regra '{rid}'. Pausando em vez de remover.")
                    stop_watching_thread_for_rule(rid)
                    with files_status_lock:
                        files_status.pop(rid, None)
                    cfg["paused"] = True
                    cfg["missing_paths"] = True
                    rules_data[rid] = cfg
                    changed = True
            except Exception as _rem_err:
                # Rate limit: loga apenas uma vez a cada 5 minutos por regra
                last_log_key = f"_last_path_error_log_{rid}"
                last_log_time = getattr(self, last_log_key, 0)
                if current_time - last_log_time > 300:
                    logging.error(f"Erro ao verificar existência de pastas para regra '{rid}': {_rem_err}")
                    setattr(self, last_log_key, current_time)
        if changed:
            save_config(rules_data)
        current_rules_data = rules_data.copy()
        existing_rule_ids_in_table = {self.rules_table.item(r, 1).text() for r in range(self.rules_table.rowCount())}
        rows_to_remove = [row_idx for row_idx in range(self.rules_table.rowCount()) if
                          self.rules_table.item(row_idx, 1).text() not in current_rules_data]
        for row_idx in sorted(rows_to_remove, reverse=True): self.rules_table.removeRow(row_idx)
        row_map = {self.rules_table.item(r, 1).text(): r for r in range(self.rules_table.rowCount())}
        sorted_rules_items = sorted(current_rules_data.items(), key=lambda item: item[0])
        fmt_display_map = {"pdf searchable": "PDF Pesquisável", "pdf": "PDF Imagem", "pdf/a": "PDF/A", "jpeg": "JPEG",
                           "tiff": "TIFF", "text": "Texto (TXT)", "rtf": "RTF", "checksum": "Checksum", "docx": "DOCX",
                           "xlsx": "XLSX", "pptx": "PPTX"}
        new_selected_row = -1
        for rid, rinfo in sorted_rules_items:
            row_idx = row_map.get(rid, -1)
            if row_idx == -1:
                row_idx = self.rules_table.rowCount()
                self.rules_table.insertRow(row_idx)
                self.rules_table.setItem(row_idx, 1, QTableWidgetItem(rid))
                self.rules_table.setItem(row_idx, 2, QTableWidgetItem(rinfo.get("output", "N/A")))
            fmt_key = rinfo.get("output_format", "N/A").lower()
            logging.debug(
                f"MainWindow: Updating format for rule '{rid}'. Stored key: '{fmt_key}'. Display text: '{fmt_display_map.get(fmt_key, fmt_key.upper())}'")
            self.rules_table.setItem(row_idx, 4, self._centered_item(fmt_display_map.get(fmt_key, fmt_key.upper())))
            is_globally_paused, is_rule_locally_paused = not engine_running.is_set() or not self.engine_enabled, rinfo.get(
                "paused", False)
            status_text = "LICENÇA INVÁLIDA" if not self.engine_enabled else (
                "MOTOR PAUSADO" if not engine_running.is_set() else (
                    "REGRA PAUSADA" if is_rule_locally_paused else "ATIVA"))
            status_icon = QStyle.SP_DialogCancelButton if (
                    is_globally_paused or is_rule_locally_paused or not self.engine_enabled) else QStyle.SP_DialogApplyButton
            item_status = self.rules_table.item(row_idx, 0) or QTableWidgetItem()
            item_status.setText(status_text)
            item_status.setIcon(self.load_icon_with_fallback(None, "dialog-error" if (
                    is_globally_paused or is_rule_locally_paused or not self.engine_enabled) else "dialog-ok-apply",
                                                             status_icon))
            item_status.setTextAlignment(Qt.AlignCenter)
            if not self.rules_table.item(row_idx, 0): self.rules_table.setItem(row_idx, 0, item_status)
            q_count = 0
            with files_status_lock:
                if rid in files_status: q_count = sum(
                    1 for f_data in files_status[rid].values() if f_data.get("progress", 0) < 100)
            logging.debug(
                f"MainWindow: Updating queue count for rule '{rid}'. Found {q_count} files in queue (progress < 100%).")
            current_q_item = self.rules_table.item(row_idx, 3)
            if not current_q_item or current_q_item.text() != str(q_count):
                self.rules_table.setItem(row_idx, 3, self._centered_item(str(q_count)))
            overall_prog, num_files_for_prog = 0, 0
            with files_status_lock:
                if rid in files_status:
                    # Considere apenas arquivos em processamento (progress < 100)
                    active_statuses = [f_data for f_data in files_status[rid].values() if
                                       f_data.get("progress", 0) < 100]
                    overall_prog = sum(f_data.get("progress", 0) for f_data in active_statuses)
                    num_files_for_prog = len(active_statuses)
            avg_prog = int(overall_prog / num_files_for_prog) if num_files_for_prog > 0 else 0
            # Actualiza a barra de progresso.  Criamos o widget apenas uma vez por linha
            # e reutilizamos nas actualizações seguintes.
            cell_widget = self.rules_table.cellWidget(row_idx, 5)
            progress_bar = None
            if cell_widget is None:
                progress_bar = QProgressBar()
                progress_bar.setAlignment(Qt.AlignCenter)
                progress_bar.setMaximum(100)
                # Formato inicial será ajustado abaixo
                container_widget = QWidget()
                layout = QVBoxLayout()
                layout.setContentsMargins(0, 0, 0, 0)
                layout.addWidget(progress_bar)
                container_widget.setLayout(layout)
                self.rules_table.setCellWidget(row_idx, 5, container_widget)
            else:
                # Procura a QProgressBar existente
                progress_bar = cell_widget.findChild(QProgressBar)
            if progress_bar:
                progress_bar.setValue(avg_prog)
                # Mostra percentagem e quantidade de arquivos activos
                progress_bar.setFormat(f"{avg_prog}% ({num_files_for_prog} Arq.)")
            # Seleciona a linha se for a actualmente selecionada
            if rid == current_selection_id:
                new_selected_row = row_idx
        # fim do loop de regras
        # Reabilita ordenação e atualização da tabela
        self.rules_table.setSortingEnabled(True)
        self.rules_table.setUpdatesEnabled(True)
        # Seleciona novamente a linha previamente seleccionada
        if new_selected_row >= 0:
            self.rules_table.selectRow(new_selected_row)
            if rid == current_selection_id: new_selected_row = row_idx
        self.rules_table.resizeRowsToContents()
        self.rules_table.setSortingEnabled(True)
        self.rules_table.setUpdatesEnabled(True)
        self.update_pause_actions_visuals()
        self.status_bar.showMessage(
            f"Pronto. Regras: {len(rules_data)}. Motor: {'ATIVO' if engine_running.is_set() and self.engine_enabled else 'PAUSADO'}",
            3000)

    def _centered_item(self, text: str) -> QTableWidgetItem:
        item = QTableWidgetItem(text);
        item.setTextAlignment(Qt.AlignCenter);
        return item

    def _poll_license(self):
        """
        Verifica periodicamente se a licença foi removida/alterada no registro.
        Se inválida, fecha a aplicação imediatamente.
        """
        # Read the current license key from the registry
        self.license_key = read_license_key()
        # Check its validity
        current_validity = is_license_valid(self.license_key)

        # If the license was previously enabled but now is not valid
        if not current_validity and self.engine_enabled:
            logging.critical("Licença se tornou inválida durante o polling! Fechando aplicação.")
            QMessageBox.critical(
                self,
                "Licença Inválida",
                "Sua licença se tornou inválida ou foi removida.\n"
                "O aplicativo será encerrado."
            )
            logging.info('Licença inválida detectada: mantendo chave no Registro (não deletar automaticamente).')
            self.close()
            return

        # Update the engine_enabled state based on current validity
        self.engine_enabled = current_validity
        # Update the GUI to reflect the new state (e.g., disable buttons, update status)
        self.update_rules_table_display()

    def _check_license_on_startup(self):
        """
        Verifica a licença na inicialização e exibe o diálogo de ativação se necessário.
        """
        self.license_key = read_license_key()
        self.engine_enabled = is_license_valid(self.license_key)

        if not self.engine_enabled:
            logging.critical("Licença inválida ou ausente na inicialização. Bloqueando acesso.")
            QMessageBox.critical(
                self,
                "Licença Inválida",
                "Sua licença expirou ou é inválida, ou não foi encontrada.\n"
                "Insira uma nova chave para continuar."
            )
            # Ensure the invalid key is removed from the registry
            logging.info('Licença inválida detectada: mantendo chave no Registro (não deletar automaticamente).')
            # The engine remains disabled
            self._show_license_dialog(startup_check=True)
        else:
            logging.info("Licença válida. Motor de OCR habilitado.")
            self._start_engine_if_needed()
        self.update_rules_table_display()  # Atualiza o status das regras na tabela

    def _show_license_dialog(self, startup_check: bool = False):
        """
        Exibe o diálogo para o usuário inserir/atualizar a chave de licença.
        Se a licença for válida, primeiro mostra os detalhes.
        """
        current_license_info = get_license_info(self.license_key)
        is_license_currently_valid = is_license_valid(self.license_key)

        info_message = ""
        if current_license_info:
            dias, expira_em, tipo_licenca = current_license_info
            info_message = f"Sua licença atual é: {tipo_licenca}\n"
            info_message += f"Expira em: {expira_em.strftime('%d/%m/%Y')}.\n"
            if "Expirada" in tipo_licenca:
                info_message += "Status: Expirada.\n"
            elif "Expira em" in tipo_licenca:
                info_message += "Status: Válida.\n"
            else:  # For production licenses that are just "Produção"
                info_message += f"Status: Válida. Restam {dias} dias.\n"
        else:
            info_message = "Nenhuma licença válida encontrada."

        # If license is valid and it's NOT a startup check, just show info and return.
        if not startup_check and is_license_currently_valid:
            QMessageBox.information(self, "Informações da Licença", info_message)
            return

        # If we reach here, either the license is NOT valid, or it IS a startup check (where we always offer to input)
        # or the user explicitly clicked "Ativar Licença" (even if valid).
        info_message += "\nClique OK para inserir uma nova chave ou Cancelar para fechar."
        QMessageBox.information(self, "Informações da Licença", info_message)  # Show info before prompting for new key

        text, ok = QInputDialog.getText(
            self, "Ativar Licença", "Insira sua chave de licença (formato AAAAA-BBBBB):")

        if not ok:  # Usuário clicou em Cancelar ou fechou o diálogo
            if startup_check and not is_license_currently_valid:  # Only log/set engine_enabled if it was invalid at startup
                logging.info("Ativação de licença cancelada na inicialização. Motor permanecerá desabilitado.")
                self.engine_enabled = False
            return

        if is_license_valid(text):
            try:
                save_license_key(text)
                self.license_key = text  # Atualiza a chave em memória
                self.engine_enabled = True  # Marca o motor como habilitado
                QMessageBox.information(self, "Sucesso", "Licença ativada com sucesso!")

                self._start_engine_if_needed()
            except OSError as e:
                QMessageBox.critical(self, "Erro de Registro",
                                     f"Não foi possível salvar a licença no Registro. Erro: {e}\n"
                                     "Tente executar o aplicativo como administrador.")
                logging.error(f"Erro ao salvar licença no Registro: {e}")
                self.engine_enabled = False  # Falha ao salvar, então o motor não pode ser ativado
            except Exception as e:
                QMessageBox.critical(self, "Erro Inesperado", f"Ocorreu um erro inesperado: {e}")
                logging.error(f"Erro inesperado ao ativar licença: {e}")
                self.engine_enabled = False
        else:
            QMessageBox.critical(self, "Erro", "Chave de licença inválida ou expirada.")
            logging.warning(f"Tentativa de ativação de licença inválida: {text}")
            self.engine_enabled = False  # Chave inválida, motor desabilitado

        self.update_rules_table_display()  # Atualiza o status das regras na tabela

    def _start_engine_if_needed(self):
        """
        Inicia o motor se a licença for válida e o motor ainda não estiver rodando.
        """
        if self.engine_enabled and not engine_running.is_set():
            logging.info("Motor estava parado. Reiniciando após validação da licença...")
            initialize_engine()
            for rid, r_config in rules_data.items():
                if not r_config.get("paused", False): start_watching_thread_for_rule(rid)
        else:
            logging.info("Motor já está rodando ou licença não está habilitada. Não é necessário iniciar.")

    def create_new_rule(self) -> None:
        if not self.engine_enabled:
            QMessageBox.warning(self, "Licença Inválida", "Não é possível criar novas regras com uma licença inválida.")
            return
        dlg = RuleDialog(self)
        if dlg.exec() == QDialog.Accepted:
            new_rid, data_for_new_rule = dlg.get_rule_data()
            rules_data[new_rid] = data_for_new_rule
            save_config(rules_data)
            try:
                Path(data_for_new_rule["input"]).mkdir(parents=True, exist_ok=True)
            except Exception as e_mkdir:
                QMessageBox.critical(self, "Erro",
                                     f"Não foi possível criar pasta de entrada '{data_for_new_rule['input']}': {e_mkdir}")
            if not data_for_new_rule.get("paused", False): start_watching_thread_for_rule(new_rid)
            self.update_rules_table_display()
            QMessageBox.information(self, "Regra Criada", f"Nova regra '{new_rid}' adicionada.")

    def edit_selected_rule(self, rule_id_to_edit: str | None = None) -> None:
        if not self.engine_enabled:
            QMessageBox.warning(self, "Licença Inválida", "Não é possível editar regras com uma licença inválida.")
            return
        if not rule_id_to_edit: rule_id_to_edit = self.get_selected_rule_id()
        if not rule_id_to_edit: QMessageBox.information(self, "Editar", "Nenhuma regra selecionada."); return
        if rule_id_to_edit not in rules_data: QMessageBox.critical(self, "Erro",
                                                                   f"Regra '{rule_id_to_edit}' não encontrada."); return
        original_rule_data = rules_data[rule_id_to_edit]
        dlg = RuleDialog(self, rule_id=rule_id_to_edit, existing_data=original_rule_data)
        if dlg.exec() == QDialog.Accepted:
            updated_rid, updated_data = dlg.get_rule_data()
            renamed = (rule_id_to_edit != updated_rid)
            if renamed:
                stop_watching_thread_for_rule(rule_id_to_edit)
                del rules_data[rule_id_to_edit]
                with files_status_lock:
                    files_status.pop(rule_id_to_edit, None)
            elif updated_data.get("paused", False) and not original_rule_data.get("paused", False):
                stop_watching_thread_for_rule(updated_rid)
            rules_data[updated_rid] = updated_data
            save_config(rules_data)
            try:
                Path(updated_data["input"]).mkdir(parents=True, exist_ok=True)
            except Exception as e_mkdir:
                QMessageBox.warning(self, "Aviso",
                                    f"Problema ao criar/verificar pasta de entrada '{updated_data['input']}': {e_mkdir}")
            if not updated_data.get("paused", False): start_watching_thread_for_rule(updated_rid)
            self.update_rules_table_display()
            self.detail_widget.populate(updated_rid if not renamed else None)
            QMessageBox.information(self, "Regra Editada", f"Regra '{updated_rid}' atualizada.")

    def remove_selected_rule(self, rule_id_to_remove: str | None = None) -> None:
        if not self.engine_enabled:
            QMessageBox.warning(self, "Licença Inválida", "Não é possível remover regras com uma licença inválida.")
            return
        if not rule_id_to_remove: rule_id_to_remove = self.get_selected_rule_id()
        if not rule_id_to_remove: QMessageBox.information(self, "Remover", "Nenhuma regra selecionada."); return
        if rule_id_to_remove not in rules_data: QMessageBox.critical(self, "Erro",
                                                                     f"Regra '{rule_id_to_remove}' não encontrada."); return
        if QMessageBox.question(self, "Confirmar Remoção",
                                f"Tem certeza que deseja remover a regra '{rule_id_to_remove}'?\nA pasta de entrada e seus arquivos serão removidos.",
                                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.Yes:
            if rule_id_to_remove in rules_data: rules_data[rule_id_to_remove]["paused"] = True
            stop_watching_thread_for_rule(rule_id_to_remove)
            with files_status_lock:
                files_status.pop(rule_id_to_remove, None)
            # Apaga a pasta de entrada no disco
            try:
                input_dir = Path(rules_data[rule_id_to_remove]['input'])
                if input_dir.exists():
                    shutil.rmtree(input_dir, ignore_errors=True)
                    logging.info(f"Pasta de entrada excluída: {input_dir}")
            except Exception as e:
                logging.error(f"Falha ao excluir pasta de entrada da regra '{rule_id_to_remove}': {e}")
            rules_data.pop(rule_id_to_remove, None)
            save_config(rules_data)
            logging.info(f"Regra '{rule_id_to_remove}' removida pelo usuário.")
            self.update_rules_table_display()
            self.detail_widget.populate(None)

    def toggle_engine_pause_state(self) -> None:
        # A declaração de global deve aparecer antes de qualquer referência ao nome
        # dentro da função. Colocamos aqui no início para evitar erros de
        # "name used prior to global declaration".
        global process_executor
        if not self.engine_enabled:
            QMessageBox.warning(self, "Licença Inválida", "Não é possível controlar o motor com uma licença inválida.")
            return
        if engine_running.is_set():
            if QMessageBox.question(self, "Pausar Motor",
                                    "Pausar todo o processamento e monitoramento de novas pastas?",
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.Yes:
                # Pausa o motor global
                engine_running.clear()
                logging.info("Motor global PAUSADO pelo usuário.")
                # NOVO: refletir pausa global em TODAS as regras (UI e coerência)
                try:
                    any_change = False
                    for _rid, _cfg in list(rules_data.items()):
                        if not _cfg.get("paused", False):
                            _cfg["paused"] = True
                            rules_data[_rid] = _cfg
                            any_change = True
                    if any_change:
                        save_config(rules_data)
                except Exception as _e_mark_pause:
                    logging.warning(f"Falha ao marcar regras como pausadas (pausa global): {_e_mark_pause}")

                # Finaliza quaisquer processamentos em andamento e impede novos enquanto pausado
                if process_executor:
                    try:
                        process_executor.shutdown(wait=False)
                    except Exception as e_shutdown:
                        logging.warning(f"Falha ao encerrar executor de processos ao pausar: {e_shutdown}")
                    process_executor = None
                # Limpa filas internas para permitir reprocessamento ao retomar
                with queued_files_lock:
                    queued_files_set.clear()
                    processing_files_set.clear()
                # Drena a fila de arquivos para que os itens enfileirados não permaneçam
                # aguardando processamento enquanto o motor está pausado. Eles serão
                # re‑enfileirados por varredura ao retomar.
                try:
                    while True:
                        prio, cnt, rid, fpath = file_queue.get_nowait()
                        file_queue.task_done()
                except queue.Empty:
                    pass
        else:
            # Retoma o motor global
            engine_running.set()
            logging.info("Motor global RESUMIDO pelo usuário.")
            # Reinicializa o executor de processos para retomar processamento
            try:
                detected_cpus = os.cpu_count() or 1
            except Exception:
                detected_cpus = 1
            process_executor = ProcessPoolExecutor(max_workers=max(1, detected_cpus))
            # NOVO: despausar TODAS as regras e limpar sinalizadores de caminhos ausentes
            try:
                any_change = False
                for rid, cfg in list(rules_data.items()):
                    if cfg.get("paused", False) or cfg.get("missing_paths", False):
                        cfg["paused"] = False
                        cfg.pop("missing_paths", None)
                        rules_data[rid] = cfg
                        any_change = True
                if any_change:
                    save_config(rules_data)
            except Exception as _e_mark_resume:
                logging.warning(f"Falha ao despausar regras ao retomar motor global: {_e_mark_resume}")
            # Reinicia observadores para TODAS as regras (sem checar paused)
            for rid in list(rules_data.keys()):
                start_watching_thread_for_rule(rid)
            # Escaneia e enfileira quaisquer arquivos existentes nas pastas de entrada
            scan_and_enqueue_existing_files()
        self.update_rules_table_display()

    def toggle_selected_rule_pause_state(self, rule_id_to_toggle: str | None = None) -> None:
        if not self.engine_enabled:
            QMessageBox.warning(self, "Licença Inválida",
                                "Não é possível pausar/resumir regras com uma licença inválida.")
            return
        if not rule_id_to_toggle: rule_id_to_toggle = self.get_selected_rule_id()
        if not rule_id_to_toggle or rule_id_to_toggle not in rules_data: return
        new_paused_state = not rules_data[rule_id_to_toggle].get("paused", False)
        rules_data[rule_id_to_toggle]["paused"] = new_paused_state
        save_config(rules_data)
        logging.info(f"Regra '{rule_id_to_toggle}' {'PAUSADA' if new_paused_state else 'RESUMIDA'} pelo usuário.")
        if not new_paused_state:
            # Reinicia observador e enfileira arquivos pendentes para a regra retomada
            start_watching_thread_for_rule(rule_id_to_toggle)
            # Escaneia arquivos existentes apenas para regras ativas
            try:
                scan_and_enqueue_existing_files()
            except Exception as e_scan:
                logging.error(f"Falha ao escanear arquivos ao retomar regra '{rule_id_to_toggle}': {e_scan}")
        self.update_rules_table_display()

    def open_credentials_dialog(self) -> None:
        NetworkCredentialDialog(self).exec()

    def open_zonal_ocr_dialog(self):
        """Abre o diálogo de configuração do reconhecimento de texto zonal."""
        if not self.engine_enabled:
            QMessageBox.warning(self, "Licença Inválida",
                                "Não é possível configurar o reconhecimento de texto zonal com uma licença inválida.")
            return
        if not rules_data:
            QMessageBox.information(self, "Reconhecimento de Texto Zonal",
                                    "Crie pelo menos uma regra antes de configurar o reconhecimento de texto zonal.")
            return
        dlg = ZonalOCRDialog(self, rules_data=rules_data)
        dlg.exec()
        # O diálogo salva a config, não precisa fazer nada aqui.

    def update_license_display(self) -> None:
        """
        Atualiza o estado interno da licença e a interface, se necessário.
        Este método não mais atualiza um QLabel na toolbar, mas mantém o estado.
        """
        self.license_key = read_license_key()
        self.engine_enabled = is_license_valid(self.license_key)
        # Nenhuma atualização de QLabel aqui. A lógica de exibição está em _show_license_dialog.

    def show_about_dialog(self) -> None:
        QMessageBox.about(self, f"Sobre {APP_NAME}",
                          f"<b>{APP_NAME} - Processador de Documentos e Gerenciador de Arquivos</b><br><br>Versão {APP_VERSION}<br>Automatiza o processamento de documentos e imagens.<br><br>Este software utiliza tecnologias de processamento de documentos e reconhecimento de texto para oferecer suas funcionalidades.")

    def on_tray_icon_activated(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        if reason in (QSystemTrayIcon.DoubleClick, QSystemTrayIcon.Trigger):
            self.show_window_from_tray()

    def show_window_from_tray(self) -> None:
        self.showNormal();
        self.activateWindow();
        self.raise_()

    def closeEvent(self, event: QCloseEvent) -> None:
        if self.is_quitting_app: event.accept(); return
        if self.tray_icon.isVisible():
            self.hide()
            self.tray_icon.showMessage(f"{APP_NAME} Minimizado", f"O {APP_NAME} continua executando em segundo plano.",
                                       self.tray_icon.icon(), 3000)
            event.ignore()
        else:
            self.confirm_quit_application()
            if not self.is_quitting_app: event.ignore()

    def confirm_quit_application(self) -> None:
        if self.is_quitting_app: return
        if QMessageBox.question(self, f'Sair do {APP_NAME}',
                                "Tem certeza que deseja fechar e parar todo o processamento?",
                                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.Yes:
            self.is_quitting_app = True
            logging.info("Usuário confirmou sair. Encerrando aplicação...")
            self.tray_icon.hide()
            self.status_bar.showMessage("Encerrando... Por favor, aguarde.", 0)
            QApplication.setOverrideCursor(Qt.WaitCursor)
            shutdown_engine()
            QApplication.restoreOverrideCursor()
            QApplication.instance().quit()
        else:
            logging.info("Usuário cancelou sair.")
            self.is_quitting_app = False

    # ==============================================================================
    # Ponto de entrada
    # ==============================================================================
    def open_split_merge_tool(self) -> None:
        """Abre o utilitário OCTecSplitMerge (EXE no mesmo diretório da instalação).
        Em ambiente de desenvolvimento, tenta rodar o script Python como fallback."""
        try:
            candidates = []
            if getattr(sys, "frozen", False):
                app_dir = Path(sys.executable).parent
                candidates.append(app_dir / "OCTecSplitMerge.exe")
            else:
                # Desenvolvimento: procurar EXE no dist e, se não existir, usar o script Python
                candidates.append(BASE_DIR / "dist" / "OCTec" / "OCTecSplitMerge.exe")

            # 1) Tenta EXEs conhecidos
            for exe_path in candidates:
                if exe_path and exe_path.exists():
                    self._splitmerge_proc = getattr(self, "_splitmerge_proc", None)
                    if self._splitmerge_proc is None:
                        self._splitmerge_proc = QProcess(self)
                    self._splitmerge_proc.setProgram(str(exe_path))
                    self._splitmerge_proc.setArguments([])
                    self._splitmerge_proc.setProcessChannelMode(QProcess.MergedChannels)
                    self._splitmerge_proc.start()
                    if not self._splitmerge_proc.waitForStarted(3000):
                        QMessageBox.warning(self, "OCTec", "Não foi possível iniciar o utilitário de Split/Merge.")
                    return

            # 2) Fallback no desenvolvimento: script Python
            dev_script = BASE_DIR / "source" / "OCTecSplitMerge.py"
            if dev_script.exists():
                self._splitmerge_proc = QProcess(self)
                self._splitmerge_proc.setProgram(sys.executable)
                self._splitmerge_proc.setArguments([str(dev_script)])
                self._splitmerge_proc.setProcessChannelMode(QProcess.MergedChannels)
                self._splitmerge_proc.start()
                if not self._splitmerge_proc.waitForStarted(3000):
                    QMessageBox.warning(self, "OCTec",
                                        "Não foi possível iniciar o utilitário de Split/Merge (fallback Python).")
                return

            QMessageBox.critical(self, "OCTec",
                                 "Utilitário OCTecSplitMerge não encontrado. Verifique a instalação.")
        except Exception as e:
            logging.exception(f"Falha ao abrir OCTecSplitMerge: {e}")
            QMessageBox.critical(self, "OCTec", f"Erro ao abrir utilitário de Split/Merge:\n{e}")


def run_octec_gui(start_in_tray: bool = False):
    logging.info(f"Iniciando {APP_NAME} v{APP_VERSION} em modo GUI...")
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # A inicialização do motor agora é feita dentro de MainWindow após a validação da licença
    window = MainWindow()
    (
        window.hide() if start_in_tray else window.show()
    )
    exit_code = app.exec()

    if engine_running.is_set():
        logging.info("Loop de eventos da aplicação finalizado, garantindo parada do motor...")
        shutdown_engine()
    logging.info(f"Aplicação {APP_NAME} encerrada com código: {exit_code}")
    sys.exit(exit_code)


def run_service_headless():
    """
    Executa o motor sem GUI. Para uso como SERVIÇO via NSSM: OCTec.exe --service
    """
    logging.info("Inicializando OCTec em modo SERVIÇO (headless)...")
    try:
        initialize_engine()
    except Exception as e:
        logging.critical(f"Falha ao inicializar engine em modo serviço: {e}", exc_info=True)
        # Em caso de falha crítica, retorna código não-zero
        sys.exit(2)

    # Handlers para parada graciosa (quando o Windows/NSSM sinalizar)
    def _graceful_stop(signum, frame):
        logging.info(f"Sinal de parada recebido ({signum}). Encerrando engine...")
        try:
            shutdown_engine()
        finally:
            pass

    for s in (getattr(signal, "SIGINT", None), getattr(signal, "SIGTERM", None), getattr(signal, "SIGBREAK", None)):
        if s is not None:
            try:
                signal.signal(s, _graceful_stop)
            except Exception:
                # Pode não estar disponível em todos os contextos
                pass

    # Loop de vida do serviço
    try:
        while engine_running.is_set():
            time.sleep(1.0)
    except Exception as e:
        logging.exception(f"Erro no loop do serviço: {e}")
    finally:
        if engine_running.is_set():
            shutdown_engine()
        logging.info("Serviço OCTec finalizado.")


# ======================= ASSINATURA PDF - HELPERS =======================

def is_windows() -> bool:
    return sys.platform.startswith("win")


def powershell_exists() -> bool:
    from shutil import which
    return which("powershell") is not None


def run_powershell(command: str, *, expect_json: bool = False, allow_ui: bool = False):
    if not powershell_exists():
        raise RuntimeError("PowerShell não encontrado no sistema.")
    base = ["powershell", "-NoProfile"]
    if not allow_ui:
        base.append("-NonInteractive")
    base += ["-Command", command]
    proc = subprocess.run(base, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip() or "Falha desconhecida no PowerShell.")
    if expect_json:
        out = proc.stdout.strip()
        return json.loads(out) if out else []
    return proc.stdout


def human_thumbprint(tp: str) -> str:
    return "".join(ch for ch in (tp or "") if ch.isalnum()).upper()


def extract_cn_from_subject(subject: str) -> str:
    if not subject: return ""
    parts = subject.replace("/", ",").split(",")
    for p in parts:
        p = p.strip()
        if p.upper().startswith("CN="):
            return p[3:].strip()
    return subject.strip()


@dataclass
class CertificateInfo:
    thumbprint: str
    subject: str
    friendly_name: str
    not_before: str
    not_after: str


def list_currentuser_my_certs() -> list[CertificateInfo]:
    ps = r"""
$certs = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.HasPrivateKey -eq $true } |
    Select-Object Thumbprint, Subject, FriendlyName, NotBefore, NotAfter
$certs | ConvertTo-Json -Depth 4
"""
    result = run_powershell(ps, expect_json=True, allow_ui=False)
    if isinstance(result, dict):
        result = [result]
    items: list[CertificateInfo] = []
    for c in result or []:
        items.append(
            CertificateInfo(
                thumbprint=human_thumbprint(c.get("Thumbprint", "")),
                subject=c.get("Subject", ""),
                friendly_name=(c.get("FriendlyName") or "").strip(),
                not_before=str(c.get("NotBefore", "")),
                not_after=str(c.get("NotAfter", "")),
            )
        )
    return items


def export_cert_to_pfx_temp(thumbprint: str) -> tuple[str, str]:
    tp = human_thumbprint(thumbprint)
    if not tp:
        raise ValueError("Thumbprint inválido.")
    tmpdir = Path(tempfile.mkdtemp(prefix="octec_pfx_"))
    pfx_path = str(tmpdir / "temp_export.pfx")

    def _do_export(allow_ui: bool) -> str | None:
        password = ''.join(
            [random.choice("ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%^&*") for _ in range(24)])
        ps = fr"""
$ErrorActionPreference = 'Stop'
$thumb = '{tp}'
$pwd = ConvertTo-SecureString '{password}' -AsPlainText -Force
$cert = Get-ChildItem Cert:\CurrentUser\My\ | Where-Object {{ $_.Thumbprint -replace '[^0-9A-Fa-f]' -eq $thumb }}
if (-not $cert) {{ throw 'Certificado não encontrado no CurrentUser\My.' }}
Export-PfxCertificate -Cert $cert -FilePath '{pfx_path}' -Password $pwd | Out-Null
"""
        try:
            run_powershell(ps, expect_json=False, allow_ui=allow_ui)
            return password
        except Exception:
            return None

    pwd = _do_export(allow_ui=True)
    if not pwd:
        pwd = _do_export(allow_ui=False)
        if not pwd:
            raise RuntimeError(
                "Falha ao exportar PFX temporário. Verifique se a chave é exportável e se a UI foi autorizada.")
    return pfx_path, pwd


def stamp_all_pages(input_pdf: str, output_pdf: str, stamp_text: str):
    doc = fitz.open(input_pdf)
    margin = 15
    stamp_width = 140
    stamp_height = 24
    border_color = (0, 0, 0)
    text_color = (0, 0, 0)
    font_size = 7
    font_name = "helv"
    try:
        # Timestamp local (BR) no formato dd/mm/aaaa HH:MM:SS
        _ts = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # Mantém a lógica original de 2 linhas: título e conteúdo.
        # Acrescenta o timestamp na mesma linha do conteúdo para manter o layout.
        parts = stamp_text.split(':', 1)
        if len(parts) == 2:
            second_line = parts[1].strip()
            text_to_draw = parts[0] + ':' + "\n" + f"{second_line} - {_ts}"
        else:
            # Caso geral: preserva o texto e adiciona o timestamp na linha de baixo.
            text_to_draw = stamp_text + "\n" + _ts

        for page in doc:
            page_box = page.bound()
            rect = fitz.Rect(
                page_box.x1 - stamp_width - margin,
                page_box.y0 + margin,
                page_box.x1 - margin,
                page_box.y0 + stamp_height + margin
            )
            page.draw_rect(rect, color=border_color, width=0.8, fill_opacity=0.6)
            page.insert_textbox(rect, text_to_draw, fontname=font_name, fontsize=font_size,
                                fontfile=None, color=text_color, align=fitz.TEXT_ALIGN_CENTER)
        doc.save(output_pdf, deflate=True, clean=True)
    finally:
        doc.close()


def sign_pdf_with_endesive(stamped_pdf_path: str, signed_pdf_path: str, pfx_path: str, pfx_password: str):
    with open(stamped_pdf_path, "rb") as f:
        pdf_bytes = f.read()
    date = datetime.now(timezone.utc).strftime("D:%Y%m%d%H%M%S+00'00'")
    dct = {
        "sigflags": 3,
        "sigflagsft": 132,
        "sigpage": 0,
        "location": "BR",
        "signingdate": date,
        "reason": "Assinatura Digital ICP-Brasil A1",
    }
    with open(pfx_path, "rb") as fp:
        key, cert, addl = pkcs12.load_key_and_certificates(fp.read(), pfx_password.encode("utf-8"))
    signature_bytes = cms.sign(pdf_bytes, dct, key, cert, addl or [], "sha256")
    with open(signed_pdf_path, "wb") as out:
        out.write(pdf_bytes)
        out.write(signature_bytes)


def _cleanup_temp_file(path_str: str | None):
    try:
        if path_str and os.path.isfile(path_str):
            tmpdir = os.path.dirname(path_str)
            try:
                os.remove(path_str)
            except Exception:
                pass
            try:
                os.rmdir(tmpdir)
            except Exception:
                pass
    except Exception:
        pass


def _pdf_has_signature(pdf_path: Path) -> bool:
    try:
        with open(pdf_path, 'rb') as f:
            data = f.read()
        return (b'/ByteRange' in data) and (b'/Type' in data and b'/Sig' in data)
    except Exception:
        return False


def maybe_sign_pdf(rule_id: str, file_name: str, rule_options: dict, final_pdf_path: Path,
                   local_processing_temp_dir: Path) -> Path:
    try:
        out_fmt = (rule_options.get("output_format", "") or "").lower()
        if out_fmt not in ("pdf", "pdf/a", "pdf searchable"):
            return final_pdf_path
        if not rule_options.get("sign_pdf"):
            return final_pdf_path
        thumb = (rule_options.get("sign_thumbprint") or "").strip()
        if not thumb:
            debug_log(rule_id, file_name, "Assinatura habilitada, mas nenhum certificado foi selecionado.")
            return final_pdf_path
        if not is_windows() or not powershell_exists():
            debug_log(rule_id, file_name, "Assinatura ignorada (requer Windows + PowerShell).")
            return final_pdf_path

        display_name = (rule_options.get("sign_display_name") or "").strip()
        if not display_name:
            for c in list_currentuser_my_certs():
                if human_thumbprint(c.thumbprint) == human_thumbprint(thumb):
                    display_name = c.friendly_name or extract_cn_from_subject(c.subject)
                    break
        if not display_name:
            display_name = "Certificado selecionado"

        update_file_status(rule_id, file_name, "Assinando PDF", progress=96)
        debug_log(rule_id, file_name, "Exportando PFX temporário…")
        pfx_path, pfx_pwd = export_cert_to_pfx_temp(thumb)
        try:
            # Se o formato de saída for PDF/A, assumimos que o carimbo já foi
            # aplicado anteriormente (durante a conversão para PDF/A).  Reaplicar
            # um carimbo aqui poderia introduzir fontes sem incorporação e
            # quebrar a conformidade PDF/A.  Portanto, apenas assinamos o PDF
            # resultante diretamente, sem adicionar um novo carimbo.
            if out_fmt == "pdf/a":
                signed = local_processing_temp_dir / f"{final_pdf_path.stem}.pdf"
                sign_pdf_with_endesive(str(final_pdf_path), str(signed), pfx_path, pfx_pwd)
                return signed
            else:
                # Para PDFs normais ou pesquisáveis, aplica um carimbo visível
                stamped = local_processing_temp_dir / f"{final_pdf_path.stem}_stamped.pdf"
                stamp_all_pages(str(final_pdf_path), str(stamped), f"Assinado digitalmente por: {display_name}")
                signed = local_processing_temp_dir / f"{final_pdf_path.stem}.pdf"
                sign_pdf_with_endesive(str(stamped), str(signed), pfx_path, pfx_pwd)
                try:
                    if stamped.exists():
                        stamped.unlink(missing_ok=True)
                except Exception:
                    pass
                return signed
        finally:
            _cleanup_temp_file(pfx_path)

    except Exception as e:
        debug_log(rule_id, file_name, f"Falha ao assinar PDF: {e}")
        return final_pdf_path


# =================== FIM - ASSINATURA PDF - HELPERS =====================


if __name__ == "__main__":
    import multiprocessing as mp

    # Necessário em executáveis (PyInstaller) no Windows
    if getattr(sys, "frozen", False):
        mp.freeze_support()

    # IMPORTANTÍSSIMO: só a instância principal inicia a GUI/motor
    if mp.current_process().name != "MainProcess":
        sys.exit(0)


    def global_exception_hook(exctype, value, tb):
        logging.critical(
            f"Exceção não tratada (global_exception_hook): {exctype.__name__}: {value}",
            exc_info=(exctype, value, tb)
        )


    sys.excepthook = global_exception_hook

    run_octec_gui()
