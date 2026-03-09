# -*- coding: utf-8 -*-
"""
OCTecSplitMerge — Unir/Dividir PDFs, com DnD real
- Tema dark/blue moderno
- UNIR: arraste PDFs (de qualquer pasta) e REORDENE por arrastar
- DIVIDIR: arraste PDFs e exporte para PASTA ou ZIP (um PDF por página)
- Log + barra de progresso
"""
import sys
import os
import shutil
import zipfile
import tempfile
import traceback
from pathlib import Path
from typing import List

import fitz  # PyMuPDF
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QProgressBar, QTextEdit, QLabel, QMessageBox, QListWidget,
    QListWidgetItem, QTabWidget, QRadioButton, QButtonGroup
)
from PySide6.QtGui import QIcon, QDragEnterEvent, QDropEvent


# ------------------------- Helpers -------------------------
def human_size(n: int) -> str:
    units = ["B", "KB", "MB", "GB"]
    s = float(n)
    for u in units:
        if s < 1024 or u == units[-1]:
            return f"{s:.1f} {u}".replace(".0 ", " ")
        s /= 1024.0


def iter_pdfs_from_paths(paths: List[Path]) -> List[Path]:
    """Expande diretórios (não recursivo) e normaliza PDFs únicos."""
    out: List[Path] = []
    for p in paths:
        if p.is_dir():
            for child in sorted(p.iterdir()):
                if child.suffix.lower() == ".pdf" and child.is_file():
                    out.append(child.resolve())
        elif p.suffix.lower() == ".pdf" and p.is_file():
            out.append(p.resolve())
    return out


def get_icon_path() -> Path | None:
    """
    IDE (projeto/OCTec/source): usa projeto/OCTec/resources/app_icon.ico
    Congelado: {app}\_internal\resources\app_icon.ico (fallback {app}\resources\app_icon.ico)
    """
    try:
        if getattr(sys, "frozen", False):
            base = Path(sys.executable).parent
            candidates = [
                base / "_internal" / "resources" / "app_icon.ico",
                base / "resources" / "app_icon.ico",
            ]
        else:
            script_dir = Path(__file__).resolve().parent           # ...\OCTec\source
            project_root = script_dir.parent                       # ...\OCTec
            candidates = [project_root / "resources" / "app_icon.ico"]
        for c in candidates:
            if c.exists():
                return c
    except Exception:
        pass
    return None


# ------------------------- Drop-enabled list -------------------------
class DropList(QListWidget):
    def __init__(self, allow_reorder: bool = False, parent=None):
        super().__init__(parent)
        self.allow_reorder = allow_reorder
        self.setAcceptDrops(True)
        self.setSelectionMode(QListWidget.ExtendedSelection)
        self.setAlternatingRowColors(False)
        self.setSpacing(2)
        self.setDropIndicatorShown(True)
        self.setDefaultDropAction(Qt.CopyAction)

        if allow_reorder:
            # Arrastar para reordenar dentro da lista
            self.setDragEnabled(True)
            self.setDragDropMode(QListWidget.InternalMove)
        else:
            self.setDragDropMode(QListWidget.NoDragDrop)

        self.setToolTip("Arraste PDFs ou pastas aqui")
        self.setMinimumHeight(180)

    def dragEnterEvent(self, e: QDragEnterEvent):
        if e.mimeData().hasUrls() or (self.allow_reorder and e.source() is self):
            e.acceptProposedAction()
        else:
            e.ignore()

    def dragMoveEvent(self, e: QDragEnterEvent):
        if e.mimeData().hasUrls() or (self.allow_reorder and e.source() is self):
            e.acceptProposedAction()
        else:
            e.ignore()

    def dropEvent(self, e: QDropEvent):
        # Se for arrasto interno (reordenação), deixa o QListWidget cuidar
        if self.allow_reorder and not e.mimeData().hasUrls():
            super().dropEvent(e)
            return

        if not e.mimeData().hasUrls():
            e.ignore()
            return

        paths: List[Path] = []
        for url in e.mimeData().urls():
            try:
                p = Path(url.toLocalFile())
                paths.append(p)
            except Exception:
                pass
        self.add_paths(paths)
        e.acceptProposedAction()

    # API utilitária
    def add_paths(self, paths: List[Path]):
        files = iter_pdfs_from_paths(paths)
        existing = {Path(self.item(i).data(Qt.UserRole)) for i in range(self.count())}
        for f in files:
            if f not in existing:
                it = QListWidgetItem(f.name)
                it.setToolTip(str(f))
                it.setData(Qt.UserRole, str(f))
                self.addItem(it)
                existing.add(f)

    def files(self) -> List[Path]:
        return [Path(self.item(i).data(Qt.UserRole)) for i in range(self.count())]

    def remove_selected(self):
        for it in self.selectedItems():
            self.takeItem(self.row(it))

    def clear_all(self):
        self.clear()


# ------------------------- Workers -------------------------
class MergeWorker(QThread):
    progressed = Signal(int)           # 0..100
    message = Signal(str)
    finished_ok = Signal(str)          # caminho do PDF
    failed = Signal(str)

    def __init__(self, inputs: List[Path], output: Path, parent=None):
        super().__init__(parent)
        self.inputs = inputs
        self.output = output

    def run(self):
        try:
            if not self.inputs:
                self.failed.emit("Nenhum PDF selecionado.")
                return
            self.message.emit(f"Mesclando {len(self.inputs)} arquivo(s).")

            # Pré-contagem de páginas
            total_pages = 0
            for p in self.inputs:
                with fitz.open(str(p)) as d:
                    total_pages += d.page_count

            out_doc = fitz.open()
            done_pages = 0
            for p in self.inputs:
                self.message.emit(f"Adicionando: {p.name}")
                with fitz.open(str(p)) as d:
                    out_doc.insert_pdf(d)
                    done_pages += d.page_count
                    self.progressed.emit(min(99, int(done_pages * 100 / max(1, total_pages))))
            self.output.parent.mkdir(parents=True, exist_ok=True)
            out_doc.save(str(self.output), garbage=3, deflate=True, clean=True)
            out_doc.close()

            self.progressed.emit(100)
            self.message.emit(f"OK: gerado {self.output.name} ({human_size(self.output.stat().st_size)})")
            self.finished_ok.emit(str(self.output))
        except Exception:
            self.failed.emit(traceback.format_exc())


class SplitWorker(QThread):
    progressed = Signal(int)
    message = Signal(str)
    finished_ok = Signal(str)   # pasta de saída OU zip gerado
    failed = Signal(str)

    def __init__(self, inputs: List[Path], out_mode: str, out_path: Path, parent=None):
        """
        out_mode: 'folder' | 'zip'
        out_path: se 'folder', é a pasta destino; se 'zip', é o arquivo .zip
        """
        super().__init__(parent)
        self.inputs = inputs
        self.out_mode = out_mode
        self.out_path = out_path

    def _split_all_to_dir(self, base_dir: Path) -> None:
        total_pages = 0
        for p in self.inputs:
            with fitz.open(str(p)) as d:
                total_pages += d.page_count

        done_pages = 0
        for pdf_path in self.inputs:
            self.message.emit(f"Lendo: {pdf_path.name}")
            with fitz.open(str(pdf_path)) as src:
                subdir = base_dir / pdf_path.stem
                subdir.mkdir(parents=True, exist_ok=True)
                for i in range(src.page_count):
                    dst = fitz.open()
                    dst.insert_pdf(src, from_page=i, to_page=i)
                    out_name = f"{pdf_path.stem}_p{i+1:04d}.pdf"
                    out_file = subdir / out_name
                    dst.save(str(out_file), garbage=3, deflate=True, clean=True)
                    dst.close()
                    done_pages += 1
                    self.progressed.emit(min(99, int(done_pages * 100 / max(1, total_pages))))
        self.progressed.emit(100)

    def run(self):
        try:
            if not self.inputs:
                self.failed.emit("Nenhum PDF selecionado.")
                return

            if self.out_mode == "folder":
                self.message.emit(f"Dividindo {len(self.inputs)} arquivo(s) em: {self.out_path}")
                self.out_path.mkdir(parents=True, exist_ok=True)
                self._split_all_to_dir(self.out_path)
                self.message.emit(f"OK: arquivos gerados em {self.out_path}")
                self.finished_ok.emit(str(self.out_path))

            elif self.out_mode == "zip":
                self.message.emit(f"Dividindo {len(self.inputs)} arquivo(s) para ZIP: {self.out_path.name}")
                temp_dir = Path(tempfile.mkdtemp(prefix="octec_split_"))
                try:
                    self._split_all_to_dir(temp_dir)
                    # Compacta mantendo estrutura por arquivo
                    with zipfile.ZipFile(self.out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                        for root, _, files in os.walk(temp_dir):
                            for f in files:
                                full = Path(root) / f
                                rel = full.relative_to(temp_dir)
                                zf.write(full, arcname=str(rel))
                    self.message.emit(f"OK: ZIP gerado em {self.out_path}")
                    self.finished_ok.emit(str(self.out_path))
                finally:
                    shutil.rmtree(temp_dir, ignore_errors=True)
            else:
                self.failed.emit(f"Modo de saída inválido: {self.out_mode}")
        except Exception:
            self.failed.emit(traceback.format_exc())


# ------------------------- UI -------------------------
class Window(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("OCTec — Unir / Dividir PDFs")
        self.setMinimumSize(900, 560)

        ic = get_icon_path()
        if ic:
            self.setWindowIcon(QIcon(str(ic)))

        self.tabs = QTabWidget()

        # --- Tab UNIR ---
        self.merge_list = DropList(allow_reorder=True)
        self.btn_add_files = QPushButton("Adicionar Arquivos…")
        self.btn_add_folder = QPushButton("Adicionar Pasta…")
        self.btn_remove = QPushButton("Remover Selecionados")
        self.btn_clear = QPushButton("Limpar")
        self.btn_merge_go = QPushButton("UNIR PDFs")

        merge_controls_top = QHBoxLayout()
        merge_controls_top.addWidget(self.btn_add_files)
        merge_controls_top.addWidget(self.btn_add_folder)
        merge_controls_top.addStretch()

        merge_controls_bottom = QHBoxLayout()
        merge_controls_bottom.addWidget(self.btn_remove)
        merge_controls_bottom.addWidget(self.btn_clear)
        merge_controls_bottom.addStretch()
        merge_controls_bottom.addWidget(self.btn_merge_go)

        merge_tab = QWidget()
        merge_layout = QVBoxLayout(merge_tab)
        title_merge = QLabel("Arraste PDFs (de qualquer pasta). Reordene por arrastar para definir a ordem de UNIR:")
        title_merge.setObjectName("title")
        merge_layout.addWidget(title_merge)
        merge_layout.addWidget(self.merge_list)
        merge_layout.addLayout(merge_controls_top)
        merge_layout.addLayout(merge_controls_bottom)

        # --- Tab DIVIDIR ---
        self.split_list = DropList(allow_reorder=False)
        self.btn_split_add_files = QPushButton("Adicionar Arquivos…")
        self.btn_split_add_folder = QPushButton("Adicionar Pasta…")
        self.btn_split_remove = QPushButton("Remover Selecionados")
        self.btn_split_clear = QPushButton("Limpar")

        self.rb_folder = QRadioButton("Exportar para PASTA")
        self.rb_zip = QRadioButton("Exportar para ZIP")
        self.rb_folder.setChecked(True)
        self.out_mode_group = QButtonGroup(self)
        self.out_mode_group.addButton(self.rb_folder)
        self.out_mode_group.addButton(self.rb_zip)

        self.btn_split_go = QPushButton("DIVIDIR PDFs")

        split_controls_top = QHBoxLayout()
        split_controls_top.addWidget(self.btn_split_add_files)
        split_controls_top.addWidget(self.btn_split_add_folder)
        split_controls_top.addStretch()

        split_controls_mid = QHBoxLayout()
        split_controls_mid.addWidget(self.btn_split_remove)
        split_controls_mid.addWidget(self.btn_split_clear)
        split_controls_mid.addStretch()
        split_controls_mid.addWidget(self.rb_folder)
        split_controls_mid.addWidget(self.rb_zip)

        split_controls_bottom = QHBoxLayout()
        split_controls_bottom.addStretch()
        split_controls_bottom.addWidget(self.btn_split_go)

        split_tab = QWidget()
        split_layout = QVBoxLayout(split_tab)
        title_split = QLabel("Arraste PDFs para dividir (um PDF por página). Escolha PASTA ou ZIP como destino:")
        title_split.setObjectName("title")
        split_layout.addWidget(title_split)
        split_layout.addWidget(self.split_list)
        split_layout.addLayout(split_controls_top)
        split_layout.addLayout(split_controls_mid)
        split_layout.addLayout(split_controls_bottom)

        # Tabs
        self.tabs.addTab(merge_tab, "UNIR")
        self.tabs.addTab(split_tab, "DIVIDIR")

        # --- Barra inferior comum ---
        self.lbl_status = QLabel("Pronto.")
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Logs de processamento aparecerão aqui…")

        root = QVBoxLayout(self)
        root.addWidget(self.tabs)
        root.addWidget(self.lbl_status)
        root.addWidget(self.progress)
        root.addWidget(self.log)

        # Eventos merge
        self.btn_add_files.clicked.connect(self._merge_add_files)
        self.btn_add_folder.clicked.connect(self._merge_add_folder)
        self.btn_remove.clicked.connect(self.merge_list.remove_selected)
        self.btn_clear.clicked.connect(self.merge_list.clear_all)
        self.btn_merge_go.clicked.connect(self._do_merge)

        # Eventos split
        self.btn_split_add_files.clicked.connect(self._split_add_files)
        self.btn_split_add_folder.clicked.connect(self._split_add_folder)
        self.btn_split_remove.clicked.connect(self.split_list.remove_selected)
        self.btn_split_clear.clicked.connect(self.split_list.clear_all)
        self.btn_split_go.clicked.connect(self._do_split)

        # Tema/estilo
        self.apply_style()

        self.worker = None  # type: ignore

    # ---------- Estilo ----------
    def apply_style(self):
        qss = """
        QWidget { background-color: #0f172a; color: #e2e8f0; }
        QTabWidget::pane { border: 1px solid #1e293b; border-radius: 10px; }
        QTabBar::tab { background: #1e293b; color: #cbd5e1; padding: 8px 16px; margin: 2px; border-radius: 8px; }
        QTabBar::tab:selected { background: #0ea5e9; color: white; }
        QPushButton {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0284c7, stop:1 #0ea5e9);
            border: none; border-radius: 12px; color: white; padding: 12px 20px; font-weight: 700;
        }
        QPushButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0ea5e9, stop:1 #38bdf8); }
        QPushButton:disabled { background: #334155; color: #94a3b8; }
        QListWidget { background: #0b1220; border: 2px dashed #334155; border-radius: 12px; }
        QListWidget::item { padding: 6px; }
        QProgressBar { background: #0b1220; border: 1px solid #334155; border-radius: 8px; text-align: center; height: 18px; }
        QProgressBar::chunk { background: #22d3ee; border-radius: 6px; }
        QTextEdit { background: #0b1220; border: 1px solid #334155; border-radius: 12px; }
        QLabel#title { font-size: 20px; color: #e2e8f0; font-weight: 700; margin: 2px 0 6px 2px; }
        """
        self.setStyleSheet(qss)

    # ---------- Utils ----------
    def write(self, text: str) -> None:
        self.log.append(text)
        QApplication.processEvents()

    def set_busy(self, busy: bool) -> None:
        widgets = [
            self.btn_add_files, self.btn_add_folder, self.btn_remove, self.btn_clear,
            self.btn_merge_go,
            self.btn_split_add_files, self.btn_split_add_folder, self.btn_split_remove,
            self.btn_split_clear, self.btn_split_go, self.tabs
        ]
        for w in widgets:
            w.setEnabled(not busy)
        if busy:
            QApplication.setOverrideCursor(Qt.WaitCursor)
        else:
            QApplication.restoreOverrideCursor()

    # ---------- Merge actions ----------
    def _merge_add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar PDFs", str(Path.home()), "Arquivos PDF (*.pdf)")
        if files:
            self.merge_list.add_paths([Path(f) for f in files])

    def _merge_add_folder(self):
        d = QFileDialog.getExistingDirectory(self, "Selecionar Pasta (não recursivo)", str(Path.home()))
        if d:
            self.merge_list.add_paths([Path(d)])

    def _do_merge(self):
        inputs = self.merge_list.files()
        if not inputs:
            QMessageBox.information(self, "Nada a unir", "Adicione PDFs à lista primeiro.")
            return
        suggested = str(Path(inputs[0]).with_name("mesclado.pdf"))
        out_file, _ = QFileDialog.getSaveFileName(self, "Salvar PDF mesclado como", suggested, "Arquivos PDF (*.pdf)")
        if not out_file:
            return

        self.log.clear()
        self.progress.setValue(0)
        self.lbl_status.setText("Mesclando…")
        self.set_busy(True)

        self.worker = MergeWorker(inputs, Path(out_file))
        self.worker.progressed.connect(self.progress.setValue)
        self.worker.message.connect(self.write)
        self.worker.finished_ok.connect(self._done_ok)
        self.worker.failed.connect(self._done_failed)
        self.worker.start()

    # ---------- Split actions ----------
    def _split_add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar PDFs", str(Path.home()), "Arquivos PDF (*.pdf)")
        if files:
            self.split_list.add_paths([Path(f) for f in files])

    def _split_add_folder(self):
        d = QFileDialog.getExistingDirectory(self, "Selecionar Pasta (não recursivo)", str(Path.home()))
        if d:
            self.split_list.add_paths([Path(d)])

    def _do_split(self):
        inputs = self.split_list.files()
        if not inputs:
            QMessageBox.information(self, "Nada a dividir", "Adicione PDFs à lista primeiro.")
            return

        out_mode = "folder" if self.rb_folder.isChecked() else "zip"
        if out_mode == "folder":
            base_dir = QFileDialog.getExistingDirectory(self, "Escolher pasta de destino", str(Path(inputs[0]).parent))
            if not base_dir:
                return
            out_path = Path(base_dir)
        else:
            suggested = str(Path(inputs[0]).with_name("dividido.zip"))
            zip_path, _ = QFileDialog.getSaveFileName(self, "Salvar ZIP como", suggested, "Arquivos ZIP (*.zip)")
            if not zip_path:
                return
            out_path = Path(zip_path)

        self.log.clear()
        self.progress.setValue(0)
        self.lbl_status.setText("Dividindo…")
        self.set_busy(True)

        self.worker = SplitWorker(inputs, out_mode, out_path)
        self.worker.progressed.connect(self.progress.setValue)
        self.worker.message.connect(self.write)
        self.worker.finished_ok.connect(self._done_ok)
        self.worker.failed.connect(self._done_failed)
        self.worker.start()

    # ---------- Callbacks ----------
    def _done_ok(self, info: str):
        self.set_busy(False)
        self.lbl_status.setText("Concluído.")
        QMessageBox.information(self, "Concluído", info)

    def _done_failed(self, err: str):
        self.set_busy(False)
        self.lbl_status.setText("Falhou.")
        self.write(f"<pre>{err}</pre>")
        QMessageBox.critical(self, "Erro", err)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("OCTecSplitMerge")
    w = Window()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
