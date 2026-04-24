"""
gui/app.py
Interface gráfica do Conversor Excel → CNAB BB.
Visual idêntico ao padrão Auto SEI (PyQt5).
"""

import os
import sys
from datetime import datetime
from typing import Optional

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFrame, QScrollArea, QSizePolicy,
    QTabWidget, QFileDialog, QMessageBox, QRadioButton,
    QButtonGroup, QGraphicsDropShadowEffect, QSplashScreen,
)
from PyQt5.QtGui import (
    QPixmap, QPainter, QColor, QFont, QPen, QBrush, QIcon,
    QPainterPath, QPalette,
)
from PyQt5.QtCore import Qt, QTimer, QRect, pyqtSignal, QThread

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from logic.config_loader import ConfigLoader, ConfigError
from logic.excel_reader import read_excel
from logic.cnab_generator import CNABGenerator


BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VISUAL_DIR = os.path.join(BASE_DIR, "visual")


def visual_path(filename: str) -> str:
    return os.path.join(VISUAL_DIR, filename)


def load_visual_pixmap(filename: str, width: int, height: int) -> QPixmap:
    px = QPixmap(visual_path(filename))
    if px.isNull():
        return QPixmap()
    return px.scaled(width, height, Qt.KeepAspectRatio, Qt.SmoothTransformation)


def load_visual_icon(filename: str) -> QIcon:
    icon = QIcon(visual_path(filename))
    if icon.isNull():
        return QIcon()
    return icon


# ══════════════════════════════════════════════════════════════
#  PALETA & ESTILOS — padrão Auto SEI
# ══════════════════════════════════════════════════════════════
BLUE       = "#1a73c8"
BLUE_DARK  = "#1558a0"
BG_APP     = "#ffffff"
BG_CARD    = "#ffffff"
BG_INPUT   = "#d4dbe6"
BG_ROW_ODD = "#f7f9fb"
BG_ROW_EVE = "#ffffff"
TEXT_MAIN  = "#1a202c"
TEXT_MUTED = "#555e6b"
BORDER     = "#d4dbe6"
RED_DEL    = "#e53e3e"
GREEN_OK   = "#228B22"

STYLE_WINDOW = f"background-color: {BG_APP};"

STYLE_CARD = f"""
    QFrame#card {{
        background-color: {BG_CARD};
        border-radius: 16px;
        border: 1px solid {BORDER};
    }}
"""

STYLE_BTN_GHOST = f"""
    QPushButton {{
        background-color: transparent;
        color: {BLUE};
        border-radius: 10px;
        font-size: 13px;
        font-family: 'Segoe UI', sans-serif;
        font-weight: 600;
        border: 1.5px solid {BLUE};
    }}
    QPushButton:hover {{ background-color: {BG_INPUT}; }}
    QPushButton:disabled {{ color: #a0aab4; border-color: #a0aab4; }}
"""

STYLE_BTN_EXEC = f"""
    QPushButton {{
        background-color: {BLUE};
        color: white;
        border-radius: 12px;
        font-size: 14px;
        font-family: 'Segoe UI', sans-serif;
        font-weight: 700;
        border: 2px solid white;
    }}
    QPushButton:hover {{ background-color: {BLUE_DARK}; }}
    QPushButton:pressed {{ background-color: #0f3f7a; }}
    QPushButton:disabled {{ background-color: #a0b4cc; color: white; border-color: #a0b4cc; }}
"""

STYLE_BTN_DANGER = f"""
    QPushButton {{
        background-color: transparent;
        color: {BLUE};
        border-radius: 10px;
        font-size: 13px;
        font-family: 'Segoe UI', sans-serif;
        font-weight: 600;
        border: 1.5px solid {BLUE};
    }}
    QPushButton:hover {{ background-color: #fff0f0; }}
"""

STYLE_TABS = f"""
    QTabWidget::pane {{
        border: none;
        background: transparent;
    }}
    QTabBar::tab {{
        background: {BG_INPUT};
        color: {TEXT_MUTED};
        padding: 8px 32px;
        font-size: 13px;
        font-weight: 600;
        font-family: 'Segoe UI', sans-serif;
        border-top-left-radius: 8px;
        border-top-right-radius: 8px;
        margin-right: 4px;
        border: none;
    }}
    QTabBar::tab:selected {{ background: {BLUE}; color: white; }}
    QTabBar::tab:hover:!selected {{ background: {BORDER}; color: {TEXT_MAIN}; }}
"""

STYLE_SCROLL = f"""
    QScrollArea {{ border: none; background: transparent; }}
    QScrollBar:vertical {{
        background: {BG_INPUT};
        width: 6px;
        border-radius: 3px;
    }}
    QScrollBar::handle:vertical {{
        background: #a0aab4;
        border-radius: 3px;
        min-height: 20px;
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
"""

STYLE_DIALOGO = f"""
    QMessageBox {{
        background-color: #ffffff;
        color: {TEXT_MAIN};
        font-family: 'Segoe UI', sans-serif;
        font-size: 13px;
    }}
    QLabel {{
        color: {TEXT_MAIN};
        background-color: transparent;
        font-size: 13px;
    }}
    QPushButton {{
        background-color: {BLUE};
        color: white;
        border-radius: 8px;
        padding: 6px 18px;
        font-size: 13px;
        font-weight: 600;
        border: none;
        min-width: 70px;
    }}
    QPushButton:hover {{ background-color: {BLUE_DARK}; }}
"""


def _paleta_clara() -> QPalette:
    pal = QPalette()
    pal.setColor(QPalette.Window,          QColor("#ffffff"))
    pal.setColor(QPalette.WindowText,      QColor(TEXT_MAIN))
    pal.setColor(QPalette.Base,            QColor("#ffffff"))
    pal.setColor(QPalette.AlternateBase,   QColor("#f5f7fa"))
    pal.setColor(QPalette.Text,            QColor(TEXT_MAIN))
    pal.setColor(QPalette.ButtonText,      QColor("#ffffff"))
    pal.setColor(QPalette.Button,          QColor(BLUE))
    pal.setColor(QPalette.Highlight,       QColor(BLUE))
    pal.setColor(QPalette.HighlightedText, QColor("#ffffff"))
    return pal


def msgbox(parent, tipo: str, titulo: str, texto: str):
    box = QMessageBox(parent)
    box.setWindowTitle(titulo)
    box.setText(texto)
    box.setPalette(_paleta_clara())
    box.setStyleSheet(STYLE_DIALOGO)
    icons = {
        "info": QMessageBox.Information,
        "warn": QMessageBox.Warning,
        "error": QMessageBox.Critical,
    }
    box.setIcon(icons.get(tipo, QMessageBox.Information))
    box.exec_()


# ══════════════════════════════════════════════════════════════
#  SplashScreen — igual ao Auto SEI
# ══════════════════════════════════════════════════════════════
class SplashScreen(QSplashScreen):
    DURACAO_MS = 2200
    PASSOS     = 55

    def __init__(self):
        self._base_px   = self._build_base_pixmap()
        self._progresso = 0
        super().__init__(self._base_px, Qt.WindowStaysOnTopHint)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        intervalo = self.DURACAO_MS // self.PASSOS
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._avancar)
        self._timer.start(intervalo)

    def _avancar(self):
        if self._progresso < self.PASSOS:
            self._progresso += 1
            self._repintar()
        else:
            self._timer.stop()

    def _repintar(self):
        px = self._base_px.copy()
        p  = QPainter(px)
        p.setRenderHint(QPainter.Antialiasing)
        W = px.width()
        bar_x, bar_y, bar_w, bar_h = 80, 222, W - 160, 5
        p.setBrush(QBrush(QColor("#1558a0")))
        p.setPen(Qt.NoPen)
        p.drawRoundedRect(bar_x, bar_y, bar_w, bar_h, 3, 3)
        fill_w = int(bar_w * self._progresso / self.PASSOS)
        if fill_w > 0:
            p.setBrush(QBrush(QColor("#ffffff")))
            p.drawRoundedRect(bar_x, bar_y, fill_w, bar_h, 3, 3)
        p.end()
        self.setPixmap(px)

    def _build_base_pixmap(self) -> QPixmap:
        W, H = 480, 300
        px = QPixmap(W, H)
        px.fill(Qt.transparent)
        p = QPainter(px)
        p.setRenderHint(QPainter.Antialiasing)

        path = QPainterPath()
        path.addRoundedRect(0, 0, W, H, 18, 18)
        p.fillPath(path, QColor(BLUE))

        p.setBrush(QBrush(QColor(BLUE_DARK)))
        p.setPen(Qt.NoPen)
        p.drawRoundedRect(0, H - 64, W, 64, 18, 18)
        p.fillRect(0, H - 64, W, 30, QColor(BLUE_DARK))

        pen = QPen(QColor("#ffffff40"), 1)
        pen.setStyle(Qt.DotLine)
        p.setPen(pen)
        p.drawLine(30, H - 68, W - 30, H - 68)

        logo = load_visual_pixmap("Icon_branco2.ico", 132, 132)
        if not logo.isNull():
            x = (W - logo.width()) // 2
            p.drawPixmap(x, 20, logo)
        else:
            p.setPen(QColor("#ffffff"))
            p.setFont(QFont("Segoe UI", 30, QFont.Bold))
            p.drawText(QRect(0, 24, W, 58), Qt.AlignCenter, "BB")

        p.setPen(QColor("#ffffff"))
        p.setFont(QFont("Segoe UI", 22, QFont.Bold))
        p.drawText(QRect(0, 140, W, 40), Qt.AlignCenter, "Conversor CNAB BB")
      
        p.setFont(QFont("Segoe UI", 11))
        p.drawText(QRect(0, 186, W, 26), Qt.AlignCenter, "Excel  ->  TXT  ·  Banco do Brasil")

        p.setPen(QColor("#9abcd8"))
        p.setFont(QFont("Segoe UI", 9))
        p.drawText(QRect(0, H - 50, W, 40), Qt.AlignCenter, "Remessa CNAB 240")
        p.end()
        return px


# ══════════════════════════════════════════════════════════════
#  DropZone — igual ao Auto SEI
# ══════════════════════════════════════════════════════════════
class DropZone(QFrame):
    files_dropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(84)
        self._set_style_normal()

        row = QHBoxLayout(self)
        row.setContentsMargins(20, 0, 20, 0)
        row.setSpacing(16)

        ico = QLabel()
        ico.setFixedSize(40, 40)
        ico.setAlignment(Qt.AlignCenter)
        ico.setStyleSheet("background: transparent; border: none;")
        px_drop = load_visual_pixmap("processo.ico", 30, 30)
        if not px_drop.isNull():
            ico.setPixmap(px_drop)
        else:
            ico.setText("XLS")
            ico.setFont(QFont("Segoe UI", 14, QFont.Bold))
        row.addWidget(ico)

        txt = QLabel(
            f'Arraste e solte ou '
            f'<a href="#" style="color:{BLUE}; font-weight:600;">clique para selecionar</a>'
            f' arquivos Excel (.xlsx, .xls)'
        )
        txt.setAlignment(Qt.AlignCenter)
        txt.setOpenExternalLinks(False)
        txt.setTextInteractionFlags(Qt.TextBrowserInteraction)
        txt.linkActivated.connect(self._on_click)
        txt.setWordWrap(True)
        txt.setStyleSheet(
            f"font-size: 13px; color: {TEXT_MAIN};"
            "font-family: 'Segoe UI'; background: transparent; border: none;"
        )
        row.addWidget(txt, 1)

    def _set_style_normal(self):
        self.setStyleSheet(f"""
            QFrame {{
                background-color: #f4f8ff;
                border: 2px dashed {BLUE};
                border-radius: 12px;
            }}
        """)

    def _set_style_hover(self):
        self.setStyleSheet(f"""
            QFrame {{
                background-color: #e8f0fe;
                border: 2px dashed {BLUE_DARK};
                border-radius: 12px;
            }}
        """)

    def _on_click(self, _=None):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Selecionar arquivos Excel", "",
            "Excel (*.xlsx *.xls);;Todos (*.*)"
        )
        if paths:
            self.files_dropped.emit(paths)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._on_click()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._set_style_hover()

    def dragLeaveEvent(self, event):
        self._set_style_normal()

    def dropEvent(self, event):
        self._set_style_normal()
        paths = [
            u.toLocalFile() for u in event.mimeData().urls()
            if u.toLocalFile().lower().endswith((".xlsx", ".xls"))
        ]
        if paths:
            self.files_dropped.emit(paths)


# ══════════════════════════════════════════════════════════════
#  TabelaBase — cabeçalho azul + scroll (padrão TabelaDocumentos)
# ══════════════════════════════════════════════════════════════
class TabelaBase(QFrame):
    """Tabela genérica com cabeçalho azul e área scrollável."""

    def __init__(self, colunas: list, altura_scroll: int = 200):
        """colunas = [(texto, largura_fixa_ou_None), ...]"""
        super().__init__()
        self.setStyleSheet("QFrame { background: transparent; border: none; }")

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ── Cabeçalho azul ─────────────────────────────────────
        header = QFrame()
        header.setObjectName("hdr")
        header.setFixedHeight(40)
        header.setStyleSheet(f"""
            QFrame#hdr {{
                background-color: {BLUE};
                border-top-left-radius:  11px;
                border-top-right-radius: 11px;
                border: 1px solid {BLUE};
                border-bottom: none;
            }}
        """)
        hrow = QHBoxLayout(header)
        hrow.setContentsMargins(12, 0, 12, 0)
        hrow.setSpacing(8)
        for texto, largura in colunas:
            lbl = QLabel(texto)
            lbl.setStyleSheet(
                "color: white; font-size: 12px; font-weight: 700;"
                "font-family: 'Segoe UI'; background: transparent; border: none;"
            )
            lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            if largura:
                lbl.setFixedWidth(largura)
            else:
                lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            hrow.addWidget(lbl)
        outer.addWidget(header)

        # ── Scroll area ────────────────────────────────────────
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFixedHeight(altura_scroll)
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll_area.setStyleSheet(
            STYLE_SCROLL + f"""
            QScrollArea {{
                border-left:   1px solid {BORDER};
                border-right:  1px solid {BORDER};
                border-bottom: 1px solid {BORDER};
                border-top: none;
                border-bottom-left-radius:  11px;
                border-bottom-right-radius: 11px;
                background: {BG_CARD};
            }}
        """
        )
        self.scroll_area.viewport().setStyleSheet(f"background: {BG_CARD}; border: none;")

        self.linhas_widget = QWidget()
        self.linhas_widget.setStyleSheet(f"background: {BG_CARD}; border: none;")
        self.linhas_layout = QVBoxLayout(self.linhas_widget)
        self.linhas_layout.setContentsMargins(0, 0, 0, 0)
        self.linhas_layout.setSpacing(0)
        self.linhas_layout.addStretch()
        self.scroll_area.setWidget(self.linhas_widget)
        outer.addWidget(self.scroll_area)

    def inserir_linha(self, widget):
        self.linhas_layout.insertWidget(self.linhas_layout.count() - 1, widget)

    def limpar(self):
        while self.linhas_layout.count() > 1:
            item = self.linhas_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def scroll_to_bottom(self):
        QTimer.singleShot(50, lambda: self.scroll_area.verticalScrollBar().setValue(
            self.scroll_area.verticalScrollBar().maximum()
        ))


# ══════════════════════════════════════════════════════════════
#  FileRow — linha de arquivo na tabela (estilo DocumentoRow)
# ══════════════════════════════════════════════════════════════
class FileRow(QFrame):
    def __init__(self, numero: int, filepath: str, on_remove, on_select, odd=True):
        super().__init__()
        self.filepath = filepath
        self._on_select = on_select
        bg = BG_ROW_ODD if odd else BG_ROW_EVE
        self._bg_normal = bg
        self.setStyleSheet(f"background-color: {bg};")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFixedHeight(50)
        self.setCursor(Qt.PointingHandCursor)

        row = QHBoxLayout(self)
        row.setContentsMargins(12, 0, 12, 0)
        row.setSpacing(10)

        # Número
        lbl_num = QLabel(str(numero))
        lbl_num.setFixedWidth(24)
        lbl_num.setAlignment(Qt.AlignCenter)
        lbl_num.setStyleSheet(
            f"font-size: 12px; color: {TEXT_MUTED}; font-weight: 700;"
            "font-family: 'Segoe UI'; background: transparent;"
        )
        row.addWidget(lbl_num)

        # Badge de extensão
        ext = os.path.splitext(filepath)[1].upper().replace(".", "")
        lbl_ext = QLabel(ext)
        lbl_ext.setFixedSize(40, 22)
        lbl_ext.setAlignment(Qt.AlignCenter)
        lbl_ext.setStyleSheet(f"""
            background-color: {BLUE};
            color: white;
            border-radius: 4px;
            font-size: 10px;
            font-weight: 700;
            font-family: 'Segoe UI';
        """)
        row.addWidget(lbl_ext)

        # Nome do arquivo
        self.lbl_nome = QLabel(os.path.basename(filepath))
        self.lbl_nome.setStyleSheet(
            f"font-size: 12px; color: {TEXT_MAIN}; font-weight: 600;"
            "font-family: 'Segoe UI'; background: transparent;"
        )
        self.lbl_nome.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        row.addWidget(self.lbl_nome, 1)

        # Status
        self.lbl_status = QLabel("")
        self.lbl_status.setFixedWidth(46)
        self.lbl_status.setAlignment(Qt.AlignCenter)
        self.lbl_status.setStyleSheet(
            "font-size: 11px; font-weight: bold; background: transparent;"
        )
        row.addWidget(self.lbl_status)

        # Botão remover
        btn_rem = QPushButton("✕")
        btn_rem.setFixedSize(28, 28)
        btn_rem.setCursor(Qt.PointingHandCursor)
        btn_rem.setStyleSheet(f"""
            QPushButton {{
                background-color: {RED_DEL};
                color: white;
                border-radius: 14px;
                font-size: 11px;
                font-weight: bold;
                border: none;
            }}
            QPushButton:hover {{ background-color: #c0392b; }}
        """)
        btn_rem.clicked.connect(lambda: on_remove(self))
        row.addWidget(btn_rem)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._on_select(self.filepath)

    def set_selected(self, sel: bool):
        if sel:
            self.setStyleSheet(f"background-color: #dceeff; border-left: 3px solid {BLUE};")
        else:
            self.setStyleSheet(f"background-color: {self._bg_normal};")

    def set_status(self, text: str, color: str):
        self.lbl_status.setText(text)
        self.lbl_status.setStyleSheet(
            f"font-size: 11px; font-weight: bold; color: {color}; background: transparent;"
        )


# ══════════════════════════════════════════════════════════════
#  PreviewTable — cabeçalho azul + grade de dados
# ══════════════════════════════════════════════════════════════
class PreviewTable(QFrame):
    def __init__(self):
        super().__init__()
        self.setStyleSheet("QFrame { background: transparent; border: none; }")
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(0)
        self._show_empty()

    def _clear(self):
        while self._layout.count():
            item = self._layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def _show_empty(self):
        self._clear()
        holder = QFrame()
        holder.setFixedHeight(160)
        holder.setStyleSheet(f"""
            QFrame {{
                background-color: {BG_ROW_ODD};
                border: 1px solid {BORDER};
                border-radius: 11px;
            }}
        """)
        v = QVBoxLayout(holder)
        v.setAlignment(Qt.AlignCenter)
        ico = QLabel()
        ico.setAlignment(Qt.AlignCenter)
        ico.setStyleSheet("background: transparent; border: none;")
        px_prev = load_visual_pixmap("eye.ico", 30, 30)
        if not px_prev.isNull():
            ico.setPixmap(px_prev)
        else:
            ico.setText("PRV")
            ico.setFont(QFont("Segoe UI", 13, QFont.Bold))
        v.addWidget(ico)
        lbl = QLabel("Nenhum arquivo selecionado")
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setStyleSheet(
            f"color: {TEXT_MUTED}; font-size: 13px; background: transparent; border: none;"
        )
        v.addWidget(lbl)
        sub = QLabel("Clique em um arquivo para visualizar")
        sub.setAlignment(Qt.AlignCenter)
        sub.setStyleSheet(
            f"color: {BORDER}; font-size: 11px; background: transparent; border: none;"
        )
        v.addWidget(sub)
        self._layout.addWidget(holder)

    def show_dataframe(self, df: pd.DataFrame, filename: str):
        self._clear()
        preview = df.head(6)
        cols = list(preview.columns)

        # Cabeçalho azul
        hdr = QFrame()
        hdr.setObjectName("pvhdr")
        hdr.setFixedHeight(36)
        hdr.setStyleSheet(f"""
            QFrame#pvhdr {{
                background-color: {BLUE};
                border-top-left-radius: 11px;
                border-top-right-radius: 11px;
            }}
        """)
        hrow = QHBoxLayout(hdr)
        hrow.setContentsMargins(10, 0, 10, 0)
        hrow.setSpacing(4)
        for col in cols:
            l = QLabel(str(col).upper()[:18])
            l.setStyleSheet(
                "color: white; font-size: 10px; font-weight: 700;"
                "font-family: 'Segoe UI'; background: transparent;"
            )
            l.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            hrow.addWidget(l)
        self._layout.addWidget(hdr)

        # Scroll com linhas
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFixedHeight(148)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setStyleSheet(
            STYLE_SCROLL + f"""
            QScrollArea {{
                border-left:   1px solid {BORDER};
                border-right:  1px solid {BORDER};
                border-bottom: 1px solid {BORDER};
                border-top: none;
                border-bottom-left-radius: 11px;
                border-bottom-right-radius: 11px;
                background: {BG_CARD};
            }}
        """
        )
        scroll.viewport().setStyleSheet(f"background: {BG_CARD};")

        inner = QWidget()
        inner.setStyleSheet(f"background: {BG_CARD};")
        ilayout = QVBoxLayout(inner)
        ilayout.setContentsMargins(0, 0, 0, 0)
        ilayout.setSpacing(0)

        for i, (_, row) in enumerate(preview.iterrows()):
            bg = BG_ROW_ODD if i % 2 == 0 else BG_ROW_EVE
            rframe = QFrame()
            rframe.setStyleSheet(f"background-color: {bg};")
            rframe.setFixedHeight(26)
            rrow = QHBoxLayout(rframe)
            rrow.setContentsMargins(10, 0, 10, 0)
            rrow.setSpacing(4)
            for col in cols:
                val = str(row.get(col, ""))[:20]
                l = QLabel(val)
                l.setStyleSheet(
                    f"color: {TEXT_MAIN}; font-size: 10px; background: transparent;"
                    "font-family: 'Segoe UI';"
                )
                l.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
                rrow.addWidget(l)
            ilayout.addWidget(rframe)

        ilayout.addStretch()
        scroll.setWidget(inner)
        self._layout.addWidget(scroll)

    def show_error(self, msg: str):
        self._clear()
        holder = QFrame()
        holder.setFixedHeight(160)
        holder.setStyleSheet(f"""
            QFrame {{
                background-color: #fff5f5;
                border: 1px solid {RED_DEL};
                border-radius: 11px;
            }}
        """)
        v = QVBoxLayout(holder)
        v.setAlignment(Qt.AlignCenter)
        lbl = QLabel(f"Erro:\n{msg[:120]}")
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setWordWrap(True)
        lbl.setStyleSheet(
            f"color: {RED_DEL}; font-size: 11px; background: transparent; border: none;"
        )
        v.addWidget(lbl)
        self._layout.addWidget(holder)


# ══════════════════════════════════════════════════════════════
#  ConversionWorker — thread de conversão
# ══════════════════════════════════════════════════════════════
class ConversionWorker(QThread):
    log_signal   = pyqtSignal(str, str)
    file_done    = pyqtSignal(str, bool)
    finished_all = pyqtSignal(int, int)

    def __init__(self, files: dict, cfg: ConfigLoader, tipo: str, output_dir: str, modo: str = "pagamentos"):
        super().__init__()
        self._files      = files
        self._cfg        = cfg
        self._tipo       = tipo
        self._output_dir = output_dir
        self._modo       = modo  # "pagamentos" ou "aplicacoes"

    def run(self):
        cfg_exec = ConfigLoader(self._cfg.config_path, tipo_pagamento=self._tipo)
        gen = CNABGenerator(cfg_exec)
        success = errors = 0
        for filepath, data in self._files.items():
            fname = os.path.basename(filepath)
            self.log_signal.emit(f"Processando: {fname}", "info")
            try:
                if self._modo == "aplicacoes":
                    content = gen.generate(data["df"], tipo_pagamento="APLICACAO")
                else:
                    content = gen.generate(data["df"], tipo_pagamento=self._tipo)
                base    = os.path.splitext(fname)[0]
                now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                out     = os.path.join(self._output_dir, f"CNAB_{base}_{now_str}.txt")
                with open(out, "w", encoding="ascii", errors="replace") as f:
                    f.write(content)
                self.log_signal.emit(f"Gerado: CNAB_{base}_{now_str}.txt", "success")
                self.file_done.emit(filepath, True)
                success += 1
            except Exception as e:
                self.log_signal.emit(f"Erro em {fname}: {e}", "error")
                self.file_done.emit(filepath, False)
                errors += 1
        self.finished_all.emit(success, errors)


# ══════════════════════════════════════════════════════════════
#  AbaPagamentos
# ══════════════════════════════════════════════════════════════
class AbaPagamentos(QWidget):
    def __init__(self, cfg: Optional[ConfigLoader], cfg_error: Optional[str]):
        super().__init__()
        self._cfg       = cfg
        self._cfg_error = cfg_error
        self._files: dict[str, dict] = {}
        self._selected: Optional[str] = None
        self._file_rows: dict[str, FileRow] = {}
        self._worker: Optional[ConversionWorker] = None
        self._tipo_pag  = "PGA"

        self.setStyleSheet("background: transparent;")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 12, 0, 8)
        layout.setSpacing(10)

        self._build_tipo_pagamento(layout)

        self._drop = DropZone()
        self._drop.files_dropped.connect(self._on_files_dropped)
        layout.addWidget(self._drop)

        # Gerenciador de arquivos
        layout.addWidget(self._section_label("Gerenciador de Arquivos", "adicionar.ico"))

        self._tabela = TabelaBase(
            [("", 24), ("", 40), ("Arquivo", None), ("Status", 46), ("", 28)],
            altura_scroll=200,
        )
        layout.addWidget(self._tabela)

        # Preview
        layout.addWidget(self._section_label("Preview dos Dados", ""))
        self._preview = PreviewTable()
        layout.addWidget(self._preview)

        self._build_botoes(layout)

    def _section_label(self, text: str, icon_file: str) -> QWidget:
        holder = QWidget()
        row = QHBoxLayout(holder)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(6)

        ico = QLabel()
        ico.setFixedSize(16, 16)
        ico.setStyleSheet("background: transparent; border: none;")
        px = load_visual_pixmap(icon_file, 14, 14)
        if not px.isNull():
            ico.setPixmap(px)
        row.addWidget(ico)

        lbl = QLabel(text)
        lbl.setStyleSheet(
            f"color: {TEXT_MAIN}; font-size: 12px; font-weight: 700;"
            "font-family: 'Segoe UI'; background: transparent;"
        )
        row.addWidget(lbl)
        row.addStretch()
        return holder

    # ── Seção: tipo de pagamento ─────────────────────────────
    def _build_tipo_pagamento(self, layout: QVBoxLayout):
        frame = QFrame()
        frame.setObjectName("tipopag")
        frame.setStyleSheet(f"""
            QFrame#tipopag {{
                background-color: {BG_ROW_ODD};
                border: 1px solid {BORDER};
                border-radius: 10px;
            }}
        """)
        row = QHBoxLayout(frame)
        row.setContentsMargins(16, 8, 16, 8)
        row.setSpacing(16)

        lbl = QLabel("Tipo de Pagamento:")
        lbl.setStyleSheet(
            f"color: {TEXT_MAIN}; font-size: 12px; font-weight: 700;"
            "font-family: 'Segoe UI'; background: transparent;"
        )
        row.addWidget(lbl)

        _radio_style = f"""
            QRadioButton {{
                color: {TEXT_MAIN};
                font-size: 12px;
                font-weight: 600;
                font-family: 'Segoe UI';
                background: transparent;
                spacing: 6px;
            }}
            QRadioButton::indicator {{
                width: 16px; height: 16px;
                border-radius: 8px;
                border: 2px solid {BLUE};
                background: white;
            }}
            QRadioButton::indicator:checked {{
                background: {BLUE};
                border: 2px solid {BLUE_DARK};
            }}
        """

        self._btn_pga = QRadioButton("  PGA  ")
        self._btn_pga.setChecked(True)
        self._btn_pga.setStyleSheet(_radio_style)
        self._btn_pga.setProperty("valor", "PGA")

        self._btn_jusmp = QRadioButton(" JUSMP ")
        self._btn_jusmp.setStyleSheet(_radio_style)
        self._btn_jusmp.setProperty("valor", "JUSMP")

        self._grp = QButtonGroup(self)
        self._grp.addButton(self._btn_pga,   0)
        self._grp.addButton(self._btn_jusmp, 1)
        self._grp.buttonClicked.connect(
            lambda btn: setattr(self, "_tipo_pag", btn.property("valor"))
        )

        row.addWidget(self._btn_pga)
        row.addWidget(self._btn_jusmp)
        row.addStretch()
        layout.addWidget(frame)

    # ── Seção: botões ────────────────────────────────────────
    def _build_botoes(self, layout: QVBoxLayout):
        row = QHBoxLayout()
        row.setSpacing(10)
        row.addStretch()

        btn_limpar = QPushButton("Limpar Tudo")
        btn_limpar.setFixedSize(148, 38)
        btn_limpar.setCursor(Qt.PointingHandCursor)
        btn_limpar.setStyleSheet(STYLE_BTN_DANGER)
        btn_limpar.setIcon(load_visual_icon("limpar.ico"))
        btn_limpar.clicked.connect(self._limpar_tudo)
        row.addWidget(btn_limpar)

        self.btn_exec = QPushButton("Processar e Baixar")
        self.btn_exec.setFixedSize(210, 46)
        self.btn_exec.setCursor(Qt.PointingHandCursor)
        self.btn_exec.setStyleSheet(STYLE_BTN_EXEC)
        self.btn_exec.setIcon(load_visual_icon("import.png"))
        self.btn_exec.clicked.connect(self._processar)
        row.addWidget(self.btn_exec)

        layout.addLayout(row)

    # ── Gerenciamento de arquivos ────────────────────────────

    def _on_files_dropped(self, paths: list):
        for p in paths:
            if p not in self._files:
                df, errs = read_excel(p, aba="pagamentos")
                self._files[p] = {"df": df, "errors": errs}
        self._refresh_lista()
        if paths:
            self._select_file(paths[-1])

    def _remove_file(self, row: FileRow):
        fp = row.filepath
        self._files.pop(fp, None)
        self._file_rows.pop(fp, None)
        if self._selected == fp:
            self._selected = None
            self._preview._show_empty()
        self._refresh_lista()

    def _select_file(self, fp: str):
        self._selected = fp
        for path, row in self._file_rows.items():
            row.set_selected(path == fp)
        data = self._files.get(fp)
        if not data:
            return
        if data["errors"] and data["df"].empty:
            self._preview.show_error(data["errors"][0])
        else:
            self._preview.show_dataframe(data["df"], os.path.basename(fp))

    def _refresh_lista(self):
        self._tabela.limpar()
        self._file_rows.clear()
        for i, (fp, data) in enumerate(self._files.items()):
            row = FileRow(i + 1, fp, self._remove_file, self._select_file, odd=i % 2 == 0)
            if data["errors"] and data["df"].empty:
                row.set_status("ERRO", RED_DEL)
            elif data["errors"]:
                row.set_status("VISUALIZAR", "#d97706")
            else:
                row.set_status("OK ", GREEN_OK)
            if fp == self._selected:
                row.set_selected(True)
            self._tabela.inserir_linha(row)
            self._file_rows[fp] = row

    def _limpar_tudo(self):
        self._files.clear()
        self._file_rows.clear()
        self._selected = None
        self._tabela.limpar()
        self._preview._show_empty()

    # ── Processamento ────────────────────────────────────────

    def _processar(self):
        if not self._cfg:
            msgbox(self, "error", "Configuração", f"config.json inválido:\n{self._cfg_error}")
            return
        valid = {fp: d for fp, d in self._files.items() if not d["df"].empty}
        if not valid:
            msgbox(self, "warn", "Sem arquivos", "Nenhum arquivo válido para processar.")
            return

        output_dir = QFileDialog.getExistingDirectory(self, "Selecionar pasta de saída")
        if not output_dir:
            return

        self.btn_exec.setEnabled(False)
        self.btn_exec.setText("Processando...")
        self.btn_exec.setIcon(load_visual_icon("troca.ico"))

        self._worker = ConversionWorker(valid, self._cfg, self._tipo_pag, output_dir)
        self._worker.file_done.connect(self._on_file_done)
        self._worker.finished_all.connect(self._on_finished)
        self._worker.start()

    def _on_file_done(self, filepath: str, ok: bool):
        if filepath in self._file_rows:
            self._file_rows[filepath].set_status(
                "✓" if ok else "ERR", GREEN_OK if ok else RED_DEL
            )

    def _on_finished(self, success: int, errors: int):
        self.btn_exec.setEnabled(True)
        self.btn_exec.setText("Processar e Baixar")
        self.btn_exec.setIcon(load_visual_icon("processo.ico"))
        if success:
            msgbox(self, "info", "Concluído",
                   f"{success} arquivo(s) CNAB 240 gerado(s) com sucesso!")
        if errors:
            msgbox(self, "error", "Erros",
                   f"{errors} arquivo(s) com erro.")


# ══════════════════════════════════════════════════════════════
#  AbaAplicacoes
# ══════════════════════════════════════════════════════════════
class AbaAplicacoes(QWidget):
    def __init__(self, cfg: Optional[ConfigLoader], cfg_error: Optional[str]):
        super().__init__()
        self._cfg       = cfg
        self._cfg_error = cfg_error
        self._files: dict[str, dict] = {}
        self._selected: Optional[str] = None
        self._file_rows: dict[str, FileRow] = {}
        self._worker: Optional[ConversionWorker] = None
        self._tipo_pag  = "BBJUMP"

        self.setStyleSheet("background: transparent;")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 12, 0, 8)
        layout.setSpacing(10)

        self._drop = DropZone()
        self._drop.files_dropped.connect(self._on_files_dropped)
        layout.addWidget(self._drop)

        layout.addWidget(self._section_label("Gerenciador de Arquivos", "adicionar.ico"))

        self._tabela = TabelaBase(
            [("", 24), ("", 40), ("Arquivo", None), ("Status", 46), ("", 28)],
            altura_scroll=200,
        )
        layout.addWidget(self._tabela)

        layout.addWidget(self._section_label("Preview dos Dados", ""))
        self._preview = PreviewTable()
        layout.addWidget(self._preview)

        self._build_botoes(layout)

    def _section_label(self, text: str, icon_file: str) -> QWidget:
        holder = QWidget()
        row = QHBoxLayout(holder)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(6)

        ico = QLabel()
        ico.setFixedSize(16, 16)
        ico.setStyleSheet("background: transparent; border: none;")
        px = load_visual_pixmap(icon_file, 14, 14)
        if not px.isNull():
            ico.setPixmap(px)
        row.addWidget(ico)

        lbl = QLabel(text)
        lbl.setStyleSheet(
            f"color: {TEXT_MAIN}; font-size: 12px; font-weight: 700;"
            "font-family: 'Segoe UI'; background: transparent;"
        )
        row.addWidget(lbl)
        row.addStretch()
        return holder

    def _build_botoes(self, layout: QVBoxLayout):
        row = QHBoxLayout()
        row.setSpacing(10)
        row.addStretch()

        btn_limpar = QPushButton("Limpar Tudo")
        btn_limpar.setFixedSize(148, 38)
        btn_limpar.setCursor(Qt.PointingHandCursor)
        btn_limpar.setStyleSheet(STYLE_BTN_DANGER)
        btn_limpar.setIcon(load_visual_icon("limpar.ico"))
        btn_limpar.clicked.connect(self._limpar_tudo)
        row.addWidget(btn_limpar)

        self.btn_exec = QPushButton("Processar e Baixar")
        self.btn_exec.setFixedSize(210, 46)
        self.btn_exec.setCursor(Qt.PointingHandCursor)
        self.btn_exec.setStyleSheet(STYLE_BTN_EXEC)
        self.btn_exec.setIcon(load_visual_icon("import.png"))
        self.btn_exec.clicked.connect(self._processar)
        row.addWidget(self.btn_exec)

        layout.addLayout(row)

    def _on_files_dropped(self, paths: list):
        for p in paths:
            if p not in self._files:
                df, errs = read_excel(p, aba="aplicacoes")
                self._files[p] = {"df": df, "errors": errs}
        self._refresh_lista()
        if paths:
            self._select_file(paths[-1])

    def _remove_file(self, row: FileRow):
        fp = row.filepath
        self._files.pop(fp, None)
        self._file_rows.pop(fp, None)
        if self._selected == fp:
            self._selected = None
            self._preview._show_empty()
        self._refresh_lista()

    def _select_file(self, fp: str):
        self._selected = fp
        for path, row in self._file_rows.items():
            row.set_selected(path == fp)
        data = self._files.get(fp)
        if not data:
            return
        if data["errors"] and data["df"].empty:
            self._preview.show_error(data["errors"][0])
        else:
            self._preview.show_dataframe(data["df"], os.path.basename(fp))

    def _refresh_lista(self):
        self._tabela.limpar()
        self._file_rows.clear()
        for i, (fp, data) in enumerate(self._files.items()):
            row = FileRow(i + 1, fp, self._remove_file, self._select_file, odd=i % 2 == 0)
            if data["errors"] and data["df"].empty:
                row.set_status("ERR", RED_DEL)
            elif data["errors"]:
                row.set_status("AVS", "#d97706")
            else:
                row.set_status("OK ", GREEN_OK)
            if fp == self._selected:
                row.set_selected(True)
            self._tabela.inserir_linha(row)
            self._file_rows[fp] = row

    def _limpar_tudo(self):
        self._files.clear()
        self._file_rows.clear()
        self._selected = None
        self._tabela.limpar()
        self._preview._show_empty()

    def _processar(self):
        if not self._cfg:
            msgbox(self, "error", "Configuracao", f"config.json invalido:\n{self._cfg_error}")
            return
        valid = {fp: d for fp, d in self._files.items() if not d["df"].empty}
        if not valid:
            msgbox(self, "warn", "Sem arquivos", "Nenhum arquivo valido para processar.")
            return

        output_dir = QFileDialog.getExistingDirectory(self, "Selecionar pasta de saida")
        if not output_dir:
            return

        self.btn_exec.setEnabled(False)
        self.btn_exec.setText("Processando...")
        self.btn_exec.setIcon(load_visual_icon("troca.ico"))

        self._worker = ConversionWorker(valid, self._cfg, self._tipo_pag, output_dir, modo="aplicacoes")
        self._worker.file_done.connect(self._on_file_done)
        self._worker.finished_all.connect(self._on_finished)
        self._worker.start()

    def _on_file_done(self, filepath: str, ok: bool):
        if filepath in self._file_rows:
            self._file_rows[filepath].set_status(
                "OK " if ok else "ERR", GREEN_OK if ok else RED_DEL
            )

    def _on_finished(self, success: int, errors: int):
        self.btn_exec.setEnabled(True)
        self.btn_exec.setText("Processar e Baixar")
        self.btn_exec.setIcon(load_visual_icon("processo.ico"))
        if success:
            msgbox(self, "info", "Concluido",
                   f"{success} arquivo(s) CNAB 240 gerado(s) com sucesso!")
        if errors:
            msgbox(self, "error", "Erros",
                   f"{errors} arquivo(s) com erro.")


# ══════════════════════════════════════════════════════════════
#  MainWindow
# ══════════════════════════════════════════════════════════════
class MainWindow(QMainWindow):
    def __init__(self, cfg: Optional[ConfigLoader], cfg_error: Optional[str]):
        super().__init__()
        self.setWindowTitle("Conversor Excel → CNAB BB")
        self.setWindowIcon(load_visual_icon("Icon.ico"))
        self.setGeometry(100, 100, 860, 840)
        self.setFixedSize(860, 840)
        self.setStyleSheet(STYLE_WINDOW)

        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(16, 12, 16, 8)
        outer.setSpacing(0)

        # Card principal com sombra — igual ao Auto SEI
        card = QFrame()
        card.setObjectName("card")
        card.setStyleSheet(STYLE_CARD)
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(36)
        shadow.setOffset(0, 6)
        shadow.setColor(QColor(0, 0, 0, 35))
        card.setGraphicsEffect(shadow)

        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(28, 22, 28, 18)
        card_layout.setSpacing(4)

        titulo = QLabel("Conversor Excel para TXT – Banco do Brasil")
        titulo.setStyleSheet(
            f"font-size: 18px; font-weight: 700; color: {TEXT_MAIN};"
            "font-family: 'Segoe UI', sans-serif;"
        )
        card_layout.addWidget(titulo)

        sub = QLabel("Converta seus arquivos Excel para formato de remessa CNAB 240")
        sub.setStyleSheet(
            f"font-size: 11px; color: {TEXT_MUTED}; font-family: 'Segoe UI';"
        )
        card_layout.addWidget(sub)

        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(STYLE_TABS)
        self.tabs.addTab(AbaPagamentos(cfg, cfg_error), "  Pagamentos  ")
        self.tabs.addTab(AbaAplicacoes(cfg, cfg_error), "  Aplicações  ")
        card_layout.addWidget(self.tabs)

        outer.addWidget(card)

        footer = QLabel(
            "Desenvolvido para facilitar a importação de arquivos no Banco do Brasil  ·  CNAB 240"
        )
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet(
            f"font-size: 11px; color: {TEXT_MUTED}; padding: 8px 0 0 0;"
        )
        outer.addWidget(footer)


# ══════════════════════════════════════════════════════════════
#  Entry point
# ══════════════════════════════════════════════════════════════
def run():
    app = QApplication(sys.argv)
    app.setWindowIcon(load_visual_icon("Icon_branco2.ico"))
    app.setFont(QFont("Segoe UI", 10))
    app.setPalette(_paleta_clara())

    cfg = None
    cfg_error = None
    try:
        cfg = ConfigLoader()
    except ConfigError as e:
        cfg_error = str(e)

    splash = SplashScreen()
    splash.show()
    app.processEvents()

    window = MainWindow(cfg, cfg_error)
    QTimer.singleShot(
        SplashScreen.DURACAO_MS,
        lambda: (splash.finish(window), window.show()),
    )
    sys.exit(app.exec_())


if __name__ == "__main__":
    run()
