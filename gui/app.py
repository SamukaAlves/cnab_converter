"""
gui/app.py  –  Conversor Excel → CNAB BB
Interface gráfica redesenhada: cards, badges, hover, estados vazios, header profissional.
"""

import os, sys
from datetime import datetime
from typing import Optional

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFrame, QScrollArea, QSizePolicy,
    QTabWidget, QFileDialog, QMessageBox, QRadioButton,
    QButtonGroup, QGraphicsDropShadowEffect, QSplashScreen, QDateEdit,
    QStackedWidget,
)
from PyQt5.QtGui import (
    QPixmap, QPainter, QColor, QFont, QPen, QBrush, QIcon,
    QPainterPath, QPalette, QLinearGradient, QFontMetrics,
)
from PyQt5.QtCore import Qt, QTimer, QRect, QRectF, pyqtSignal, QThread, QDate, QPoint

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from logic.config_loader import ConfigLoader, ConfigError
from logic.excel_reader import read_excel
from logic.cnab_generator import CNABGenerator

BASE_DIR   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VISUAL_DIR = os.path.join(BASE_DIR, "visual")

def vp(f):  return os.path.join(VISUAL_DIR, f)
def vpx(f, w, h):
    px = QPixmap(vp(f))
    return px.scaled(w, h, Qt.KeepAspectRatio, Qt.SmoothTransformation) if not px.isNull() else QPixmap()
def vico(f):
    ic = QIcon(vp(f)); return ic if not ic.isNull() else QIcon()

# ── Paleta ────────────────────────────────────────────────────────
BLUE        = "#1a73c8"
BLUE_DARK   = "#1558a0"
BLUE_LIGHT  = "#e8f0fe"
BLUE_MID    = "#4a90d9"
BG          = "#f0f4f8"
CARD        = "#ffffff"
BORDER      = "#dde3ed"
BORDER_DARK = "#c5cdd9"
TEXT        = "#1a202c"
TEXT2       = "#4a5568"
TEXT3       = "#8a96a8"
GREEN       = "#1e8a3c"
GREEN_BG    = "#e6f4ea"
RED         = "#d93025"
RED_BG      = "#fce8e6"
AMBER       = "#b45309"
AMBER_BG    = "#fef3c7"
PURPLE      = "#6d28d9"
PURPLE_BG   = "#ede9fe"

# ── Helpers de sombra ─────────────────────────────────────────────
def _shadow(widget, blur=20, dy=4, alpha=20):
    s = QGraphicsDropShadowEffect()
    s.setBlurRadius(blur); s.setOffset(0, dy)
    s.setColor(QColor(0, 0, 0, alpha))
    widget.setGraphicsEffect(s)

# ── Estilo global ─────────────────────────────────────────────────
STYLE_APP = f"""
    QWidget {{ font-family: 'Segoe UI', sans-serif; }}
    QScrollBar:vertical {{
        background: {BORDER}; width: 6px; border-radius: 3px;
    }}
    QScrollBar::handle:vertical {{
        background: #a0aab4; border-radius: 3px; min-height: 24px;
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
    QScrollBar:horizontal {{ height: 0px; }}
"""

STYLE_TABS = f"""
    QTabWidget::pane {{ border: none; background: transparent; }}
    QTabBar::tab {{
        background: {CARD};
        color: {TEXT2};
        padding: 10px 28px;
        font-size: 13px; font-weight: 600;
        border-top-left-radius: 10px;
        border-top-right-radius: 10px;
        margin-right: 3px;
        border: 1px solid {BORDER};
        border-bottom: none;
    }}
    QTabBar::tab:selected {{ background: {BLUE}; color: white; border-color: {BLUE}; }}
    QTabBar::tab:hover:!selected {{ background: {BLUE_LIGHT}; color: {BLUE}; }}
"""

STYLE_MSGBOX = f"""
    QMessageBox {{ background: {CARD}; }}
    QLabel {{ color: {TEXT}; font-size: 13px; background: transparent; }}
    QPushButton {{
        background: {BLUE}; color: white; border-radius: 8px;
        padding: 6px 20px; font-size: 13px; font-weight: 600; border: none; min-width: 70px;
    }}
    QPushButton:hover {{ background: {BLUE_DARK}; }}
"""

# ──────────────────────────────────────────────────────────────────
#  Componentes reutilizáveis
# ──────────────────────────────────────────────────────────────────

class Card(QFrame):
    """Card branco com borda e sombra."""
    def __init__(self, radius=14, parent=None):
        super().__init__(parent)
        self.setStyleSheet(f"""
            QFrame {{
                background: {CARD};
                border-radius: {radius}px;
                border: 1px solid {BORDER};
            }}
        """)
        _shadow(self, 18, 3, 18)


class Badge(QLabel):
    """Pill badge colorido."""
    COLORS = {
        "ok":         (GREEN,  GREEN_BG),
        "erro":       (RED,    RED_BG),
        "aviso":      (AMBER,  AMBER_BG),
        "processando":(BLUE,   BLUE_LIGHT),
        "info":       (PURPLE, PURPLE_BG),
    }
    def __init__(self, text: str, kind: str = "ok", parent=None):
        super().__init__(text, parent)
        fg, bg = self.COLORS.get(kind, (TEXT2, BORDER))
        self.setStyleSheet(f"""
            QLabel {{
                background: {bg}; color: {fg};
                border-radius: 10px; padding: 2px 10px;
                font-size: 11px; font-weight: 700;
                border: 1px solid {fg}40;
            }}
        """)
        self.setFixedHeight(22)
        self.setAlignment(Qt.AlignCenter)


class Divider(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.HLine)
        self.setStyleSheet(f"border: none; border-top: 1px solid {BORDER}; background: transparent;")
        self.setFixedHeight(1)


class IconLabel(QLabel):
    """Label com ícone à esquerda e texto à direita."""
    def __init__(self, icon_file: str, text: str, icon_size=14, font_size=12, bold=True, color=TEXT, parent=None):
        super().__init__(parent)
        px = vpx(icon_file, icon_size, icon_size)
        self._ico = px
        self._txt = text
        self._fs  = font_size
        self._bold = bold
        self._color = color
        self._update()

    def _update(self):
        self.setText(self._txt)
        self.setStyleSheet(f"""
            QLabel {{
                color: {self._color};
                font-size: {self._fs}px;
                font-weight: {'700' if self._bold else '400'};
                background: transparent;
            }}
        """)


def _btn(text, style="primary", icon_file="", w=None, h=40):
    b = QPushButton(text)
    b.setCursor(Qt.PointingHandCursor)
    if h: b.setFixedHeight(h)
    if w: b.setFixedWidth(w)
    if icon_file:
        ic = vico(icon_file)
        if not ic.isNull(): b.setIcon(ic)

    if style == "primary":
        b.setStyleSheet(f"""
            QPushButton {{
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 {BLUE_MID}, stop:1 {BLUE});
                color: white; border-radius: 10px;
                font-size: 13px; font-weight: 700; border: none;
                padding: 0 20px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 {BLUE}, stop:1 {BLUE_DARK});
            }}
            QPushButton:pressed {{ background: {BLUE_DARK}; }}
            QPushButton:disabled {{ background: #c5cdd9; color: #8a96a8; }}
        """)
    elif style == "ghost":
        b.setStyleSheet(f"""
            QPushButton {{
                background: transparent; color: {BLUE};
                border-radius: 10px; font-size: 13px; font-weight: 600;
                border: 1.5px solid {BLUE}; padding: 0 16px;
            }}
            QPushButton:hover {{ background: {BLUE_LIGHT}; }}
            QPushButton:disabled {{ color: {TEXT3}; border-color: {TEXT3}; }}
        """)
    elif style == "danger":
        b.setStyleSheet(f"""
            QPushButton {{
                background: transparent; color: {RED};
                border-radius: 10px; font-size: 13px; font-weight: 600;
                border: 1.5px solid {RED}40; padding: 0 16px;
            }}
            QPushButton:hover {{ background: {RED_BG}; border-color: {RED}; }}
        """)
    elif style == "icon_del":
        b.setFixedSize(26, 26)
        b.setStyleSheet(f"""
            QPushButton {{
                background: {RED_BG}; color: {RED};
                border-radius: 13px; font-size: 13px; font-weight: 900;
                border: none; padding: 0;
            }}
            QPushButton:hover {{ background: {RED}; color: white; }}
        """)
    return b


# ──────────────────────────────────────────────────────────────────
#  SplashScreen
# ──────────────────────────────────────────────────────────────────
class SplashScreen(QSplashScreen):
    DURACAO_MS = 2000
    PASSOS     = 50

    def __init__(self):
        self._base = self._build()
        self._prog = 0
        super().__init__(self._base, Qt.WindowStaysOnTopHint)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        t = QTimer(self); t.timeout.connect(self._tick); t.start(self.DURACAO_MS // self.PASSOS)
        self._t = t

    def _tick(self):
        if self._prog < self.PASSOS: self._prog += 1; self._paint()
        else: self._t.stop()

    def _paint(self):
        px = self._base.copy(); p = QPainter(px)
        p.setRenderHint(QPainter.Antialiasing)
        W = px.width(); bx, by, bw, bh = 70, 228, W - 140, 4
        p.setBrush(QBrush(QColor("#ffffff30"))); p.setPen(Qt.NoPen)
        p.drawRoundedRect(bx, by, bw, bh, 2, 2)
        fw = int(bw * self._prog / self.PASSOS)
        if fw:
            p.setBrush(QBrush(QColor("#ffffff"))); p.drawRoundedRect(bx, by, fw, bh, 2, 2)
        p.end(); self.setPixmap(px)

    def _build(self):
        W, H = 460, 280; px = QPixmap(W, H); px.fill(Qt.transparent)
        p = QPainter(px); p.setRenderHint(QPainter.Antialiasing)
        grad = QLinearGradient(0, 0, W, H)
        grad.setColorAt(0, QColor(BLUE_MID)); grad.setColorAt(1, QColor(BLUE_DARK))
        path = QPainterPath(); path.addRoundedRect(0, 0, W, H, 20, 20)
        p.fillPath(path, grad)
        p.fillRect(0, H - 56, W, 30, QColor("#00000020"))
        p.setBrush(QBrush(QColor("#00000018"))); p.setPen(Qt.NoPen)
        p.drawRoundedRect(0, H - 56, W, 56, 20, 20)
        logo = vpx("Icon_branco2.ico", 90, 90)
        if not logo.isNull(): p.drawPixmap((W - logo.width()) // 2, 22, logo)
        else:
            p.setPen(QColor("white")); p.setFont(QFont("Segoe UI", 28, QFont.Bold))
            p.drawText(QRect(0, 22, W, 60), Qt.AlignCenter, "BB")
        p.setPen(QColor("white")); p.setFont(QFont("Segoe UI", 20, QFont.Bold))
        p.drawText(QRect(0, 128, W, 36), Qt.AlignCenter, "Conversor CNAB BB")
        p.setFont(QFont("Segoe UI", 10)); p.setPen(QColor("#c8dcf0"))
        p.drawText(QRect(0, 170, W, 24), Qt.AlignCenter, "Excel  →  Remessa CNAB 240  ·  Banco do Brasil")
        p.setPen(QColor("#7eabd4")); p.setFont(QFont("Segoe UI", 9))
        p.drawText(QRect(0, H - 48, W, 40), Qt.AlignCenter, "Inicializando...")
        p.end(); return px


# ──────────────────────────────────────────────────────────────────
#  DropZone  (estados: normal / hover / active)
# ──────────────────────────────────────────────────────────────────
class DropZone(QFrame):
    files_dropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(110)
        self._state = "normal"
        self._apply_style()

        vbox = QVBoxLayout(self)
        vbox.setAlignment(Qt.AlignCenter)
        vbox.setSpacing(6)

        # Ícone circular
        self._ico_frame = QFrame()
        self._ico_frame.setFixedSize(42, 42)
        self._ico_frame.setStyleSheet(f"""
            QFrame {{
                background: {BLUE_LIGHT};
                border-radius: 21px;
                border: 2px solid {BLUE}60;
            }}
        """)
        ico_layout = QHBoxLayout(self._ico_frame)
        ico_layout.setContentsMargins(0, 0, 0, 0)
        ico_lbl = QLabel()
        ico_lbl.setAlignment(Qt.AlignCenter)
        ico_lbl.setStyleSheet("background: transparent; border: none;")
        px = vpx("export.png", 20, 20)
        if not px.isNull(): ico_lbl.setPixmap(px)
        else: ico_lbl.setText("↑"); ico_lbl.setFont(QFont("Segoe UI", 16, QFont.Bold))
        ico_layout.addWidget(ico_lbl)

        ico_row = QHBoxLayout(); ico_row.addWidget(self._ico_frame)
        vbox.addLayout(ico_row)

        self._lbl_main = QLabel(
            'Arraste e solte ou '
            f'<a href="#" style="color:{BLUE}; font-weight:700; text-decoration:underline;">clique para selecionar</a>'
            ' arquivos Excel (.xlsx, .xls)'
        )
        self._lbl_main.setAlignment(Qt.AlignCenter)
        self._lbl_main.setOpenExternalLinks(False)
        self._lbl_main.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self._lbl_main.linkActivated.connect(self._browse)
        self._lbl_main.setStyleSheet(f"font-size: 13px; color: {TEXT2}; background: transparent; border: none;")
        vbox.addWidget(self._lbl_main)

        self._lbl_sub = QLabel("Tamanho máximo: 10MB por arquivo")
        self._lbl_sub.setAlignment(Qt.AlignCenter)
        self._lbl_sub.setStyleSheet(f"font-size: 11px; color: {TEXT3}; background: transparent; border: none;")
        vbox.addWidget(self._lbl_sub)

    def _apply_style(self):
        styles = {
            "normal": f"background: {BG}; border: 2px dashed {BORDER_DARK}; border-radius: 12px;",
            "hover":  f"background: {BLUE_LIGHT}; border: 2px dashed {BLUE}; border-radius: 12px;",
            "active": f"background: {BLUE_LIGHT}; border: 2px solid {BLUE}; border-radius: 12px;",
        }
        self.setStyleSheet(f"QFrame {{ {styles[self._state]} }}")

    def _browse(self, _=None):
        paths, _ = QFileDialog.getOpenFileNames(self, "Selecionar arquivos Excel", "", "Excel (*.xlsx *.xls)")
        if paths: self.files_dropped.emit(paths)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton: self._browse()

    def enterEvent(self, e):
        self._state = "hover"; self._apply_style()

    def leaveEvent(self, e):
        self._state = "normal"; self._apply_style()

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self._state = "active"; self._apply_style()

    def dragLeaveEvent(self, e):
        self._state = "normal"; self._apply_style()

    def dropEvent(self, e):
        self._state = "normal"; self._apply_style()
        paths = [u.toLocalFile() for u in e.mimeData().urls()
                 if u.toLocalFile().lower().endswith((".xlsx", ".xls"))]
        if paths: self.files_dropped.emit(paths)


# ──────────────────────────────────────────────────────────────────
#  FileCard — card informativo por arquivo
# ──────────────────────────────────────────────────────────────────
class FileCard(QFrame):
    """Card horizontal com ícone, nome, tamanho, badge e botão remover."""

    STATUS_DEF = {
        "pronto":      ("Ok",      "ok"),
        "erro":        ("Erro",        "erro"),
        "aviso":       ("Aviso",       "aviso"),
        "processando": ("Processando", "processando"),
        "concluido":   ("✓", "ok"),
        "falhou":      ("Falhou ✗",   "erro"),
    }

    def __init__(self, filepath: str, on_remove, on_select, numero: int, odd=True, parent=None):
        super().__init__(parent)
        self.filepath   = filepath
        self._on_select = on_select
        self._odd       = odd
        self._selected  = False
        self._bg        = "#f8fafc" if odd else CARD
        self._apply_style(False)
        self.setCursor(Qt.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFixedHeight(58)

        row = QHBoxLayout(self)
        row.setContentsMargins(14, 0, 14, 0)
        row.setSpacing(12)

        # Ícone do arquivo
        ico = QLabel()
        ico.setFixedSize(34, 34)
        ico.setAlignment(Qt.AlignCenter)
        ico.setStyleSheet(f"""
            background: {BLUE_LIGHT}; border-radius: 8px; border: 1px solid {BLUE}30;
            font-size: 10px; font-weight: 700; color: {BLUE};
        """)
        ext = os.path.splitext(filepath)[1].upper().replace(".", "")
        px_ext = vpx("processo.ico", 18, 18)
        if not px_ext.isNull(): ico.setPixmap(px_ext)
        else: ico.setText(ext)
        row.addWidget(ico)

        # Info: nome + tamanho
        info = QVBoxLayout(); info.setSpacing(1)
        nome = os.path.basename(filepath)
        self._lbl_nome = QLabel(nome)
        self._lbl_nome.setStyleSheet(f"font-size: 12px; font-weight: 600; color: {TEXT}; background: transparent;")
        self._lbl_nome.setMaximumWidth(340)
        metrics = QFontMetrics(self._lbl_nome.font())
        elided = metrics.elidedText(nome, Qt.ElideMiddle, 340)
        self._lbl_nome.setText(elided)
        info.addWidget(self._lbl_nome)

        try:
            size_kb = os.path.getsize(filepath) // 1024
            size_str = f"{size_kb} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB"
        except Exception:
            size_str = "—"
        lbl_size = QLabel(size_str)
        lbl_size.setStyleSheet(f"font-size: 10px; color: {TEXT3}; background: transparent;")
        info.addWidget(lbl_size)

        row.addLayout(info, 1)

        # Badge de status
        self._badge = Badge("Carregando…", "processando")
        row.addWidget(self._badge)

        # Botão remover
        btn_del = _btn("✕", "icon_del")
        btn_del.clicked.connect(lambda: on_remove(self))
        row.addWidget(btn_del)

    def _apply_style(self, sel: bool):
        if sel:
            self.setStyleSheet(f"""
                QFrame {{
                    background: {BLUE_LIGHT};
                    border-left: 3px solid {BLUE};
                    border-top: none; border-right: none; border-bottom: none;
                }}
            """)
        else:
            self.setStyleSheet(f"QFrame {{ background: {self._bg}; border: none; }}")

    def set_selected(self, sel: bool):
        self._selected = sel; self._apply_style(sel)

    def set_status(self, key: str):
        txt, kind = self.STATUS_DEF.get(key, (key, "info"))
        self._badge.setText(txt)
        fg, bg = Badge.COLORS.get(kind, (TEXT2, BORDER))
        self._badge.setStyleSheet(f"""
            QLabel {{
                background: {bg}; color: {fg};
                border-radius: 10px; padding: 2px 10px;
                font-size: 11px; font-weight: 700; border: 1px solid {fg}40;
            }}
        """)

    def enterEvent(self, e):
        if not self._selected:
            self.setStyleSheet(f"QFrame {{ background: {BLUE_LIGHT}; border: none; }}")

    def leaveEvent(self, e):
        if not self._selected:
            self._apply_style(False)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton: self._on_select(self.filepath)


# ──────────────────────────────────────────────────────────────────
#  GerenciadorArquivos — header com título + botão + lista de cards
# ──────────────────────────────────────────────────────────────────
class GerenciadorArquivos(Card):
    def __init__(self, on_clear, parent=None):
        super().__init__(radius=12, parent=parent)
        self._on_clear = on_clear
        self._file_rows: dict[str, FileCard] = {}
        self.setStyleSheet(f"""
            QFrame {{
        background: {CARD};
        border-top-left-radius: 0px;
        border-top-right-radius: 0px;
        border-bottom-left-radius: 0px;
        border-bottom-right-radius: 0px;
        border: none;
            }}
        """)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ── Header do gerenciador ───────────────────────────────
        hdr = QFrame()
        hdr.setFixedHeight(46)
        hdr.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 {BLUE}, stop:1 {BLUE_MID});
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
                border: none;
            }}
        """)
        hrow = QHBoxLayout(hdr)
        hrow.setContentsMargins(16, 0, 12, 0)
        hrow.setSpacing(8)

        ico_hdr = QLabel()
        ico_hdr.setStyleSheet("background: transparent; border: none;")
        px_h = vpx("adicionar.ico", 16, 16)
        if not px_h.isNull(): ico_hdr.setPixmap(px_h)
        hrow.addWidget(ico_hdr)

        lbl_hdr = QLabel("Gerenciador de Arquivos")
        lbl_hdr.setStyleSheet("color: white; font-size: 13px; font-weight: 700; background: transparent; border: none;")
        hrow.addWidget(lbl_hdr, 1)

        self._lbl_count = QLabel("0 arquivo(s)")
        self._lbl_count.setStyleSheet("color: #c8dcf0; font-size: 11px; background: transparent; border: none;")
        hrow.addWidget(self._lbl_count)

        outer.addWidget(hdr)

        # ── Scroll ──────────────────────────────────────────────
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setFixedHeight(170)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setStyleSheet(f"""
            QScrollArea {{ border: none; background: {CARD}; }}
            QScrollBar:vertical {{
                background: {BORDER}; width: 5px; border-radius: 2px;
            }}
            QScrollBar::handle:vertical {{
                background: #a0aab4; border-radius: 2px; min-height: 20px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        """)

        self._body = QWidget()
        self._body.setStyleSheet(f"background: {CARD};")
        self._body_layout = QVBoxLayout(self._body)
        self._body_layout.setContentsMargins(0, 0, 0, 0)
        self._body_layout.setSpacing(0)
        self._body_layout.addStretch()
        self._scroll.setWidget(self._body)
        outer.addWidget(self._scroll)

        # ── Estado vazio ─────────────────────────────────────────
        self._empty = QFrame()
        self._empty.setFixedHeight(170)
        self._empty.setStyleSheet(f"background: {CARD}; border: none;")
        ev = QVBoxLayout(self._empty)
        ev.setAlignment(Qt.AlignCenter); ev.setSpacing(4)
        el = QLabel("Nenhum arquivo adicionado")
        el.setAlignment(Qt.AlignCenter)
        el.setStyleSheet(f"color: {TEXT3}; font-size: 12px; background: transparent;")
        es = QLabel("Arraste arquivos Excel para a área acima")
        es.setAlignment(Qt.AlignCenter)
        es.setStyleSheet(f"color: {TEXT3}; font-size: 11px; background: transparent;")
        ev.addWidget(el); ev.addWidget(es)
        outer.addWidget(self._empty)

        # Rodapé arredondado
        foot = QFrame()
        foot.setFixedHeight(8)
        foot.setStyleSheet(f"""
            QFrame {{
                background: {CARD};
                border-bottom-left-radius: 12px;
                border-bottom-right-radius: 12px;
                border: none;
            }}
        """)
        outer.addWidget(foot)

        self._update_view()

    def _update_view(self):
        n = len(self._file_rows)
        self._lbl_count.setText(f"{n} arquivo{'s' if n != 1 else ''}") 
        self._scroll.setVisible(n > 0)
        self._empty.setVisible(n == 0)

    def add_card(self, fc: FileCard):
        self._file_rows[fc.filepath] = fc
        self._body_layout.insertWidget(self._body_layout.count() - 1, fc)
        self._update_view()

    def remove_card(self, fc: FileCard):
        self._file_rows.pop(fc.filepath, None)
        self._body_layout.removeWidget(fc)
        fc.deleteLater()
        self._update_view()

    def clear_all(self):
        for fc in list(self._file_rows.values()):
            self._body_layout.removeWidget(fc); fc.deleteLater()
        self._file_rows.clear()
        self._update_view()

    def get_row(self, filepath: str) -> Optional["FileCard"]:
        return self._file_rows.get(filepath)

    def select_only(self, filepath: str):
        for fp, fc in self._file_rows.items():
            fc.set_selected(fp == filepath)


# ──────────────────────────────────────────────────────────────────
#  PreviewPanel — cabeçalho + tabela com hover
# ──────────────────────────────────────────────────────────────────
class PreviewPanel(Card):
    def __init__(self, parent=None):
        super().__init__(radius=12, parent=parent)
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(0)
        self._show_empty()

    def _clear(self):
        while self._layout.count():
            item = self._layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()

    def _header_bar(self, title: str, subtitle: str = ""):
        hdr = QFrame()
        hdr.setFixedHeight(46)
        hdr.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 {BLUE}, stop:1 {BLUE_MID});
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
                border-bottom-left-radius: 0px;
                border-bottom-right-radius: 0px;
                border: none;
            }}
        """)
        row = QHBoxLayout(hdr)
        row.setContentsMargins(16, 0, 12, 0); row.setSpacing(8)
        t = QLabel(title)
        t.setStyleSheet("color: white; font-size: 13px; font-weight: 700; background: transparent; border: none;")
        row.addWidget(t, 1)
        if subtitle:
            s = QLabel(subtitle)
            s.setTextFormat(Qt.RichText)
            s.setStyleSheet("color: #c8dcf0; font-size: 11px; background: transparent; border: none;")
            row.addWidget(s)
        return hdr

    def _show_empty(self):
        self._clear()
        body = QFrame()
        body.setStyleSheet(f"""
            QFrame {{
                background: {CARD};
                border-radius: 12px;
                border: 1px solid {BORDER};
            }}
        """)
        body.setMinimumHeight(140)
        v = QVBoxLayout(body); v.setAlignment(Qt.AlignCenter); v.setSpacing(8)
        ico = QLabel()
        ico.setAlignment(Qt.AlignCenter)
        ico.setStyleSheet("background: transparent; border: none;")
        px = vpx("eye.png", 28, 28)
        if not px.isNull(): ico.setPixmap(px)
        v.addWidget(ico)
        l1 = QLabel("Nenhum arquivo selecionado")
        l1.setAlignment(Qt.AlignCenter)
        l1.setStyleSheet(f"color: {TEXT2}; font-size: 13px; font-weight: 600; background: transparent; border: none;")
        v.addWidget(l1)
        l2 = QLabel("Clique em um arquivo no Gerenciador para visualizar os dados")
        l2.setAlignment(Qt.AlignCenter)
        l2.setStyleSheet(f"color: {TEXT3}; font-size: 11px; background: transparent; border: none;")
        v.addWidget(l2)
        self._layout.addWidget(body)

    def show_dataframe(self, df: pd.DataFrame, filename: str, total_valor: str = ""):
        self._clear()
        preview = df.head(8)
        cols    = list(preview.columns)
        n_rows  = len(preview)

        subtitle = f"{len(df)} registros encontrados"
        if total_valor:
            subtitle = f"{subtitle} · Valor total: <b><span style='color:#ffffff'>{total_valor}</span></b>"
        hdr = self._header_bar("Preview dos Dados", subtitle)
        self._layout.addWidget(hdr)

        # Cabeçalho das colunas
        col_hdr = QFrame()
        col_hdr.setFixedHeight(32)
        col_hdr.setStyleSheet(f"background: {BG}; border: none;")
        crow = QHBoxLayout(col_hdr)
        crow.setContentsMargins(14, 0, 14, 0); crow.setSpacing(6)
        for col in cols:
            l = QLabel(str(col).upper()[:16])
            l.setStyleSheet(f"color: {TEXT2}; font-size: 10px; font-weight: 700; background: transparent;")
            l.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            crow.addWidget(l)
        self._layout.addWidget(col_hdr)

        self._layout.addWidget(Divider())

        # Scroll com linhas hover
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFixedHeight(min(n_rows * 28 + 4, 196))
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet(f"""
            QScrollArea {{ border: none; background: {CARD}; border-bottom-left-radius: 12px; border-bottom-right-radius: 12px; }}
            QScrollBar:vertical {{ background: {BORDER}; width: 5px; border-radius: 2px; }}
            QScrollBar::handle:vertical {{ background: #a0aab4; border-radius: 2px; min-height: 20px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        """)
        inner = QWidget(); inner.setStyleSheet(f"background: {CARD};")
        il = QVBoxLayout(inner); il.setContentsMargins(0, 0, 0, 0); il.setSpacing(0)

        for i, (_, row) in enumerate(preview.iterrows()):
            bg = "#f8fafc" if i % 2 == 0 else CARD
            rf = _HoverRow(bg)
            rf.setFixedHeight(28)
            rrow = QHBoxLayout(rf); rrow.setContentsMargins(14, 0, 14, 0); rrow.setSpacing(6)
            for col in cols:
                val = str(row.get(col, ""))[:22]
                l = QLabel(val)
                l.setStyleSheet(f"color: {TEXT}; font-size: 10px; background: transparent;")
                l.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
                rrow.addWidget(l)
            il.addWidget(rf)

        il.addStretch()
        scroll.setWidget(inner)
        self._layout.addWidget(scroll)

    def show_error(self, msg: str):
        self._clear()
        hdr = self._header_bar("Preview dos Dados")
        self._layout.addWidget(hdr)
        body = QFrame()
        body.setMinimumHeight(100)
        body.setStyleSheet(f"background: {RED_BG}; border: none; border-bottom-left-radius: 12px; border-bottom-right-radius: 12px;")
        v = QVBoxLayout(body); v.setAlignment(Qt.AlignCenter)
        l = QLabel(f"⚠  {msg[:140]}")
        l.setWordWrap(True); l.setAlignment(Qt.AlignCenter)
        l.setStyleSheet(f"color: {RED}; font-size: 12px; background: transparent;")
        v.addWidget(l)
        self._layout.addWidget(body)


class _HoverRow(QFrame):
    def __init__(self, bg, parent=None):
        super().__init__(parent); self._bg = bg
        self.setStyleSheet(f"QFrame {{ background: {bg}; border: none; }}")
    def enterEvent(self, e):
        self.setStyleSheet(f"QFrame {{ background: {BLUE_LIGHT}; border: none; }}")
    def leaveEvent(self, e):
        self.setStyleSheet(f"QFrame {{ background: {self._bg}; border: none; }}")


# ──────────────────────────────────────────────────────────────────
#  ControlPanel — seção de controles (tipo, forma, data)
# ──────────────────────────────────────────────────────────────────
class ControlPanel(Card):
    def __init__(self, parent=None):
        super().__init__(radius=12, parent=parent)
        self.setStyleSheet(f"""
            QFrame {{
        background: {CARD};
        border-top-left-radius: 12px;
        border-top-right-radius: 12px;
        border-bottom-left-radius: 0px;
        border-bottom-right-radius: 0px;
        border: none;
            }}
        """)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # Header gradiente
        hdr = QFrame()
        hdr.setFixedHeight(42)
        hdr.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 {BLUE}, stop:1 {BLUE_MID});
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
                border: none;
            }}
        """)
        hrow = QHBoxLayout(hdr)
        hrow.setContentsMargins(16, 0, 16, 0)
        lbl = QLabel("Configurações do Pagamento")
        lbl.setStyleSheet("color: white; font-size: 13px; font-weight: 700; background: transparent; border: none;")
        hrow.addWidget(lbl)
        outer.addWidget(hdr)

        # Corpo dos controles em grid 3 colunas
        body = QFrame()
        body.setStyleSheet(f"background: {CARD}; border: none; border-bottom-left-radius: 12px; border-bottom-right-radius: 12px;")
        row = QHBoxLayout(body)
        row.setContentsMargins(20, 14, 20, 14)
        row.setSpacing(32)

        # Coluna 1: Tipo de Pagamento
        self._tipo_pag = "PGA"
        row.addLayout(self._col_tipo())

        # Separador
        sep1 = QFrame()
        sep1.setFrameShape(QFrame.VLine)
        sep1.setStyleSheet(f"border: none; border-left: 1px solid {BORDER}; background: transparent;")
        row.addWidget(sep1)

        # Coluna 2: Forma de Lançamento
        self._forma = "03"
        row.addLayout(self._col_forma())

        # Separador
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.VLine)
        sep2.setStyleSheet(f"border: none; border-left: 1px solid {BORDER}; background: transparent;")
        row.addWidget(sep2)

        # Coluna 3: Data de Pagamento
        row.addLayout(self._col_data())
        row.addStretch()

        outer.addWidget(body)

    def _lbl_group(self, text: str) -> QLabel:
        l = QLabel(text)
        l.setStyleSheet(f"color: {TEXT2}; font-size: 11px; font-weight: 700; background: transparent; letter-spacing: 0.5px;")
        return l

    def _radio_style(self):
        return f"""
            QRadioButton {{
                color: {TEXT}; font-size: 12px; font-weight: 600;
                background: transparent; spacing: 8px;
            }}
            QRadioButton::indicator {{
                width: 16px; height: 16px; border-radius: 8px;
                border: 2px solid {BORDER_DARK}; background: white;
            }}
            QRadioButton::indicator:checked {{
                background: {BLUE}; border: 2px solid {BLUE};
            }}
            QRadioButton::indicator:hover {{
                border-color: {BLUE};
            }}
        """

    def _col_tipo(self) -> QVBoxLayout:
        col = QVBoxLayout(); col.setSpacing(8)
        col.addWidget(self._lbl_group("TIPO DE PAGAMENTO"))
        rs = self._radio_style()
        self._btn_pga   = QRadioButton("PGA");   self._btn_pga.setChecked(True);  self._btn_pga.setStyleSheet(rs)
        self._btn_jusmp = QRadioButton("JUSMP"); self._btn_jusmp.setStyleSheet(rs)
        self._grp_tipo  = QButtonGroup()
        self._grp_tipo.addButton(self._btn_pga,   0)
        self._grp_tipo.addButton(self._btn_jusmp, 1)
        self._grp_tipo.buttonClicked.connect(
            lambda b: setattr(self, "_tipo_pag", "JUSMP" if b is self._btn_jusmp else "PGA")
        )
        col.addWidget(self._btn_pga)
        col.addWidget(self._btn_jusmp)
        return col

    def _col_forma(self) -> QVBoxLayout:
        col = QVBoxLayout(); col.setSpacing(8)
        col.addWidget(self._lbl_group("FORMA DE LANÇAMENTO"))
        rs = self._radio_style()
        self._btn_f03 = QRadioButton("03 – DOC / TED"); self._btn_f03.setChecked(True); self._btn_f03.setStyleSheet(rs)
        self._btn_f01 = QRadioButton("01 – Transf. BB"); self._btn_f01.setStyleSheet(rs)
        self._grp_forma = QButtonGroup()
        self._grp_forma.addButton(self._btn_f03, 0)
        self._grp_forma.addButton(self._btn_f01, 1)
        self._grp_forma.buttonClicked.connect(
            lambda b: setattr(self, "_forma", "01" if b is self._btn_f01 else "03")
        )
        col.addWidget(self._btn_f03)
        col.addWidget(self._btn_f01)
        return col

    def _col_data(self) -> QVBoxLayout:
        col = QVBoxLayout(); col.setSpacing(8)
        col.addWidget(self._lbl_group("DATA DE PAGAMENTO"))
        self._dt = QDateEdit()
        self._dt.setCalendarPopup(True)
        self._dt.setDisplayFormat("dd/MM/yyyy")
        self._dt.setDate(QDate.currentDate())
        self._dt.setFixedHeight(34)
        self._dt.setStyleSheet(f"""
            QDateEdit {{
                background: white; color: {TEXT2};
                border: 1.5px solid {BORDER_DARK}; border-radius: 8px;
                padding: 4px 10px; font-size: 13px; font-weight: 600;
                min-width: 138px;
            }}
            QDateEdit:focus {{ border-color: {BLUE}; }}
            QDateEdit::drop-down {{
                subcontrol-origin: padding; subcontrol-position: top right;
                width: 28px; border: none;
            }}
            QDateEdit::down-arrow {{
                image: url("{vp('seta_colorida.png').replace(chr(92), '/')}");
                width: 14px; height: 14px;
            }}
            QCalendarWidget {{ background: white; color: black; }}
            QCalendarWidget QWidget {{ background: white; color: black; }}
            QCalendarWidget QAbstractItemView {{
                background: white; color: black;
                selection-background-color: {BLUE}; selection-color: white;
            }}
            QCalendarWidget QToolButton {{
                background: white; color: black;
                border: 1px solid {BORDER}; padding: 4px; margin: 2px;
            }}
            QCalendarWidget QSpinBox {{
                background: white; color: black; border: 1px solid {BORDER};
            }}
        """)
        col.addWidget(self._dt)
        col.addStretch()
        return col

    @property
    def tipo_pag(self) -> str: return self._tipo_pag
    @property
    def forma_lancamento(self) -> str: return self._forma
    @property
    def payment_date(self) -> str: return self._dt.date().toString("dd/MM/yyyy")


# ──────────────────────────────────────────────────────────────────
#  ConversionWorker
# ──────────────────────────────────────────────────────────────────
class ConversionWorker(QThread):
    log_signal   = pyqtSignal(str, str)
    file_done    = pyqtSignal(str, bool)
    finished_all = pyqtSignal(int, int)

    def __init__(self, files, cfg, tipo, output_dir,
                 modo="pagamentos", payment_date=None, forma_lancamento=None):
        super().__init__()
        self._files = files; self._cfg = cfg; self._tipo = tipo
        self._output_dir = output_dir; self._modo = modo
        self._payment_date = payment_date; self._forma = forma_lancamento

    def run(self):
        cfg_exec = ConfigLoader(self._cfg.config_path, tipo_pagamento=self._tipo)
        if self._forma:
            cfg_exec.lote["forma_lancamento"] = str(self._forma).zfill(2)
        gen = CNABGenerator(cfg_exec)
        success = errors = 0
        for filepath, data in self._files.items():
            fname = os.path.basename(filepath)
            try:
                if self._modo == "aplicacoes":
                    content = gen.generate(data["df"], tipo_pagamento="APLICACAO")
                else:
                    content = gen.generate(data["df"], tipo_pagamento=self._tipo,
                                           payment_date=self._payment_date)
                base    = os.path.splitext(fname)[0]
                now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                out     = os.path.join(self._output_dir, f"CNAB_{base}_{now_str}.txt")
                with open(out, "w", encoding="ascii", errors="replace") as f:
                    f.write(content)
                self.file_done.emit(filepath, True); success += 1
            except Exception as e:
                self.log_signal.emit(f"Erro em {fname}: {e}", "error")
                self.file_done.emit(filepath, False); errors += 1
        self.finished_all.emit(success, errors)


# ──────────────────────────────────────────────────────────────────
#  AbaPagamentos
# ──────────────────────────────────────────────────────────────────
class AbaPagamentos(QWidget):
    def __init__(self, cfg, cfg_error):
        super().__init__()
        self._cfg = cfg; self._cfg_error = cfg_error
        self._files: dict[str, dict] = {}
        self._selected: Optional[str] = None
        self._worker = None
        self.setStyleSheet(f"background: transparent;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 10, 0, 6)
        layout.setSpacing(10)

        # Controles
        self._ctrl = ControlPanel()
        layout.addWidget(self._ctrl)

        # Drop zone
        self._drop = DropZone()
        self._drop.files_dropped.connect(self._on_dropped)
        layout.addWidget(self._drop)

        # Gerenciador
        self._ger = GerenciadorArquivos(on_clear=self._limpar_tudo)
        layout.addWidget(self._ger)

        # Preview
        self._prev = PreviewPanel()
        layout.addWidget(self._prev)

        # Botões
        layout.addLayout(self._build_botoes())

    def _build_botoes(self):
        row = QHBoxLayout(); row.setSpacing(10); row.addStretch()
        self._btn_clear = _btn("  Limpar Tudo", "ghost", "limpar.ico", w=150, h=44)
        self._btn_clear.clicked.connect(self._limpar_tudo)
        row.addWidget(self._btn_clear)
        self._btn_exec = _btn("  Processar e Gerar CNAB", "primary", "import.png", w=220, h=44)
        self._btn_exec.setEnabled(False)
        self._btn_exec.clicked.connect(self._processar)
        row.addWidget(self._btn_exec)
        return row

    def _on_dropped(self, paths: list):
        for p in paths:
            if p not in self._files:
                df, errs = read_excel(p, aba="pagamentos")
                self._files[p] = {"df": df, "errors": errs}
                fc = FileCard(p, self._ger.remove_card, self._select_file,
                              numero=len(self._files), odd=len(self._files) % 2 == 1)
                self._ger.add_card(fc)
                status = "erro" if (errs and df.empty) else ("aviso" if errs else "pronto")
                fc.set_status(status)
        self._update_btn()
        if paths: self._select_file(paths[-1])

    def _select_file(self, fp: str):
        self._selected = fp
        self._ger.select_only(fp)
        data = self._files.get(fp)
        if not data: return
        if data["errors"] and data["df"].empty:
            self._prev.show_error(data["errors"][0])
        else:
            df = data["df"]
            total_valor = ""
            if "valor" in df.columns and not df.empty:
                total = pd.to_numeric(df["valor"], errors="coerce").fillna(0).sum()
                total_valor = f"R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            self._prev.show_dataframe(df, os.path.basename(fp), total_valor=total_valor)

    def _limpar_tudo(self):
        self._files.clear(); self._selected = None
        self._ger.clear_all(); self._prev._show_empty()
        self._update_btn()

    def _update_btn(self):
        has_valid = any(not d["df"].empty for d in self._files.values())
        self._btn_exec.setEnabled(has_valid)

    def _processar(self):
        if not self._cfg:
            _msgbox(self, "error", "Configuração", f"config.json inválido:\n{self._cfg_error}"); return
        valid = {fp: d for fp, d in self._files.items() if not d["df"].empty}
        if not valid:
            _msgbox(self, "warn", "Sem arquivos", "Nenhum arquivo válido."); return
        out_dir = QFileDialog.getExistingDirectory(self, "Pasta de saída")
        if not out_dir: return

        self._btn_exec.setEnabled(False)
        self._btn_exec.setText("  Processando…")
        for fp in valid: self._ger.get_row(fp) and self._ger.get_row(fp).set_status("processando")

        self._worker = ConversionWorker(
            valid, self._cfg, self._ctrl.tipo_pag, out_dir,
            payment_date=self._ctrl.payment_date,
            forma_lancamento=self._ctrl.forma_lancamento,
        )
        self._worker.file_done.connect(self._on_done)
        self._worker.finished_all.connect(self._on_finished)
        self._worker.start()

    def _on_done(self, fp: str, ok: bool):
        fc = self._ger.get_row(fp)
        if fc: fc.set_status("concluido" if ok else "falhou")

    def _on_finished(self, success: int, errors: int):
        self._btn_exec.setEnabled(True)
        self._btn_exec.setText("  Processar e Gerar CNAB")
        if success: _msgbox(self, "info", "Concluído", f"{success} arquivo(s) CNAB gerado(s) com sucesso!")
        if errors:  _msgbox(self, "error", "Erros", f"{errors} arquivo(s) com erro.")


# ──────────────────────────────────────────────────────────────────
#  AbaAplicacoes
# ──────────────────────────────────────────────────────────────────
class AbaAplicacoes(QWidget):
    def __init__(self, cfg, cfg_error):
        super().__init__()
        self._cfg = cfg; self._cfg_error = cfg_error
        self._files: dict[str, dict] = {}
        self._selected = None; self._worker = None
        self.setStyleSheet("background: transparent;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 10, 0, 6)
        layout.setSpacing(10)

        self._drop = DropZone()
        self._drop.files_dropped.connect(self._on_dropped)
        layout.addWidget(self._drop)

        self._ger = GerenciadorArquivos(on_clear=self._limpar_tudo)
        layout.addWidget(self._ger)

        self._prev = PreviewPanel()
        layout.addWidget(self._prev)

        layout.addLayout(self._build_botoes())

    def _build_botoes(self):
        row = QHBoxLayout(); row.setSpacing(10); row.addStretch()
        self._btn_clear = _btn("  Limpar Tudo", "ghost", "limpar.ico", w=150, h=44)
        self._btn_clear.clicked.connect(self._limpar_tudo)
        row.addWidget(self._btn_clear)
        self._btn_exec = _btn("  Processar e Gerar CNAB", "primary", "import.png", w=220, h=44)
        self._btn_exec.setEnabled(False)
        self._btn_exec.clicked.connect(self._processar)
        row.addWidget(self._btn_exec)
        return row

    def _on_dropped(self, paths: list):
        for p in paths:
            if p not in self._files:
                df, errs = read_excel(p, aba="aplicacoes")
                self._files[p] = {"df": df, "errors": errs}
                fc = FileCard(p, self._ger.remove_card, self._select_file,
                              numero=len(self._files), odd=len(self._files) % 2 == 1)
                self._ger.add_card(fc)
                fc.set_status("erro" if (errs and df.empty) else ("aviso" if errs else "pronto"))
        self._update_btn()
        if paths: self._select_file(paths[-1])

    def _select_file(self, fp: str):
        self._selected = fp; self._ger.select_only(fp)
        data = self._files.get(fp)
        if not data: return
        if data["errors"] and data["df"].empty: self._prev.show_error(data["errors"][0])
        else:
            df = data["df"]
            total_valor = ""
            if "valor" in df.columns and not df.empty:
                total = pd.to_numeric(df["valor"], errors="coerce").fillna(0).sum()
                total_valor = f"R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            self._prev.show_dataframe(df, os.path.basename(fp), total_valor=total_valor)

    def _limpar_tudo(self):
        self._files.clear(); self._selected = None
        self._ger.clear_all(); self._prev._show_empty(); self._update_btn()

    def _update_btn(self):
        self._btn_exec.setEnabled(any(not d["df"].empty for d in self._files.values()))

    def _processar(self):
        if not self._cfg:
            _msgbox(self, "error", "Configuração", f"config.json inválido:\n{self._cfg_error}"); return
        valid = {fp: d for fp, d in self._files.items() if not d["df"].empty}
        if not valid:
            _msgbox(self, "warn", "Sem arquivos", "Nenhum arquivo válido."); return
        out_dir = QFileDialog.getExistingDirectory(self, "Pasta de saída")
        if not out_dir: return
        self._btn_exec.setEnabled(False); self._btn_exec.setText("  Processando…")
        for fp in valid:
            fc = self._ger.get_row(fp)
            if fc: fc.set_status("processando")
        self._worker = ConversionWorker(valid, self._cfg, "BBJUMP", out_dir, modo="aplicacoes")
        self._worker.file_done.connect(self._on_done)
        self._worker.finished_all.connect(self._on_finished)
        self._worker.start()

    def _on_done(self, fp, ok):
        fc = self._ger.get_row(fp)
        if fc: fc.set_status("concluido" if ok else "falhou")

    def _on_finished(self, s, e):
        self._btn_exec.setEnabled(True); self._btn_exec.setText("  Processar e Gerar CNAB")
        if s: _msgbox(self, "info", "Concluído", f"{s} arquivo(s) gerado(s)!")
        if e: _msgbox(self, "error", "Erros", f"{e} arquivo(s) com erro.")


# ──────────────────────────────────────────────────────────────────
#  MainWindow
# ──────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self, cfg, cfg_error):
        super().__init__()
        self.setWindowTitle("Conversor Excel → CNAB BB")
        self.setWindowIcon(vico("Icon.ico"))
        self.setFixedSize(880, 900)
        self.setStyleSheet(f"QMainWindow {{ background: {BG}; }} {STYLE_APP}")
        
        central = QWidget()
        central.setStyleSheet(f"background: {BG};")
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(20, 16, 20, 12)
        root.setSpacing(12)

        # ── Header profissional ─────────────────────────────────
        root.addWidget(self._build_header(cfg_error))

        # ── Tabs ────────────────────────────────────────────────
        tabs_card = QFrame()
        tabs_card.setStyleSheet(f"""
            QFrame {{
                 background: {CARD};
                 border-radius: 14px;
                 border: none;
            }}
        """)
        _shadow(tabs_card, 24, 4, 16)
        tc_layout = QVBoxLayout(tabs_card)
        tc_layout.setContentsMargins(18, 10, 18, 14)
        tc_layout.setSpacing(0)

        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(STYLE_TABS)
        self.tabs.addTab(AbaPagamentos(cfg, cfg_error), "  Pagamentos  ")
        self.tabs.addTab(AbaAplicacoes(cfg, cfg_error), "  Aplicações  ")
        tc_layout.addWidget(self.tabs)
        root.addWidget(tabs_card, 1)

        # ── Footer ──────────────────────────────────────────────
        footer = QLabel("Banco do Brasil  ·  CNAB 240  ·  Remessa de Pagamentos")
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet(f"font-size: 10px; color: {TEXT3}; background: transparent;")
        root.addWidget(footer)

    def _build_header(self, cfg_error) -> QFrame:
        hdr = QFrame()
        hdr.setFixedHeight(72)
        hdr.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 {BLUE_DARK}, stop:0.6 {BLUE}, stop:1 {BLUE_MID});
                border-radius: 14px;
                border: none;
            }}
        """)
        _shadow(hdr, 28, 6, 30)

        row = QHBoxLayout(hdr)
        row.setContentsMargins(20, 0, 20, 0)
        row.setSpacing(14)

        # Logo
        logo_frame = QFrame()
        logo_frame.setFixedSize(46, 46)
        logo_frame.setStyleSheet(f"""
            QFrame {{
                background: transparent;
                border-radius: 12px;
                border: none;
            }}
        """)
        ll = QHBoxLayout(logo_frame); ll.setContentsMargins(6, 6, 6, 6)
        logo_lbl = QLabel()
        logo_lbl.setAlignment(Qt.AlignCenter)
        logo_lbl.setStyleSheet("background: transparent; border: none;")
        px_logo = vpx("Icon_branco.ico", 28, 28)
        if not px_logo.isNull(): logo_lbl.setPixmap(px_logo)
        else: logo_lbl.setText("BB"); logo_lbl.setFont(QFont("Segoe UI", 13, QFont.Bold)); logo_lbl.setStyleSheet("color: white; background: transparent; border: none;")
        ll.addWidget(logo_lbl)
        row.addWidget(logo_frame)

        # Textos
        txt = QVBoxLayout(); txt.setSpacing(2)
        t1 = QLabel("Conversor Excel → CNAB BB")
        t1.setStyleSheet("color: white; font-size: 16px; font-weight: 700; background: transparent;")
        t2 = QLabel("Gere arquivos de remessa CNAB 240 para o Banco do Brasil a partir de planilhas Excel")
        t2.setStyleSheet("color: #c8dcf0; font-size: 11px; background: transparent;")
        txt.addWidget(t1); txt.addWidget(t2)
        row.addLayout(txt, 1)

        return hdr


# ──────────────────────────────────────────────────────────────────
#  Helpers globais
# ──────────────────────────────────────────────────────────────────
def _msgbox(parent, tipo, titulo, texto):
    box = QMessageBox(parent)
    box.setWindowTitle(titulo); box.setText(texto)
    box.setStyleSheet(STYLE_MSGBOX)
    box.setIcon({"info": QMessageBox.Information, "warn": QMessageBox.Warning,
                 "error": QMessageBox.Critical}.get(tipo, QMessageBox.Information))
    box.exec_()


def _paleta():
    pal = QPalette()
    pal.setColor(QPalette.Window,          QColor(BG))
    pal.setColor(QPalette.WindowText,      QColor(TEXT))
    pal.setColor(QPalette.Base,            QColor(CARD))
    pal.setColor(QPalette.AlternateBase,   QColor(BG))
    pal.setColor(QPalette.Text,            QColor(TEXT))
    pal.setColor(QPalette.Button,          QColor(BLUE))
    pal.setColor(QPalette.ButtonText,      QColor("white"))
    pal.setColor(QPalette.Highlight,       QColor(BLUE))
    pal.setColor(QPalette.HighlightedText, QColor("white"))
    return pal


# ──────────────────────────────────────────────────────────────────
#  Entry point
# ──────────────────────────────────────────────────────────────────
def run():
    app = QApplication(sys.argv)
    app.setWindowIcon(vico("Icon.ico"))
    app.setFont(QFont("Segoe UI", 10))
    app.setPalette(_paleta())

    cfg = None; cfg_error = None
    try:
        cfg = ConfigLoader()
    except ConfigError as e:
        cfg_error = str(e)

    splash = SplashScreen()
    splash.show(); app.processEvents()

    window = MainWindow(cfg, cfg_error)
    QTimer.singleShot(
        SplashScreen.DURACAO_MS,
        lambda: (splash.finish(window), window.show()),
    )
    sys.exit(app.exec_())


if __name__ == "__main__":
    run()