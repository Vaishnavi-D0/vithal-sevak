"""
marathi_keyboard.py
On-screen Marathi (Devanagari) keyboard for correcting auto-translated text
directly inside the app, without needing an OS-level Marathi keyboard.
"""

from PyQt5.QtWidgets import QDialog, QGridLayout, QPushButton, QVBoxLayout, QTextEdit
from PyQt5.QtGui import QFont

ROWS = [
    "अ आ इ ई उ ऊ ए ऐ ओ औ अं अः".split(),
    "क ख ग घ ङ च छ ज झ ञ".split(),
    "ट ठ ड ढ ण त थ द ध न".split(),
    "प फ ब भ म य र ल व".split(),
    "श ष स ह ळ क्ष ज्ञ".split(),
    "ा ि ी ु ू े ै ो ौ ं ः ्".split(),
    "० १ २ ३ ४ ५ ६ ७ ८ ९".split(),
]


class MarathiKeyboard(QDialog):
    """A shared, reusable virtual keyboard. Call attach_to(line_edit) to
    redirect keystrokes to a particular field before showing it."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("मराठी कीबोर्ड (Marathi Keyboard)")
        self.target = None

        layout = QVBoxLayout()
        font = QFont("Noto Sans Devanagari", 13)

        for row in ROWS:
            grid = QGridLayout()
            for i, ch in enumerate(row):
                btn = QPushButton(ch)
                btn.setFont(font)
                btn.setFixedSize(42, 36)
                btn.clicked.connect(lambda _, c=ch: self.insert_char(c))
                grid.addWidget(btn, 0, i)
            layout.addLayout(grid)

        controls = QGridLayout()
        space_btn = QPushButton("Space")
        space_btn.clicked.connect(lambda: self.insert_char(" "))
        back_btn = QPushButton("⌫ Backspace")
        back_btn.clicked.connect(self.backspace)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        controls.addWidget(space_btn, 0, 0)
        controls.addWidget(back_btn, 0, 1)
        controls.addWidget(close_btn, 0, 2)
        layout.addLayout(controls)

        self.setLayout(layout)

    def attach_to(self, field):
        """Set which field (QLineEdit or QTextEdit) receives the keyboard's input."""
        self.target = field

    def insert_char(self, ch):
        if self.target is None:
            return
        if isinstance(self.target, QTextEdit):
            self.target.textCursor().insertText(ch)
        else:
            self.target.insert(ch)
        self.target.setFocus()

    def backspace(self):
        if self.target is None:
            return
        if isinstance(self.target, QTextEdit):
            self.target.textCursor().deletePreviousChar()
        else:
            cursor_pos = self.target.cursorPosition()
            if cursor_pos > 0:
                text = self.target.text()
                new_text = text[:cursor_pos - 1] + text[cursor_pos:]
                self.target.setText(new_text)
                self.target.setCursorPosition(cursor_pos - 1)
        self.target.setFocus()
