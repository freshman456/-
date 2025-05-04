import sys

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QKeySequence
from PyQt6.QtWidgets import QApplication, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QVBoxLayout, QWidget

from LinesUtils import LinesUtils


class EditRectWindow(QWidget):
	closeSignal = pyqtSignal()
	modifySignal = pyqtSignal()

	def __init__(self, parent=None):
		super().__init__(parent)
		self.viewText = None
		self.h_line = None
		self.w_line = None
		self.content_edit = None
		self.text_label = None
		self.text_label_title = None
		self.main_layout = None
		self.setWindowTitle("编辑窗口")
		self.setFixedSize(300, 280)
		self.move(950, 300)
		self.setStyleSheet("""
	        QWidget {
	        border: 1px solid black;
	        background-color: rgb(255, 255, 255);
	        }
	    """)
		# 设置窗口属性
		self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)  # 关键属性
		self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)

	def initUI(self):
		self.main_layout = QVBoxLayout()
		self.text_label_title = QLabel("单元格信息:")
		self.text_label_title.setStyleSheet("""
			QLabel {
				font-size: 12px;
				font-weight: bold;
				border:none;
			}
		""")
		self.text_label = QLabel("单元格信息展示")
		self.text_label.setFixedHeight(25)
		self.text_label.setWordWrap(True)  #

		self.main_layout.addWidget(self.text_label_title)
		self.main_layout.addWidget(self.text_label)

		title_layout = QHBoxLayout()

		title_label = QLabel()
		title_label.setText("内容")
		title_label.setFixedHeight(22)
		title_label.setStyleSheet("""
			QLabel {
				font-size: 12px;
				font-weight: bold;
				border:none;
			}
			""")
		title_layout.addWidget(title_label)
		title_layout.addStretch(1)

		self.display_label = QTextEdit()
		self.display_label.setReadOnly(True)
		self.display_label.setFixedHeight(45)

		self.main_layout.addLayout(title_layout)
		self.main_layout.addWidget(self.display_label)

		content_layout = QVBoxLayout()
		content_label = QLabel("编辑框:")
		content_label.setStyleSheet("""
			QLabel {
				font-size: 12px;
				font-weight: bold;
				border:none;
			}
		""")

		content_layout.addWidget(content_label)
		self.content_edit = QTextEdit()
		self.content_edit.setPlaceholderText("内容编辑框")
		self.content_edit.setFixedHeight(45)
		content_layout.addWidget(self.content_edit)

		btn_layout = QHBoxLayout()
		btn_layout.setContentsMargins(0, 0, 0, 0)
		# btn_layout.addStretch(1)
		confirm_btn = QPushButton("确认")
		confirm_btn.setFixedSize(100, 30)
		confirm_btn.setStyleSheet("""
			QPushButton {
				background-color: #00aaff;
				border-radius:5px;
				border-color:white;
				font-size: 12px;
				font-weight: bold;
				color: black;
				text-align: center;
			}
		""")
		confirm_btn.clicked.connect(self.emitModify)
		btn_layout.addWidget(confirm_btn)

		cancel_btn = QPushButton()
		cancel_btn.setFixedSize(100, 30)
		cancel_btn.setText("关闭窗口")
		cancel_btn.setStyleSheet("""
			QPushButton {
				background-color: #ff0000;
				border-radius:5px;
				border-color:white;
				font-size: 12px;
				font-weight: bold;
				color: black;
				text-align: center;
			}
		""")
		cancel_btn.pressed.connect(self.close)
		btn_layout.addWidget(cancel_btn)
		btn_layout.setSpacing(15)

		self.main_layout.addLayout(content_layout)
		self.main_layout.addLayout(btn_layout)
		self.main_layout.addStretch(1)
		self.main_layout.setSpacing(10)
		self.setLayout(self.main_layout)

	def closeEvent(self, event):
		self.content_edit.setPlainText("")
		self.closeSignal.emit()
		event.accept()

	def displayBlock(self, viewText=None, content=None, width=None, height=None):
		if viewText:
			self.h_line = height
			self.w_line = width
			self.text_label.setText(f"宽度:{width}px, 高度:{height} 内容: {viewText}")
		if content is not None:
			self.display_label.setPlainText(content)
		self.content_edit.setFocus()

	def emitModify(self):
		self.modifySignal.emit()
		self.get_view_text(self.content_edit.toPlainText())
		self.text_label.setText(f"宽度:{self.w_line}px, 高度:{self.h_line}px 内容: {self.viewText}")
		self.display_label.setPlainText(self.content_edit.toPlainText())
		self.close()

	def getContent(self):
		return self.content_edit.toPlainText()

	def get_view_text(self, text):
		print("enter setViewText")
		if len(text) <= 3:
			self.viewText = text
		# 160->8个汉字 16个英文字母
		else:
			width = self.w_line
			res = LinesUtils.count_chinese_and_non_chinese_regex(text)
			max_len = int(width / 160.0 * 8)
			if res > max_len:
				self.viewText = text[:max_len - 3]
				if len(self.viewText) == 0:
					self.viewText = self.viewText + "..."
				else:
					self.viewText = text[0] + "..."
			else:
				self.viewText = text

		print("end setViewText")
