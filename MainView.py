import os

from PyQt6.QtGui import QColor, QIcon, QImage, QPainter, QPen, QPixmap
from PyQt6.QtWidgets import QCheckBox, QComboBox, QFileDialog, QGraphicsItem, QGraphicsLineItem, QGraphicsRectItem, \
	QLabel, QLineEdit, QMessageBox, QPushButton, QSizePolicy, QSpacerItem, QSpinBox, QStatusBar, QTextEdit
from PyQt6.QtCore import QDir, QLineF, QPoint, QRect, QRectF, Qt

import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QGraphicsView, QGraphicsScene,
                             QHBoxLayout, QVBoxLayout, QWidget)
from matplotlib.hatch import SouthEastHatch

from Block import Block
from project.BackgroundScene import BackgroundScene
from project.LinesUtil import LinesUtil
from project.MovableLineItem import MovableLineItem
from project.MovableRectItem import MovableRectItem
from project.EditRectWindow import EditRectWindow
from project.RectDetection import RectDetection
from project.CustomView import VisualizedView
from project.table_convertor import Convertor


class TableDetectionWindow(QMainWindow):
	def __init__(self):
		super().__init__()
		self.btn_height_minus = None
		self.btn_height_plus = None
		self.btn_width_minus = None
		self.btn_width_plus = None
		self.toggle_box = None
		self.combox = None
		self.view = VisualizedView()
		self.table = Block()
		self.detector = RectDetection()
		self.edit_window = None
		self.status_label = None
		self.widgetLayout = None
		self.background_pixmap = None
		self.bottom_widget = None
		self.right_widget = None
		self.pen = None
		self.scene = None
		self.label = None
		self.rect_isvisible = None
		self.edit_window_show = False
		self.selected_item = None

		self.view_rect_btn = None
		self.line_buttons = []
		self.rect_buttons = []
		self.sys_buttons = []
		self.script_dir = os.path.dirname(os.path.abspath(__file__))

	def initUI(self):
		self.setWindowTitle("表格结构检测")
		self.setGeometry(50, 50, 1300, 800)
		# 创建主窗口的中央部件和布局
		central_widget = QWidget()

		self.setCentralWidget(central_widget)
		self.status_bar = self.statusBar()
		self.status_bar.showMessage("开始界面")
		main_hbox = QHBoxLayout(central_widget)

		# 左边部分 - QGraphicsView 和 QGraphicsScene
		self.initForwardScene()
		self.view.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
		main_hbox.addWidget(self.view, 5)  # 占据3份空间

		# 右边部分 - 垂直布局
		right_vbox = QVBoxLayout()
		main_hbox.addLayout(right_vbox, 2)  # 占据2份空间
		right_vbox.setContentsMargins(0, 0, 0, 0)
		# 右边控件
		self.right_widget = QWidget()
		# 设置最小高度为200px
		right_vbox.addWidget(self.right_widget)

		self.right_widget.setObjectName("RightWidget")  # 设置唯一标识

		self.initToolWidget()
		self.right_widget.setLayout(self.widgetLayout)
		self.right_widget.setStyleSheet("""
			QWidget#RightWidget {
                border:1px solid black;
            }
		    """)

	def initForwardScene(self):
		self.init_scene = QGraphicsScene()
		file_btn = QPushButton("上传文件")
		file_btn.resize(300, 100)
		file_btn.setStyleSheet("""
			QPushButton {
				background-color: #54a9fd;
				border-radius: 10px;
				font-size: 20px;
				font-weight: bold;
				text-align: center;
				color: white;
			}
			QPushButton:hover {
				background-color: #3a95e6;
			}
		""")
		pixmap = QPixmap("imgs/fileUpload.png")
		file_btn.setIcon(QIcon(pixmap))
		file_btn.clicked.connect(self.fileUpload)
		proxy_file_btn = self.init_scene.addWidget(file_btn)
		proxy_file_btn.setPos(100, 100)
		self.view.setScene(self.init_scene)
		self.view.setRenderHints(
			QPainter.RenderHint.Antialiasing |
			QPainter.RenderHint.SmoothPixmapTransform
		)

	def initToolWidget(self):
		print("enter init tool widget")
		self.widgetLayout = QVBoxLayout()
		self.widgetLayout.setContentsMargins(5, 0, 5, 0)
		tool_layout = QHBoxLayout()
		tool_layout.setContentsMargins(0, 0, 0, 0)
		tool_layout.addStretch(1)
		tool_label = QLabel("工具栏")
		tool_label.setFixedSize(380, 50)
		tool_label.setStyleSheet("""
			QLabel {
				border:1px red solid;    
				font-size:23px;
				font-weight: bold;
				color: black;
			}
		""")
		tool_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
		tool_layout.addWidget(tool_label)
		tool_layout.addStretch(1)

		self.widgetLayout.addLayout(tool_layout)
		self.init_line_toolbar()
		self.init_rect_toolbar()
		self.init_save_convertor_toolbar()
		self.init_system_tool()

		for button in self.rect_buttons:
			button.setEnabled(False)
		for button in self.line_buttons:
			button.setEnabled(False)
		for button in self.sys_buttons:
			button.setEnabled(False)

		self.widgetLayout.addStretch(1)
		print("end init tool widget")

	def init_rect_toolbar(self):
		main_layout = QVBoxLayout()

		title_layout = QHBoxLayout()
		title_label = QLabel()
		title_label.setText("单元格工具栏")
		title_label.setFixedSize(105, 25)
		title_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
		title_label.setStyleSheet("""
				QLabel {
					font-size: 16px;
					font-weight: bold;
					border:none;
				}
			""")
		title_layout.addWidget(title_label)
		title_layout.addStretch(1)

		main_layout.addLayout(title_layout)

		btn_layout = QHBoxLayout()

		addBtn = QPushButton("增加矩形")
		addBtn.setStyleSheet("border: 1px solid white;")
		pixmap = QPixmap("imgs/rect.png")
		addBtn.setFixedSize(80, 25)
		addBtn.setIcon(QIcon(pixmap))
		addBtn.clicked.connect(lambda: self.setMode(pre_mode=self.scene.current_mode, mode='box'))
		addBtn.setStyleSheet("""
							QPushButton {
								background-color: #55aaff;
								border-radius:5px;
								border-color:white;
								font-size: 12px;
								font-weight: bold;
								color: black;
								text-align: center;
							}
							QPushButton:hover {
								background-color: #00aaff;
							}
						""")

		stop_button = QPushButton("编辑文字")
		stop_button.setFixedSize(80, 25)
		stop_button.setStyleSheet("""
					QPushButton {
						background-color: #55aaff;
						border-radius:5px;
						border-color:white;
						font-size: 12px;
						font-weight: bold;
						color: black;
						text-align: center;
					}
					QPushButton:hover {
						background-color: #00aaff;
					}
				""")
		stop_button.clicked.connect(lambda: self.setMode(pre_mode=self.scene.current_mode, mode='view'))

		self.view_rect_btn = QPushButton("显示文字方框")
		self.view_rect_btn.setFixedSize(80, 25)
		self.view_rect_btn.setStyleSheet("""
			QPushButton {
						background-color: #55aaff;
						border-radius:5px;
						border-color:white;
						font-size: 12px;
						font-weight: bold;
						color: black;
						text-align: center;
					}
					QPushButton:hover {
						background-color: #00aaff;
					}
		""")
		self.view_rect_btn.clicked.connect(self.displayRect)

		standardize_btn = QPushButton()
		standardize_btn.setFixedSize(80, 25)
		standardize_btn.setText("标准化表格")
		standardize_btn.setStyleSheet("""
				QPushButton {
						background-color: #55aaff;
						border-radius:5px;
						border-color:white;
						font-size: 12px;
						font-weight: bold;
						color: black;
						text-align: center;
					}
					QPushButton:hover {
						background-color: #00aaff;
					}
			""")
		standardize_btn.clicked.connect(self.standardizeLengthTable)

		btn_layout.addWidget(addBtn)
		btn_layout.addWidget(stop_button)
		btn_layout.addWidget(self.view_rect_btn)
		btn_layout.addWidget(standardize_btn)
		btn_layout.setSpacing(8)

		main_layout.addLayout(btn_layout)

		self.label = QTextEdit("单元格信息框")
		self.label.setReadOnly(True)
		self.label.setFixedSize(380, 90)
		self.label.setStyleSheet("border-color:black")
		main_layout.addWidget(self.label)

		size_layout = QHBoxLayout()

		width_label = QLabel("宽度")
		width_label.setStyleSheet("border:none")
		width_plus_btn = QPushButton()
		width_plus_btn.setIcon(QIcon(QPixmap("./imgs/add.png")))
		width_plus_btn.setStyleSheet("border:none")
		width_plus_btn.clicked.connect(lambda: self.adjust_rect_size_step(dimension="width", operation="+"))

		self.width_number = QLineEdit()
		self.width_number.setStyleSheet("border:none")
		self.width_number.setFixedSize(35, 20)

		width_minus_btn = QPushButton()
		width_minus_btn.setIcon(QIcon(QPixmap("./imgs/sub.png")))
		width_minus_btn.setStyleSheet("border:none")
		width_minus_btn.clicked.connect(lambda: self.adjust_rect_size_step(dimension="width", operation="-"))

		height_label = QLabel("高度")
		height_label.setStyleSheet("border:none")

		height_plus_btn = QPushButton()
		height_plus_btn.setIcon(QIcon(QPixmap("./imgs/add.png")))
		height_plus_btn.setStyleSheet("border:none")
		height_plus_btn.clicked.connect(lambda: self.adjust_rect_size_step(dimension="height", operation="+"))

		self.height_number = QLineEdit()
		self.height_number.setStyleSheet("border:none")
		self.height_number.setFixedSize(35, 20)

		height_minus_btn = QPushButton()
		height_minus_btn.setIcon(QIcon(QPixmap("./imgs/sub.png")))
		height_minus_btn.setStyleSheet("border:none")
		height_minus_btn.clicked.connect(lambda: self.adjust_rect_size_step(dimension="height", operation="-"))

		adjustBtn = QPushButton("调整")
		adjustBtn.setFixedSize(60, 25)
		adjustBtn.setStyleSheet("""
			QPushButton {
				background-color: #54a9fd;
				border-radius:8px;
				font-size: 13px;
				font-weight: bold;
				text-align: center;
				color: black;
			}
			QPushButton:hover {
				background-color: #3a95e6;
			}
		""")
		adjustBtn.clicked.connect(self.adjustRectSize)

		step_label = QLabel("步长")
		step_label.setStyleSheet("border:none")
		self.rect_step_box = QSpinBox()
		self.rect_step_box.setStyleSheet("border-color:black")
		self.rect_step_box.setMinimum(1)
		self.rect_step_box.setMaximum(20)
		self.rect_step_box.setValue(3)
		self.rect_step_box.setFixedHeight(25)

		size_layout.addWidget(width_label)
		size_layout.addWidget(width_plus_btn)
		size_layout.addWidget(self.width_number)
		size_layout.addWidget(width_minus_btn)

		size_layout.addWidget(height_label)
		size_layout.addWidget(height_plus_btn)
		size_layout.addWidget(self.height_number)
		size_layout.addWidget(height_minus_btn)

		size_layout.addWidget(step_label)
		size_layout.addWidget(self.rect_step_box)
		size_layout.addWidget(adjustBtn)

		main_layout.addLayout(size_layout)

		self.rect_buttons.append(addBtn)
		self.rect_buttons.append(stop_button)
		self.rect_buttons.append(self.view_rect_btn)
		self.rect_buttons.append(standardize_btn)

		self.widgetLayout.addLayout(main_layout)

	def init_line_toolbar(self):
		main_layout = QVBoxLayout()

		title_layout = QHBoxLayout()
		title_label = QLabel()
		title_label.setText("线条工具栏")
		title_label.setFixedSize(95, 25)
		title_label.setStyleSheet("""
			QLabel {
				font-size: 16px;
				font-weight: bold;
				border:none;
			}
		""")

		title_label.setAlignment(Qt.AlignmentFlag.AlignLeft)

		title_layout.addWidget(title_label)
		title_layout.addStretch(1)

		main_layout.addLayout(title_layout)

		tool_layout = QHBoxLayout()

		drawBtn = QPushButton("绘制线条")
		pixmap = QPixmap("imgs/pen.png")
		drawBtn.setFixedSize(80, 25)
		drawBtn.setIcon(QIcon(pixmap))
		drawBtn.setStyleSheet("""
					QPushButton {
						background-color: #55aaff;
						border-radius:5px;
						border-color:white;
						font-size: 12px;
						font-weight: bold;
						color: black;
						text-align: center;
					}
					QPushButton:hover {
						background-color: #00aaff;
					}
				""")
		drawBtn.clicked.connect(lambda: self.setMode(pre_mode=self.scene.current_mode, mode='line'))

		special_btn = QPushButton()
		special_btn.setText("线条特殊")
		special_btn.setFixedSize(80, 25)
		special_btn.setStyleSheet("""
			QPushButton {
				background-color: #55aaff;
				border-radius:5px;
				border-color:white;
				font-size: 12px;
				font-weight: bold;
				color: black;
				text-align: center;
			}
			QPushButton:hover {
				background-color: #00aaff;
			}
		""")
		special_btn.clicked.connect(self.set_line_special)

		tool_layout.addWidget(drawBtn)
		tool_layout.addWidget(special_btn)

		step_layout = QHBoxLayout()
		length_label = QLabel("长度调整")
		length_label.setStyleSheet("border:none")
		length_label.setFixedSize(55, 20)

		length_plus_btn = QPushButton()
		length_plus_btn.setIcon(QIcon("./imgs/add.png"))
		length_plus_btn.setStyleSheet("border:none")
		length_plus_btn.setFixedSize(20, 20)
		length_plus_btn.clicked.connect(lambda: self.adjust_special_length_step(operation="+"))

		self.length_number = QLabel()
		self.length_number.setStyleSheet("""
			QLabel {
			background-color: white;
			border-color:black;
			}
		""")
		self.length_number.setFixedSize(30, 20)

		length_minus_btn = QPushButton()
		length_minus_btn.setFixedSize(20, 20)
		length_minus_btn.setStyleSheet("border:none")
		length_minus_btn.setIcon(QIcon("./imgs/sub.png"))
		length_minus_btn.clicked.connect(lambda: self.adjust_special_length_step(operation="-"))

		step_label = QLabel("步长")
		step_label.setStyleSheet("border:none")

		self.line_step_box = QSpinBox()
		self.line_step_box.setStyleSheet("border-color:black;")
		self.line_step_box.setMinimum(1)
		self.line_step_box.setMaximum(20)
		self.line_step_box.setValue(3)
		self.line_step_box.setFixedHeight(25)

		step_layout.addWidget(length_label)
		step_layout.addWidget(length_plus_btn)
		step_layout.addWidget(self.length_number)
		step_layout.addWidget(length_minus_btn)
		step_layout.addWidget(step_label)
		step_layout.addWidget(self.line_step_box)

		step_layout.setSpacing(4)

		tool_layout.addLayout(step_layout)
		tool_layout.addStretch(1)
		tool_layout.setSpacing(5)

		main_layout.addLayout(tool_layout)

		left_layout = QVBoxLayout()
		left_widget = QWidget()

		vertical_layout = QHBoxLayout()

		vertical_label = QLabel("竖线长度调整")
		vertical_label.setContentsMargins(0, 0, 0, 0)
		vertical_label.setFixedSize(80, 25)
		vertical_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
		vertical_label.setStyleSheet("border:none")

		self.vertical_edit = QLineEdit()
		self.vertical_edit.setStyleSheet("border-color:black")
		self.vertical_edit.setPlaceholderText("正减负增")
		self.vertical_edit.setFixedSize(80, 25)

		vertical_layout.addWidget(vertical_label)
		vertical_layout.addWidget(self.vertical_edit)
		vertical_layout.setSpacing(5)

		left_layout.addLayout(vertical_layout)

		btn_layout = QHBoxLayout()
		btn_layout.addStretch(1)
		# # 确认按钮
		vertical_btn = QPushButton()
		vertical_btn.setFixedSize(80, 25)
		vertical_btn.setText("确认")
		vertical_btn.setStyleSheet("""
			QPushButton {
				background-color: #55aaff;
				border-radius:5px;
				border-color:white;
				font-size: 12px;
				font-weight: bold;
				color: black;
				text-align: center;
			}
			QPushButton:hover {
				background-color: #00aaff;
			}
		""")
		vertical_btn.clicked.connect(self.adjustVerticalLines)
		btn_layout.addWidget(vertical_btn)
		btn_layout.addStretch(1)

		left_layout.addLayout(btn_layout)
		left_widget.setLayout(left_layout)

		right_layout = QVBoxLayout()
		right_widget = QWidget()

		# 输入框
		horizontal_layout = QHBoxLayout()

		horizontal_label = QLabel("横线长度调整")
		horizontal_label.setContentsMargins(0, 0, 0, 0)
		horizontal_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
		horizontal_label.setStyleSheet("border:none")
		horizontal_label.setFixedSize(80, 30)

		self.horizontal_edit = QLineEdit()
		self.horizontal_edit.setStyleSheet("border-color:black")
		self.horizontal_edit.setPlaceholderText("正减负增")
		self.horizontal_edit.setFixedSize(80, 25)

		horizontal_layout.addWidget(horizontal_label)
		horizontal_layout.addWidget(self.horizontal_edit)
		horizontal_layout.setSpacing(5)
		right_layout.addLayout(horizontal_layout)

		btn_layout = QHBoxLayout()
		btn_layout.addStretch(1)
		# 确认按钮
		horizontal_btn = QPushButton()
		horizontal_btn.setFixedSize(80, 25)
		horizontal_btn.setText("确认")
		horizontal_btn.setStyleSheet("""
			QPushButton {
				background-color: #55aaff;
				border-radius:5px;
				border-color:white;
				font-size: 12px;
				font-weight: bold;
				color: black;
				text-align: center;
			}
			QPushButton:hover {
				background-color: #00aaff;
			}
		""")
		horizontal_btn.clicked.connect(self.adjustHorizontalLines)

		btn_layout.addWidget(horizontal_btn)
		btn_layout.addStretch(1)

		right_layout.addLayout(btn_layout)
		right_widget.setLayout(right_layout)

		bottom_layout = QHBoxLayout()
		bottom_layout.addWidget(left_widget)
		bottom_layout.addWidget(right_widget)

		main_layout.addLayout(bottom_layout)

		self.line_buttons.append(drawBtn)
		self.line_buttons.append(special_btn)
		self.line_buttons.append(vertical_btn)
		self.line_buttons.append(horizontal_btn)

		self.widgetLayout.addLayout(main_layout)

	def init_save_convertor_toolbar(self):
		main_layout = QVBoxLayout()

		title_layout = QHBoxLayout()

		title_label = QLabel("数据保存与表格转换")
		title_label.setFixedSize(150, 25)
		title_label.setStyleSheet("""
				QLabel {
					font-size: 16px;
					font-weight: bold;
					border:none;
				}
			""")
		title_layout.addWidget(title_label)
		title_layout.addStretch(1)

		main_layout.addLayout(title_layout)

		file_name_layout = QHBoxLayout()
		title_label = QLabel("输入保存文件名")
		title_label.setStyleSheet("border:none")

		self.file_name = QLineEdit()
		self.file_name.setFixedSize(190, 25)
		self.file_name.setPlaceholderText("文件名不能含有空格")

		file_name_layout.addWidget(title_label)
		file_name_layout.addWidget(self.file_name)
		file_name_layout.addStretch(1)
		file_name_layout.setSpacing(8)

		main_layout.addLayout(file_name_layout)

		save_layout = QHBoxLayout()

		save_rect_btn = QPushButton("保存")
		save_rect_btn.setIcon(QIcon("imgs/save.png"))
		save_rect_btn.setStyleSheet("""
			QPushButton {
						background-color:#43cd7f;
						border-radius:5px;
						border-color:white;
						font-size: 12px;
						font-weight: bold;
						color: black;
						text-align: center;
					}
			QPushButton:hover {
				background-color: #3ab46f;
			}
		""")
		save_rect_btn.clicked.connect(self.saveAsFile)
		save_rect_btn.setFixedSize(80, 25)

		self.combox = QComboBox()
		self.combox.setStyleSheet("border-color:black")
		self.combox.setPlaceholderText("选择保存文件格式")
		self.combox.addItem("word")
		self.combox.addItem("pdf")
		self.combox.addItem("html")
		self.combox.setCurrentIndex(-1)
		self.combox.setFixedSize(190, 25)

		save_layout.addWidget(save_rect_btn)
		save_layout.addWidget(self.combox)
		save_layout.setSpacing(10)
		save_layout.addStretch(1)

		main_layout.addLayout(save_layout)

		file_path_layout = QHBoxLayout()
		title_label = QLabel("输出路径")
		title_label.setFixedSize(80, 25)
		title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
		title_label.setStyleSheet("border:none")
		self.path_label = QLabel()
		self.path_label.setFixedSize(240, 25)
		self.path_label.setStyleSheet("""
			QLabel {
				background-color:white;
				font-size: 12px;
			}
		""")
		file_path_layout.addWidget(title_label)
		file_path_layout.addWidget(self.path_label)
		file_path_layout.setSpacing(8)
		file_path_layout.addStretch(1)

		main_layout.addLayout(file_path_layout)

		btn_layout = QHBoxLayout()
		word_btn = QPushButton()
		word_btn.setFixedSize(80, 30)
		word_btn.setText("转为word表格")
		word_btn.setStyleSheet("""
			QPushButton {
				background-color: #f8a500;
				border-radius:5px;
				border-color:white;
				font-size: 12px;
				font-weight: bold;
				color: black;
				text-align: center;
			}
			QPushButton:hover {
				background-color: #e99b00;
			}
		""")
		word_btn.clicked.connect(self.con_to_word)

		pdf_btn = QPushButton()
		pdf_btn.setFixedSize(80, 30)
		pdf_btn.setText("转为pdf表格")
		pdf_btn.setStyleSheet("""
			QPushButton {
				background-color: #f8a500;
				border-radius:5px;
				border-color:white;
				font-size: 12px;
				font-weight: bold;
				color: black;
				text-align: center;
			}
			QPushButton:hover {
				background-color: #e99b00;
			}
		""")
		pdf_btn.clicked.connect(self.con_to_pdf)

		html_btn = QPushButton()
		html_btn.setFixedSize(80, 30)
		html_btn.setText("转为html表格")
		html_btn.setStyleSheet("""
			QPushButton {
				background-color: #f8a500;
				border-radius:5px;
				border-color:white;
				font-size: 12px;
				font-weight: bold;
				color: black;
				text-align: center;
			}
			QPushButton:hover {
				background-color: #e99b00;
			}
		""")
		html_btn.clicked.connect(self.con_to_html)

		btn_layout.addWidget(word_btn)
		btn_layout.addWidget(pdf_btn)
		btn_layout.addWidget(html_btn)
		btn_layout.setSpacing(10)
		btn_layout.addStretch(1)

		main_layout.addLayout(btn_layout)

		self.sys_buttons.append(save_rect_btn)
		self.widgetLayout.addLayout(main_layout)

	def init_system_tool(self):
		main_layout = QVBoxLayout()

		title_layout = QHBoxLayout()

		title_label = QLabel()
		title_label.setFixedSize(100, 20)
		title_label.setText("系统工具")
		title_label.setStyleSheet("""
			QLabel {
			font-size: 16px;
			font-weight: bold;
			border:none;
			}
		""")

		title_layout.addWidget(title_label)
		title_layout.addStretch(1)
		main_layout.addLayout(title_layout)

		btn_layout = QHBoxLayout()

		delete_btn = QPushButton("删除")
		delete_btn.setFixedSize(80, 25)
		pixmap = QPixmap("imgs/delete.png")
		delete_btn.setIcon(QIcon(pixmap))
		delete_btn.setStyleSheet("""
			QPushButton {
				background-color: #fc0000;
				border-radius:5px;
				border-color:white;
				font-size: 12px;
				font-weight: bold;
				color: black;
				text-align: center;
			}
			QPushButton:hover {
				background-color: #ed0000;
			}
		""")
		delete_btn.clicked.connect(self.deleteSceneItems)
		delete_btn.setShortcut("Ctrl+D")

		transfer_btn = QPushButton()
		transfer_btn.setFixedSize(80, 25)
		transfer_btn.setText("图片预处理")
		transfer_btn.setStyleSheet("""
			QPushButton {
					background-color: #55aaff;
					border-radius:5px;
					border-color:white;
					font-size: 12px;
					font-weight: bold;
					color: black;
					text-align: center;
				}
			QPushButton:hover {
				background-color: #00aaff;
			}
		""")
		transfer_btn.clicked.connect(self.imageTrans)

		exit_btn = QPushButton()
		exit_btn.setFixedSize(80, 25)
		exit_btn.setText("结束程序")
		exit_btn.setStyleSheet("""
			QPushButton {
				background-color: #aaaafe;
				border-radius:5px;
				border-color:white;
				font-size: 12px;
				font-weight: bold;
				color: black;
				text-align: center;
			}
			QPushButton:hover {
				background-color: #8787c9;
			}
		""")
		exit_btn.clicked.connect(self.pre_close)

		init_btn = QPushButton()
		init_btn.setFixedSize(80, 25)
		init_btn.setText("回到开始")
		init_btn.setStyleSheet("""
			QPushButton {
				background-color: #aaaafe;
				border-radius:5px;
				border-color:white;
				font-size: 12px;
				font-weight: bold;
				color: black;
				text-align: center;
			}
			QPushButton:hover {
				background-color: #8787c9;
			}
		""")
		init_btn.clicked.connect(self.back_forward_scene)

		btn_layout.addWidget(delete_btn)
		btn_layout.addWidget(transfer_btn)
		btn_layout.addWidget(init_btn)
		btn_layout.addWidget(exit_btn)
		btn_layout.addStretch(1)
		btn_layout.setSpacing(10)

		main_layout.addLayout(btn_layout)

		convertor_layout = QVBoxLayout()
		convertor_label = QLabel("文件保存相关设置")
		convertor_label.setFixedSize(100, 20)
		convertor_label.setStyleSheet("""
			QLabel {
			font-size: 12px;
			font-weight: bold;
			border:none;}
		""")
		convertor_layout.addWidget(convertor_label)

		self.toggle_box = QCheckBox()
		self.toggle_box.setFixedSize(190, 25)
		self.toggle_box.setText("使用Word接口转换docx至pdf")
		self.toggle_box.setStyleSheet("""
				    QCheckBox::indicator {
				        width: 16px;
				        height: 16px;
				    }
				    QCheckBox::indicator:unchecked {
				        image: url("imgs/unchecked.png");
				    }
				    QCheckBox::indicator:checked {
				        image: url("imgs/checked.png");
				    }
				    QCheckBox{
				        border:none;
				    }
				""")

		self.pdf_table_box = QCheckBox()
		self.pdf_table_box.setChecked(True)
		self.pdf_table_box.setFixedSize(190, 25)
		self.pdf_table_box.setText("保存时使用标准化pdf表格")
		self.pdf_table_box.setStyleSheet("""
						    QCheckBox::indicator {
						        width: 16px;
						        height: 16px;
						    }
						    QCheckBox::indicator:unchecked {
						        image: url("imgs/unchecked.png");
						    }
						    QCheckBox::indicator:checked {
						        image: url("imgs/checked.png");
						    }
						    QCheckBox{
				                border:none;
				            }
						""")

		self.html_title_box = QCheckBox()
		self.html_title_box.setFixedSize(210, 25)
		self.html_title_box.setText("保存为html表格时首行作为标题")
		self.html_title_box.setStyleSheet("""
					    QCheckBox::indicator {
					        width: 16px;
					        height: 16px;
					    }
					    QCheckBox::indicator:unchecked {
					        image: url("imgs/unchecked.png");
					    }
					    QCheckBox::indicator:checked {
					        image: url("imgs/checked.png");
					    }
					    QCheckBox{
					        border:none;
					    }
					""")

		convertor_layout.addWidget(self.toggle_box)
		convertor_layout.addWidget(self.pdf_table_box)
		convertor_layout.addWidget(self.html_title_box)
		convertor_layout.setSpacing(5)
		main_layout.addLayout(convertor_layout)

		self.sys_buttons.append(delete_btn)
		self.widgetLayout.addLayout(main_layout)

	def initEditScene(self):
		if self.background_pixmap.isNull():
			print("NULL")

		self.scene = BackgroundScene()
		self.scene.setSceneRect(0, 0, 900, 740)
		self.pen = QPen(QColor(0, 0, 255), 2)  # 蓝色，2像素宽

		scene = self.scene.sceneRect()
		self.scene.scene_border = QGraphicsRectItem(scene)
		self.scene.scene_border.setPen(QPen(QColor(255, 0, 0), 2))
		self.scene.scene_border.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False)
		self.scene.scene_border.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, False)
		self.scene.addItem(self.scene.scene_border)

		self.scene.setBackgroundPixmap(self.background_pixmap)

		for button in self.rect_buttons:
			button.setEnabled(True)
		for button in self.line_buttons:
			button.setEnabled(True)
		for button in self.sys_buttons:
			button.setEnabled(True)

		self.status_bar.showMessage("表格检测")
		self.view.setScene(self.scene)
		self.view.setRenderHints(
			QPainter.RenderHint.Antialiasing |
			QPainter.RenderHint.SmoothPixmapTransform
		)
		self.view.end_draw_rect.connect(lambda: self.setMode(pre_mode="box", mode="view"))
		self.initTableStructure()
		self.initRect()
		result_text = "\n".join(["; ".join(sublist) for sublist in self.table.block_texts])
		self.label.setPlainText(result_text)

	def initRect(self):
		if self.table.rect_points is not None:
			print("开始进行点的计算")
			for point, text, block in zip(self.table.rect_points, self.table.text, self.table.data):
				rect = QRect(QPoint(point[0][0], point[0][1]), QPoint(point[2][0], point[2][1]))
				rect_item = MovableRectItem(text=text, rect=rect)
				if self.scene.start_point_x is not None:
					x = self.scene.start_point_x + block[0][0][0]
					y = self.scene.start_point_y + block[0][0][1]
					rect_item.setPos(x, y)
					rect_item.setPen(QPen(Qt.GlobalColor.green, 2))
					rect_item.adjust_text_position()
					rect_item.setVisible(False)

					self.scene.addItem(rect_item)
					self.scene.rect_items.append(rect_item)

		self.scene.sceneClickedSignal.connect(self.updateLabel)
		self.scene.endLineSignal.connect(lambda: self.setMode(pre_mode="line", mode="view"))
		self.scene.editSignal.connect(self.openEditWindow)
		self.setMode(pre_mode="view", mode="view")
		print("end initRect")

	def initTableStructure(self):
		print("enter initTableStructure")
		rows = LinesUtil.identify_rows(self.detector.table_info, 25)
		columns = LinesUtil.identify_columns(self.detector.table_info, 25)
		row_lines = LinesUtil.drawRowLines(rows)
		col_lines = LinesUtil.drawColLines(columns)
		base_x = self.scene.start_point_x
		base_y = self.scene.start_point_y
		# 绘制垂直线
		for vl in col_lines:
			x = base_x + vl[0]
			y_start = base_y
			y_end = base_y + vl[1]
			line = MovableLineItem(x, y_start, x, y_end)
			line.setPen(QPen(Qt.GlobalColor.red, 2))
			line.type = "type_y"
			line.special = False
			line.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, True)
			line.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
			line.setFlag(QGraphicsItem.GraphicsItemFlag.ItemSendsScenePositionChanges, True)

			self.scene.addItem(line)  # 红色
			self.scene.col_lines.append(line)

		# 绘制水平线
		for hl in row_lines:
			y = base_y + hl[0]
			x_start = base_x
			x_end = base_x + hl[1]

			line = MovableLineItem(x_start, y, x_end, y)
			line.type = "type_x"
			line.special = False
			line.setPen(QPen(Qt.GlobalColor.red, 2))
			line.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, True)
			line.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
			line.setFlag(QGraphicsItem.GraphicsItemFlag.ItemSendsScenePositionChanges, True)

			self.scene.addItem(line)  # 红色
			self.scene.row_lines.append(line)

		print("end initTableStructure")

	def standardizeLengthTable(self):
		print("enter standardizeTable")
		self.scene.col_lines = LinesUtil.sortLines(self.scene.col_lines, 1)
		self.scene.row_lines = LinesUtil.sortLines(self.scene.row_lines, 2)

		origin_x1 = self.scene.col_lines[0].line().x1()
		origin_x2 = self.scene.col_lines[-1].line().x1()
		delta_x1 = self.scene.col_lines[0].scenePos().x()
		delta_x2 = self.scene.col_lines[-1].scenePos().x()
		horizontal_line_length = origin_x2 - origin_x1 - delta_x1 + delta_x2
		pos_x = origin_x1 + delta_x1
		for line_item in self.scene.row_lines:
			if not line_item.special:
				line = line_item.line()
				line.setLength(horizontal_line_length)
				line_item.setLine(line)

		origin_y1 = self.scene.row_lines[0].line().y1()
		origin_y2 = self.scene.row_lines[-1].line().y1()
		delta_y1 = self.scene.row_lines[0].scenePos().y()
		delta_y2 = self.scene.row_lines[-1].scenePos().y()

		vertical_line_length = origin_y2 - origin_y1 - delta_y1 + delta_y2

		for line_item in self.scene.col_lines:
			if not line_item.special:
				line = line_item.line()
				line.setLength(vertical_line_length)
				line_item.setLine(line)

		pos_y = origin_y1 + delta_y1
		self.standardizePosTable(pos_x=pos_x, pos_y=pos_y)
		print("end standardizeTable")

	def standardizePosTable(self, pos_x, pos_y):
		print("enter standardizePosTable")
		for item in self.scene.row_lines:
			if not item.special:
				delta_x = pos_x - item.line().x1()
				if abs(delta_x) >= 0.1:
					delta_y = item.scenePos().y()
					item.setPos(delta_x, delta_y)

		for item in self.scene.col_lines:
			if not item.special:
				delta_y = pos_y - item.line().y1()
				if abs(delta_y) >= 0.1:
					delta_x = item.scenePos().x()
					item.setPos(delta_x, delta_y)
		self.standardize_special_lines_pos()
		print("end standardizePosTable")

	def setMode(self, pre_mode, mode):
		print("start mode")
		self.scene.current_mode = mode
		if mode == "line":
			self.status_bar.showMessage("绘制线条中...")
			list_items = self.scene.items()
			if list_items is not None:
				for item in list_items:
					if isinstance(item, MovableRectItem) or isinstance(item, MovableLineItem):
						item.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False)
		else:
			if mode == "box":
				self.status_bar.showMessage("绘制矩形框中...")
				self.scene.setLinesNotMov()
			else:
				if pre_mode == 'line':
					items = self.scene.items()
					if items is not None:
						for item in items:
							if isinstance(item, QGraphicsLineItem):
								if item.line().length() < 5.0:
									self.scene.removeItem(item)
				self.status_bar.showMessage("表格检测")
				list_items = self.scene.items()
				if list_items is not None:
					for item in list_items:
						if isinstance(item, MovableRectItem) or isinstance(item, MovableLineItem):
							item.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
		self.set_button_style(mode=mode)
		print("end mode")

	def fileUpload(self):
		print("start fileUpload")
		# 打开文件选择对话框，只允许选择图片
		file_dialog = QFileDialog()
		file_dialog.setNameFilter("图片文件 (*.png *.jpg *.jpeg)")  # 设置过滤器，只显示图片文件
		directory_path = "./trans"

		# 检查目录是否存在
		if QDir(directory_path).exists():
			file_dialog.setDirectory(directory_path)
		else:
			file_dialog.setDirectory(QDir.currentPath())

		if file_dialog.exec():
			selected_files = file_dialog.selectedFiles()
			if selected_files:
				file_path = selected_files[0]
				if file_path is not None:
					QMessageBox.information(self, "消息通知", f"图片上传成功!\n正在检测图片表格,图片路径:{file_path}")
					self.table.ocr(file_path)
					self.background_pixmap = QPixmap(file_path)
					# print(f"选择的文件路径: {file_path}")
					self.detector.detectRect(file_path)
					self.initEditScene()
				else:
					QMessageBox.warning(self, "消息提示", "未找到文件!")

	def updateLabel(self):
		list_items = self.scene.items()
		if list_items is not None:
			for item in list_items:
				if isinstance(item, MovableRectItem):
					if item.isSelected():
						self.selected_item = item
						text = item.text
						width = item.rect().width()
						height = item.rect().height()
						self.label.setText(f"宽度为: {width}px, 高度为: {height}px\n内容: \n{text}")
						self.width_number.setText(str(width))
						self.height_number.setText(str(height))
						pen = QPen(QColor(Qt.GlobalColor.red))
						pen.setWidth(2)
						item.setPen(pen)
						if self.edit_window_show:
							self.edit_window.displayBlock(content=text, width=width, height=height)
					else:
						if item.pen().color() != Qt.GlobalColor.green and item.pen().color() != Qt.GlobalColor.transparent:
							if self.edit_window_show or self.rect_isvisible is False:
								pen = QPen(QColor(Qt.GlobalColor.transparent))
							else:
								pen = QPen(QColor(Qt.GlobalColor.green))
							pen.setWidth(2)
							item.setPen(pen)
				elif isinstance(item, MovableLineItem):
					if item.isSelected():
						self.selected_item = item
						pen = QPen(QColor(Qt.GlobalColor.blue))
						pen.setWidth(2)
						item.setPen(pen)
						if item in self.scene.special_col_lines or item in self.scene.special_row_lines:
							length = item.line().length()
							self.length_number.setText(str(length))
					else:
						if not item.special:
							if item.pen().color() != Qt.GlobalColor.red:
								pen = QPen(QColor(Qt.GlobalColor.red))
								pen.setWidth(2)
								item.setPen(pen)
						else:
							if item.pen().color() != Qt.GlobalColor.black:
								pen = QPen(QColor(Qt.GlobalColor.black))
								pen.setWidth(2)
								item.setPen(pen)

		select_items = self.scene.selectedItems()
		if len(select_items) == 0:
			self.selected_item = None

	# self.scene.clearSelection()
	def set_button_style(self, mode):
		if mode == "line":
			self.line_buttons[0].setStyleSheet("""
						QPushButton {
							background-color:  #ff5500;
							border-radius:5px;
							border-color:white;
							font-size: 12px;
							font-weight: bold;
							color: black;
							text-align: center;
						}
					""")
			self.rect_buttons[0].setStyleSheet("""
						QPushButton {
								background-color: #55aaff;
								border-radius:5px;
								border-color:white;
								font-size: 12px;
								font-weight: bold;
								color: black;
								text-align: center;
							}
						QPushButton:hover {
							background-color: #ff5500;
						}
					""")
			self.rect_buttons[1].setStyleSheet("""
						QPushButton {
								background-color: #55aaff;
								border-radius:5px;
								border-color:white;
								font-size: 12px;
								font-weight: bold;
								color: black;
								text-align: center;
							}
						QPushButton:hover {
							background-color: #ff5500;
						}
					""")
		elif mode == "box":
			self.line_buttons[0].setStyleSheet("""
				QPushButton {
						background-color: #55aaff;
						border-radius:5px;
						border-color:white;
						font-size: 12px;
						font-weight: bold;
						color: black;
						text-align: center;
					}
				QPushButton:hover {
					background-color: #ff5500;
				}
			""")
			self.rect_buttons[0].setStyleSheet("""
				QPushButton {
					background-color:  #ff5500;
					border-radius:5px;
					border-color:white;
					font-size: 12px;
					font-weight: bold;
					color: black;
					text-align: center;
				}
			""")
			self.rect_buttons[1].setStyleSheet("""
				QPushButton {
						background-color: #55aaff;
						border-radius:5px;
						border-color:white;
						font-size: 12px;
						font-weight: bold;
						color: black;
						text-align: center;
					}
				QPushButton:hover {
					background-color: #ff5500;
				}
			""")
		else:
			self.line_buttons[0].setStyleSheet("""
							QPushButton {
									background-color: #55aaff;
									border-radius:5px;
									border-color:white;
									font-size: 12px;
									font-weight: bold;
									color: black;
									text-align: center;
								}
							QPushButton:hover {
								background-color: #ff5500;
							}
						""")
			self.rect_buttons[0].setStyleSheet("""
				QPushButton {
						background-color: #55aaff;
						border-radius:5px;
						border-color:white;
						font-size: 12px;
						font-weight: bold;
						color: black;
						text-align: center;
					}
				QPushButton:hover {
					background-color: #ff5500;
				}
			""")
			self.rect_buttons[1].setStyleSheet("""
				QPushButton {
					background-color:  #ff5500;
					border-radius:5px;
					border-color:white;
					font-size: 12px;
					font-weight: bold;
					color: black;
					text-align: center;
				}
			""")

	def saveAsFile(self):
		print("enter save")
		"""可以导出为pdf"""
		flag = False
		if self.file_name.text() != "":
			flag = True
		if flag and ' ' in self.file_name.text():
			QMessageBox.warning(self, "消息提示", "文件名不能包含空格!")
			return

		current_index = self.combox.currentIndex()
		# 导出为word
		if current_index == 0:
			if len(self.scene.special_col_lines) > 0 or len(self.scene.special_row_lines) > 0:
				QMessageBox.warning(self, "消息提示", "可能含有合并单元格，暂不支持合并单元格导出为word文件")
			else:
				res = False
				if flag:
					res = self.scene.exportWord(file_name=self.file_name.text() + ".docx")
				else:
					res = self.scene.exportWord()
				if res:
					QMessageBox.information(self, "消息提示", "保存为word文件成功!")
					self.path_label.setText(self.script_dir + "\\out_word\\" + self.file_name.text())
				else:
					QMessageBox.warning(self, "消息提示", "文件保存失败!")
		# 导出为pdf
		elif current_index == 1:
			res = False
			if flag:
				res = self.scene.exportPDF(standard=self.pdf_table_box.isChecked(),
				                           file_name=self.file_name.text() + ".pdf")
			else:
				res = self.scene.exportPDF(standard=self.pdf_table_box.isChecked())
			if res:
				QMessageBox.information(self, "消息提示", "保存为pdf文件成功!!")
				self.path_label.setText(self.script_dir + "\\out_pdf\\" + self.file_name.text())
			else:
				QMessageBox.warning(self, "消息提示", "文件保存失败!")
		elif current_index == 2:
			if len(self.scene.special_col_lines) > 0 or len(self.scene.special_row_lines) > 0:
				QMessageBox.warning(self, "消息提示", "可能含有合并单元格，暂不支持合并单元格导出为html文件")
			else:
				res = False
				if flag:
					res = self.scene.exportHtml5Table(special=self.html_title_box.isChecked(),
					                                  file_name=self.file_name.text() + ".html")
				else:
					res = self.scene.exportHtml5Table(special=self.html_title_box.isChecked())
				if res:
					QMessageBox.information(self, "消息提示", "保存为html文件成功!")
					self.path_label.setText(self.script_dir + "\\out_html\\" + self.file_name.text())
				else:
					QMessageBox.warning(self, "消息提示", "文件保存失败!")
		else:
			QMessageBox.warning(self, "消息提示", "请选择一种保存格式!")
		print("end save")

	def openEditWindow(self):
		print("enter openEditWindow")
		if isinstance(self.selected_item, MovableRectItem):
			self.updateLabel()
			if self.edit_window is None:
				self.edit_window = EditRectWindow()
				self.edit_window.initUI()
				self.edit_window.closeSignal.connect(self.updateEditWindowStatus)
				self.edit_window.modifySignal.connect(self.modifyBlock)
				self.edit_window_show = True

			if self.rect_isvisible:
				self.hideRectBorder()
				self.rect_isvisible = False

			if self.edit_window.isVisible() is False:
				self.edit_window_show = True
				self.changeLinesSelected(self.edit_window_show)
				text = self.selected_item.text
				width = self.selected_item.rect().width()
				height = self.selected_item.rect().height()
				viewText = self.selected_item.viewText
				self.edit_window.displayBlock(viewText=viewText, content=text, width=width, height=height)
				self.edit_window.show()
				self.edit_window.activateWindow()
		else:
			pass
		print("end openEditWindow")

	def updateEditWindowStatus(self):
		print("enter updateEditWindowStatus")
		self.edit_window_show = False
		self.changeLinesSelected(self.edit_window_show)
		self.updateLabel()
		self.scene.clearSelection()
		self.selected_item = None
		if not self.rect_isvisible:
			self.displayRectBorder()
			self.rect_isvisible = True

		print("end updateEditWindowStatus")

	def changeLinesSelected(self, state):
		if state is True:
			if self.scene.col_lines:
				for item in self.scene.col_lines:
					item.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, False)
			if self.scene.row_lines:
				for item in self.scene.row_lines:
					item.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, False)
			if self.scene.bias_lines:
				for item in self.scene.row_lines:
					item.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, False)
		elif state is False:
			if self.scene.col_lines:
				for item in self.scene.col_lines:
					item.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, True)
			if self.scene.row_lines:
				for item in self.scene.row_lines:
					item.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, True)
			if self.scene.bias_lines:
				for item in self.scene.row_lines:
					item.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, True)

	def deleteSceneItems(self):
		items = self.scene.selectedItems()
		if items:
			for item in items:
				if isinstance(item, MovableLineItem):
					if item.type == "type_y":
						if not item.special:
							self.scene.col_lines.remove(item)
						else:
							self.scene.special_col_lines.remove(item)
					elif item.type == "type_x":
						if not item.special:
							self.scene.row_lines.remove(item)
						else:
							self.scene.special_row_lines.remove(item)
					else:
						self.scene.bias_lines.remove(item)
				elif isinstance(item, MovableRectItem):
					self.scene.rect_items.remove(item)
				self.scene.removeItem(item)
		self.scene.clearSelection()

	def displayRect(self):
		print("enter displayRect")
		if self.rect_isvisible is None:
			self.scene.setLinesNotMov()
			self.view_rect_btn.setText("隐藏绿色方框")
			self.viewRect()
			self.rect_isvisible = True
		else:
			if self.rect_isvisible is True:
				self.hideRectBorder()
				self.rect_isvisible = False
				self.view_rect_btn.setText("显示绿色方框")

			else:
				self.displayRectBorder()
				self.rect_isvisible = True
				self.view_rect_btn.setText("隐藏绿色方框")
		print("end displayRect")

	def viewRect(self):
		items = self.scene.items()
		if items is not None:
			for item in items:
				if isinstance(item, MovableRectItem):
					item.setVisible(True)

	def adjustRectSize(self):
		""" 调整选中矩形的尺寸 """
		try:
			width = float(self.width_number.text())
			height = float(self.height_number.text())
		except ValueError:
			QMessageBox.warning(self, "输入错误", "请输入有效的数字")
			return
		if self.scene is not None:
			selected_rects = [
				item for item in self.scene.selectedItems()
				if isinstance(item, MovableRectItem)
			]

			if not selected_rects:
				QMessageBox.information(self, "提示", "请先选择要调整的矩形")
				return

			for rect in selected_rects:
				rect.reSize(width, height)

			self.scene.update()  # 刷新场景
		else:
			QMessageBox.warning(self,"操作提示","请先选择一张图片!")
			
	def adjust_rect_size_step(self, dimension, operation):
		print("enter adjust_rect_size_step")
		"""调整矩形尺寸"""
		if isinstance(self.selected_item, MovableRectItem):
			rect = self.selected_item.rect()
			current_value = 0
			if dimension == "width":
				current_value = rect.width()
			elif dimension == "height":
				current_value = rect.height()
			new_value = current_value
			step = self.rect_step_box.value()
			try:
				if dimension == "width":
					if operation == "+":
						new_value = current_value + step
					elif operation == "-":
						if current_value <= step:
							raise ValueError("尺寸不足无法减少")
						new_value = current_value - step
					self.selected_item.setRect(rect.x(), rect.y(), new_value, rect.height())
					self.width_number.setText(str(new_value))
				elif dimension == "height":
					if operation == "+":
						new_value = current_value + step
					elif operation == "-":
						if current_value <= step:
							raise ValueError("尺寸不足无法减少")
						new_value = current_value - step
					self.height_number.setText(str(new_value))
					self.selected_item.setRect(rect.x(), rect.y(), rect.width(), new_value)
				self.updateLabel()
			except ValueError as e:
				QMessageBox.warning(self, "操作错误", str(e))
				return
		else:
			pass
		print("end adjust_rect_size_step")

	# 更新矩形尺寸（保持左上角位置不变）

	def hideRectBorder(self):
		print("enter hideRectBorder")
		list_items = self.scene.items()
		if list_items is not None:
			if self.rect_isvisible:
				self.rect_isvisible = False
				for item in list_items:
					if isinstance(item, MovableRectItem):
						pen = QPen()
						if item.isSelected() is False:
							pen.setColor(QColor("transparent"))
						else:
							pen.setColor(QColor(Qt.GlobalColor.red))
						pen.setWidth(2)
						item.setPen(pen)
		print("end hideRectBorder")

	def displayRectBorder(self):
		print("enter displayRectBorder")
		list_items = self.scene.items()
		if list_items is not None:
			for item in list_items:
				if isinstance(item, MovableRectItem):
					pen = QPen(QColor(Qt.GlobalColor.green))
					pen.setWidth(2)
					item.setPen(pen)
		print("end displayRectBorder")

	def modifyBlock(self):
		print("enter modifyBlock")
		if self.selected_item is not None and isinstance(self.selected_item, MovableRectItem):
			self.selected_item.text = self.edit_window.getContent()
			self.selected_item.setViewText(self.selected_item.text)

			self.selected_item.adjust_text_position()
		print("end modifyBlock")

	def adjustVerticalLines(self):
		input_text = self.vertical_edit.text()
		try:
			input_length = float(input_text)
		except ValueError:
			QMessageBox.warning(self, "错误", "请输入有效数字")
			return
		self.reSizeLinesLength(input_length=input_length, items=self.scene.col_lines)
		self.vertical_edit.setText("")
		self.vertical_edit.setFocus()

	def adjustHorizontalLines(self):
		input_text = self.horizontal_edit.text()
		try:
			input_length = float(input_text)
		except ValueError:
			QMessageBox.warning(self, "错误", "请输入有效数字")
			return
		self.reSizeLinesLength(input_length=input_length, items=self.scene.row_lines)
		self.horizontal_edit.setText("")
		self.horizontal_edit.setFocus()

	def reSizeLinesLength(self, input_length, items):
		for line_item in items:
			line = line_item.line()
			length = line.length()
			line.setLength(length - input_length)
			line_item.setLine(line)

	def adjust_special_length_step(self, operation):
		if self.selected_item is None or isinstance(self.selected_item, MovableRectItem):
			return
		if (self.selected_item in self.scene.special_col_lines or
		    self.selected_item in self.scene.special_row_lines) or self.selected_item in self.scene.bias_lines:
			step = self.line_step_box.value()
			line = self.selected_item.line()
			current_value = line.length()
			if operation == "+":
				new_value = current_value + step
				line.setLength(new_value)
				self.length_number.setText(str(int(new_value)))
				self.selected_item.setLine(line)
			else:
				if current_value <= step:
					QMessageBox.warning(self, "消息提示", "长度不能再减少了!")
				else:
					new_value = current_value - step
					line.setLength(new_value)
					self.length_number.setText(str(new_value))
					self.selected_item.setLine(line)

	def imageTrans(self):
		# 创建文件对话框
		file_dialog = QFileDialog()
		file_dialog.setNameFilter("图片文件 (*.png *.jpg *.jpeg)")
		file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)  # 修改1：允许多选
		file_dialog.setViewMode(QFileDialog.ViewMode.Detail)

		# 显示文件对话框
		if file_dialog.exec():
			selected_files = file_dialog.selectedFiles()
			success_count = 0  # 成功计数器
			fail_list = []  # 失败文件列表
			file_name = ""
			# if len(selected_files) > 0:
			try:
				for file_path in selected_files:
					print(file_path)
					# file_path = selected_files[0]
					file_name = os.path.basename(file_path)
					base_name, original_ext = os.path.splitext(file_name)
					original_ext = original_ext.lstrip('.')  # 去掉点，得到 "jpg"
					# 创建 QImage 对象
					image = QImage(file_path)
					if image.isNull():
						QMessageBox.warning(self, "提示信息", "图片未成功加载!")
						return

					# 缩放图像
					scaled_image = image.scaled(
						int(900 * 0.8),
						int(740 * 0.8),
						Qt.AspectRatioMode.KeepAspectRatio,
						Qt.TransformationMode.SmoothTransformation
					)
					folder_path = r".\trans"
					os.makedirs(folder_path, exist_ok=True)
					# 保存图像
					output_ext = original_ext if original_ext.lower() in ['png', 'jpg', 'jpeg'] else 'png'
					output_file = os.path.join(folder_path, f"{base_name}_trans.{output_ext}")
					# 格式特殊处理
					quality = 90  # JPG 质量参数
					if output_ext.lower() in ['jpg', 'jpeg']:
						success = scaled_image.save(output_file, "JPEG", quality)
					else:
						success = scaled_image.save(output_file, output_ext.upper())
					if success:
						success_count += 1
					else:
						fail_list.append(f"{file_name} (保存失败)")
			except Exception as e:
				fail_list.append(f"{file_name} ({str(e)})")

			# 汇总显示结果
			msg = []
			if success_count > 0:
				msg.append(f"成功处理 {success_count} 张图片")
			if fail_list:
				msg.append("失败文件:\n" + "\n".join(fail_list))

			QMessageBox.information(
				self,
				"处理结果",
				"\n\n".join(msg) if msg else "没有选择任何文件"
			)

	# if success:
	# 	QMessageBox.information(self, "成功!", "图片已处理!")
	# else:
	# 	QMessageBox.warning(self, "提示信息", f"图片保存失败，可能不支持 {output_ext} 格式")

	def con_to_pdf(self):
		flag = False
		if self.file_name.text() != "":
			flag = True
		file_dialog = QFileDialog()
		file_dialog.setNameFilter("表格文件 (*.docx *.html)")  # 设置过滤器，只显示图片文件
		if file_dialog.exec():
			selected_files = file_dialog.selectedFiles()
			if selected_files:
				file_path = selected_files[0]
				file_ext = os.path.splitext(file_path)[1].lower()
				if file_ext == ".docx":
					if self.toggle_box.isChecked():
						self.status_bar.showMessage("正在转换中......")
						res = False
						if flag:
							res = Convertor.convert_to_pdf_micro(file_path, file_name=self.file_name.text() + ".pdf")
						else:
							res = Convertor.convert_to_pdf_micro(file_path)
						if res:
							QMessageBox.information(self, "消息提示", "成功转为pdf文件")
							self.path_label.setText(self.script_dir + "\\out_pdf\\" + self.file_name.text())
						else:
							QMessageBox.information(self, "消息提示", "格式转换失败", )
						self.status_bar.showMessage("表格检测")
					else:
						res = False
						if flag:
							res = Convertor.convert_to_pdf(file_path, 1, file_name=self.file_name.text() + ".pdf")
						else:
							res = Convertor.convert_to_pdf(file_path, 1)
						if res:
							QMessageBox.information(self, "消息提示", "成功转为pdf文件")
							self.path_label.setText(self.script_dir + "\\out_pdf\\" + self.file_name.text())
						else:
							QMessageBox.information(self, "消息提示", "格式转换失败", )
				elif file_ext == ".html":
					Convertor.convert_to_pdf(file_path, 2)
				else:
					QMessageBox.warning(self, "出现未知错误!", "消息提示")

	def con_to_word(self):
		flag = False
		if self.file_name.text() != "":
			flag = True
		file_dialog = QFileDialog()
		file_dialog.setNameFilter("表格文件 (*.pdf *.html)")  # 设置过滤器，只显示图片文件
		if file_dialog.exec():
			selected_files = file_dialog.selectedFiles()
			if selected_files:
				file_path = selected_files[0]
				file_ext = os.path.splitext(file_path)[1].lower()
				if file_ext == ".pdf":
					res = False
					if flag:
						res = Convertor.convert_to_word(file_path, 1, file_name=self.file_name.text() + ".docx")
					else:
						res = Convertor.convert_to_word(file_path, 1)
					if res:
						QMessageBox.information(self, "消息提示", "文件转换成功!")
						self.path_label.setText(self.script_dir + "\\out_word\\" + self.file_name.text())
					else:
						QMessageBox.warning(self, "消息提示", "文件转换失败!")
				elif file_ext == ".html":
					if flag:
						res = Convertor.convert_to_word(file_path, 2, file_name=self.file_name.text() + ".docx")
					else:
						res = Convertor.convert_to_word(file_path, 2)
					if res:
						QMessageBox.information(self, "消息提示", "文件转换成功!")
						self.path_label.setText(self.script_dir + "\\out_word\\" + self.file_name.text())
					else:
						QMessageBox.warning(self, "消息提示", "文件转换失败!")
				else:
					QMessageBox.warning(self, "消息提示", "出现未知错误!")

	def con_to_html(self):
		flag = False
		if self.file_name.text() != "":
			flag = True
		file_dialog = QFileDialog()
		file_dialog.setNameFilter("表格文件 (*.docx *.pdf)")  # 设置过滤器，只显示图片文件
		if file_dialog.exec():
			selected_files = file_dialog.selectedFiles()
			if selected_files:
				file_path = selected_files[0]
				file_ext = os.path.splitext(file_path)[1].lower()
				if file_ext == ".docx":
					res = False
					if flag:
						res = Convertor.convert_to_html(file_path, 1, file_name=self.file_name.text() + ".html")
					else:
						res = Convertor.convert_to_html(file_path, 1)
					if res:
						QMessageBox.information(self, "消息提示", "文件转换成功!")
						self.path_label.setText(self.script_dir + "\\out_html\\" + self.file_name.text())
					else:
						QMessageBox.warning(self, "消息提示", "文件转换失败!")
				elif file_ext == ".pdf":
					if flag:
						res = Convertor.convert_to_html(file_path, 2, file_name=self.file_name.text() + ".html")
					else:
						res = Convertor.convert_to_html(file_path, 2)
					if res:
						QMessageBox.information(self, "消息提示", "文件转换成功!")
						self.path_label.setText(self.script_dir + "\\out_html\\" + self.file_name.text())
					else:
						QMessageBox.warning(self, "消息提示", "文件转换失败!")
				else:
					QMessageBox.warning(self, "消息提示", "出现未知错误!")

	def set_line_special(self):
		print("enter set special")
		items = self.scene.selectedItems()
		if len(items) and isinstance(items[0], MovableLineItem):
			items[0].special = not items[0].special
			if items[0].special:
				items[0].setPen(QPen(Qt.GlobalColor.black, 2))
				if items[0].type == "type_y":
					self.scene.special_col_lines.append(items[0])
					self.scene.col_lines.remove(items[0])
				elif items[0].type == "type_x":
					self.scene.special_col_lines.append(items[0])
					self.scene.row_lines.remove(items[0])
			else:
				items[0].setPen(QPen(Qt.GlobalColor.red, 2))
				if items[0].type == "type_y":
					self.scene.special_col_lines.remove(items[0])
					self.scene.col_lines.append(items[0])
				elif items[0].type == "type_x":
					self.scene.special_col_lines.remove(items[0])
					self.scene.row_lines.append(items[0])
		print("end set special")

	def standardize_special_lines_pos(self):
		print("enter standardize special lines pos")
		for line_item in self.scene.special_col_lines:
			line_item.getCorrectPosY()
			(flag, pos_y, length) = self.is_near_row(self.scene.row_lines, line_item)
			if flag:
				delta = pos_y - line_item.line().y1()
				if abs(delta) >= 0.1:
					line_item.setPos(line_item.scenePos().x(), delta)
				if length != 0:
					line = line_item.line()
					line.setLength(length)
					line_item.setLine(line)
				line_item.getCorrectPosY()
		print("end standardize special lines pos")

	def is_near_row(self, row_lines, line_item):
		print("enter is_near_row")
		r1 = False
		pos_y = line_item.point_y
		pos_y_end = pos_y + line_item.line().length()
		length = 0
		for row_item in row_lines:
			if abs(row_item.point_y - pos_y) < 5:
				r1 = True
				pos_y = row_item.point_y
			if abs(row_item.point_y - pos_y_end) < 5:
				length = row_item.point_y - pos_y

		print("end is_near_row")
		return r1, pos_y, length

	def pre_close(self):
		msg_box = QMessageBox()
		msg_box.setWindowTitle("操作提示")
		msg_box.setText("确定结束程序吗？")
		# 添加完全自定义的按钮（文本和角色）
		confirm_btn = msg_box.addButton("确认", QMessageBox.ButtonRole.YesRole)
		cancel_btn = msg_box.addButton("取消", QMessageBox.ButtonRole.NoRole)
		# 显示弹框
		msg_box.exec()
		# 通过按钮角色判断用户操作
		if msg_box.clickedButton() == confirm_btn:
			self.close()
		elif msg_box.clickedButton() == cancel_btn:
			pass

	def back_forward_scene(self):
		if self.scene is not None:
			msg_box = QMessageBox()
			msg_box.setWindowTitle("操作提示")
			msg_box.setText("确定重新选择文件吗？")
			# 添加完全自定义的按钮（文本和角色）
			confirm_btn = msg_box.addButton("确认", QMessageBox.ButtonRole.YesRole)
			cancel_btn = msg_box.addButton("取消", QMessageBox.ButtonRole.NoRole)
			# 显示弹框
			msg_box.exec()
			# 通过按钮角色判断用户操作
			if msg_box.clickedButton() == confirm_btn:
				self.view.setScene(self.init_scene)
				self.scene.clear_data()
				self.scene = None
				self.rect_isvisible = None
				self.edit_window_show = False
				self.selected_item = None
				self.rect_buttons[2].setText("显示文字方框")
				for button in self.rect_buttons:
					button.setEnabled(False)
				for button in self.line_buttons:
					button.setEnabled(False)
				for button in self.sys_buttons:
					button.setEnabled(False)
			elif msg_box.clickedButton() == cancel_btn:
				pass
		else:
			QMessageBox.warning(self, "消息提示", "已经是文件上传界面!")


if __name__ == "__main__":
	app = QApplication(sys.argv)
	window = TableDetectionWindow()
	window.initUI()
	window.show()
	sys.exit(app.exec())
