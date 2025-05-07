import os

from PyQt6.QtCore import QLineF, QMarginsF, QRect, Qt, pyqtSignal
from PyQt6.QtGui import QPageSize, QPainter, QPen
from PyQt6.QtPrintSupport import QPrinter
from PyQt6.QtWidgets import QGraphicsItem, QGraphicsLineItem, QGraphicsScene
from docx import Document
from docx.shared import Cm

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

from project.MovableLineItem import MovableLineItem
from project.table_convertor import Convertor
from reportlab.pdfbase.pdfmetrics import stringWidth


class CustomizeScene(QGraphicsScene):
	sceneClickedSignal = pyqtSignal()  # 自定义信号
	editSignal = pyqtSignal()
	endLineSignal = pyqtSignal()

	def __init__(self):
		super().__init__()
		self.scene_border = None
		self.drawing_line = None
		self.start_pos = None
		self.current_mode = None
		self.background_pixmap = None
		self.export_flag = False
		self.col_lines = []
		self.row_lines = []
		self.bias_lines = []
		self.start_point_x = None
		self.start_point_y = None
		self.rect_items = []
		self.cells = []
		self.special_col_lines = []
		self.special_row_lines = []

	def set_background_pixmap(self, pixmap):
		self.background_pixmap = pixmap
		rect = self.background_pixmap.rect()
		c1 = int(self.sceneRect().width() / 2)
		c2 = int(self.sceneRect().height() / 2)
		x = int(rect.width() / 2)
		y = int(rect.height() / 2)
		self.start_point_x = int(c1 - x)
		self.start_point_y = int(c2 - y)

	def drawBackground(self, painter, rect):
		super().drawBackground(painter, rect)
		# 绘制背景图片
		if self.background_pixmap is not None and self.export_flag is False:
			rect = self.background_pixmap.rect()
			rect = QRect(self.start_point_x, self.start_point_y, rect.width(), rect.height())
			painter.drawPixmap(rect, self.background_pixmap, self.background_pixmap.rect())

	def mousePressEvent(self, event):
		super().mousePressEvent(event)
		if event.button() == Qt.MouseButton.LeftButton:
			if self.current_mode == "line":
				# 开始绘制线条
				self.start_pos = event.scenePos()
				self.drawing_line = MovableLineItem(self.start_pos.x(), self.start_pos.y(), self.start_pos.x(),
				                                    self.start_pos.y())
				self.drawing_line.setPen(QPen(Qt.GlobalColor.red, 2))
				self.addItem(self.drawing_line)
			elif self.current_mode == "view":
				self.sceneClickedSignal.emit()  # 发射信号

	def mouseDoubleClickEvent(self, event):
		super().mouseDoubleClickEvent(event)
		self.editSignal.emit()

	# super().mousePressEvent(event)

	def mouseMoveEvent(self, event):
		if self.current_mode == "line" and self.drawing_line:
			# 更新线条终点
			self.drawing_line.setLine(QLineF(self.start_pos, event.scenePos()))
		super().mouseMoveEvent(event)

	def mouseReleaseEvent(self, event):
		if self.drawing_line:
			self.drawing_line.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, True)
			self.drawing_line.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False)
			self.drawing_line.setFlag(QGraphicsItem.GraphicsItemFlag.ItemSendsScenePositionChanges, True)

			line = self.drawing_line.line()
			dx = line.dx()  # x方向变化量
			dy = line.dy()  # y方向变化量
			epsilon = 5.0
			if abs(dx) < epsilon:
				new_line = QLineF(line.x1(), line.y1(), line.x1(), line.y2())
				self.drawing_line.setLine(new_line)
				self.drawing_line.type = "type_y"
				self.col_lines.append(self.drawing_line)

			elif abs(dy) < epsilon:
				corrected_line = QLineF(line.p1().x(), line.p1().y(), line.p2().x(), line.p1().y())
				self.drawing_line.setLine(corrected_line)
				self.drawing_line.type = "type_x"
				self.row_lines.append(self.drawing_line)
			else:
				self.drawing_line.type = "type_bias"
				self.bias_lines.append(self.drawing_line)
			# 完成线条绘制
			self.drawing_line = None
			self.start_pos = None
			self.set_line_selectable()
			self.endLineSignal.emit()
		super().mouseReleaseEvent(event)

	def set_line_selectable(self):
		if self.col_lines:
			for line in self.col_lines:
				line.setFlag(QGraphicsLineItem.GraphicsItemFlag.ItemIsSelectable, True)
		if self.row_lines:
			for line in self.row_lines:
				line.setFlag(QGraphicsLineItem.GraphicsItemFlag.ItemIsSelectable, True)
		if self.bias_lines:
			for line in self.bias_lines:
				line.setFlag(QGraphicsLineItem.GraphicsItemFlag.ItemIsSelectable, True)

	def set_lines_not_mov(self):
		if self.col_lines:
			for line in self.col_lines:
				line.setFlag(QGraphicsLineItem.GraphicsItemFlag.ItemIsMovable, False)
		if self.row_lines:
			for line in self.row_lines:
				line.setFlag(QGraphicsLineItem.GraphicsItemFlag.ItemIsMovable, False)
		if self.bias_lines:
			for line in self.bias_lines:
				line.setFlag(QGraphicsLineItem.GraphicsItemFlag.ItemIsMovable, False)
		if self.special_col_lines:
			for line in self.special_col_lines:
				line.setFlag(QGraphicsLineItem.GraphicsItemFlag.ItemIsMovable, False)
		if self.special_row_lines:
			for line in self.special_row_lines:
				line.setFlag(QGraphicsLineItem.GraphicsItemFlag.ItemIsMovable, False)

	def export_pdf(self, standard=False, file_name="output.pdf"):
		if standard:
			self.get_row_and_col()
			return self.create_standard_pdf_table(file_name=file_name)
		else:
			self.scene_border.setPen(QPen(Qt.GlobalColor.transparent, 2))
			if self.rect_items:
				for rect_item in self.rect_items:
					rect_item.setPen(QPen(Qt.GlobalColor.transparent, 2))
					rect_item.text_item.setPlainText(rect_item.text)
					rect_item.text_item.setDefaultTextColor(Qt.GlobalColor.black)

			if self.row_lines:
				for line in self.row_lines:
					line.setPen(QPen(Qt.GlobalColor.black, 2))
			if self.bias_lines:
				for line in self.bias_lines:
					line.setPen(QPen(Qt.GlobalColor.black, 2))
			if self.col_lines:
				for line in self.col_lines:
					line.setPen(QPen(Qt.GlobalColor.black, 2))

			printer = QPrinter(QPrinter.PrinterMode.HighResolution)
			printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)

			folder_path = r".\output_pdf"

			os.makedirs(folder_path, exist_ok=True)
			full_path = os.path.join(folder_path, file_name)

			printer.setOutputFileName(full_path)

			printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
			printer.setPageMargins(QMarginsF(0, 0, 0, 0))

			painter = QPainter()
			if not painter.begin(printer):
				return False
			try:
				# 获取场景边界和页面边界
				scene_rect = self.sceneRect()
				# 乘以小于1的数才能方法 因为后面/width /height
				scene_rect.setWidth(self.width() * 0.9)
				scene_rect.setHeight(self.height() * 0.9)

				page_rect = printer.pageRect(QPrinter.Unit.Point)  # 显式指定单位为点

				# 计算缩放比例，确保内容完全适应页面
				scale_factor = min(
					page_rect.width() / scene_rect.width(),  # 595/900 ≈ 0.661
					page_rect.height() / scene_rect.height()  # 842/740 ≈ 1.138
				)
				# 计算页面中心坐标
				page_center_x = page_rect.width() / 2.0  # 595/2 = 297.5
				page_center_y = page_rect.height() / 2.0  # 842/2 = 421
				painter.translate(
					page_center_x * 2.33 + scene_rect.width() / 2.0,
					page_center_y * 2.33 + scene_rect.height() / 2.0  # 421 - 244.5 = 176.5（垂直居中上方留白）
				)

				painter.scale(scale_factor, scale_factor)

				self.export_flag = True
				# 可以改变绘制的表格的位置
				self.render(painter, source=scene_rect)
			finally:
				painter.end()
				self.export_flag = False
				self.scene_border.setPen(QPen(Qt.GlobalColor.black, 2))
				if self.rect_items:
					for rect_item in self.rect_items:
						rect_item.setPen(QPen(Qt.GlobalColor.green, 2))
						rect_item.text_item.setDefaultTextColor(Qt.GlobalColor.blue)
						rect_item.set_view_text(rect_item.text)
				if self.row_lines:
					for line in self.row_lines:
						line.setPen(QPen(Qt.GlobalColor.red, 2))
				if self.bias_lines:
					for line in self.bias_lines:
						line.setPen(QPen(Qt.GlobalColor.red, 2))
				if self.col_lines:
					for line in self.col_lines:
						line.setPen(QPen(Qt.GlobalColor.red, 2))
			return True

	def create_standard_pdf_table(self, file_name="output.pdf"):
		Convertor.register_font()
		folder_path = r"./output_pdf"
		print(self.cells)
		os.makedirs(folder_path, exist_ok=True)
		full_path = os.path.join(folder_path, file_name)
		cols = len(self.col_lines)
		rows = len(self.row_lines)
		if cols < 2 or rows < 2:
			return False
		font_name = "simhei"
		font_size = 11
		padding = 12
		col_widths = []
		# 获取字符串宽度函数

		# 遍历每列
		for col_idx in range(len(self.cells[0])):
			max_width = 0
			# 遍历每行该列的内容
			for row in self.cells:
				cell_content = str(row[col_idx])
				# 计算文字宽度（考虑换行符）
				lines = cell_content.split('\n')
				for line in lines:
					width = stringWidth(line, font_name, font_size)
					if width > max_width:
						max_width = width
			# 列宽 = 最大文字宽度 + 左右边距
			col_widths.append(max_width + 2 * padding)

		# ------------------------- 计算行高-------------------------
		row_heights = []
		line_height = font_size * 1.2  # 单行高度（含行距）
		padding = 8  # 上下边距各 8pt

		for row in self.cells:
			max_lines = 1  # 最小行数为 1
			for cell in row:
				cell_content = str(cell)
				# 关键修复：直接使用 split('\n') 的显式换行符分割结果
				explicit_lines = cell_content.split('\n')  # 显式换行符分割的行数
				line_count = len(explicit_lines)

				# 计算自动换行（根据列宽和文字长度）
				if col_widths is not None:  # 需要先计算列宽
					col_idx = row.index(cell)  # 获取当前单元格的列索引
					col_width = col_widths[col_idx] - 2 * padding  # 可用宽度
					# 计算自动换行后的行数（根据字体宽度）
					text_width = stringWidth(cell_content.replace('\n', ' '), font_name, font_size)
					if text_width > col_width:
						auto_line_count = int(text_width / col_width) + 1
						line_count = max(line_count, auto_line_count)

				if line_count > max_lines:
					max_lines = line_count

			# 行高 = 行数 * 单行高度 + 上下边距
			row_heights.append(max_lines * line_height + 2 * padding)

		doc = SimpleDocTemplate(full_path, pagesize=A4)
		table = Table(self.cells, colWidths=col_widths, rowHeights=row_heights)
		style = TableStyle([
			# 表格边框
			('GRID', (0, 0), (-1, -1), 1, colors.black),  # 全表格边框
			# 对齐方式
			('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # 居中对齐
			('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # 垂直居中
			('FONT', (0, 0), (-1, -1), "simhei"),
			('FONTSIZE', (0, 0), (- 1, - 1), 12)
		])
		table.setStyle(style)
		table.hAlign = 'CENTER'  # 设置表格在页面中水平居中
		# 构建PDF文档
		elements = [table]
		doc.build(elements)
		return True

	def export_word(self, file_name="output.docx"):
		self.get_row_and_col()
		widths = []
		length = len(self.col_lines)
		index = 0
		while length >= 2 and index + 1 < length:
			line1 = self.col_lines[index]
			line2 = self.col_lines[index + 1]
			x1 = line1.scenePos().x() + line1.line().x1()
			x2 = line2.scenePos().x() + line2.line().x1()
			width = (x2 - x1) / 37.8
			index += 1
			if width > 1:
				widths.append(width)
			else:
				widths.append(1)

		heights = []
		length = len(self.row_lines)
		index = 0

		while length >= 2 and index + 1 < length:
			line1 = self.row_lines[index]
			line2 = self.row_lines[index + 1]
			y1 = line1.scenePos().x() + line1.line().y1()
			y2 = line2.scenePos().x() + line2.line().y2()
			height = (y2 - y1) / 37.8
			index += 1
			if height > 1:
				heights.append(height)
			else:
				heights.append(1)
		try:
			if len(self.col_lines) < 2 or len(self.row_lines) < 2:
				return False
			doc = Document()
			table = doc.add_table(rows=len(heights), cols=len(widths), style="Table Grid")
			# 设置每列宽度（按列遍历）
			for col_idx, width in enumerate(widths):
				column = table.columns[col_idx]
				column.width = Cm(width)  # 直接设置整列宽度
			# 设置每行高度（按行遍历）
			for row_idx, height in enumerate(heights):
				row = table.rows[row_idx]
				row.height = Cm(height)  # 直接设置整行高度

			for row_idx, row in enumerate(table.rows):
				for col_idx, cell in enumerate(row.cells):
					# 检查行索引是否在 data 的范围内
					if row_idx < len(self.cells):
						row_data = self.cells[row_idx]
						# 检查列索引是否在该行的范围内
						if col_idx < len(row_data):
							cell.text = row_data[col_idx]
						else:
							cell.text = ""  # 列超出 data 当前行的范围
					else:
						cell.text = ""  # 行超出 data 的范围
			# 指定保存路径
			folder_path = r".\output_word"
			os.makedirs(folder_path, exist_ok=True)
			full_path = os.path.join(folder_path, file_name)
			doc.save(full_path)
		except PermissionError:
			return False
		except Exception as e:
			return False
		return True

	def get_row_and_col(self):
		if self.rect_items is not None:
			if len(self.col_lines) >= 2:
				for line in self.col_lines:
					line.get_correct_pos_x()

			if len(self.row_lines) >= 2:
				for line in self.row_lines:
					line.get_correct_pos_y()
			col_points = []
			cols = 0
			for line in self.col_lines:
				col_points.append(line.point_x)
			if col_points:
				cols = len(col_points) - 1

			row_points = []
			rows = 0
			for line in self.row_lines:
				row_points.append(line.point_y)
			if row_points:
				rows = len(row_points) - 1

			col_points.sort()
			row_points.sort()

			for rect in self.rect_items:
				pos = rect.scenePos()
				border = rect.rect()
				x = pos.x() + border.width() / 2
				y = pos.y() + border.height() / 2
				index = 0
				while index < len(col_points) and x > col_points[index]:
					index += 1
				if index < len(col_points) and x < col_points[index]:
					rect.col = index
				index = 0
				while index < len(row_points) and y > row_points[index]:
					index += 1
				if index < len(row_points) and y < row_points[index]:
					rect.row = index

			self.cells = [["" for _ in range(cols)] for _ in range(rows)]
			for item in self.rect_items:
				if item.row and item.col:
					self.cells[item.row - 1][item.col - 1] += item.text + '\n'

	def export_html5_table(self, file_name="output.html", special=True):
		self.get_row_and_col()
		if self.cells and len(self.row_lines) >= 2 and len(self.col_lines) >= 2:
			# 构建 HTML5 内容
			html = [
				'<!DOCTYPE html>',
				'<html lang="zh">',
				'<head>',
				'  <meta charset="UTF-8">',
				'  <title>导出表格</title>',
				'  <style>',
				'    table { '
				'              border-collapse: collapse; '
				'              width: 60%;'
				'              margin:20px auto;'
				'           }',
			]
			if special:  # 新增判断条件
				# 仅在特殊处理时添加第一行样式
				html.extend([
					'    tr:first-child th {',
					'      background-color: #cccccc; /* 深灰色背景 */',
					'      color: white;          /* 文字白色 */',
					'      font-weight: bold;     /* 强制加粗 */',
					'      padding: 12px;',
					'    }'
				])
			# 公共样式
			html.extend([
				'    td { padding: 10px; }',
				'    th, td { border: 2px solid black; text-align: center; }',
				'  </style>',
				'</head>',
				'<body>',
				'  <table>',
				'    <thead>',
				'      <tr>'
			])
			try:
				first_row_tag = 'th' if special else 'td'
				# 添加表头（仅第一行特殊处理为 th）
				html.extend(f'        <{first_row_tag}>{cell}</{first_row_tag}>' for cell in self.cells[0])
				html.append('      </tr>')
				html.append('    </thead>')

				# 添加数据行（全部使用 td）
				html.append('    <tbody>')
				for row in self.cells[1:]:
					html.append('      <tr>')
					html.extend(f'        <td>{cell}</td>' for cell in row)
					html.append('      </tr>')
				html.append('    </tbody>')

				# 闭合标签
				html.extend([
					'  </table>',
					'</body>',
					'</html>'
				])

				folder_path = r".\output_html"
				os.makedirs(folder_path, exist_ok=True)
				filename = os.path.join(folder_path, file_name)
				print(filename)
				# 写入文件
				with open(filename, 'w', encoding='utf-8') as f:
					f.write('\n'.join(html))
			except PermissionError:
				return False
			except Exception as e:
				return False
			return True

	def clear_data(self):
		self.scene_border = None
		self.drawing_line = None
		self.start_pos = None
		self.current_mode = None
		self.background_pixmap = None
		self.export_flag = False
		self.col_lines.clear()
		self.row_lines.clear()
		self.bias_lines.clear()
		self.start_point_x = None
		self.start_point_y = None
		self.rect_items.clear()
		self.cells.clear()
		self.special_col_lines.clear()
		self.special_row_lines.clear()
