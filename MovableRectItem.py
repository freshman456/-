from PyQt6.QtCore import QRect, QRectF, Qt, pyqtSignal
from PyQt6.QtGui import QBrush, QFont, QPainter
from PyQt6.QtWidgets import QGraphicsItem, QGraphicsRectItem, QGraphicsTextItem, QLabel
import re

from project.LinesUtil import LinesUtil


class MovableRectItem(QGraphicsRectItem):
	def __init__(self, text="文本(双击进行编辑)", rect=None, parent=None):
		if rect is None:
			super().__init__(parent)  # 默认方框尺寸80x60
		else:
			super().__init__(0, 0, rect.width(), rect.height(), parent)

		self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsMovable |
		              QGraphicsItem.GraphicsItemFlag.ItemIsSelectable |
		              QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges |
		              QGraphicsItem.GraphicsItemFlag.ItemSendsScenePositionChanges)
		self.setAcceptHoverEvents(True)
		self.text = text
		self.viewText = ""
		self.text_item = QGraphicsTextItem(self)
		self.text_item.setDefaultTextColor(Qt.GlobalColor.blue)

		font = QFont()
		font.setPointSize(14)  # 字体大小（单位：点）
		font.setWeight(QFont.Weight.Medium)  # PyQt6 必须使用 QFont.Weight 枚举
		self.text_item.setFont(font)

		self.text_item.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
		self.text_item.setZValue(1)  # 确保文本在矩形上层
		if rect is not None:
			self.setViewText(text)

		self.col = 0
		self.row = 0

	def paint(self, painter: QPainter, option, widget=None):
		# 自定义绘制样式
		painter.setPen(self.pen())
		painter.setBrush(QBrush(Qt.GlobalColor.transparent))
		painter.drawRoundedRect(self.rect(), 0, 0)

	def adjust_text_position(self):
		# 文本居中处理
		text_rect = self.text_item.boundingRect()
		rect = self.rect()
		self.text_item.setPos(
			rect.center().x() - text_rect.width() / 2,
			rect.center().y() - text_rect.height() / 2
		)

	def mouseDoubleClickEvent(self, event):
		super().mouseDoubleClickEvent(event)

	def itemChange(self, change, value):
		# 位置/尺寸变化时调整文本位置
		if change == QGraphicsItem.GraphicsItemChange.ItemPositionHasChanged or \
				change == QGraphicsItem.GraphicsItemChange.ItemTransformHasChanged:
			self.adjust_text_position()
			self.setViewText(self.text)
		return super().itemChange(change, value)

	def reSize(self, width, height):
		""" 设置新尺寸并保持中心点不变 """
		old_center = self.rect().center()
		new_rect = QRectF(0, 0, width, height)
		new_rect.moveCenter(old_center)
		self.setRect(new_rect)

	def setViewText(self, text):
		print("enter setViewText")
		if len(text) <= 3:
			self.viewText = text
		# 160->8个汉字 16个英文字母
		else:
			width = self.rect().width()
			res = LinesUtil.count_chinese_and_non_chinese_regex(text)
			max_len = int(width * 8 / 160.0 )
			if res > max_len:
				self.viewText = text[:max_len - 3]
				if len(self.viewText) != 0:
					self.viewText = self.viewText + "..."
				else:
					self.viewText = self.text[0] + "..."
			else:
				self.viewText = text

		self.text_item.setPlainText(self.viewText)
		print("end setViewText")
