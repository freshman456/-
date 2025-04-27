from PyQt6.QtCore import QLineF, Qt
from PyQt6.QtGui import QPainter, QPainterPath, QPainterPathStroker
from PyQt6.QtWidgets import QGraphicsLineItem, QStyle


class MovableLineItem(QGraphicsLineItem):
	def __init__(self, x1, y1, x2, y2, parent=None):
		super().__init__(QLineF(x1, y1, x2, y2), parent)
		self.point_x = None
		self.point_y = None
		self.special = False
		self.type = None

	def getCorrectPosX(self):
		self.point_x = self.line().x1() + self.scenePos().x()

	def getCorrectPosY(self):
		self.point_y = self.line().y1() + self.scenePos().y()

	def shape(self):
		# 创建路径（原始线条）
		path = QPainterPath()
		path.moveTo(self.line().p1())
		path.lineTo(self.line().p2())
		# 使用 QPainterPathStroker 加宽路径
		stroker = QPainterPathStroker()
		stroker.setWidth(12)  # 设置点击区域的宽度（敏感区域）
		return stroker.createStroke(path)

	def paint(self, painter: QPainter, option, widget=None):
		# 保存原始画笔和选项
		original_flags = option.state
		# 禁用选中状态的绘制
		option.state &= ~QStyle.StateFlag.State_Selected # 移除选中状态标志
		super().paint(painter, option, widget)
		# 恢复选项状态（可选）
		option.state = original_flags