from PyQt6.QtCore import QLineF, QPointF, QRectF, Qt, pyqtSignal
from PyQt6.QtGui import QPainterPath, QPainterPathStroker, QPen
from PyQt6.QtWidgets import QGraphicsLineItem, QGraphicsRectItem, QGraphicsView, QLabel
from matplotlib.hatch import SouthEastHatch

from project.CustomizeScene import CustomizeScene
from project.MovableRectItem import MovableRectItem


class CustomizeView(QGraphicsView):
	end_draw_rect = pyqtSignal()

	def __init__(self, parent=None):
		super().__init__(parent)
		# self.cursor_x = 0.0
		# self.cursor_y = 0.0
		self.scene_coord_label = QLabel(self)
		self.scene_coord_label.setGeometry(10, 10, 220, 20)
		self.current_rect_item = None
		self.start_point = 0.0
		self.dragging = False
		self.current_scene = None

	def mousePressEvent(self, event):
		if event.button() == Qt.MouseButton.LeftButton:
			self.current_scene = self.scene()
			if isinstance(self.current_scene, CustomizeScene):
				if "box" == self.current_scene.current_mode:
					point = self.mapToScene(event.pos())
					self.start_point = self.mapToScene(event.pos())
					self.dragging = True
					self.current_rect_item = MovableRectItem()
					self.current_rect_item.setPen(QPen(Qt.GlobalColor.green, 2))
					self.current_scene.addItem(self.current_rect_item)
		super().mousePressEvent(event)

	def mouseMoveEvent(self, event):
		# 获取视图坐标
		view_pos = event.pos()
		# 转换为场景坐标
		scene_pos = self.mapToScene(view_pos)
		# 更新标签显示场景坐标
		self.scene_coord_label.setText(f"Scene 坐标: ({scene_pos.x():.2f}, {scene_pos.y():.2f})")
		if isinstance(self.current_scene, CustomizeScene):
			if self.current_rect_item is not None:
				end_point = self.mapToScene(event.pos())
				rect = QRectF(self.start_point, end_point).normalized()
				self.current_rect_item.setRect(rect)
		super().mouseMoveEvent(event)

	def mouseReleaseEvent(self, event):
		if event.button() == Qt.MouseButton.LeftButton and self.dragging:
			final_rect = self.current_rect_item.rect()
			if final_rect.width() < 10 or final_rect.height() < 10:
				print("矩形框太小 不与创建")
				self.scene().removeItem(self.current_rect_item)
			else:
				self.current_rect_item.adjust_text_position()
				self.current_rect_item.setViewText(text="文本")
				rect = QRectF(0, 0, final_rect.width(), final_rect.height())
				self.current_rect_item.setRect(rect)
				self.current_rect_item.setPos(final_rect.x(), final_rect.y())
				self.current_scene.rect_items.append(self.current_rect_item)
				self.end_draw_rect.emit()
			self.current_rect_item = None
			self.dragging = False
			self.start_point = QPointF(0, 0)
			self.current_scene.clearSelection()
		super().mouseReleaseEvent(event)
