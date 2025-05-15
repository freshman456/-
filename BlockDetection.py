
import numpy as np
from PyQt6.QtGui import QImage
from paddleocr import PaddleOCR


class BlockDetection:
	def __init__(self):
		self.data = None
		self.text = []
		self.rect_points = []
		self.block_texts = []

	def ocr(self, image):
		# 初始化 OCR 模型，设置参数以适应核显环境
		ocr = PaddleOCR(
			use_angle_cls=True,  # 启用角度分类
			lang="ch",  # 使用中文模型
			show_log=False,  # 不显示日志信息，减少资源占用
			det_db_thresh=0.4,  # 设置文字像素点得分阈值
			det_db_box_thresh=0.8,  # 设置文字区域平均得分阈值
			det_db_unclip_ratio=1.0,  # 设置文字区域扩张系数
			max_batch_size=10,  # 设置预测的 batch size
			use_dilation=False,  # 不对分割结果进行膨胀
			det_db_score_mode="slow",  # 使用快速得分计算方法
			ocr_version="PP-OCRv3"  # 使用 PP-OCRv3 模型
		)
		# 进行文字识别
		if isinstance(image,QImage):
			if image.format() != QImage.Format.Format_RGB888:
				image = image.convertToFormat(QImage.Format.Format_RGB888)
			buffer = image.constBits()
			buffer.setsize(image.sizeInBytes())
			byte_data = bytes(buffer)

			# 关键参数计算
			height, width = image.height(), image.width()
			bytes_per_line = image.bytesPerLine()  # 每行实际字节数（含填充）

			# 第一步：将数据视为二维数组 [行数, 每行字节数]
			np_image = (np.frombuffer(byte_data, dtype=np.uint8)
			            .reshape((height, bytes_per_line)))

			# 第二步：裁剪掉每行的对齐填充部分
			np_image = np_image[:, :width * 3]  # 每行保留有效像素数据
			# 第三步：转换为三维数组 [高, 宽, 通道]
			np_image = np_image.reshape((height, width, 3))
			# 通道顺序调整（RGB → BGR）
			np_image = np_image[..., ::-1]  # 等同于 cv2.cvtColor(np_image, cv2.COLOR_RGB2BGR)
			result = ocr.ocr(np_image)
			self.clear_data()
			# 提取文本块信息
			self.data = result[0]  # 假设只处理第一张图片的结果
			if self.data is None or len(self.data) == 0 :
				return
			# 按文本块的垂直位置（y 坐标）排序
			self.data.sort(key=lambda block: block[0][0][1])  # 按第一个点的 y 坐标排序
			# 存储每一行的数据
			current_row = []
			# 当前行的纵坐标
			previous_y = None
			# 用于判断是否属于同一行的阈值
			threshold = 10
			index = 1
			block_text = [f"第{index}行"]
			for block in self.data:
				y = block[0][0][1]
				# 如果是第一行，直接添加到当前行
				# 如果current_row为空
				if not current_row:
					current_row.append(block)
					previous_y = y
				else:
					# 如果当前文本块的 y 坐标与上一个文本块的 y 坐标差值小于阈值，认为属于同一行
					if abs(y - previous_y) < threshold:
						current_row.append(block)
					# block_text.append(block[1][0])
					else:
						# lambda  关键字 b表示接受第一个参数 b[0][0][0]取得是 x1的值 key= 是sort方法中的参数
						current_row.sort(key=lambda block: block[0][0][0])
						for cell in current_row:
							block_text.append(cell[1][0])
						self.block_texts.append(block_text)
						index += 1
						block_text = [f"第{index}行"]
						# rows.append(current_row)
						current_row = [block]
						previous_y = y
			# 处理最后一行

			if current_row:
				current_row.sort(key=lambda block: block[0][0][0])
				for cell in current_row:
					block_text.append(cell[1][0])
				self.block_texts.append(block_text)
			self.get_point_and_text()

	def get_point_and_text(self):
		scale_factor = 1  # 缩放因子
		for block in self.data:
			point = [[int(x * (1 / scale_factor)), int(y * (1 / scale_factor))] for [x, y] in block[0]]
			self.rect_points.append(point)
			block[0] = [[int(x * scale_factor), int(y * scale_factor)] for [x, y] in block[0]]
			self.text.append(block[1][0])

	def clear_data(self):
		self.data = None
		self.text.clear()
		self.rect_points.clear()
		self.block_texts.clear()
