import numpy as np
from PyQt6.QtGui import QImage
from rapid_table import RapidTable, RapidTableInput


class StructureDetection:
	def __init__(self):
		input_args = RapidTableInput(
			model_type="slanet_plus",  # 选择模型类型
			use_cuda=False,  # 是否使用 GPU
			device="cpu"  # CPU 设备
		)
		self.table_engine = RapidTable(input_args)
		self.table_info = []

	def detect_structure(self, image):
		self.table_info = []
		if image is None:
			return
		if isinstance(image, QImage):
			if image.format() != QImage.Format.Format_RGB888:
				image = image.convertToFormat(QImage.Format.Format_RGB888)
			buffer = image.constBits()
			buffer.setsize(image.sizeInBytes())
			byte_data = bytes(buffer)

			# 关键参数计算
			height, width = image.height(), image.width()
			bytes_per_line = image.bytesPerLine()  # 每行实际字节数（含填充）

			# 正确reshape步骤（处理内存对齐）
			# 第一步：将数据视为二维数组 [行数, 每行字节数]
			np_image = np.frombuffer(byte_data, dtype=np.uint8).reshape((height, bytes_per_line))
			# 第二步：裁剪掉每行的对齐填充部分
			np_image = np_image[:, :width * 3]  # 每行保留有效像素数据
			# 第三步：转换为三维数组 [高, 宽, 通道]
			np_image = np_image.reshape((height, width, 3))
			# 通道顺序调整（RGB → BGR）
			np_image = np_image[..., ::-1]  # 等同于 cv2.cvtColor(np_image, cv2.COLOR_RGB2BGR)
			results = self.table_engine(np_image)
			self.table_info = results.cell_bboxes
			rounded_table_info = []
			if self.table_info is None or len(self.table_info) == 0 :
				return
			for row in self.table_info:
				rounded_row = [round(num, 2) for num in row]
				rounded_table_info.append(rounded_row)
			# 将结果赋值回 self.structure_detector.table_info
			self.table_info = rounded_table_info
