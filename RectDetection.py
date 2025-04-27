from rapid_table import RapidTable, RapidTableInput


class RectDetection:
	def __init__(self):
		input_args = RapidTableInput(
			model_type="slanet_plus",  # 选择模型类型
			use_cuda=False,  # 是否使用 GPU
			device="cpu"  # CPU 设备
		)
		self.table_engine = RapidTable(input_args)
		self.table_info = []

	def detectRect(self, path):
		self.table_info = []
		print("enter detect rect")
		if path is None:
			print("path is None")
			return
		results = self.table_engine(path)
		self.table_info = results.cell_bboxes
		rounded_table_info = []
		for row in self.table_info:
			rounded_row = [round(num, 2) for num in row]
			rounded_table_info.append(rounded_row)

		# 将结果赋值回 self.detector.table_info
		self.table_info = rounded_table_info
		print("end detect rect")
