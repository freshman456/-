import re


class LinesUtils:
	@staticmethod
	def identify_rows(cells, threshold=10):
		start_y = cells[0][1]
		rows = []
		temp_row = []
		for cell in cells:
			if abs(cell[1] - start_y) < threshold:
				temp_row.append(cell)
			else:
				rows.append(temp_row)
				temp_row = [cell]
				start_y = cell[1]
		if temp_row:
			rows.append(temp_row)
		return rows

	@staticmethod
	def identify_columns(cells, threshold=15):
		sorted_cells = sorted(cells, key=lambda x: x[0])
		start_x = sorted_cells[0][0]
		columns = []
		temp_column = []
		for cell in sorted_cells:
			if abs(cell[0] - start_x) < threshold:
				temp_column.append(cell)
			else:
				columns.append(temp_column)
				temp_column = [cell]
				start_x = cell[0]
		if temp_column:
			columns.append(temp_column)
		return columns

	@staticmethod
	def drawColLines(columns):
		count = len(columns)
		col_lines = []
		line = []
		for column in columns:
			line.append(column[0][0])
			eight_column_values = [row[7] for row in column]
			# 找到最大值作为该列的长度
			length = max(eight_column_values)
			line.append(length)
			col_lines.append(line)
			line = []
			count -= 1

			if count == 0:
				x_points = [row[2] for row in column]
				start_x = min(x_points)
				line.append(start_x)
				line.append(length)
				col_lines.append(line)

		return col_lines

	@staticmethod
	def drawRowLines(rows):
		count = len(rows)
		row_lines = []
		line = []
		for row in rows:
			# 取y坐标
			second_column_values = [row[1] for row in row]
			# 取水平方向上的x值
			length_values = [row[2] for row in row]
			# 最小值作为该行水平线的起始y
			start_y = min(second_column_values)
			# x中的最大值作为该行水平线的长度
			length = max(length_values)

			line.append(start_y)
			line.append(length)
			row_lines.append(line)
			line = []
			count -= 1

			if count == 0:
				# 最后一行的话 取右下角的y作为结束的水平线的起始y值
				last_row = [row[7] for row in row]
				start_y = min(last_row)
				line.append(start_y)
				line.append(length)
				row_lines.append(line)

		return row_lines

	@staticmethod
	def sortLines(lines, mode):
		print("enter sortLines")
		if mode == 1:
			for line in lines:
				line.get_correct_pos_x()
			sort_lines = sorted(lines, key=lambda item: item.point_x)
			return sort_lines
		elif mode == 2:
			for line in lines:
				line.get_correct_pos_y()
			sort_lines = sorted(lines, key=lambda item: item.point_y)
			return sort_lines
		else:
			print("end sortLines3")
			return lines

	@staticmethod
	def count_chinese_and_non_chinese_regex(text):
		# 匹配所有汉字
		chinese_pattern = re.compile(r'[\u4e00-\u9fff]')
		chinese_count = len(chinese_pattern.findall(text))
		non_chinese_count = len(text) - chinese_count
		res = chinese_count + non_chinese_count / 2
		return res
