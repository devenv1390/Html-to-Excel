from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, RGBColor


# 填充标题表格
def fill_title_table(table, data_list):
    print("==========================================")
    print("------ 正在处理标题数据 ------")
    for row_index, row in enumerate(table.rows):
        for col_index, cell in enumerate(row.cells):
            if col_index == 2:
                for data_index, data in enumerate(data_list):
                    if table.cell(row_index, col_index - 1).text in data[2]:
                        if data[4] == 'warning':
                            cell.text = "N/A"
                        else:
                            if data[3] == 'pass':
                                cell.text = "OK"
                            else:
                                cell.text = "NOK"
                        data_list.pop(data_index)
                if cell.text != 'AUTOSAR网络管理测试' and cell.text != '物理层测试' and cell.text != '数据链路层测试' and cell.text != '网络管理测试' and cell.text != '应用层测试':
                    if cell.text == '':
                        cell.text = "N/A"
                    set_result_type(cell)

    print("------ 完成标题数据处理 ------")
    print("==========================================")


# 填充普通表格
def fill_normal_table(table, data_list):
    for row_index, row in enumerate(table.rows):
        for col_index, cell in enumerate(row.cells):
            if row_index >= 2 and row_index - 2 < len(data_list):
                if col_index == 2:
                    cell.text = str(data_list[row_index - 2][0])

                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER  # 设置单元格垂直居中对齐

                    # 遍历单元格内的段落并设置水平居中对齐
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 设置段落水平居中对齐
                        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER  # 设置单元格垂直居中对齐
                        paragraph.paragraph_format.widow_control = True  # 设置自动换行
                        run = paragraph.runs[0]
                        run.font.size = Pt(10.5)

                        # 遍历段落内的run，并设置字体
                        for run in paragraph.runs:
                            for char in run.text:
                                if 0x4E00 <= ord(char) <= 0x9FFF:
                                    # 中文设置为宋体
                                    run.font.name = '宋体'
                                else:
                                    # 英文、数字和符号设置为Times New Roman
                                    run.font.name = 'Times New Roman'

                elif col_index == 4:
                    cell.text = data_list[row_index - 2][1]
                    if cell.text == 'warning':
                        cell.text = "N/A"
                    else:
                        if cell.text == 'pass':
                            cell.text = "OK"
                        else:
                            cell.text = "NOK"
                    set_result_type(cell)


# 预填充表格
def pre_fill_normal_table(table):
    for row_index, row in enumerate(table.rows):
        for col_index, cell in enumerate(row.cells):
            if row_index >= 2:
                if col_index == 4:
                    cell.text = "N/A"
                    set_result_type(cell)


# 设置测试结果单元格样式
def set_result_type(cell):
    # 设置垂直居中对齐
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    content = cell.text.strip()

    shading_color = None  # 默认为无色
    if content == "OK":
        shading_color = RGBColor(0, 128, 0)  # 绿色
        # print("OK")
    elif content == "NOK":
        shading_color = RGBColor(255, 0, 0)  # 红色
        # print("NOK")
    elif content == "N/A":
        shading_color = RGBColor(255, 255, 255)  # 白色
        # print("N/A")

    # 添加或修改单元格的背景色
    if shading_color is not None:
        if cell._element.tcPr is None:
            cell._element.tcPr = parse_xml(f'<w:tcPr {nsdecls("w")}/>')
        shading_element = parse_xml(
            f'<w:shd {nsdecls("w")} w:fill="{shading_color}"/>')
        cell._element.tcPr.append(shading_element)

    # 遍历单元格内的段落并设置水平居中对齐
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragraph.runs[0]
        run.font.size = Pt(10.5)
        run.font.name = 'Times New Roman'


# 读出单元格内容，主要用于测试
def find_text_with_read_table(docx_file, target_text):
    doc = Document(docx_file)

    paragraphs = doc.paragraphs
    all_tables = doc.tables
    target_text = target_text.encode('utf-8').decode('utf-8')

    for aPara in paragraphs:  # 遍历段落找表格
        if target_text in aPara.text:
            ele = aPara._p.getnext()
            while ele.tag != '' and ele.tag[-3:] != 'tbl':
                ele = ele.getnext()
            if ele.tag != '':
                for table in all_tables:
                    if table._tbl == ele:
                        for i in range(len(table.rows)):
                            for j in range(len(table.columns)):
                                print(table.cell(i, j).text)


# 填充word表格
def find_text_with_fill_table(docx_file, target_text, data_list,
                              file_path, index):
    doc = Document(docx_file)

    paragraphs = doc.paragraphs
    all_tables = doc.tables
    target_text = target_text.encode('utf-8').decode('utf-8')

    if index == 1:
        for table in all_tables:
            pre_fill_normal_table(table)
    #     for aPara in paragraphs:  # 遍历段落找表格
    #         # print(aPara.text)
    #         if title_target_text in aPara.text:
    #             ele = aPara._p.getnext()
    #             while ele.tag != '' and ele.tag[-3:] != 'tbl':
    #                 ele = ele.getnext()
    #             if ele.tag != '':
    #                 for table in all_tables:
    #                     if table._tbl == ele:
    #                         table.autofit = False
    #                         fill_title_table(table, title_data_list)

    print("==========================================")
    print("------ 正在处理第" + index.__str__() + "个测试数据 ------")

    for aPara in paragraphs:  # 遍历段落找表格
        if target_text in aPara.text:
            ele = aPara._p.getnext()
            while ele.tag != '' and ele.tag[-3:] != 'tbl':
                ele = ele.getnext()
            if ele.tag != '':
                for table in all_tables:
                    if table._tbl == ele:
                        table.autofit = False
                        print("Find " + target_text)
                        fill_normal_table(table, data_list)

    doc.save(file_path)

    print("------ 完成第" + index.__str__() + "个测试数据的处理 ------")
    print("==========================================")


# 填充word标题表格
def find_text_with_fill_title(docx_file, title_target_text, title_data_list, file_path):
    doc = Document(docx_file)

    paragraphs = doc.paragraphs
    all_tables = doc.tables

    for aPara in paragraphs:  # 遍历段落找表格
        # print(aPara.text)
        if title_target_text in aPara.text:
            ele = aPara._p.getnext()
            while ele.tag != '' and ele.tag[-3:] != 'tbl':
                ele = ele.getnext()
            if ele.tag != '':
                for table in all_tables:
                    if table._tbl == ele:
                        table.autofit = False
                        fill_title_table(table, title_data_list)

    doc.save(file_path)


# 检测特定字符串
def find_same_context(cell, target_list):
    # 使用循环遍历找到包含特定字符串的元素
    for _str in target_list:
        if cell.text.strip() in _str[2]:
            # print("find " + cell.text.strip())
            return True
    return False


# 解析HTML中的表格信息
def process_nested_table(table):
    nested_data = []

    rows = table.find_all('tr')

    # 获取表头
    header_row = rows[0]
    headers = []
    for th in header_row.find_all('th'):
        headers.append(th.text.strip())
    # 添加表头到数据
    nested_data.append(headers)

    # 处理表格行
    for row in rows[1:]:
        row_data = []
        for cell in row.find_all(['th', 'td']):
            # 如果单元格内包含其他嵌套表格，则递归处理
            if cell.find('table'):
                nested_table = cell.find('table')
                nested_table_data = process_nested_table(nested_table)
                row_data.append(nested_table_data)
            else:
                row_data.append(cell.get_text().strip())

        # 将行数据添加到嵌套数据中
        nested_data.append(row_data)

    return nested_data


# 解析HTML中的标题信息
def process_table(table):
    data = []

    rows = table.find_all('tr')

    # 处理表格行
    for row in rows[0:]:
        row_data = []
        for cell in row.find_all(['th', 'td']):
            # 如果单元格内包含其他嵌套表格，则递归处理
            if cell.find('table'):
                nested_table = cell.find('table')
                nested_table_data = process_nested_table(nested_table)
                row_data.append(nested_table_data)
            else:
                row_data.append(cell.get_text().strip())

        # 将行数据添加到嵌套数据中
        data.append(row_data)

    return data


# 替换文本
def replace_at_symbol(lst):
    result = []
    for item in lst:
        if isinstance(item, list):
            result.append(replace_at_symbol(item))
        elif isinstance(item, str) and item == '<br>' and item == '</br>':
            result.append(';')
        else:
            result.append(item)
    return result


# 处理数据，整理后合并成一个final_list
def connect_data_type_zero(table, title_data):
    final_data = []
    flag_count = 0  # 计数，第几个测试项目
    for cell in table:
        for _cell in cell:
            for __cell in _cell:
                if __cell == 'Timestamp':
                    temp_data = cell
                    final_data.append(['序号', '题号', '测试大类项目名称', '大类项目测试结果'])
                    if len(title_data[flag_count]) > 4:
                        title_data[flag_count].pop()
                        final_data.append(title_data[flag_count])
                    else:
                        final_data.append(title_data[flag_count])
                    flag_count += 1
                    final_data.append(['测试项目', '测试标准', '测试数值', '测试结果'])
                    for i in range(4, len(temp_data) + 1, 4):
                        temp_cell_data = temp_data[i]
                        temp_cell_data.append(temp_data[i - 3].pop())
                        final_data.append(temp_cell_data)
                    final_data.append([' '])
                    final_data.append([' '])
    return final_data


def connect_data_type_one(table, title_data):
    final_data = []
    flag_count = 0  # 计数，第几个测试项目
    for cell in table:
        for _cell_index, _cell in enumerate(cell):
            if _cell == 'Timestamp':
                temp_data = cell
                final_data.append(['序号', '题号', '测试大类项目名称', '大类项目测试结果'])
                if len(title_data[flag_count]) > 4:
                    title_data[flag_count].pop()
                    final_data.append(title_data[flag_count])
                else:
                    final_data.append(title_data[flag_count])
                flag_count += 1
                final_data.append(['测试项目', '测试标准', '测试数值', '测试结果'])
                for i in range(4, len(temp_data) + 1, 4):
                    temp_cell_data = temp_data[i]
                    temp_cell_data.append(temp_data[i - 3].pop())
                    final_data.append(temp_cell_data)
                final_data.append([' '])
                final_data.append([' '])
    return final_data


# 提取final_list中的单个list元素作为data_list填入表格
def get_list_from_final(final_list):
    data_list = []

    for data_index, data in enumerate(final_list):
        temp_list = []
        if len(data) > 2:
            if data[1] == '题号':
                temp_list.append([final_list[data_index + 1][1] + " " + final_list[data_index + 1][2]])
                for _data in range(data_index + 3, len(final_list) - 1):
                    if final_list[_data] == [' ']:
                        break
                    else:
                        temp_list.append([final_list[_data][2], final_list[_data][3]])
                data_list.append(temp_list)
    return data_list


# list内元素计数
def count_element(lst, target):
    count = 0
    for item in lst:
        if isinstance(item, list):
            count += count_element(item, target)  # 递归调用处理嵌套列表
        elif item == target:
            count += 1
    return count

# #填充cell
# def fill_cell(ws):
#     prev_row_values = []  # 用于存储上一行的值
#     for row in ws.rows:
#         for cell in row:
#             cell.alignment = alignment
#             if cell.value == 'pass':
#                 cell.fill = green_fill
#                 cell.font = data_font
#                 cell.value = 'OK'
#                 cell.border = thin_border
#             elif cell.value == 'warning':
#                 cell.fill = yellow_fill
#                 cell.font = data_font
#                 cell.value = 'N/A'
#                 cell.border = thin_border
#             elif cell.value == 'fail':
#                 cell.fill = red_fill
#                 cell.font = data_font
#                 cell.value = 'NOK'
#                 cell.border = thin_border
#             elif cell.value is not None:
#                 if cell.value != ' ':
#                     cell.border = thin_border
#                     cell.font = value_font
#                     for _i in header_list:
#                         if cell.value == _i:
#                             cell.font = header_font
#                             cell.fill = blue_fill
#                             cell.border = thin_border
#                     for header_value in ['序号', '题号', '测试大类项目名称', '大类项目测试结果']:
#                         if header_value in prev_row_values:
#                             cell.font = header_font
#         # 更新上一行的值
#         prev_row_values = [cell.value for cell in row]
#
# # 创建一个新的 DataFrame 来存储嵌套数据
#     df = pd.DataFrame(final_list)
#
#     # 写入 Excel 文件
#     df.to_excel(output_filepath, index=False)
#
#     # 修改 Excel 文件格式
#     wb = openpyxl.load_workbook(output_filepath)
#
#     # 获取默认的活动工作表
#     ws_detail = wb.active
#     ws_detail.title = "详细信息"
#
#     # 创建“总览”工作簿
#     ws_overview = wb.create_sheet(title="总览")
#     for i in range(len(title_data)):
#         for j in range(len(title_data[i])):
#             if isinstance(title_data[i][j], list):
#                 ws_overview.append(title_data[i][j])
#             else:
#                 ws_overview.cell(i + 1, j + 1).value = title_data[i][j]  # 写入数据
#                 ws_overview.cell(i + 1, j + 1).alignment = Alignment(horizontal='center', vertical='center')  # 居中对齐
#
#     # 设置列宽
#     ws_detail_column_widths = [20, 60, 70, 10, 10]
#     for i, width in enumerate(ws_detail_column_widths):
#         col_letter = chr(65 + i)  # A, B, C...
#         ws_detail.column_dimensions[col_letter].width = width
#
#     ws_overview_column_widths = [26, 10, 26, 10]
#     for i, width in enumerate(ws_overview_column_widths):
#         col_letter = chr(65 + i)  # A, B, C...
#         ws_overview.column_dimensions[col_letter].width = width
#
#     # 定义单元格边框样式
#     thin_border = Border(left=Side(style='thin'),
#                          right=Side(style='thin'),
#                          top=Side(style='thin'),
#                          bottom=Side(style='thin'))
#
#     # 设置单元格对齐方式 Alignment(horizontal=水平对齐模式,vertical=垂直对齐模式,text_rotation=旋转角度,wrap_text=是否自动换行)
#     alignment = Alignment(horizontal='center', vertical='center', text_rotation=0, wrap_text=True)
#
#     # 定义单元格样式
#     header_font = Font(bold=True, size=14)
#     data_font = Font(bold=True, size=14)
#     value_font = Font(size=14)
#     yellow_fill = PatternFill(fill_type='solid', fgColor='FFFF00')
#     green_fill = PatternFill(fill_type='solid', fgColor='00FF00')
#     red_fill = PatternFill(fill_type='solid', fgColor='FF0000')
#     blue_fill = PatternFill(fill_type='solid', fgColor='00B0F0')
#
#     # 拟造条件list
#     header_list = ['测试项目', '测试标准', '测试数值', '测试结果', '序号', '题号', '测试大类项目名称',
#                    '大类项目测试结果']
#     # 填充表格数据
#     fill_cell(ws_overview)
#     fill_cell(ws_detail)
#
#     temp_cell1 = ws_overview["c4"]
#     temp_cell2 = ws_overview["c5"]
#
#     temp_cell1.fill = green_fill
#     temp_cell2.fill = red_fill
#     temp_cell1.font = data_font
#     temp_cell2.font = data_font
#
#     # 将工作表名称排序
#     sort_lst = sorted(wb.sheetnames)
#
#     # 重新排列工作表
#     for sheet_name in sort_lst:
#         ws = wb[sheet_name]
#         wb.remove(ws)
#         wb._add_sheet(ws)
#
#     # 保存 Excel 文件
#     wb.save(output_filepath)
