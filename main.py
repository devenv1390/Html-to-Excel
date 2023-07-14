import os

from bs4 import BeautifulSoup

# 工具
from tools import find_text_with_fill_table, process_nested_table, process_table, count_element, connect_data, \
    replace_at_symbol

# 创建 output 文件夹（如果不存在）
if not os.path.exists('output'):
    os.makedirs('output')

# 获取 input 文件夹下的所有 HTML 文件
html_files = os.listdir('html_input')
docx_file = os.listdir('model_input')
print("------ 共有" + len(html_files).__str__() + "个文件待处理 ------")
i = 1
for filename in html_files:
    print("------ 正在处理第" + i.__str__() + "个文件 ------")
    i += 1
    # 拼接文件路径
    html_input_filepath = os.path.join('html_input', filename)
    model_input_filepath = os.path.join('model_input', "高速CAN单节点测试报告模板.docx")
    output_filepath = os.path.join('output', filename.replace('.html', '.docx'))

    # 读取 HTML 文件
    with open(html_input_filepath, 'r', encoding='utf-8') as file:
        html_content = file.read()

    # 使用 BeautifulSoup 解析 HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # 处理多级嵌套表格数据
    tables = soup.find_all('table')
    nested_tables_data = []
    nested_temp_data = []
    for table in tables:
        table_data = process_nested_table(table)
        nested_tables_data.append(table_data)
        temp_data = process_table(table)
        nested_temp_data.append(temp_data)

    # 算通过、不通过、没测试
    total = len(nested_temp_data[9])
    warning_number = count_element(nested_temp_data[9], 'warning')
    fail_number = count_element(nested_temp_data[9], 'fail')
    pass_number = total - fail_number - warning_number
    complete_number = pass_number + fail_number

    per_complete = round(complete_number / total, 2) * 100
    per_not_complete = 100 - per_complete
    per_pass = round(pass_number / total, 2) * 100
    per_fail = round(fail_number / total, 2) * 100

    str_pre_complete = per_complete.__str__() + "%"
    str_per_not_complete = per_not_complete.__str__() + "%"
    str_per_pass = per_pass.__str__() + "%"
    str_per_fail = per_fail.__str__() + "%"

    title_data = [['测试项目总数', total, ''], ['已完成的测试项目', complete_number, "完成率" + str_pre_complete],
                  ['未完成的测试', warning_number, "未完成率" + str_per_not_complete],
                  ['通过的测试数量', pass_number, "通过率" + str_per_pass],
                  ['未通过的测试数量', fail_number, "未通过率" + str_per_fail], [' '], nested_temp_data[9]]

    # 新建一个list存储修改的数据
    final_list = connect_data(nested_tables_data)
    final_list[1] = title_data[6][0]
    final_list = replace_at_symbol(final_list)

    # 指定要搜索的文本和数据列表
    target_text = "6.1.2"
    data_list = [["哈哈哈哈3.51V", "OK"], ["1.51V", "NOK"], ["2.00V", "N/A"], ["2.51V", "NOK"]]

    # # 执行搜索并填充表格
    find_text_with_fill_table(model_input_filepath, target_text, data_list, output_filepath)



print("------ 已处理完成所有文件 ------")
# os.system("pause")
