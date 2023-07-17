import os

from bs4 import BeautifulSoup

# 工具
from tools import find_text_with_fill_table, process_nested_table, process_table, count_element, connect_data_type_zero, \
    replace_at_symbol, find_text_with_read_table, get_list_from_final, find_text_with_fill_title, connect_data_type_one

# 创建 output 文件夹（如果不存在）
if not os.path.exists('output'):
    os.makedirs('output')

# 获取 input 文件夹下的所有 HTML 文件
html_files = os.listdir('html_input')
docx_file = os.listdir('model_input')
print("------ 共有" + len(html_files).__str__() + "个文件待处理 ------")
i = 1
for filename in html_files:
    print(" ")
    print("------ 正在处理第" + i.__str__() + "个文件 ------")
    # 拼接文件路径
    html_input_filepath = os.path.join('html_input', filename)
    model_input_filepath = os.path.join('model_input', docx_file[0])
    output_filepath = os.path.join('output', filename.replace('.html', '.docx'))

    # 读取 HTML 文件
    with open(html_input_filepath, 'rb') as file:
        html_content = file.read()

    # 使用 BeautifulSoup 解析 HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # HTML 文件的种类
    file_type = 0

    if soup.find('table', class_='HeadingTable') is not None:
        file_type = 0
    elif soup.find('table', class_="MsoNormalTable") is not None:
        file_type = 1

    print("此次导入的HTML文件类型为 " + file_type.__str__() + " 类型")

    # 根据不同种类的 HTML 采用不同的标题表格和实验内容表格解析方式
    title_data = []
    if file_type == 0:
        # 查找第一个class为"Heading4"的<div>标签
        heading = soup.find('div', class_='Heading4', string='Test Case Results')

        # 查找紧接着该<div>标签的第一个<table>标签
        table = heading.find_next('table')

        # 用一个title_list存储标题表格的数据
        nested_temp_data = process_table(table)
        title_data = nested_temp_data
        print(title_data)

    elif file_type == 1:
        # 查找第一个带有指定style属性的<span>标签
        span = soup.find('span', string='Test Case Results')

        # 查找紧接着该<span>标签的第一个<table>标签
        table = span.find_next('table')

        # 用一个title_list存储标题表格的数据
        nested_temp_data = process_table(table)
        title_data = nested_temp_data
        print(title_data)

    # 新建一个list存储修改的数据
    final_list = []

    # 处理多级嵌套表格数据
    tables = soup.find_all('table')
    nested_tables_data = []

    if file_type == 0:

        for table in tables:
            table_data = process_nested_table(table)
            nested_tables_data.append(table_data)

        final_list = connect_data_type_zero(nested_tables_data, title_data)
        final_list = replace_at_symbol(final_list)
        print(final_list)

    elif file_type == 1:

        for table in tables:
            table_data = process_table(table)
            nested_tables_data.append(table_data)

        final_list = connect_data_type_one(nested_tables_data, title_data)
        print(final_list)

    # # 算通过、不通过、没测试
    # total = len(nested_temp_data[9])
    # warning_number = count_element(nested_temp_data[9], 'warning')
    # fail_number = count_element(nested_temp_data[9], 'fail')
    # pass_number = total - fail_number - warning_number
    # complete_number = pass_number + fail_number
    #
    # per_complete = round(complete_number / total, 2) * 100
    # per_not_complete = 100 - per_complete
    # per_pass = round(pass_number / total, 2) * 100
    # per_fail = round(fail_number / total, 2) * 100
    #
    # str_pre_complete = per_complete.__str__() + "%"
    # str_per_not_complete = per_not_complete.__str__() + "%"
    # str_per_pass = per_pass.__str__() + "%"
    # str_per_fail = per_fail.__str__() + "%"

    # nested_tables_data = []
    # nested_temp_data = []
    #
    # data_list = get_list_from_final(final_list)
    #
    # # 指定要搜索的文本和数据列表
    # title_target_text = "测试项目总览"
    # find_text_with_fill_title(model_input_filepath, title_target_text, title_data, output_filepath)
    #
    # # 执行搜索并填充表格
    # j = 1
    # print("------ 共有" + len(data_list).__str__() + "个测试数据待处理 ------")
    # for data in data_list:
    #     target_text = data[0][0].split("@")[0]
    #     temp_data = []
    #     for _data_index, _data in enumerate(data):
    #         if _data_index == 0:
    #             continue
    #         else:
    #             temp_data.append(_data)
    #     find_text_with_fill_table(output_filepath, target_text, temp_data, output_filepath, j)
    #     j += 1
    #
    # print("------ 完成第" + i.__str__() + "个文件的处理 ------")
    # i += 1
    # print(" ")

print("------ 已处理完成所有文件 ------")
# os.system("pause")
