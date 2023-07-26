import os
import time

from bs4 import BeautifulSoup

# 工具
from tools import find_text_with_fill_table, process_nested_table, process_table, connect_data_type_zero, \
    replace_at_symbol, get_list_from_final, find_text_with_fill_title, connect_data_type_one, \
    special_duel_with_title, delete_enter

# 创建 output 文件夹（如果不存在）
if not os.path.exists('output'):
    os.makedirs('output')

# 计时器
start_all_time = time.time()

try:
    # 获取 input 文件夹下的所有 HTML 文件
    html_files = os.listdir('html_input')
    docx_file = os.listdir('model_input')
    print(f"------ 共有 {len(html_files)} 个文件待处理 ------")

    # 在所有 HTML 文件中遍历处理
    for i, filename in enumerate(html_files, start=1):
        start_time = time.time()
        print("\n------ 正在处理第{i}个文件 ------".format(i=i))
        print(f"------ 文件名为：{filename} ------")

        # 拼接文件路径
        html_input_filepath = os.path.join('html_input', filename)  # HTML 文件放置的文件夹
        model_input_filepath = os.path.join('model_input', docx_file[0])  # 模板文件放置的文件夹（一般只允许放置一个模板文件）
        output_filepath = os.path.join('output', filename.replace('.html', '.docx'))  # 输出的文件夹

        # 读取 HTML 文件
        with open(html_input_filepath, 'rb') as file:
            html_content = file.read()

        # 使用 BeautifulSoup 解析 HTML
        soup = BeautifulSoup(html_content, 'html.parser')

        # HTML 文件的种类
        file_type = 0

        if soup.find('table', class_='HeadingTable') is not None:
            file_type = 0
            if soup.find('big', string='Preparation of Test Module') is not None:
                file_type = 0.5

        elif soup.find('table', class_="MsoNormalTable") is not None:
            file_type = 1

        elif soup.find('h1') is not None:
            file_type = 2

        print(f"------ 此次导入的HTML文件类型为 {file_type} 类型 ------")

        # 根据不同种类的 HTML 采用不同的标题表格和实验内容表格解析方式
        title_data = []
        if file_type == 0 or file_type == 0.5:
            # 查找第一个class为"Heading4"的<div>标签
            heading = soup.find('div', class_='Heading4', string='Test Case Results')

            # 查找紧接着该<div>标签的第一个<table>标签
            table = heading.find_next('table')

            # 用一个title_list存储标题表格的数据
            nested_temp_data = process_table(table)
            title_data = nested_temp_data
            replace_at_symbol(title_data)
            # print(title_data)

        elif file_type == 1:
            # 查找第一个带有指定string属性的<span>标签
            span = soup.find('span', string='Test Case Results')

            # 查找紧接着该<span>标签的第一个<table>标签
            table = span.find_next('table')

            # 用一个title_list存储标题表格的数据
            nested_temp_data = process_table(table)
            title_data = nested_temp_data
            replace_at_symbol(title_data)
            # print(title_data)

        elif file_type == 2:
            # 查找第一个带有指定string属性的<font>标签
            font = soup.find('font', string='3 测试结果目录')

            # 查找紧接着该<span>标签的第一个<table>标签
            table = font.find_next('table')

            title_data = special_duel_with_title(table)

        # 新建一个list存储修改的数据
        final_list = []

        # 处理多级嵌套表格数据
        nested_tables_data = []

        if file_type == 0:
            tables = soup.find_all('table')

            for table in tables:
                table_data = process_nested_table(table)
                nested_tables_data.append(table_data)

            final_list = connect_data_type_zero(nested_tables_data, title_data)
            data_list = get_list_from_final(final_list)

        elif file_type == 1 or file_type == 0.5:
            tables = soup.find_all('table')

            for table in tables:
                table_data = process_table(table)
                nested_tables_data.append(table_data)

            final_list = connect_data_type_one(nested_tables_data, title_data)
            data_list = get_list_from_final(final_list)
            print(data_list)

        elif file_type == 2:
            font = soup.find('font', string='6 物理层', size="5")
            tables = font.find_all_next('table')

            temp_result = []

            result_index = 0
            for table in tables:
                # print(table.text)
                # print("---------------")
                temp_result.append(delete_enter(table))

            for data in title_data:
                if data[0] == '6.1' or data[0] == '6.9' or data[0] == '6.8' or data[0] == '6.2':
                    data[2] = ''
                else:
                    if temp_result[result_index][1] == 'N/T':
                        data[2] = 'N/A'
                    else:
                        data[2] = temp_result[result_index][1]
                    result_index += 1

            # print(title_data)

        # 指定要搜索的文本和数据列表
        # 对标题表格进行搜索和填充
        title_target_text = "测试项目总览"
        find_text_with_fill_title(model_input_filepath, title_target_text, title_data, output_filepath, file_type)

        # 定义目标文本的映射字典
        target_text_mapping = {
            '6.1': '6.1 BUS-OFF下NM状态转换测试',
            '9.8': '9.8 Check Sum行为检测',
            '6.2.3': '6.2.3 内阻-接地断开',
            '7.4': '7.4 100%总线负载下的报文接收',
            '8.2': '8.2 节点首次完成所有周期型数据帧发送的时间',
            '6.8.5': '6.8.5 CAN_H对CAN_L短路',
            '6.8.4': '6.8.4 CAN_H与/或CAN_L对地短路',
            '6.8.3': '6.8.3 CAN_H与/或CAN_L对电源短路',
            '5.17': '5.17 BSM-RMS-NOS-RSS-PBSM-BSM-RMS-NOS-RSS-PBSM-BSM转换测试'
        }

        if file_type == 2:
            continue

        # 执行搜索并填充测试表格
        j = 1  # 测试数据数量迭代器
        print("------ 共有" + len(data_list).__str__() + "个测试数据待处理 ------")
        for data in data_list:
            target_text = data[0][0].split("@")[0]
            if file_type == 1 or file_type == 0.5:
                special_target_result = data[0][0].split("@")[1]
            else:
                special_target_result = ''
            if target_text == '6.4 位上升下降时间':
                target_text = '6.4 位上升/下降时间'
            else:
                temp_target_text = target_text.split(" ")[0]
                target_text = target_text_mapping.get(temp_target_text, target_text)
            temp_data = [_data for _data_index, _data in enumerate(data) if _data_index != 0]
            # print(target_text)
            # print(special_target_result)
            find_text_with_fill_table(output_filepath, target_text, temp_data, output_filepath, j,
                                      special_target_result)
            j += 1

        print(f"------ 完成第{i}个文件的处理 ------")
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"------ 耗时: {elapsed_time:.2f} 秒 ------")
        print(" ")
    print("------ 已处理完成所有文件 ------")
    end_all_time = time.time()
    elapsed_all_time = end_all_time - start_all_time
    print(f"------ 总耗时: {elapsed_all_time:.2f} 秒 ------")
except Exception as e:
    print("发生错误，报错信息为:", e)

os.system("pause")
