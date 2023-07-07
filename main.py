import pandas as pd
from bs4 import BeautifulSoup
from alive_progress import alive_bar

with alive_bar(100) as bar:
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


    # 读取 HTML 文件
    with open('input.html', 'r', encoding='utf-8') as file:
        html_content = file.read()

    # 使用 BeautifulSoup 解析 HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # 处理多级嵌套表格数据
    tables = soup.find_all('table')
    nested_tables_data = []
    for table in tables:
        table_data = process_nested_table(table)
        nested_tables_data.append(table_data)

    # 创建一个新的 DataFrame 来存储嵌套数据
    df = pd.DataFrame(nested_tables_data)

    # 创建 Excel 文件
    df.to_excel('output.xlsx', index=False)
