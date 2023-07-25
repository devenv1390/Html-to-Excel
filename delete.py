import os

from docx import Document


def remove_tables_from_docx(input_file_path, output_file_path):
    try:
        # 加载输入的docx文件
        doc = Document(input_file_path)

        title_data = find_title_table(doc)
        # print(title_data)
        for target in title_data:
            find_and_delete_table(doc, target)

        # 将修改后的文档保存到输出文件
        doc.save(output_file_path)

        print("Tables removed successfully.")
    except Exception as e:
        print("An error occurred:", e)


# 查找并填充word标题表格
def find_title_table(doc):
    title_data = []

    paragraphs = doc.paragraphs
    all_tables = doc.tables

    for aPara in paragraphs:  # 遍历段落找表格
        # print(aPara.text)
        if "测试项目总览" in aPara.text:
            ele = aPara._p.getnext()
            while ele.tag != '' and ele.tag[-3:] != 'tbl':
                ele = ele.getnext()
            if ele.tag != '':
                for table in all_tables:
                    if table._tbl == ele:
                        for row_index, row in enumerate(table.rows):
                            for col_index, cell in enumerate(row.cells):
                                if col_index == 1 and 1 < row_index < 77:
                                    temp_cell = table.cell(row_index, col_index + 1)
                                    if cell.text != 'AUTOSAR网络管理测试' and cell.text != '物理层测试' \
                                            and cell.text != '数据链路层测试' and cell.text != '网络管理测试' \
                                            and cell.text != '应用层测试':
                                        if temp_cell.text == "N/A":
                                            title_data.append(cell.text)
    return title_data


# 查找并删除文档中的表格
def find_and_delete_table(doc, target_text):
    paragraphs = doc.paragraphs
    all_tables = doc.tables

    for aPara in paragraphs:  # 遍历段落找表格
        # print(aPara.text)
        if target_text in aPara.text:
            ele = aPara._p.getnext()
            while ele.tag != '' and ele.tag[-3:] != 'tbl':
                ele = ele.getnext()
            if ele.tag != '':
                for table in all_tables:
                    if table._tbl == ele:
                        table._element.getparent().remove(table._element)


if __name__ == "__main__":
    docx_file = os.listdir('output')
    # 在所有 docx 文件中遍历处理
    for i, filename in enumerate(docx_file, start=1):
        print(f"正在处理{filename}")
        input_file_path = os.path.join('output', filename)
        output_file_path = os.path.join('output', filename)
        remove_tables_from_docx(input_file_path, output_file_path)
        print("当前文件处理完成")
        print("==================")
