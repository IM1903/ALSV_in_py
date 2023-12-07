#!/Users/otaku/Documents/PythonProjects/Playground
# Filename: ALWV.py

import dominate
import os
import pandas as pd
import loadImageFromCell
import subprocess
import argparse
from dominate.tags import *

os.chdir(os.path.dirname(__file__))


def sheet_renderer(file_path):
    ioPath = file_path
    excel_file = pd.ExcelFile(ioPath)
    work_book = pd.read_excel(excel_file, sheet_name=None)

    multiTabCounter = 1
    sheet_name_array = []
    sheet_data_array = []
    # 遍历取出表头名称

    for sheet_name, sheet_data in work_book.items():
        sheet_name_array.append(sheet_name)
        print("正在解析: {0}".format(sheet_name))
    sheet_data_array = loadImageFromCell.handle_dataframe(ioPath)

    html = dominate.document(title="testData")
    with html:
        with div(id="tabs"):
            with ul():
                for sheetName in sheet_name_array:
                    # 遍历表页清单, 创建Tabs多标签头
                    li(a(sheetName, href="#tab-{0}".format(sheet_name_array.index(sheetName) + 1),
                         onclick='tabSwitch()'))
                # 遍历表页清单, 创建Tabs多标签页
            for sheetData in sheet_data_array:
                sheet_header = sheetData.columns[0:].tolist()
                # print(sheet_header)
                with div(id="tab-{0}".format(multiTabCounter)):
                    with div(cls="col-md-12 ms-sm-auto px-md-4"):
                        with div(cls="table-responsive small"):
                            with table(cls="table table-striped table-sm",
                                       id="dataTable-{0}".format(multiTabCounter)):
                                with thead():
                                    tableHeader = tr()
                                    for item in sheet_header:
                                        tableHeader.add(th(item))
                                tableContent = tbody()
                                for index, row in sheetData.iterrows():
                                    tableContentRow = tr()
                                    data = row.fillna(0).to_dict()
                                    for key, value in data.items():
                                        if data.get(key) != 0:
                                            # <img src=data:image/png; base64,{data} alt=cell_coordinate />
                                            if str(data.get(key)).startswith("b64:"):
                                                print("found pic in cell")
                                                tableContentRow.add(
                                                    td(img(src="data:image/png; base64,{0}".format(
                                                        data.get(key).replace("b64:", "")))))
                                            else:
                                                tableContentRow.add(td(data.get(key)))
                                        else:
                                            tableContentRow.add(td(' '))
                                    tableContent.add(tableContentRow)
                multiTabCounter += 1
            multiTabCounter = 1
    with open("doc.html", 'w') as file:
        file.write(html.render())


def main():
    parser = argparse.ArgumentParser(description='接收工作表并处理,最后生成可视化页面便于分析对比.')
    parser.add_argument('-p', type=str, help='简短化的PATH参数')
    parser.add_argument('-path', type=str, help='工作表的路径')
    try:
        args = parser.parse_args()
        sheet_renderer(args.path)
        subprocess.run(['python3', 'server.py'])
    except FileNotFoundError:
        print('路径错误')


if __name__ == '__main__':
    main()
