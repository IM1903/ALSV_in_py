import base64
import io

import openpyxl
import pandas as pd
from openpyxl_image_loader import SheetImageLoader


def image2base64(image):
    buffetStream = io.BytesIO()
    image.save(buffetStream, format="PNG")
    buffetStream.seek(0)
    img_bytes = buffetStream.read()
    img_str = base64.b64encode(img_bytes)
    return img_str


def handle_dataframe(dataframe_file_path: str) -> list:
    excel_file = pd.ExcelFile(dataframe_file_path)
    work_book = pd.read_excel(excel_file, sheet_name=None)
    print("正在处理表格中的图像")
    sheetname_array = []
    dataframe_array = []

    for sheet_name, sheet_data in work_book.items():
        sheetname_array.append(sheet_name)
    # By default, it appears that pandas does not read images, as it uses only openpyxl to read
    # the file.  As a result we need to load into memory the dataframe and explicitly load in
    # the images, and then convert all of this to HTML and put it back into the normal
    # dataframe, ready for use.
    for sheetname in sheetname_array:
        print("图像处理: 正在搜索 {0} 页".format(sheetname))
        raw_workbook = openpyxl.load_workbook(dataframe_file_path)
        raw_sheet = raw_workbook[sheetname]
        image_loader = SheetImageLoader
        image_loader._images = {}
        loader = image_loader(raw_sheet)
        dataframe = pd.read_excel(dataframe_file_path, sheet_name=sheetname)
        for row_index, row in dataframe.iterrows():
            for column_index, cell_data in enumerate(row):
                cell_coordinate = raw_sheet.cell(int(row_index) + 2, column_index + 1).coordinate
                if loader.image_in(cell_coordinate):
                    print("'{0}' 页 {1} 单元格中找到一张图片".format(sheetname, cell_coordinate))
                    fetched_image = loader.get(cell_coordinate)
                    # print(str(fetched_image))
                    print("正在保存图像...")
                    # print(image_to_b64str)

                    # dataframe.iat[row_index, column_index] = ' <img src="data:image/png;base64,' + image2base64(fetched_image).decode('utf-8') + f'" alt="{cell_coordinate}" /> '
                    dataframe.iat[row_index, column_index] = "b64:" + image2base64(
                        fetched_image).decode('utf-8')

        dataframe_array.append(dataframe)

    return dataframe_array
