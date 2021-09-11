import json  # import json module
import pandas as pd
import datetime
import xlwings as xw
import numpy as np
import os
import sys


def is_nan(x):
    return (x != x)


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def excel_trans_print(files, output_path, p_var, progress_bar):
    print("files : ", files)
    with open(resource_path("./mapping/toggle.json"), 'r', encoding='UTF-8') as wms_file:
        wms_data = json.load(wms_file)

    new_df = pd.DataFrame(columns=wms_data["header"])

    count = 0
    for i, file in enumerate(files):
        # 수정
        print("file : ", file)
        if file["type"] == "쿠팡":
            json_file_name = "./mapping/coupang.json"
        elif file["type"] == "11번가":
            json_file_name = "./mapping/eleven.json"
        elif file["type"] == "위메프":
            json_file_name = "./mapping/wemakeprice.json"
        elif file["type"] == "네이버":
            json_file_name = "./mapping/naver.json"
        elif file["type"] == "티몬":
            json_file_name = "./mapping/tmon.json"
        elif file["type"] == "롯데온":
            json_file_name = "./mapping/lotte.json"
        elif file["type"] == "ESM+":
            json_file_name = "./mapping/esm+.json"

        with open(resource_path(json_file_name), 'r', encoding='UTF-8') as json_file:
            json_data = json.load(json_file)
        print("1111111111111111")
        if file["type"] == "11번가":
            df = pd.read_excel(file["file"], 'Sheet1', engine="xlrd", skiprows=1)
        else:
            # openpyxl 사용 버젼
            # df = pd.read_excel(file["file"], 'Sheet1', engine="openpyxl")

            # xlwings 사용 버젼, 이건 PC의 엑셀 사용하게 된다.
            app = xw.App(visible=False)
            wb = xw.Book(file["file"])
            sheet = wb.sheets[0]
            df = sheet.range('A1').options(pd.DataFrame, index=False, expand='table').value
            app.quit()
        df = df.replace(np.nan, '', regex=True)
        print("2222222222222222")
        print(df)
        print("new_df  :", new_df)
        for j, row in df.iterrows():
            print("row : ", row)
            tmp_row = []
            for index, header in enumerate(wms_data["header"]):
                print("header  :", header)
                if "literal" in json_data["mapping"][header]:
                    print("aaaaa")
                    # print("literal : ", json_data["mapping"][header]["literal"])
                    item = json_data["mapping"][header]["literal"]
                    tmp_row.append(item)
                elif "date" in json_data["mapping"][header]:
                    print("aaaaa")
                    print("literal : ", datetime.date.today())
                    today = datetime.date.today()
                    d1 = today.strftime("%Y-%m-%d")
                    tmp_row.append(d1)
                elif "or" in json_data["mapping"][header]:
                    set_flag = False
                    for d in json_data["mapping"][header]["or"]:
                        print("d : ", row[d])
                        print("is_nan(row[d]) : ", is_nan(row[d]))
                        print("str(row[d]).isspace() : ", str(row[d]).isspace())
                        print("len(row[d]) : ", len(row[d]))
                        if len(str(row[d])) != 0 and set_flag is False:
                            print("DDDDDDDDDDDDDDDDDDD: ", row[d])
                            item = row[d]
                            tmp_row.append(item)
                            set_flag = True
                elif json_data["mapping"][header]:
                    print("bbbbb")
                    # print("data : ", row[json_data["mapping"][header]])
                    if header == "주문번호":
                        item = str(row[json_data["mapping"][header]])
                    else :
                        item = row[json_data["mapping"][header]]
                    tmp_row.append(item)
                else:
                    print("ccccc")
                    # print("Non:", json_data["mapping"][header])
                    item = json_data["mapping"][header]
                    tmp_row.append(item)
            print("item : ", tmp_row)
            new_df.loc[count] = tmp_row
            count = count + 1
            progress = (i + 1) / len(files) * 100
            p_var.set(progress)
            progress_bar.update()

    print(json_data)
    print("new_df : ", new_df)
    today = datetime.datetime.now()
    now_date = today.strftime("%Y%m%d")
    now_time = today.strftime('%H%M%S')
    print(now_time)
    out_file_name = output_path + "/toggle_output" + "_" + now_date + "_" + now_time + ".xlsx"

    print(out_file_name)

    # new_df.to_excel(out_file_name, engine='openpyxl', sheet_name='Sheet1', header=True, index=False)
    writer = pd.ExcelWriter(out_file_name, engine="xlsxwriter")
    new_df.to_excel(writer, sheet_name='Sheet1', header=True, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]
    #number_format = workbook.add_format({'num_format': '###'})
    worksheet.set_column("D:D", 20)
    worksheet.set_column("G:G", 20)
    writer.close()
