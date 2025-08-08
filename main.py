import pandas as pd

# pd.set_option('display.max_columns', None)
SHEET_NAMES = ["Chuyenbay", "Thongtinchung", "Hanhkhach", "PNR"]
LABEL_NAMES = [
    {
        "sheet": "Chuyenbay",
        "headers": ["Thông tin chuyến bay"],
        "first_row": [
            "Hãng vận chuyển",
            "Nhãn hiệu quốc tịch",
            "Số chuyến bay",
            "Ngày bay",
            "Nơi đi",
            "Nơi đến",
            "Đường bay",
            "Nơi quá cảnh",
        ],
        "second_row": None,
        "first_skiprow": 2,
        "second_skiprow": None,
        "first_nrows": None,
        "second_nrows": None,
    },
    {
        "sheet": "Thongtinchung",
        "headers": ["Thông tin chung", "Danh sách thành viên tổ bay"],
        "first_row": [
            "Số khách",
            "Nơi đi",
            "Nơi xuất cảnh",
            "Through on same flight",
            "Nơi đến",
            "Nơi nhập cảnh",
            "Through on same flight",
        ],
        "second_row": [
            "Họ và tên",
            "Giới tính",
            "Quốc tịch",
            "Ngày sinh",
            "Số giấy tờ",
            "Loại giấy tờ",
            "Nơi cấp",
            "Ngày hết hạn",
        ],
        "first_skiprow": 2,
        "second_skiprow": 7,
        "first_nrows": 1,
        "second_nrows": None,
    },
    {
        "sheet": "Hanhkhach",
        "headers": ["Danh sách hành khách"],
        "first_row": [
            "Số ghế",
            "Họ và tên",
            "Giới tính",
            "Quốc tịch",
            "Ngày sinh",
            "Loại giấy tờ",
            "Số giấy tờ",
            "Nơi cấp",
            "Quốc gia cư trú",
            "Nơi đi",
            "Nơi đến",
            "Cảng hàng không đầu tiên",
            "Hành lý",
            "Ngày hết hạn",
        ],
        "second_row": None,
        "first_skiprow": 2,
        "second_skiprow": None,
        "first_nrows": None,
        "second_nrows": None,
    },
    {
        "sheet": "PNR",
        "headers": ["Danh sách hành khách PNR"],
        "first_row": [
            "Mã đặt chỗ",
            "Ngày đặt chỗ",
            "Thông tin vé",
            "Tên hành khách",
            "Tên khác",
            "Hành trình bay",
            "Địa chỉ",
            "Điện thoại/Email",
            "Thông tin liên hệ",
            "Số lượng hành khách chung mã đặt chỗ",
            "Mã người đặt chỗ",
            "Số ghế",
            "Thông tin hành lý",
            "Ghi chú",
        ],
        "second_row": None,
        "first_skiprow": 2,
        "second_skiprow": None,
        "first_nrows": None,
        "second_nrows": None,
    },
]


def select_sheet(path, sheet_name):
    sheet = None
    first_result = None
    second_result = None
    for elem in LABEL_NAMES:
        for detail in elem:
            if sheet_name == elem[detail]:
                sheet = elem
    if sheet:
        if sheet["first_row"]:
            if sheet["first_nrows"]:
                first_result = pd.read_excel(
                    path,
                    sheet_name=sheet["sheet"],
                    skiprows=sheet["first_skiprow"],
                    usecols=sheet["first_row"],
                    nrows=sheet["first_nrows"],
                )
            else:
                first_result = pd.read_excel(
                    path,
                    sheet_name=sheet["sheet"],
                    skiprows=sheet["first_skiprow"],
                    usecols=sheet["first_row"],
                )
        if sheet["second_row"]:
            if sheet["second_nrows"]:
                second_result = pd.read_excel(
                    path,
                    sheet_name=sheet["sheet"],
                    skiprows=sheet["second_skiprow"],
                    usecols=sheet["second_row"],
                    nrows=sheet["second_nrows"],
                )
            else:
                second_result = pd.read_excel(
                    path,
                    sheet_name=sheet["sheet"],
                    skiprows=sheet["second_skiprow"],
                    usecols=sheet["second_row"],
                )
    if first_result is not None:
        print(first_result)

    if second_result is not None:
        print(second_result)


def show_detail(df):
    for index, row in df.iterrows():
        for col in df.columns:
            print(f" {col}: {row[col]}")
        print("-" * 50)


for sheet in SHEET_NAMES:
    select_sheet("./files/API_05.04.2023.xls.xlsx", sheet)
