import requests
from datetime import datetime, timedelta
import pandas as pd
import os

# pip install requests pandas xlrd xlsxwriter

cookies = {
    "user_lang": "en-GB",
    "JSESSIONID": "6BB03A85706AE003F5AAA32F627505D9",
}


def download(country_ids, start, end):
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Cache-Control": "max-age=0",
        "Connection": "keep-alive",
        "Content-Type": "application/x-www-form-urlencoded",
        # 'Cookie': 'user_lang=en-GB; JSESSIONID=B9E3AA348C78E7ACDD24944253431D51',
        "DNT": "1",
        "Origin": "https://eudcs.byd.com",
        "Referer": "https://eudcs.byd.com/main.html",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
        "sec-ch-ua": '"Chromium";v="128", "Not;A=Brand";v="24", "Google Chrome";v="128"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
    }

    data = {
        "CREATE_DATE_gte": start,
        "CREATE_DATE_lte": end,
        "COUNTRY_IDS": country_ids,
        "VIN": "",
        "SERIES": "",
        "CUST_NAME": "",
        "PURCHASE_DATE_gte": "",
        "PURCHASE_DATE_lte": "",
        "PRODUCT_DATE_gte": "",
        "PRODUCT_DATE_lte": "",
    }

    expCommercialVehicle = requests.post(
        "https://eudcs.byd.com/servicemng/customerquery/CommercialVehicleCustomerAction/expCommercialVehicle.json",
        cookies=cookies,
        headers=headers,
        data=data,
    )

    if expCommercialVehicle.status_code == 200:
        # 获取文件名
        filename = (
            expCommercialVehicle.headers.get("Content-disposition")
            .split("filename=")[1]
            .strip('"')
        )

        # 保存文件
        with open(filename, "wb") as f:
            f.write(expCommercialVehicle.content)
        print(f"文件已保存为: {filename}")
    else:
        print(f"请求失败，状态码: {expCommercialVehicle.status_code}")


def generate_weekly_date_pairs(start: str, end: str):
    # 将字符串转换为日期对象
    start_date = datetime.strptime(start, "%Y-%m-%d")
    end_date = datetime.strptime(end, "%Y-%m-%d")

    # 初始化结果列表
    date_pairs = []
    # 以周为间隔生成日期对
    current_start_date = start_date
    delta = 30 #DMS单次导出极限约为5W个VIN
    while current_start_date < end_date:
        # 计算结束日期
        current_end_date = current_start_date + timedelta(delta)  # delta为30天, 大于该值会出现数据丢失，原因不明
        # 检查结束日期是否超出范围
        if current_end_date > end_date:
            current_end_date = end_date
        # 添加日期对到列表
        date_pairs.append(
            (
                current_start_date.strftime("%Y-%m-%d"),
                current_end_date.strftime("%Y-%m-%d"),
            )
        )
        # 更新开始日期
        current_start_date += timedelta(delta + 1)
    return date_pairs


def merge_excel_files(output_file):
    # 获取当前工作目录
    current_directory = os.getcwd()

    # 列出所有 .xlsx 文件
    file_list = [
        file for file in os.listdir(current_directory) if file.endswith(".xls")
    ]

    print(file_list)

    # 创建一个空的 DataFrame 用于存储合并后的数据
    combined_df = pd.DataFrame()

    for file in file_list:
        # 读取 Excel 文件
        df = pd.read_excel(file)

        # 如果 combined_df 为空，直接赋值
        if combined_df.empty:
            combined_df = df
        else:
            # 合并数据，忽略重复的列名（即第一行）
            combined_df = pd.concat([combined_df, df], ignore_index=True)

    # 保存合并后的 DataFrame 到新的 Excel 文件
    combined_df.to_excel(output_file, index=False)
    print(f"合并完成，输出文件: {output_file}")


country_ids = ["500019","500026","500042","500043","500045","500055","500056","500064","500072","500221","500222","500090","500168","500226","500234","500237",'500171']
start = "2022-01-01"
end = "2024-11-22"
for country in country_ids:
    for s, e in generate_weekly_date_pairs(start, end):
        print(country, s, e)
        download(country, s, e)

merge_excel_files("result.xlsx")
