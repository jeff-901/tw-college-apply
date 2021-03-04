from selenium import webdriver
from bs4 import BeautifulSoup
import time
from selenium.webdriver.support.ui import Select
import re
import csv
import random
import json
import requests
import sys

attr = [
    "學校",
    "學系",
    "科目",
    "科目檢定",
    "篩選倍率",
    "學測成績採計方式",
    "一階佔甄選總成績比例",
    "指定項目",
    "指定項目檢定",
    "佔甄選總成績比例",
    "甄選總成績同分參酌之順序",
    "指定項目甄試費",
    "寄發(或公告)指定項目甄試通知",
    "繳交資料截止",
    "指定項目甄試日期",
    "榜示",
    "總成績複查截止",
    "同級分(分數)超額篩選方式",
    "招生名額",
    "性別要求",
    "預計甄試人數",
    "原住民外加名額",
    "離島外加名額",
    "願景計畫外加名額",
    "校系代碼",
]


def crawl_data(_id, year):

    r = requests.get(
        f"https://www.cac.edu.tw/apply{year}/system/{year}_aColQry4qy_forapply_o5wp6ju/html/{_id}.htm?v=1.0"
    )

    r.encoding = "utf-8"
    # print(r.text)
    data_dict = {}
    for i in range(2, 10):
        data_dict[attr[i]] = []
    # print(data_dict)
    soup = BeautifulSoup(r.text, "lxml")
    tables = soup.select("table")
    table1 = tables[0]
    col_dep = table1.select("p")[0].font.text
    # print(col_dep.split("\n"))
    data_dict["學校"] = col_dep.split("\n")[0]
    data_dict["學系"] = col_dep.split("\n")[1].strip()
    table2 = tables[1]
    trs = table2.select("tr")
    for i in range(len(trs)):
        # trs = table2.select("tr")
        tr = trs[i]
        tds = tr.select("td")
        # print(len(tds))
        k = 0
        for j in range(len(tds)):
            # print(tds[j].select("font")[0].text)

            if i == 0 and j == len(tds) - 1:
                data_dict["甄選總成績同分參酌之順序"] = tds[j].select("font")[0].text
            elif i == 0 and j == 6:
                data_dict["一階佔甄選總成績比例"] = tds[j].select("font")[0].text
            elif j == 1:
                data_dict[tds[j - 1].select("font")[0].text] = (
                    tds[j].select("font")[0].text
                )
            elif j > 1 and j < 6:
                data_dict[attr[j]].append(tds[j].select("font")[0].text)
            elif j >= 6 and k < 3:
                if tds[j].select("font")[0].text.strip() != "":
                    data_dict[attr[k + 7]].append(tds[j].select("font")[0].text)
                k += 1
        # print("--------------------------------------------")
        # print(data_dict)
    # print(len(trs))
    table3 = tables[2]
    trs = table3.select("tr")
    for i in range(len(trs) - 1):
        tr = trs[i]
        tds = tr.select("td")
        text = tds[1].select("font")[0].text
        if i == len(trs) - 2:
            text = text.replace("\u3000", " ")
        data_dict[tds[0].select("font")[0].text] = text
    # print(data_dict)
    for key in data_dict:
        if (key) not in attr:
            print(data_dict["學校"])
            print(data_dict["學系"])
            print(key)
    return data_dict


year = "110"
browser = webdriver.Chrome()
all_data = []
browser.get(
    f"https://www.cac.edu.tw/apply{year}/system/{year}_aColQry4qy_forapply_o5wp6ju/TotalGsdShow.htm"
)

table = browser.find_element_by_xpath("/html/body/center/h3/div/table/tbody")
colleges = table.find_elements_by_tag_name("td")

for i in range(len(colleges)):
    colleges[i].find_element_by_tag_name("a").click()
    try:
        departments = browser.find_elements_by_xpath("/html/body/table/tbody/tr")
        for j in range(2, len(departments) + 1):
            _id = browser.find_element_by_xpath(
                f"/html/body/table/tbody/tr[{j}]/td/b/font"
            ).text
            all_data.append(crawl_data(f"{year}_" + _id[1:-1], year))
            print("total: " + str(len(all_data)), end="\r")
        # time.sleep(2)
        back = browser.find_element_by_xpath(
            "/html/body/center[2]/table/tbody/tr/td[3]/input"
        )
        back.click()
        # time.sleep(2)
        table = browser.find_element_by_xpath("/html/body/center/h3/div/table/tbody")
        colleges = table.find_elements_by_tag_name("td")
    except:
        print("selenium go wrong")
        with open(
            f"apply_first_{len(all_data)}.csv", "w", encoding="utf-8", newline=""
        ) as f:
            writer = csv.DictWriter(f, fieldnames=attr)
            writer.writeheader()
            for data in all_data:
                try:
                    writer.writerow(data)
                except:
                    print(data)

with open(f"apply_{year}.csv", "w", encoding="utf-8", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=attr)
    writer.writeheader()
    for data in all_data:
        try:
            writer.writerow(data)
        except:
            print(data)