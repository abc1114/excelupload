import openpyxl
import pymysql
import os
import sys

result = []

def fileCheck(filename):
	return os.path.isfile("./"+filename)


def excel_to_list(filename):
    wb = openpyxl.load_workbook(filename)
    ws = wb.active

    tmp_data = []
    for r in ws.rows:
        id = r[0].value
        version_id = r[1].value
        itemplan_id = r[2].value
        ic_idlc = r[3].value
        ic_dlfc = r[4].value
        ic_idohc = r[5].value
        ic_ohdfe = r[6].value
        ic_ohdfd = r[7].value
        ic_dlvc = r[8].value
        ic_ohdvc = r[9].value
        proq = r[10].value
        proamt_unit = r[11].value
        proamt_acc = r[12].value


        tmp_data.append(id)
        tmp_data.append(version_id)
        tmp_data.append(itemplan_id)
        tmp_data.append(ic_idlc)
        tmp_data.append(ic_dlfc)
        tmp_data.append(ic_idohc)
        tmp_data.append(ic_ohdfe)
        tmp_data.append(ic_ohdfd)
        tmp_data.append(ic_dlvc)
        tmp_data.append(ic_ohdvc)
        tmp_data.append(proq)
        tmp_data.append(proamt_unit)
        tmp_data.append(proamt_acc)

        if len(tmp_data) == 13:
            result.append(tuple(tmp_data))
            tmp_data = []

def main():
    conn = pymysql.connect(
        host="223.194.46.212",
        user="root",
        passwd="12345!",
        database="knowhow",
        charset='utf8'
    )

    while True:
        filename = input("1. 엑셀파일명 입력: ")
        if fileCheck(filename) is True:
            excel_to_list(filename)
            del(result[0])
            print(result)

            cursor = conn.cursor()
            sql = "INSERT INTO cc_costbill1 (id , version_id ,itemplan_id, ic_idlc, ic_dlfc, ic_idohc, ic_ohdfe, " \
                  "ic_ohdfd, ic_dlvc, ic_ohdvc, proq, proamt_unit, proamt_acc) VALUES (%s, %s, %s, %s, %s, %s, %s, " \
                  "%s, %s, %s, %s, %s, %s) "
            cursor.executemany(sql, result)
            conn.commit()
            print("완료")

        else:
            print("파일존재x")
            break
if __name__ == '__main__':
    main()