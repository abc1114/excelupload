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
        ei_elc = r[3].value
        ei_einv = r[4].value
        ei_erm = r[5].value
        ei_eoh = r[6].value
        ei_sum = r[7].value
        bi_blc = r[8].value
        bi_brm = r[9].value
        bi_idohc = r[10].value
        bi_dohc = r[11].value
        bi_sum = r[12].value
        uc_ulc_sum = r[13].value
        uc_dlc = r[14].value
        uc_idlc = r[15].value
        uc_srw = r[16].value
        uc_idohc = r[17].value
        uc_aoh_sum = r[18].value
        uc_dohc = r[19].value
        uc_price = r[20].value
        uc_tsum = r[21].value
        uc_tudc = r[22].value
        ic_idlc = r[23].value
        ic_alc_sum = r[24].value
        ic_dlfc = r[25].value
        ic_dlvc = r[26].value
        ic_arm = r[27].value
        ic_idohc = r[28].value
        ic_aoh_sum = r[29].value
        ic_ohdfe = r[30].value
        ic_ohdfd = r[31].value
        ic_ohdvc = r[32].value
        ic_sum = r[33].value
        proamt_unit = r[34].value
        proamt_acc = r[35].value
        st_unit = r[36].value
        st_acc = r[37].value
        proq = r[38].value
        ra_rm = r[39].value


        tmp_data.append(id)
        tmp_data.append(version_id)
        tmp_data.append(str(itemplan_id))
        tmp_data.append(ei_elc)
        tmp_data.append(ei_einv)
        tmp_data.append(ei_erm)
        tmp_data.append(ei_eoh)
        tmp_data.append(ei_sum)
        tmp_data.append(bi_blc)
        tmp_data.append(bi_brm) #10
        tmp_data.append(bi_idohc)
        tmp_data.append(bi_dohc)
        tmp_data.append(bi_sum)
        tmp_data.append(uc_ulc_sum)
        tmp_data.append(uc_dlc)
        tmp_data.append(uc_idlc)
        tmp_data.append(uc_srw)
        tmp_data.append(uc_idohc)
        tmp_data.append(uc_aoh_sum)
        tmp_data.append(uc_dohc) #20
        tmp_data.append(uc_price)
        tmp_data.append(uc_tsum)
        tmp_data.append(uc_tudc)
        tmp_data.append(ic_idlc)
        tmp_data.append(ic_alc_sum)
        tmp_data.append(ic_dlfc)
        tmp_data.append(ic_dlvc)
        tmp_data.append(ic_arm)
        tmp_data.append(ic_idohc)
        tmp_data.append(ic_aoh_sum) #30
        tmp_data.append(ic_ohdfe)
        tmp_data.append(ic_ohdfd)
        tmp_data.append(ic_ohdvc)
        tmp_data.append(ic_sum)
        tmp_data.append(proamt_unit)
        tmp_data.append(proamt_acc)
        tmp_data.append(st_unit)
        tmp_data.append(st_acc)
        tmp_data.append(proq)
        tmp_data.append(ra_rm)

        if len(tmp_data) == 40:
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
            sql = "INSERT INTO cc_costbill (id , version_id ,itemplan_id, ei_elc, ei_einv, ei_erm, ei_eoh, ei_sum, " \
                  "bi_blc, bi_brm, bi_idohc, bi_dohc, bi_sum, uc_ulc_sum, uc_dlc, uc_idlc, uc_srw, uc_idohc, " \
                  "uc_aoh_sum, uc_dohc, uc_price, uc_tsum, uc_tudc, ic_idlc, ic_alc_sum, ic_dlfc, ic_dlvc, ic_arm, " \
                  "ic_idohc, ic_aoh_sum, ic_ohdfe, ic_ohdfd, ic_ohdvc, ic_sum, proamt_unit, proamt_acc, st_unit, " \
                  "st_acc, proq, ra_rm) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, " \
                  "%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) "
            cursor.executemany(sql, result)
            conn.commit()
            print("완료")

        else:
            print("파일존재x")
            break
if __name__ == '__main__':
    main()