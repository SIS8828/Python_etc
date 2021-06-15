
import json
import pandas as pd
import datetime
import pymssql
import sys
import os

# excel 파일 주소
PROCESS_DIR = 'Dir'

# insert 쿼리
insert_sql = """
    insert 쿼리    
        """
# select 쿼리
select_sql = """
        SELECT 쿼리
"""
# delete 쿼리
delete_sql = """
        DELETE 쿼리
"""
update_sql = """
        UPDATE 쿼리
"""


def updateDB():
    conn = pymssql.connect(host='서버주소', database='테이블명', charset='utf8', autocommit=True)

    try:
        cursor = conn.cursor()
        print('update sql 문 : ' + update_sql)
        cursor.execute(update_sql)
        print('update 완료')
        print(datetime.datetime.now())
        print('커밋완료')
    finally:
        conn.close()


def deleteDB():
    conn = pymssql.connect(host='서버주소', database='테이블명', charset='utf8', autocommit=True)
    try:
        cursor = conn.cursor()
        print('delete sql 문 : ' + delete_sql)
        cursor.execute(delete_sql)
        print('delete 완료')
        print(datetime.datetime.now())
        print('커밋완료')
    finally:
        conn.close()


def insertDB(info):
    cursor.execute(insert_sql.format(info))
    #print(insert_sql.format(info))


def main(argv):
    # 결과를 json으로 보내주기
    result = {'code': None}
    # file 이름 지정
    # file_list = ['특정엑셀']
    # PROCESS_DIR 경로에 있는 모든 파일을 읽어온다.
    file_list = os.listdir(PROCESS_DIR)
    # 엑셀 내 공백존재시 사용
    number = 100000000
    dbServer = '서버명'
    conn = pymssql.connect(host=dbServer, database='FRWEB_KR', charset='utf8', autocommit=True)
    cursor = conn.cursor()
    print("작업할 파일 List : " + str(file_list))
    try:
        for i in file_list:
            # 전처리 작업
            print("파일명 : " + str(i))
            ex1 = pd.read_excel(PROCESS_DIR + i, header=3)
            # 필요없는 열 제거
            ex1.drop([''], axis=1,inplace=True)
            # 필요없는 행 제거
            ex1.drop([0, 1], axis=0, inplace=True)
            # 필요없는 데이터 제거
            ex1.drop(ex1.tail(4).index, inplace=True)
            # 전처리된 데이터에 NA값 존재시 0으로 기입
            ex1 = ex1.fillna(0)
            # 엑셀 읽어오기
            ex2 = pd.read_excel(PROCESS_DIR + i, header=None)
            # 특정 데이터만 가져오기
            ex2 = ex2.loc[[3], 4:]
            # 가공된 데이터
            date_list = ex2.values[0].tolist()[5:]
            # 날짜변수
            eow_date = []
            # 엑셀의 날짜데이터 가져오기
            for z in range(0, len(date_list)):
                time = date_list[z].split(" ")[4] + '-' + date_list[z].split(" ")[3] + '-' + \
                       date_list[z].split(" ")[2]
                date = datetime.datetime.strptime(time, "%Y-%m-%d").strftime("%Y-%m-%d")
                eow_date.append(date)
            # 데이터 추출
            eng_code = ex1['데이터1'].tolist()
            gpn = ex1['데이터2'].tolist()
            name = ex1['데이터3'].tolist()
            smu = ex1['데이터4'].tolist()
            rank = ex1['데이터5'].tolist()
            # insert data 추출 작업
            for j in range(0, len(eng_code)):
                # 특정 조건 만족시 넘김
                if eng_code[j][:1] == '조건':
                    continue

                eng_code_data = eng_code[j]
                gpn_data = gpn[j]
                name_data = name[j]
                smu_data = str(smu[j]).split(" ")[0]
                rank_data = rank[j].split(" ")[0]

                # 공란시 ZZ********* 으로 처리
                if gpn_data == 0.0:
                    gpn_data = "ZZ" + str(number)
                    number += 1
                # 가공된 데이터 DB에 넣어주기
                for z in range(6, len(date_list), 7):
                    eow_date_data = eow_date[z]
                    hour1 = ex1[date_list[z - 6]].tolist()[j]
                    hour2 = ex1[date_list[z - 5]].tolist()[j]
                    hour3 = ex1[date_list[z - 4]].tolist()[j]
                    hour4 = ex1[date_list[z - 3]].tolist()[j]
                    hour5 = ex1[date_list[z - 2]].tolist()[j]
                    hour6 = ex1[date_list[z - 1]].tolist()[j]
                    hour7 = ex1[date_list[z]].tolist()[j]
                    sum_hours = hour1 + hour2 + hour3 + hour4 + hour5 + hour6 + hour7
                    # 일주일 시간의 합이 0시간이면 스킵
                    if sum_hours == 0.0:
                        continue
                    else:
                        print(datetime.datetime.now())
                        insertDB(cursor,info)
            print('한개 파일 종료')
        result['code'] = 'success'
    except Exception as e:
        _, _, tb = sys.exc_info()
        print('error line No = {}'.format(tb.tb_lineno))
        print(e)
        result['code'] = 'fail'
        json_str = json.dumps(result, ensure_ascii=False)
        print("<peon>")
        print(json_str)
        print("</peon>")
    finally:
        conn.close()
        if result['code'] == "success":
            print(":::::::::::::::::::::::::::::INSERT 완료되어 Delete 진행:::::::::::::::::::::::::::::::::::::::")
            deleteDB()
            print(":::::::::::::::::::::::::::::Delete 완료되어 Update 진행:::::::::::::::::::::::::::::::::::::::")
            updateDB()
            print(":::::::::::::::::::::::::::::Update 완료:::::::::::::::::::::::::::::::::::::::::::::::::::::")

        print('finish')
        # 결과 json파일로 생성하여 다른프로그램으로 산출
        json_str = json.dumps(result, ensure_ascii=False)
        print("<peon>")
        print(json_str)
        print("</peon>")


if __name__ == "__main__":
    main(sys.argv)
    print("프로세스 종료")