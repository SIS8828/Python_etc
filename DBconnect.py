import datetime
import pandas as pd
import pymssql
import sys
import traceback
import time

dir = "시작주소"

sql = """
조회할 쿼리
"""
def readExcel(name):
    excel = pd.read_excel(dir + name)
    EngList = excel["Engagement"].tolist()

    return EngList

def select_sql(cursor, i):
    cursor.execute(sql.format(i))
    result = cursor.fetchone()
    return result

def main(argv):
    # DB Connection
    conn = pymssql.connect(host='서버명', database='DB명', charset='utf8', autocommit=True)
    # List Data
    EngList = readExcel("엑셀명")
    cursor = conn.cursor()
    eng_list = []
    t = 0
    try:
        for i in EngList:
            # 정보 수집하기
            eng_info = pd.DataFrame({'ENG_CODE': EngList[t],
                                     'E_NM': [select_sql(cursor, i)[0]]})
            # 수집된 정보 list에 적재
            eng_list.append(eng_info)
            t += 1
        # list에 담긴 정보 DataFrame형식으로 변환
        output = pd.concat(eng_list)
        # 추출된 정보를 Excel로 산출하기
        output.to_excel('엑셀명',
                        index=False,
                        header=True,
                        startrow=0)
    except Exception as e:
        # 에러발생시 처리를 위한 구문
        print("에러 발생")
        print("Exceoption : ", e)
        print(traceback.format_exc())
    finally:
        conn.close()

if __name__ == "__main__":
    print(datetime.datetime.now())
    main(sys.argv)
    print(datetime.datetime.now())
    print("프로세스 종료")
