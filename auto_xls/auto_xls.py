import pandas as pd
import random as r
import xlwings as xw
import openpyxl as ox
import os

def call_file(folderlocation):
    os.chdir(folderLocation)  # 파일이 있는 곳으로 위치 변경
    while(True):
        fileName= str(input("파일명을 확장자명과 같이 입력해주세요(ex:a.xls): "))#파일명 입력
        if os.path.exists(fileName)==False: #파일이 존재하지 않을 경우 다시 입력을 요구한다
            print("다시 입력하세요")
            continue
        else:
            work(fileName)
            break

def work(fileName):
    f=craft(fileName)#결제 완료&결제 대기인 사람을 분리하여 파일에 담는다
    while(True):
        print("<기능> \n\n1)상품 항목별 개수 카운트\n2)결제일&시간 순서로 나열\n")
        n= int(input("원하는 기능을 선택해주세요: "))
        if n==1:
            first_work(f)
            break
        elif n==2:
            second_work(f)
            break
        else:
            print("다시 입력해주세요")
            continue

def craft(f):
    file=pd.read_excel(f, sheet_name=None)#전체 시트를 읽어 df저장
    print(file)
    n_f = pd.concat(file, ignore_index=True)#전체 시트를 하나의 시트로 합침

    condition_c = (n_f['결제상태'] == '결제 완료')#결제 상태-'결제 완료' 일때
    condition_d = (n_f['결제상태'] == '결제 대기')#결제 상태-'결제 대기' 일때
    n_f=n_f[condition_c]#결제 완료가 된 사람만 필터링
    wait=n_f[condition_d]#결제 대기인 사람만 필터링

    final="결제완료_"+str(r.randint(0,10000000))+".xlsx"
    delay = "결제대기_" + str(r.randint(0, 10000000)) + ".xlsx"

    n_f.to_excel(final, index=False)#완료인 사람만 모아 새로운 파일로 저장(인덱스는 추가하지 않는다)
    wait.to_excel(delay, index=False)#대기인 사람만 모아 새로운 파일로 저장(인덱스는 추가하지 않는다)

    return final#필터링한 파일명 리턴

def first_work(f):
    print("작성중")

def second_work(f):
    print("작성중")

folderLocation=input("파일이 있는 폴더 위치를 입력하세요(주소로 입력해야 합니다!): ")
call_file(folderLocation)#파일 위치 입력





