import pandas as pd
import time
import ast
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
    f=craft(fileName)#결제 완료&결제 대기인 목록을 분리하여 각각의 파일에 담는다

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

def craft(f):#결제 누락 항목 제거& 결제 대기 항목은 다른 엑셀 파일로 저장/그룹별 이름 항목으로 저장
    i=1
    file=pd.read_excel(f,sheet_name=0)#가장 첫번째 시트 호출해서 데이터프레임으로 저장
    file['그룹명'] = "#선택" + str(0)
    while(True):#다음 시트들을 호출해서 시트명 칼럼 추가
        try:
          call=pd.read_excel(f, sheet_name=i)
          call['그룹명']="#선택"+str(i)
          file=pd.concat([file,call],ignore_index=True)#호출된 시트의 데이터프레임을 첫번째 시트의 데이터프레임과 합침
          i+=1
        except:
            break#예외가 생길 경우 반복문 멈춤

    file.fillna(0,inplace=True)#결측지는 0으로 대체
    condition_c = (file['결제상태'] == '결제 완료')#결제 상태-'결제 완료' 일때
    condition_w = (file['결제상태'] == '결제 대기')#결제 상태-'결제 대기' 일때
    wait=file[condition_w]#결제 대기인 사람만 필터링
    file=file[condition_c]#결제 완료가 된 사람만 필터링


    t=time.ctime(time.time()).replace(" ","").replace(":","")
    final="결제완료_"+t+".xlsx"
    delay = "결제대기_" +t+".xlsx"

    file.to_excel(final, index=False)#완료인 사람만 모아 새로운 파일로 저장(인덱스는 추가하지 않는다)
    wait.to_excel(delay, index=False)#대기인 사람만 모아 새로운 파일로 저장(인덱스는 추가하지 않는다)

    return final#필터링한 파일명 리턴



def first_work(f):#상품 항목별 개수 카운트
    file = pd.read_excel(f)#파일을 가져온다
    li=[]
    for i,s in enumerate (file.columns.tolist()):#파일에서 "옵션"이라는 문자열이 있는 칼럼명 추출
        if '옵션' in s:
            li.append(str(s))

    count=dict()#개수를 저장할 딕셔너리 생성
    for i in li:#'옵션'이 있는 열만 체크
        li2=file[i].tolist()#열을 리스트로 가져온다
        for i2 in li2:
            if i2 not in count: count[i2]=1#이전에 없던 것이면 딕셔너리에 키와 벨류 추가
            else :count[i2]+=1#이전에 있던 것이면 딕셔너리에 벨류만 추가

    del count[0]#결측치 제거
    df=pd.DataFrame({'상품명':count.keys(),'개수':count.values()})#딕셔너리를 데이터프레임으로 변경

    with pd.ExcelWriter(f,mode='a', engine='openpyxl') as writer:
        #openpyxl을 엔진으로 이용해서 파일에 시트를 추가하고 상품 개수라는 시트명 붙임 데이터프레임 저장
        df.to_excel(writer, index=False, sheet_name="상품 개수")


def second_work(f):#결제일&결제 시간 데이터가 주어졌을 경우 입금한 순서대로 정렬하는 기능 구현
    print("작성중")



folderLocation=input("파일이 있는 폴더 위치를 입력하세요(주소로 입력해야 합니다!): ")
#try:
call_file(folderLocation)#파일 위치 입력
#except:
    #folderLocation = input("파일이 있는 폴더 위치를 다시 입력하세요(주소로 입력해야 합니다!): ")
    #call_file(folderLocation)  # 파일 위치 입력




