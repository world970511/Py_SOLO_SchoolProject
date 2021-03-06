import pandas as pd
import time
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
    f = craft(fileName)  # 결제 완료&결제 대기인 목록을 분리하여 각각의 파일에 담는다
    second_work(first_work(f))


def craft(f):#결제 누락 항목 제거& 결제 대기 항목은 다른 엑셀 파일로 저장/그룹별 이름 항목으로 저장
    i=1
    file=pd.read_excel(f,sheet_name=0)#가장 첫번째 시트 호출해서 데이터프레임으로 저장
    if '결제완료' not in f:
        #같은 파일을 반복해서 사용할 경우 오류가 생겨 이를 막기위해 사용
        #결제완료(예전에 처리한거) 가 아닌 경우만 다음시트까지 확인, 아닐 경우 첫번째 시트만
        file['그룹명'] = "#선택" + str(0)
        while(True):#다음 시트들을 호출해서 시트명 칼럼 추가
            try:
                call=pd.read_excel(f, sheet_name=i)
                call['그룹명']="#선택_"+str(i)
                file=pd.concat([file,call],ignore_index=True)#호출된 시트의 데이터프레임을 첫번째 시트의 데이터프레임과 합침
                i+=1
            except:
                break#예외가 생길 경우 반복문 멈춤

    file.fillna(" ",inplace=True)#결측지는 "non"으로 대체
    wait=file[file['결제상태'] == '결제 대기']#결제 대기인 사람만 필터링
    file=file[file['결제상태'] == '결제 완료']#결제 완료가 된 사람만 필터링

    final = f.split(".")[0] + '_결제완료.xlsx'

    if not os.path.exists(final):            #'원본파일명_결제완료.xlsx'가 폴더에 없을 경우 새 파일을 만든다
        with pd.ExcelWriter( final, mode='w', engine='openpyxl') as writer:
            file.to_excel(writer, index=False,sheet_name="결제 완료")
        with pd.ExcelWriter( final, mode='a', engine='openpyxl') as writer:
                wait.to_excel(writer, index=False,sheet_name="결제 대기")
        return  final#필터링한 파일명 리턴

    else:            #'원본파일명_결제완료.xlsx'가 폴더에 있을 경우 원본파일명_결제완료_현재시각을 표시한 새 파일을 제공한다
        t = time.ctime(time.time()).replace(" ", "").replace(":", "_")
        final2 = f+"_결제완료_" + t + ".xlsx"
        with pd.ExcelWriter(final2, mode='w', engine='openpyxl') as writer:
            file.to_excel(writer, index=False,sheet_name="결제 완료")
        with pd.ExcelWriter(final2, mode='a', engine='openpyxl') as writer:
                wait.to_excel(writer, index=False, sheet_name="결제 대기")
        return final2 #필터링한 파일명 리턴


def first_work(f):#상품 항목별 개수 카운트
    file = pd.read_excel(f,sheet_name="결제 완료")#파일을 가져온다
    li=[]
    for i,s in enumerate (file.columns.tolist()):#파일에서 "옵션"이라는 문자열이 있는 칼럼명 추출
        if '옵션' in s:
            li.append(str(s))

    count=dict()#개수를 저장할 딕셔너리 생성

    if li != []:#옵션 항목이 있을 경우
        for i in li:#'옵션'이 있는 열만 체크
            li2=file[i].tolist()#열을 리스트로 가져온다
            for i2 in li2:
                if i2 not in count: count[i2]=1#이전에 없던 것이면 딕셔너리에 키와 벨류 추가
                else :count[i2]+=1#이전에 있던 것이면 딕셔너리에 벨류만 추가
        del count[" "]#결측치 제거
        df=pd.DataFrame({'상품명':count.keys(),'개수':count.values()})#딕셔너리를 데이터프레임으로 변경
        with pd.ExcelWriter(f, mode='a', engine='openpyxl') as writer:
            # openpyxl을 엔진으로 이용해서 파일에 상품 개수라는 시트명 붙이고 시트를 추가해 데이터프레임을 저장
            df.to_excel(writer, index=False, sheet_name="옵션 상품 개수")
            return f

    else:#옵션 항목이 없을 경우
        li2=file["그룹명"].tolist()#그룹명 칼럼을 리스트로 가져온다
        for i2 in li2:
            if i2 not in count: count[i2]=1#이전에 없던 것이면 딕셔너리에 키와 벨류 추가
            else :count[i2]+=1
        df = pd.DataFrame({'선물 그룹': count.keys(), '개수': count.values()})  # 딕셔너리를 데이터프레임으로 변경
        with pd.ExcelWriter(f, mode='a', engine='openpyxl') as writer:
            # openpyxl을 엔진으로 이용해서 파일에 모든 선물 그룹별 인원이라는 시트명 붙이고 시트를 추가해 데이터프레임을 저장
            df.to_excel(writer, index=False, sheet_name="선물 그룹별 인원")
        return f


def second_work(f):#입금한 순서대로 정렬
    file=pd.read_excel(f,sheet_name="결제 완료")#엑셀 파일을 가져와 데이터 프레임으로 정리
    type(file)
    try:
        F= file.sort_values(by='후원번호', ascending=True)#가져온 데이터를 '후원번호'(=입금시간)를 기준으로 정렬
    except:print("항목을 확인할 수 없습니다.")#항목을 확인할 수 없어 에러가 날 경우
    else:#에러가 나지 않을 경우 파일에 입금순으로 나열한 파일 입력
            with pd.ExcelWriter(f, mode='a', engine='openpyxl') as writer:
                F.to_excel(writer, index=False, sheet_name="입금순서")
    finally:
        print("\n완료되었습니다. 파일이 있는 폴더를 확인해주세요")


if __name__ == '__main__':
    while(True):
            folderLocation=input("파일이 있는 폴더 위치를 입력하세요(주소로 입력해야 합니다!): ")
            if os.path.isdir(folderLocation)==False:
                print("다시 입력하세요.")
            else:
                call_file(folderLocation)
            ans=str(input("계속 실행하시겠습니까(y/n)?: "))
            if ans=='n':
                break;
            else:
                ans2=str(input("\n폴더 위치를 변경하시겠습니까?(y/n): "))
                if ans2=='y':
                    continue
                else: call_file(folderLocation)

