import pandas as pd
from tkinter import filedialog
from tkinter import *
import tkinter.messagebox as tkm
import time
import os


def work(fileName):
    f=craft(os.path.join(os.path.abspath(fileName)))#결제 완료&결제 대기인 목록을 분리하여 각각의 파일에 담는다
    tkm.showwarning("안내", "잠시 기다려주세요, 데이터가 많을수록 시간이 걸립니다.")
    second_work(first_work(f))


def craft(f):#결제 누락 항목 제거& 결제 대기 항목은 다른 엑셀 파일로 저장/그룹별 이름 항목으로 저장
    i=1
    file=pd.read_excel(f,sheet_name=0)#가장 첫번째 시트 호출해서 데이터프레임으로 저장
    if '결제완료' not in f:
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
    except:
        tkm.showwarning("경고", "항목을 확인할 수 없습니다.")#항목을 확인할 수 없어 에러가 날 경우
    else:#에러가 나지 않을 경우 파일에 입금순으로 나열한 파일 입력
            with pd.ExcelWriter(f, mode='a', engine='openpyxl') as writer:
                F.to_excel(writer, index=False, sheet_name="입금순서")
    finally:
        tkm.showwarning("적용 완료", "완료되었습니다. 파일이 있는 폴더를 확인해주세요")


def load():
    root = Tk()
    root.withdraw()
    try:#폴더에서 파일 선택
        f=filedialog.askopenfilename(parent=root, initialdir='\\', title="Open Data files",initialfile='tmp', \
                                    filetypes=(("data files","*.xls;*.xlsx"), ("all files", "*.*")))
        work(f)
    except:tkm.showwarning("경고", "잘못된 파일이거나 선택되지 않았습니다")


def main():
    root = Tk()
    root.title("엑셀자동처리_텀블벅")
    root.geometry("300x100")
    btn = Button(root,text="실행",command=load,padx=24)#실행
    btn2 = Button(root,text="실행 종료", command=root.quit,padx=10)#실행창 종료
    btn.pack()
    btn2.pack()
    root.mainloop()


if __name__ == '__main__':
    main()