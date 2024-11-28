import pandas as pd     # 데이터 처리
import os               # 파일 읽어오기
import re               # 데이터 정규화 
from PyQt5.QtWidgets import QLabel

def collect(folder_path, locations, out_file, state_label):
    
    # 폴더 이름
    # folder_path = "excel_files"
    # 폴더 안의 모든 파일 읽어오기
    try:
        all_files = os.listdir(folder_path)
    except:
        state_label.setText("folder 이름 틀렸어!!")
        return

    # 정해진 행열 입력하는 부분
    # locations = []
    # locations.append(('A',1))
    # locations.append(('B',5))
    # locations.append(('A',2))
    # locations.append(('C',6))
    # locations.append(('A',3))

    # 합쳐질 데이터
    combined_df = pd.DataFrame()

    # 파일 마다 읽어오기
    for file in all_files:
        df = pd.read_excel(folder_path+"/"+file, header=None)
        df.columns = [chr(65 + i) for i in range(len(df.columns))]
        print(df)
        print("----------------------------")

        try:
            values = [df.loc[loc[1], loc[0]] for loc in locations]
        except:
            state_label.setText("행열 위치 잘써라!!")
            return
        new_df = pd.DataFrame([values])
        print(new_df)
        print("----------------------------")

        combined_df = pd.concat([combined_df, new_df], axis=0, ignore_index=True)


    print(combined_df)

    # 결과 저장 (덮어쓰기 됨 주의)
    # out_file = "collected.xlsx"

    try:
        combined_df.to_excel(out_file, index=False, header=None)
    except:
        state_label.setText("저장될 파일 이름 잘못됨")
        return
    
    state_label.setText(f"{out_file}에 저장됨")


def parse_input(input_text):
    locations = []
    
    # 소문자는 대문자로
    input_text = input_text.upper() 

    # 정규식 패턴: 문자+숫자 조합
    pattern = r"([A-Z]+)(\d+)"
    
    # 모든 매칭된 그룹 찾기
    matches = re.findall(pattern, input_text)
    
    # 결과를 locations 리스트에 추가
    for match in matches:
        letter = match[0]  # 문자 부분 (예: 'A', 'B', 'AB')
        number = int(match[1])  # 숫자 부분 (예: 1, 5, 12)
        locations.append((letter, number))
    
    return locations