import pandas as pd
def process_excel_files(user_info_path, completion_list_path):
    # 엑셀 파일의 모든 시트 읽기
    user_info_sheets = pd.read_excel(user_info_path, sheet_name=None)
    completion_list_sheets = pd.read_excel(completion_list_path, sheet_name=None)

    # 결과를 저장할 딕셔너리
    result_sheets = {}

    for sheet_name in user_info_sheets.keys():
        if sheet_name not in completion_list_sheets:
            print(f"경고: '{sheet_name}' 시트가 수료자 명단에 없습니다. 이 시트는 처리되지 않습니다.")
            continue

        users = user_info_sheets[sheet_name]
        completion_list = completion_list_sheets[sheet_name]

        # user_info의 이름+전화번호 뒤 4자리 열에서 이름과 전화번호 추출
        users['name'] = users['이름+전화번호 뒤 4자리'].str.extract(r'([가-힣]+)')
        users['phone'] = users['이름+전화번호 뒤 4자리'].str.extract(r'(\d{4})$')

        # 수료자 명단에서 이름+전화번호 형식으로 된 항목 처리
        completion_list['name'] = completion_list['이름'].str.extract(r'([가-힣]+)')
        completion_list['phone'] = completion_list['이름'].str.extract(r'(\d{4})$')

        # 이름으로 먼저 매칭
        merged = pd.merge(users, completion_list[['name', 'phone', '수료 여부']], 
                            on=['name'], how='left', suffixes=('', '_completion'))

        # 동명이인 처리: 전화번호가 일치하는 경우에만 수료 여부 유지
        mask = (merged['phone'] != merged['phone_completion']) & merged['phone_completion'].notna()
        merged.loc[mask, '수료 여부'] = None

        # 불필요한 열 제거
        merged = merged.drop('phone_completion', axis=1)

        # 수료 여부가 NaN인 경우 '미수료'로 채우기
        merged['수료 여부'] = merged['수료 여부'].fillna('미수료')

        merged = merged.drop(['name', 'phone'], axis=1)
        
        # 결과를 딕셔너리에 저장
        result_sheets[sheet_name] = merged

    # 결과를 새로운 엑셀 파일로 저장
    with pd.ExcelWriter('result.xlsx') as writer:
        for sheet_name, df in result_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print("처리가 완료되었습니다. 결과는 'result.xlsx' 파일에 저장되었습니다.")

# 스크립트 실행
if __name__ == "__main__":
    user_info_path = "./user_info/users.xlsx"
    completion_list_path = "./user_info/pre_certified.xlsx"
    process_excel_files(user_info_path, completion_list_path)