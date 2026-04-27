import pandas as pd

# 1. 파일 경로 설정 (D드라이브 VBA 폴더 안의 파일)
file_path = r'D:\VBA\Stability master plan.xlsx'

try:
    # 5행부터 데이터이므로 header=4 (0부터 시작)
    # 파일이 무거우므로 engine='openpyxl' 추가
    df = pd.read_excel(file_path, header=4, engine='openpyxl')

    # 2. EQ열(147번째 열) 필터링
    # 열 이름이 'EQ'인지 확인하고, 아니면 146번째 인덱스로 필터링합니다.
    if 'EQ' in df.columns:
        result = df[df['EQ'].notna()].copy()
    else:
        result = df[df.iloc[:, 146].notna()].copy()

    # 3. 날짜순 정렬 (날짜 열 이름을 'Date' 또는 실제 열 이름으로 바꿔주세요)
    # 만약 열 이름을 모른다면 첫 번째 열(0번)을 기준으로 정렬합니다.
    date_col = result.columns[0] # 첫 번째 열이 날짜라고 가정
    result[date_col] = pd.to_datetime(result[date_col], errors='coerce')
    result = result.sort_values(by=date_col).dropna(subset=[date_col])

    # 4. 결과 저장
    save_path = r'D:\VBA\Stability_Filtered_Result.xlsx'
    result.to_excel(save_path, index=False)
    
    print(f"✅ 추출 성공! {len(result)}건의 데이터를 날짜순으로 정리했습니다.")
    print(f"📍 저장 위치: {save_path}")

except Exception as e:
    print(f"❌ 에러 발생: {e}")