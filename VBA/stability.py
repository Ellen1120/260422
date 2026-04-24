import pandas as pd
import numpy as np
import os

# 1. 파일 경로 설정
file_path = r"D:\VBA\Stability master plan.xlsx"
save_path = r"D:\VBA\Stability_Result.xlsx"

print(f"🚀 분석을 시작합니다 (정확도 우선 모드)...")

try:
    if not os.path.exists(file_path):
        print(f"❌ 파일을 찾을 수 없습니다: {file_path}")
    else:
        # 엑셀 읽기
        df = pd.read_excel(file_path, header=None)

        # 2. C2, E2에서 검색 조건 추출 (부동소수점 제거 처리)
        def clean_val(v):
            return str(v).split('.')[0].strip()

        target_year = clean_val(df.iloc[1, 2])   # C2
        target_month = clean_val(df.iloc[1, 4])  # E2
        print(f"🔎 검색 대상: {target_year}년 {target_month}월")

        # 3. 목표 열(Column) 찾기
        target_col = -1
        for col in range(13, df.shape[1]):
            # 4행(연도)과 5행(월)의 텍스트를 정밀 대조
            h_text = clean_val(df.iloc[3, col]) + clean_val(df.iloc[4, col])
            if target_year in h_text and target_month in h_text:
                target_col = col
                break

        if target_col == -1:
            print(f"❌ {target_year}년 {target_month}월 열을 찾지 못했습니다.")
        else:
            print(f"✅ 열 찾기 성공! 데이터를 정밀 스캔합니다.")
            
            results = []
            # A~M열의 병합된 정보를 실시간으로 채우며 내려갑니다.
            memo = [None] * 13 

            for i in range(5, len(df)):
                # 제품명 등 기본정보(A~M열) 업데이트 (병합 셀 대응)
                for j in range(13):
                    v = df.iloc[i, j]
                    if pd.notna(v) and str(v).strip() != "":
                        memo[j] = str(v).strip()

                # 해당 월 칸의 값 확인
                raw_cell = df.iloc[i, target_col]
                cell_val = str(raw_cell).strip() if pd.notna(raw_cell) else ""
                
                # 칸에 무언가 적혀 있다면 추출!
                if cell_val != "" and cell_val.lower() != "nan":
                    # '12M/10' 등에서 날짜 숫자(10) 추출
                    try:
                        day_sort = int(cell_val.split('/')[-1]) if '/' in cell_val else 0
                    except:
                        day_sort = 0
                    
                    # 위에서부터 채워져 내려온 제품 정보 + 불출값
                    row_data = list(memo) + [cell_val, day_sort]
                    results.append(row_data)

            # 4. 정렬 및 저장
            if results:
                res_df = pd.DataFrame(results)
                # 날짜순 정렬 후 임시 열 삭제
                res_df = res_df.sort_values(by=res_df.columns[-1]).iloc[:, :-1]
                
                # 컬럼명 자동 설정
                headers = ["No.", "제품명", "STP No.", "개시일", "구분", "배치사이즈", "단위", 
                          "타입", "안정성", "기간", "배치번호", "제조일", "유효기간", f"{target_month}월 상세"]
                
                res_df.to_excel(save_path, index=False, header=headers)
                print(f"✨ 성공! 총 {len(results)}건의 리스트를 날짜순으로 뽑았습니다.")
                print(f"💾 결과 확인: {save_path}")
            else:
                print(f"💡 해당 월 열에 데이터가 없습니다. 엑셀을 확인하세요.")

except Exception as e:
    print(f"🚨 오류 발생: {e}")