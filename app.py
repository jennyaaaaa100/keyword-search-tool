import streamlit as st
import requests
import pandas as pd
import io
import time

st.set_page_config(page_title="구역별 검색량 추출기", layout="wide")

st.title("📊 네이버 키워드 월별 검색량 자동 추출기")
st.caption("키워드를 입력하면 월별 검색량 추이를 수집해 엑셀로 변환합니다.")

keywords_input = st.text_area(
    "키워드 입력 (쉼표 또는 줄바꿈으로 구분)",
    "힐스테이트가장더퍼스트, 도마변동1구역, 호반써밋그랜드센트럴",
    height=120
)

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Referer": "https://surffing.net/"
}

if st.button("🚀 데이터 수집 시작", type="primary"):
    raw_list = [kw.strip() for kw in keywords_input.replace("\n", ",").split(",") if kw.strip()]
    
    if not raw_list:
        st.warning("키워드를 1개 이상 입력해주세요.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        all_data = []

        for idx, kw in enumerate(raw_list):
            status_text.text(f"[{idx+1}/{len(raw_list)}] '{kw}' 수집 중...")
            url = f"https://surffing.net/api/datalab/graph?keyword={kw}"
            
            try:
                res = requests.get(url, headers=headers, timeout=10)
                if res.status_code == 200:
                    data = res.json()
                    if data.get("success") and "chartData" in data:
                        results = data["chartData"].get("results", [])
                        if results and "data" in results[0]:
                            for item in results[0]["data"]:
                                all_data.append({
                                    "키워드": kw,
                                    "기준연월": item["period"][:7],
                                    "검색량": int(item["volume"])
                                })
            except Exception as e:
                st.error(f"'{kw}' 수집 에러: {e}")

            progress_bar.progress((idx + 1) / len(raw_list))
            time.sleep(0.3)

        status_text.text("수집 완료!")

        if all_data:
            df = pd.DataFrame(all_data)
            
            pivot_df = df.pivot(index="기준연월", columns="키워드", values="검색량").fillna(0).astype(int).sort_index(ascending=False)
            
            st.success(f"총 {len(raw_list)}개 키워드 수집이 완료되었습니다.")
            st.dataframe(pivot_df, use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                pivot_df.sort_index(ascending=True).to_excel(writer, sheet_name="키워드비교_피벗")
                df.to_excel(writer, sheet_name="원본데이터", index=False)
            output.seek(0)

            st.download_button(
                label="📥 엑셀 파일 다운로드 (.xlsx)",
                data=output,
                file_name="월별검색량_추이_수집결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("데이터를 수집하지 못했습니다. 키워드를 확인해주세요.")
