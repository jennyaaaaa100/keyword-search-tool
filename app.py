import streamlit as st
import cloudscraper
import pandas as pd
import io
import time
import urllib.parse

st.set_page_config(page_title="구역별 검색량 추출기", layout="wide")

st.title("📊 네이버 키워드 월별 검색량 자동 추출기")
st.caption("키워드를 입력하면 월별 검색량 추이를 수집해 엑셀로 변환합니다.")

keywords_input = st.text_area(
    "키워드 입력 (쉼표 또는 줄바꿈으로 구분)",
    "힐스테이트가장더퍼스트, 도마변동1구역, 호반써밋그랜드센트럴",
    height=120
)

if st.button("🚀 데이터 수집 시작", type="primary"):
    raw_list = [kw.strip() for kw in keywords_input.replace("\n", ",").split(",") if kw.strip()]
    
    if not raw_list:
        st.warning("키워드를 1개 이상 입력해주세요.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        all_data = []

        # Cloudflare 보안 차단 우회 세션 생성
        scraper = cloudscraper.create_scraper(
            browser={
                'browser': 'chrome',
                'platform': 'windows',
                'desktop': True
            }
        )

        for idx, kw in enumerate(raw_list):
            status_text.text(f"[{idx+1}/{len(raw_list)}] '{kw}' 수집 중...")
            encoded_kw = urllib.parse.quote(kw)
            url = f"https://surffing.net/api/datalab/graph?keyword={encoded_kw}"
            
            try:
                res = scraper.get(url, timeout=20)
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
                        else:
                            st.warning(f"'{kw}' 검색 데이터가 존재하지 않습니다.")
                    else:
                        st.warning(f"'{kw}' 응답 형식 불일치: {data}")
                else:
                    st.error(f"'{kw}' 호출 실패 (상태 코드: {res.status_code})")
            except Exception as e:
                st.error(f"'{kw}' 통신 오류: {e}")

            progress_bar.progress((idx + 1) / len(raw_list))
            time.sleep(0.5)

        status_text.text("수집 작업 완료!")

        if all_data:
            df = pd.DataFrame(all_data)
            
            # 월별/키워드별 피벗 테이블 생성
            pivot_df = df.pivot(index="기준연월", columns="키워드", values="검색량").fillna(0).astype(int).sort_index(ascending=False)
            
            st.success(f"총 {len(all_data)}건의 데이터 수집이 완료되었습니다.")
            st.dataframe(pivot_df, use_container_width=True)

            # 엑셀 다운로드 버퍼 생성
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                pivot_df.sort_index(ascending=True).to_excel(writer, sheet_name="키워드비교_피벗")
                df.to_excel(writer, sheet_name="전체_RawData", index=False)
            output.seek(0)

            st.download_button(
                label="📥 엑셀 파일 다운로드 (.xlsx)",
                data=output,
                file_name="월별검색량_수집결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("데이터를 수집하지 못했습니다. 키워드를 확인해주세요.")
