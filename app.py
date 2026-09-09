import streamlit as st
import pandas as pd
import json
import io

st.set_page_config(page_title="검색량 데이터 변환기", layout="wide")

st.title("📊 네이버 키워드 월별 검색량 엑셀 변환기")
st.caption("Surffing API 링크를 새 탭에서 열고, 보이는 내용을 복사해 붙여넣으면 엑셀로 생성됩니다.")

st.markdown("""
### 사용 방법
1. 아래 링크를 새 창에서 엽니다:  
   👉 `https://surffing.net/api/datalab/graph?keyword=원하는키워드`
2. 화면에 뜨는 텍스트를 전체 복사(`Ctrl + A` → `Ctrl + C`)합니다.
3. 아래 입력창에 붙여넣고 버튼을 누르면 엑셀이 다운로드됩니다.
""")

raw_json = st.text_area("JSON 데이터 붙여넣기", height=180)

if st.button("🚀 엑셀 변환 및 다운로드", type="primary"):
    if not raw_json.strip():
        st.warning("데이터를 입력해주세요.")
    else:
        try:
            data = json.loads(raw_json)
            results = data.get("chartData", {}).get("results", [])
            
            if results and "data" in results[0]:
                kw_title = results[0].get("title", "키워드")
                rows = []
                for item in results[0]["data"]:
                    rows.append({
                        "키워드": kw_title,
                        "기준연월": item["period"][:7],
                        "검색량": int(item["volume"])
                    })
                
                df = pd.DataFrame(rows)
                st.success(f"'{kw_title}' 데이터 변환 완료!")
                st.dataframe(df, use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, sheet_name="검색량", index=False)
                output.seek(0)

                st.download_button(
                    label="📥 엑셀 파일 다운로드 (.xlsx)",
                    data=output,
                    file_name=f"{kw_title}_월별검색량.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("데이터 형식이 올바르지 않거나 결과가 비어 있습니다.")
        except Exception as e:
            st.error(f"파싱 오류: {e}")
