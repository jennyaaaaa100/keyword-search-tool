import streamlit as st
import streamlit.components.v1 as components
import json
import pandas as pd
import io

st.set_page_config(page_title="구역별 검색량 일괄 추출기", layout="wide")

st.title("📊 네이버 키워드 월별 검색량 일괄 추출기")
st.caption("키워드를 여러 개 입력하면 브라우저에서 Cloudflare 차단 없이 일괄 수집해 엑셀로 변환합니다.")

# 1. 키워드 입력 필드
default_keywords = "힐스테이트가장더퍼스트, 도마변동1구역, 호반써밋그랜드센트럴"
keywords_input = st.text_area(
    "조회할 키워드 입력 (쉼표 또는 줄바꿈으로 여러 개 입력 가능)",
    default_keywords,
    height=110
)

# 2. 브라우저에서 직접 일괄 조회하는 JavaScript 엔진 임베딩
html_code = f"""
<div style="background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 8px; padding: 16px; margin-bottom: 20px;">
    <button id="runBtn" style="background-color: #ff4b4b; color: white; border: none; padding: 10px 20px; font-size: 15px; font-weight: bold; border-radius: 6px; cursor: pointer;">
        🚀 브라우저에서 일괄 데이터 수집 시작
    </button>
    <div id="status" style="margin-top: 12px; font-size: 14px; color: #333; font-weight: 500;"></div>
    <div id="progress-bar" style="width: 0%; height: 6px; background-color: #ff4b4b; border-radius: 3px; margin-top: 8px; transition: width 0.3s;"></div>
    <textarea id="output" style="width: 100%; height: 120px; margin-top: 12px; display: none;"></textarea>
</div>

<script>
document.getElementById('runBtn').addEventListener('click', async () => {{
    const rawInput = `{keywords_input}`;
    const keywords = rawInput.split(/[\\n,]+/).map(k => k.trim()).filter(k => k.length > 0);
    
    if (keywords.length === 0) {{
        alert("키워드를 1개 이상 입력해주세요.");
        return;
    }}

    const statusEl = document.getElementById('status');
    const progressEl = document.getElementById('progress-bar');
    const runBtn = document.getElementById('runBtn');
    
    runBtn.disabled = true;
    runBtn.style.opacity = "0.6";
    
    let combinedData = [];
    
    for (let i = 0; i < keywords.length; i++) {{
        const kw = keywords[i];
        statusEl.innerText = `[${{i + 1}}/${{keywords.length}}] '${{kw}}' 데이터 조회 중...`;
        progressEl.style.width = `${{Math.round(((i + 1) / keywords.length) * 100)}}%`;
        
        try {{
            const url = `https://surffing.net/api/datalab/graph?keyword=${{encodeURIComponent(kw)}}`;
            const res = await fetch(url, {{
                method: "GET",
                headers: {{ "Accept": "application/json" }}
            }});
            
            if (res.ok) {{
                const data = await res.json();
                if (data && data.chartData && data.chartData.results && data.chartData.results[0]?.data) {{
                    data.chartData.results[0].data.forEach(item => {{
                        combinedData.push({{
                            "키워드": kw,
                            "기준연월": item.period.substring(0, 7),
                            "검색량": parseInt(item.volume)
                        }});
                    }});
                }}
            }} else {{
                console.error(`Error loading ${{kw}}: HTTP ${{res.status}}`);
            }}
        }} catch (err) {{
            console.error(err);
        }}
        // 요청 간 0.3초 대기
        await new Promise(r => setTimeout(r, 300));
    }}

    statusEl.innerHTML = `<span style="color: #28a745; font-weight: bold;">✔ 총 ${{keywords.length}}개 키워드 수집 완료! 아래 확인 버튼을 눌러주세요.</span>`;
    runBtn.disabled = false;
    runBtn.style.opacity = "1";

    // 부모 Streamlit 세션으로 데이터 전송용
    const outBox = document.getElementById('output');
    outBox.style.display = 'block';
    outBox.value = JSON.stringify(combinedData);
    outBox.select();
    navigator.clipboard.writeText(outBox.value);
    alert("데이터 수집 완료! 클립보드에 자동 복사되었습니다. 아래 [데이터 변환] 상자에 붙여넣고 완료하세요.");
}});
</script>
"""

components.html(html_code, height=220)

# 3. 데이터 수신 및 엑셀 다운로드 변환 파트
st.markdown("---")
raw_data_input = st.text_area(
    "수집 결과 데이터 (위 버튼 완료 후 자동으로 복사되니 'Ctrl + V'로 붙여넣으세요)",
    height=90,
    placeholder="수집 완료 후 여기에 붙여넣기(Ctrl + V) 하시면 됩니다."
)

if st.button("📊 엑셀 표 생성 및 다운로드", type="primary"):
    if not raw_data_input.strip():
        st.warning("데이터가 비어 있습니다. 위 버튼으로 수집을 진행한 뒤 붙여넣어 주세요.")
    else:
        try:
            items = json.loads(raw_data_input)
            if items:
                df = pd.DataFrame(items)
                pivot_df = df.pivot(index="기준연월", columns="키워드", values="검색량").fillna(0).astype(int).sort_index(ascending=False)

                st.success(f"총 {len(items)}건 데이터 정리 완료!")
                st.dataframe(pivot_df, use_container_width=True)

                # 엑셀 파일 생성
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    pivot_df.sort_index(ascending=True).to_excel(writer, sheet_name="키워드비교_피벗")
                    df.to_excel(writer, sheet_name="전체_RawData", index=False)
                output.seek(0)

                st.download_button(
                    label="📥 엑셀 파일 다운로드 (.xlsx)",
                    data=output,
                    file_name="월별검색량_일괄추출결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("수집된 데이터가 없습니다.")
        except Exception as e:
            st.error(f"데이터 변환 오류: {e}")
