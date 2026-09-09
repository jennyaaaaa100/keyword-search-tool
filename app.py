import streamlit as st
import streamlit.components.v1 as components
import json
import pandas as pd
import io

st.set_page_config(page_title="구역별 검색량 일괄 추출기", layout="wide")

st.title("📊 네이버 키워드 월별 검색량 일괄 추출기")
st.caption("키워드를 여러 개 입력하면 보안 차단(CORS/403)을 우회해 일괄 수집한 뒤 엑셀로 변환합니다.")

# 1. 키워드 입력 필드
default_keywords = "힐스테이트가장더퍼스트, 도마변동1구역, 호반써밋그랜드센트럴"
keywords_input = st.text_area(
    "조회할 키워드 입력 (쉼표 또는 줄바꿈으로 구분)",
    default_keywords,
    height=100
)

# 2. CORS 및 보안 우회 프록시 엔진
html_code = f"""
<div style="background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px; padding: 16px;">
    <button id="runBtn" style="background-color: #ff4b4b; color: white; border: none; padding: 10px 22px; font-size: 15px; font-weight: bold; border-radius: 6px; cursor: pointer;">
        🚀 데이터 일괄 수집 시작
    </button>
    <div id="status" style="margin-top: 12px; font-size: 14px; color: #333; font-weight: 500;"></div>
    <div id="progress-bar" style="width: 0%; height: 6px; background-color: #ff4b4b; border-radius: 3px; margin-top: 8px; transition: width 0.3s;"></div>
    <textarea id="output" style="width: 100%; height: 90px; margin-top: 12px; display: none; font-size: 11px;"></textarea>
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
        statusEl.innerText = `[${{i + 1}}/${{keywords.length}}] '${{kw}}' 수집 중...`;
        progressEl.style.width = `${{Math.round(((i + 1) / keywords.length) * 100)}}%`;
        
        try {{
            const targetUrl = `https://surffing.net/api/datalab/graph?keyword=${{encodeURIComponent(kw)}}`;
            // CORS 우회 프록시 통과
            const proxyUrl = `https://api.allorigins.win/get?url=${{encodeURIComponent(targetUrl)}}`;
            
            const res = await fetch(proxyUrl);
            if (res.ok) {{
                const proxyData = await res.json();
                const data = JSON.parse(proxyData.contents);
                
                if (data && data.chartData && data.chartData.results && data.chartData.results[0]?.data) {{
                    data.chartData.results[0].data.forEach(item => {{
                        combinedData.push({{
                            "키워드": kw,
                            "기준연월": item.period.substring(0, 7),
                            "검색량": parseInt(item.volume)
                        }});
                    }});
                }}
            }}
        }} catch (err) {{
            console.error("수집 오류:", err);
        }}
        await new Promise(r => setTimeout(r, 400));
    }}

    const outBox = document.getElementById('output');
    outBox.style.display = 'block';
    outBox.value = JSON.stringify(combinedData);
    
    runBtn.disabled = false;
    runBtn.style.opacity = "1";

    if (combinedData.length > 0) {{
        statusEl.innerHTML = `<span style="color: #28a745; font-weight: bold;">✔ 수집 완료! 아래 상자에 붙여넣고 [엑셀 표 생성]을 눌러주세요.</span>`;
        outBox.select();
        try {{
            await navigator.clipboard.writeText(outBox.value);
            alert("총 " + combinedData.length + "건 수집 완료! 클립보드에 자동 복사되었습니다. 아래 입력창에 붙여넣기(Ctrl+V) 하세요.");
        }} catch(e) {{
            alert("수집 완료! 검은 상자의 텍스트를 복사(Ctrl+C)하여 아래에 붙여넣으세요.");
        }}
    }} else {{
        statusEl.innerHTML = `<span style="color: #dc3545; font-weight: bold;">✖ 데이터를 가져오지 못했습니다. 키워드를 확인해주세요.</span>`;
    }}
}});
</script>
"""

components.html(html_code, height=210)

# 3. 엑셀 표 변환 및 다운로드 파트
st.markdown("---")
raw_data_input = st.text_area(
    "수집 결과 데이터 (위에서 자동 복사된 내용을 'Ctrl + V'로 붙여넣으세요)",
    height=80,
    placeholder="수집 완료 후 여기에 붙여넣기(Ctrl + V)"
)

if st.button("📊 엑셀 표 생성 및 다운로드", type="primary"):
    if not raw_data_input.strip() or raw_data_input.strip() == "[]":
        st.warning("데이터가 비어 있습니다. 위 버튼으로 먼저 수집을 진행해 주세요.")
    else:
        try:
            items = json.loads(raw_data_input)
            if items:
                df = pd.DataFrame(items)
                pivot_df = df.pivot(index="기준연월", columns="키워드", values="검색량").fillna(0).astype(int).sort_index(ascending=False)

                st.success(f"총 {len(items)}건의 데이터 변환 완료!")
                st.dataframe(pivot_df, use_container_width=True)

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
                st.error("데이터 항목이 비어 있습니다.")
        except Exception as e:
            st.error(f"데이터 변환 오류: {e}")
