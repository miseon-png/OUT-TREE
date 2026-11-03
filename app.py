
import io
import datetime
import streamlit as st
import pandas as pd
import openpyxl
import gspread

st.set_page_config(
    page_title="원재료 수불부 출고 자동 연동기",
    page_icon="📦",
    layout="wide"
)

st.title("📦 현장 작업기준서 ➔ 원재료 수불부 자동 연동")
st.markdown("""
현장 작업기준서 엑셀 파일을 업로드하면 **야채 원재료 수불부** 규격에 맞춰 
일일 출고 데이터를 자동 추출·정리하고, 구글 스프레드시트에 바로 반영할 수 있습니다.
""")

# 원재료 시트 열 매핑 (열 인덱스: 4=D, 7=G, 10=J, 13=M, 16=P, 19=S, 22=V, 25=Y)
MATERIALS_MAP = {
    4: "카이피라",
    7: "프릴아이스",
    10: "버터헤드",
    13: "레드오크",
    16: "로메인",
    19: "치커리",
    22: "적근대",
    25: "케일"
}

def parse_work_standard(file_bytes, partner_name="생산투입"):
    """현장 작업기준서에서 수불부 규격의 출고 데이터프레임 생성"""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    valid_sheets = [s for s in wb.sheetnames if s != '무지']
    
    rows = []
    for sheet in valid_sheets:
        ws = wb[sheet]
        
        # 날짜 추출 (A3 셀)
        d_val = ws.cell(3, 1).value
        if isinstance(d_val, (datetime.datetime, datetime.date)):
            date_str = d_val.strftime("%Y-%m-%d")
        else:
            date_str = str(d_val)[:10] if d_val else f"2026-10-{sheet.zfill(2)}"
            
        # 각 원료별 22행 투입합계 수량 추출
        for col_idx, mat_name in MATERIALS_MAP.items():
            kg = ws.cell(22, col_idx).value
            if kg is not None and kg != "":
                kg_float = round(float(kg), 1)
                if kg_float > 0:  # 출고량이 있는 품목만 기록
                    rows.append({
                        "일자": date_str,
                        "구분": "출고",
                        "거래처": partner_name,
                        "원료명": mat_name,
                        "수량(kg)": kg_float,
                        "단가": "",
                        "총금액": "",
                        "비고": f"작업기준서 {sheet}일자 출고"
                    })
                    
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["일자", "원료명"]).reset_index(drop=True)
    return df

# 사이드바 설정
st.sidebar.header("⚙️ 연동 설정")
partner = st.sidebar.text_input("출고 거래처 표기명", value="생산투입")
uploaded_file = st.sidebar.file_uploader("현장 작업기준서 엑셀 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_result = parse_work_standard(uploaded_file.read(), partner_name=partner)
        
        st.subheader("📋 생성된 원재료 출고 수불부 데이터 (미리보기)")
        st.write(f"총 **{len(df_result)}건**의 일일 출고 내역이 추출되었습니다.")
        st.dataframe(df_result, use_container_width=True)
        
        # 통계 요약
        st.markdown("### 📊 10월 원재료별 총 출고 소요량")
        summary = df_result.groupby("원료명")["수량(kg)"].sum().reset_index()
        col1, col2 = st.columns([1, 2])
        with col1:
            st.dataframe(summary, use_container_width=True)
        with col2:
            st.bar_chart(summary.set_index("원료명")["수량(kg)"])

        # 다운로드 버튼
        st.markdown("---")
        csv_data = df_result.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            label="📥 수불부 형식 CSV 다운로드 (구글 시트에 직접 붙여넣기용)",
            data=csv_data,
            file_name=f"원재료수불부_출고데이터_{datetime.date.today()}.csv",
            mime="text/csv"
        )
        
        # 구글 스프레드시트 직접 자동 전송 (선택 옵션)
        st.markdown("### 🚀 구글 시트 원클릭 자동 추가")
        sheet_url = "https://docs.google.com/spreadsheets/d/1WM472gdj_7HBUvJdLRgwIU0t7GAMnZ47GpsCJfB5pBo/edit"
        st.caption(f"대상 시트: {sheet_url}")
        
        if "gcp_service_account" in st.secrets:
            if st.button("구글 스프레드시트에 지금 추가하기"):
                with st.spinner("구글 스프레드시트에 데이터를 전송하는 중..."):
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(sheet_url)
                    ws = sh.worksheet("시트1")
                    
                    # 수불부 데이터 추가 (헤더 제외 리스트 변환)
                    append_rows = df_result.values.tolist()
                    ws.append_rows(append_rows)
                    st.success("✅ 구글 스프레드시트 수불부 맨 아래에 출고 데이터가 성공적으로 추가되었습니다!")
        else:
            st.info("💡 Streamlit Secrets에 `gcp_service_account`(구글 서비스 계정 키)를 등록하시면 버튼 클릭 한 번으로 구글 시트에 자동 Append 됩니다.")

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
else:
    st.info("👈 왼쪽 사이드바에서 '현장 작업기준_26년 10월.xlsx' 파일을 업로드해 주세요.")
