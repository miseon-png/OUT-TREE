import io
import datetime
from decimal import Decimal, ROUND_HALF_UP
import streamlit as st
import pandas as pd
import openpyxl
import gspread

# 페이지 기본 설정
st.set_page_config(
    page_title="원재료 수불부 자동 연동기",
    page_icon="🥬",
    layout="wide"
)

st.title("🥬 현장 작업기준서 ➔ 원재료 수불부 자동 연동 시스템")
st.markdown("""
현장 작업기준서 엑셀 파일을 업로드하면 **[가공 투입합계(22행) + 원물(23행)]**을 합산하고,
배합비로 인한 소수점은 **전부 반올림(정수)**하여 **야채 원재료 수불부**로 자동 전송합니다.
""")

# 원재료 시트 열 매핑
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

def round_half_up(value):
    """사사오입(0.5 이상 올림) 정수 반올림"""
    if value is None or value == "":
        return 0
    d = Decimal(str(value))
    return int(d.quantize(Decimal('1'), rounding=ROUND_HALF_UP))

def parse_work_standard(file_bytes, partner_name="생산투입"):
    """현장 작업기준서에서 [투입합계 + 원물]을 합산하고 정수 반올림하여 수불부 데이터 생성"""
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
            
        for col_idx, mat_name in MATERIALS_MAP.items():
            # 22행: 가공 투입합계
            val_22 = ws.cell(22, col_idx).value
            # 23행: 원물 추가량
            val_23 = ws.cell(23, col_idx).value
            
            kg_22 = float(val_22) if val_22 not in (None, "") else 0.0
            kg_23 = float(val_23) if val_23 not in (None, "") else 0.0
            
            total_kg = kg_22 + kg_23
            final_kg = round_half_up(total_kg)
            
            if final_kg > 0:
                note = f"작업기준서 {sheet}일자 출고"
                if kg_23 > 0:
                    note += f" (원물 {int(kg_23)}kg 포함)"
                    
                rows.append({
                    "일자": date_str,
                    "구분": "출고",
                    "거래처": partner_name,
                    "원료명": mat_name,
                    "수량(kg)": final_kg,
                    "단가": "",
                    "총금액": "",
                    "비고": note
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
        
        st.subheader("📋 출고 수불부 변환 결과 (반올림 완료)")
        st.caption(f"총 {len(df_result)}건의 일일 출고 내역 (원물 포함 및 정수 반올림 적용)")
        
        # 순수 테이블 출력
        st.table(df_result.head(30))
        if len(df_result) > 30:
            st.caption(f"※ 화면에는 상위 30건만 미리보기로 표시됩니다. (전체 {len(df_result)}건은 하단 다운로드 및 구글 시트로 모두 전송됩니다)")
        
        # 원재료별 총 누적 출고량
        st.markdown("### 📊 10월 원재료별 총 출고 수량 합계")
        summary = df_result.groupby("원료명")["수량(kg)"].sum().reset_index()
        summary.columns = ["원료명", "총 출고량(kg)"]
        st.table(summary)

        # CSV 다운로드
        st.markdown("---")
        csv_data = df_result.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            label="📥 수불부 형식 CSV 다운로드",
            data=csv_data,
            file_name=f"원재료수불부_출고데이터_{datetime.date.today()}.csv",
            mime="text/csv"
        )
        
        # 구글 스프레드시트 원클릭 연동
        st.markdown("### 🚀 구글 스프레드시트 원클릭 자동 전송")
        sheet_url = "https://docs.google.com/spreadsheets/d/1WM472gdj_7HBUvJdLRgwIU0t7GAMnZ47GpsCJfB5pBo/edit"
        st.caption(f"연동 대상: {sheet_url}")
        
        if "gcp_service_account" in st.secrets:
            if st.button("구글 스프레드시트에 지금 추가하기", type="primary"):
                with st.spinner("구글 스프레드시트에 출고 데이터를 기록하는 중..."):
                    # Secrets 딕셔너리 복사 및 private_key 줄바꿈 안전 보정
                    credentials_dict = dict(st.secrets["gcp_service_account"])
                    if "private_key" in credentials_dict:
                        credentials_dict["private_key"] = credentials_dict["private_key"].replace("\\n", "\n")
                    
                    gc = gspread.service_account_from_dict(credentials_dict)
                    sh = gc.open_by_url(sheet_url)
                    ws = sh.worksheet("시트1")
                    
                    # 수불부 데이터 일괄 추가
                    append_rows = df_result.values.tolist()
                    ws.append_rows(append_rows)
                    st.success("✅ [가공 + 원물 반올림] 출고 데이터가 구글 시트 맨 아래에 성공적으로 반영되었습니다!")
        else:
            st.info("💡 Streamlit Secrets에 서비스 계정 키가 설정되어 있어야 자동 전송이 가능합니다.")

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
else:
    st.info("👈 왼쪽 사이드바에서 '현장 작업기준_26년 10월.xlsx' 파일을 업로드해 주세요.")
