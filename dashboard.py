"""생산 현황 대시보드: logic.py의 검증·집계 함수를 이용하는 Streamlit 화면."""
from datetime import datetime
from zoneinfo import ZoneInfo
import streamlit as st
import pandas as pd
import logic

st.set_page_config(page_title="생산 현황 대시보드", page_icon="📊", layout="wide")
st.title("📊 생산 현황 대시보드")
st.caption("엑셀 파일 업로드 → 검증 → 라인 선택 → 집계·차트 → 다운로드")

# 1단계: 업로드·검증. 실제 계산은 logic.py에 맡깁니다.
업로드파일 = st.file_uploader("생산실적 엑셀 파일을 업로드하세요 (.xlsx)", type=["xlsx"])
if 업로드파일 is None:
    st.info("먼저 엑셀 파일을 업로드해 주세요. 필수 열: " + ", ".join(logic.필수열))
    st.stop()

try:
    원본df = logic.엑셀읽기(업로드파일)
except ValueError as 오류:
    st.error(f"업로드 파일을 확인해 주세요: {오류}")
    st.stop()
except Exception as 오류:
    st.error(f"데이터 처리 중 오류가 발생했습니다: {오류}")
    st.stop()

st.success(f"데이터 검증 완료: {len(원본df):,}건")

# 2단계: 조회 조건. 기본값은 전체 라인입니다.
전체라인 = sorted(원본df["라인"].unique().tolist())
왼쪽, 오른쪽 = st.columns([2, 1])
with 왼쪽:
    선택라인 = st.multiselect("조회할 생산 라인을 선택하세요", 전체라인, default=전체라인)
with 오른쪽:
    불량기준 = st.slider("불량률 경고 기준 (%)", min_value=0.0, max_value=100.0,
                      value=3.0, step=0.1)

if not 선택라인:
    st.warning("선택한 라인이 없습니다. 조회할 라인을 하나 이상 선택해 주세요.")
    st.stop()

조회df = 원본df.loc[원본df["라인"].isin(선택라인)].copy()
if 조회df.empty:
    st.info("선택한 조건에 해당하는 데이터가 없습니다.")
    st.stop()

# 화면 위쪽 큰 숫자: 반드시 필터링된 조회 데이터에 대해 계산합니다.
요약값 = logic.요약(조회df)
지표1, 지표2, 지표3, 지표4 = st.columns(4)
지표1.metric("조회 건수", f'{요약값["건수"]:,}건')
지표2.metric("총 생산량", f'{요약값["생산량"]:,}개')
지표3.metric("총 불량수", f'{요약값["불량수"]:,}개')
지표4.metric("전체 불량률", "계산불가" if 요약값["불량률"] is None else f'{요약값["불량률"]:.2f}%')
if 요약값["생산량"] == 0:
    st.warning("조회된 총 생산량이 0이어서 전체 불량률을 계산할 수 없습니다.")

집계기준 = st.selectbox("집계 기준을 선택하세요", ["라인", "제품", "날짜"])
집계df = logic.집계하기(조회df, 집계기준)
초과df = 집계df.loc[집계df["불량률"] > 불량기준]
if 초과df.empty:
    st.success(f"집계 항목 중 불량률 {불량기준:.1f}%를 초과한 항목이 없습니다.")
else:
    st.warning(f"불량률 {불량기준:.1f}% 초과: {len(초과df)}개 항목")
    st.dataframe(초과df[[집계기준, "생산량", "불량수", "불량률"]].style.format({"불량률": "{:.2f}%"}),
                 use_container_width=True, hide_index=True)
if 집계df["불량률"].isna().any():
    st.info("집계 생산량이 0인 항목의 불량률은 계산불가(NaN)로 표시됩니다.")

탭1, 탭2, 탭3 = st.tabs(["집계 표·막대 차트", "날짜별 추이", "조회 데이터"])
with 탭1:
    st.subheader(f"{집계기준}별 생산 현황")
    st.dataframe(집계df.style.format({"불량률": "{:.2f}%", "생산량": "{:,.0f}", "불량수": "{:,.0f}"}),
                 use_container_width=True, hide_index=True)
    차트df = 집계df.set_index(집계기준)[["생산량", "불량수"]]
    st.bar_chart(차트df)
with 탭2:
    st.subheader("날짜별 생산량·불량수 추이")
    날짜df = logic.집계하기(조회df, "날짜").sort_values("날짜")
    st.line_chart(날짜df.set_index("날짜")[["생산량", "불량수"]])
    st.dataframe(날짜df, use_container_width=True, hide_index=True)
with 탭3:
    st.subheader("현재 조회된 생산 데이터")
    st.dataframe(조회df, use_container_width=True, hide_index=True)

# 3단계: 다운로드. 파일명은 서버 시간대가 아닌 한국 표준시를 기준으로 합니다.
st.subheader("조회·집계 결과 다운로드")
오늘 = datetime.now(ZoneInfo("Asia/Seoul")).strftime("%Y%m%d")
CSV데이터 = 조회df.to_csv(index=False).encode("utf-8-sig")
엑셀데이터 = logic.엑셀바이트(집계df)
다운1, 다운2 = st.columns(2)
with 다운1:
    st.download_button("조회 데이터 CSV 다운로드", data=CSV데이터,
                       file_name=f"조회결과_{오늘}.csv", mime="text/csv",
                       use_container_width=True)
with 다운2:
    st.download_button("집계 결과 엑셀 다운로드", data=엑셀데이터,
                       file_name=f"집계결과_{오늘}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
st.caption("다운로드 데이터는 현재 조회 조건을 반영하며, 원본 엑셀 파일은 변경하지 않습니다.")
