
"""생산 대시보드: ① 업로드·요약 → ② 필터·차트 → ③ 다운로드"""
from datetime import datetime, timedelta, timezone

import streamlit as st

import logic

KST = timezone(timedelta(hours=9))  # 한국 날짜로 파일명 생성

st.set_page_config(page_title="생산 대시보드", page_icon="🏭", layout="wide")
st.title("🏭 생산 대시보드")

# ① 업로드 · 검증
file = st.file_uploader("생산실적 엑셀 파일(xlsx)", type="xlsx")
use_sample = st.checkbox("파일 없이 샘플 데이터로 보기")

if file is None and not use_sample:
    st.info("엑셀 파일을 올리거나 샘플 데이터를 선택하세요. 필수 열: " + " · ".join(logic.필수열))
    st.stop()

try:
    if file is not None:
        df = logic.엑셀읽기(file)                  # 읽기 + 7개 열·값 검증
    else:
        df = logic.검증하기(logic.샘플데이터())     # 샘플도 같은 검증을 거침
except ValueError as e:
    st.error(str(e))
    st.stop()

# ② 라인 선택 · 집계 기준 · 불량 기준
st.sidebar.header("조회 조건")
lines = sorted(df["라인"].unique())
sel_lines = st.sidebar.multiselect("라인", lines, default=lines)
group_col = st.sidebar.radio("집계 기준", ["라인", "제품", "날짜"])
limit = st.sidebar.slider("불량 기준(%)", min_value=0.5, max_value=10.0, value=3.0, step=0.1)

if not sel_lines:                                   # 라인 미선택 → 빈 결과 안내
    st.info("선택된 라인이 없습니다. 왼쪽에서 라인을 하나 이상 선택하세요.")
    st.stop()

filtered = df[df["라인"].isin(sel_lines)]

# 요약 (큰 숫자 4개) — 필터된 데이터 기준
s = logic.요약(filtered)
c1, c2, c3, c4 = st.columns(4)
c1.metric("건수", f"{s['건수']:,}건")
c2.metric("총 생산", f"{s['생산량']:,}개")
c3.metric("총 불량", f"{s['불량수']:,}개")
c4.metric("전체 불량률", "계산불가" if s["불량률"] is None else f"{s['불량률']:.2f}%")

if s["불량률"] is None:
    st.warning("선택한 데이터의 총 생산량이 0이라 불량률을 계산할 수 없습니다.")

# 집계
agg = logic.집계하기(filtered, group_col)

# 합계 확인: 원본 → 필터 → 집계 합계가 서로 맞는지
ok = (s["생산량"] == int(agg["생산량"].sum()) and s["불량수"] == int(agg["불량수"].sum())
      and s["건수"] == int(agg["건수"].sum()))
st.caption(f"{'✅' if ok else '❌'} 합계 확인: 원본 {len(df)}행 · 생산 {int(df['생산량'].sum()):,}  →  "
           f"조회 {len(filtered)}행 · 생산 {s['생산량']:,}  =  {group_col}별 집계 합계 {int(agg['생산량'].sum()):,}")

# 계산불가 · 기준 초과 안내
no_calc = agg[agg["불량률"].isna()]
if not no_calc.empty:
    st.warning(f"생산량이 0이라 불량률을 계산할 수 없는 {group_col}: "
               + ", ".join(no_calc[group_col].astype(str)))

over = agg[agg["불량률"] > limit]                   # 반올림 전 값으로 비교
if over.empty:
    st.success(f"{group_col}별 집계에서 불량 기준 {limit:.1f}%를 초과한 항목이 없습니다.")
else:
    items = ", ".join(f"{r[group_col]}({r['불량률']:.2f}%)" for _, r in over.iterrows())
    st.error(f"불량 기준 {limit:.1f}% 초과 {group_col} {len(over)}건: {items}")

# 탭: 집계 표·막대 차트 / 날짜별 추이 / 조회 데이터
tab1, tab2, tab3 = st.tabs(["📊 집계 표·막대 차트", "📈 날짜별 추이", "📋 조회 데이터"])

with tab1:
    st.subheader(f"{group_col}별 집계")
    st.dataframe(agg.round({"불량률": 2}), hide_index=True)
    st.caption(f"{group_col}별 불량률(%) · 기준 {limit:.1f}%")
    st.bar_chart(agg.set_index(group_col)["불량률"])

with tab2:
    daily = filtered.groupby(["날짜", "라인"], as_index=False)[["생산량", "불량수"]].sum()
    daily["불량률"] = daily["불량수"] / daily["생산량"].where(daily["생산량"] > 0) * 100
    left, right = st.columns(2)
    with left:
        st.subheader("날짜별 생산량")
        st.line_chart(daily.pivot(index="날짜", columns="라인", values="생산량"))
    with right:
        st.subheader("날짜별 불량률(%)")
        st.line_chart(daily.pivot(index="날짜", columns="라인", values="불량률"))

with tab3:
    view = logic.판정하기(filtered, 불량기준=limit, 주의기준=min(2.5, limit - 0.1))
    st.dataframe(view.round({"불량률": 2, "달성률": 1}), hide_index=True)

# ③ 다운로드
today = datetime.now(KST).strftime("%Y%m%d")
csv_bytes = filtered.to_csv(index=False).encode("utf-8-sig")   # 엑셀에서 한글 깨짐 방지
excel_bytes = logic.엑셀바이트(agg)

# 합계 확인: 내려받을 데이터의 합계가 화면 수치와 같은지
ok = int(filtered["생산량"].sum()) == s["생산량"] and int(agg["생산량"].sum()) == s["생산량"]
st.caption(f"{'✅' if ok else '❌'} 다운로드 합계 확인: 조회 CSV {len(filtered)}행 · 생산 {int(filtered['생산량'].sum()):,} / "
           f"집계 엑셀 {len(agg)}행 · 생산 {int(agg['생산량'].sum()):,}")

d1, d2 = st.columns(2)
d1.download_button("📄 조회 데이터 CSV", data=csv_bytes,
                   file_name=f"조회결과_{today}.csv", mime="text/csv")
d2.download_button("📊 집계 결과 엑셀", data=excel_bytes,
                   file_name=f"집계결과_{group_col}_{today}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
