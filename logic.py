"""생산실적 검증·집계 함수. 화면 코드와 분리해 Jupyter에서도 확인합니다."""
import io
import math
import pandas as pd

필수열 = ["날짜", "라인", "제품", "계획수량", "생산량", "불량수", "가동시간"]

def 샘플데이터():
    # 기존 교재와 같은 15일 × 3개 라인의 가상 생산실적입니다.
    rows = []
    for day in range(1, 16):
        for i, (line, product, plan) in enumerate([
            ("1라인", "모터", 500), ("2라인", "펌프", 600), ("3라인", "밸브", 700)
        ]):
            base = {"1라인": 8, "2라인": 12, "3라인": 22}[line]
            rows.append([f"2026-07-{day:02d}", line, product, plan,
                         plan + ((day*7+i*11) % 60)-25,
                         base+(day*5+i*3) % 9, 10 if day % 5 == 0 else 8])
    return pd.DataFrame(rows, columns=필수열)

def 검증하기(df):
    """잘못된 입력은 이유를 알리고 중단합니다. 원본은 변경하지 않습니다."""
    missing = [col for col in 필수열 if col not in df.columns]
    if missing:
        raise ValueError("필요한 열이 없습니다: " + ", ".join(missing))
    if df.empty:
        raise ValueError("데이터가 0행입니다. 제목 아래에 데이터 행을 넣으세요.")
    d = df.copy()
    for col in ["라인", "제품"]:
        if d[col].isna().any() or d[col].astype(str).str.strip().eq("").any():
            raise ValueError(f"{col}에 빈 값이 있습니다.")
        d[col] = d[col].astype(str).str.strip()
    for col in ["계획수량", "생산량", "불량수", "가동시간"]:
        d[col] = pd.to_numeric(d[col], errors="coerce")
        if d[col].isna().any() or not d[col].map(math.isfinite).all():
            raise ValueError(f"{col}에는 빈칸 없이 유한한 숫자를 넣으세요.")
        if (d[col] < 0).any():
            raise ValueError(f"{col}에는 음수를 넣을 수 없습니다.")
    for col in ["계획수량", "생산량", "불량수"]:
        if (d[col] % 1 != 0).any():
            raise ValueError(f"{col}은 개수이므로 정수여야 합니다.")
    if (d["불량수"] > d["생산량"]).any():
        raise ValueError("불량수는 생산량보다 클 수 없습니다.")
    dates = pd.to_datetime(d["날짜"], errors="coerce")
    if dates.isna().any():
        raise ValueError("날짜에는 2026-07-01과 같은 유효한 날짜를 넣으세요.")
    d["날짜"] = dates.dt.strftime("%Y-%m-%d")
    return d

def 엑셀읽기(file):
    # 업로드 객체·파일 경로·BytesIO를 모두 pandas에 전달할 수 있습니다.
    try:
        df = pd.read_excel(file, engine="openpyxl")
    except Exception as exc:
        raise ValueError("엑셀 파일을 읽지 못했습니다. 정상적인 .xlsx 파일인지 확인하세요.") from exc
    return 검증하기(df)

def 지표추가(df):
    d = df.copy()
    # 분모가 0이면 계산 불가(NaN). 정상 0%로 잘못 표시하지 않습니다.
    d["불량률"] = d["불량수"] / d["생산량"].where(d["생산량"] > 0) * 100
    d["달성률"] = d["생산량"] / d["계획수량"].where(d["계획수량"] > 0) * 100
    return d

def 요약(df):
    production = int(df["생산량"].sum())
    defect = int(df["불량수"].sum())
    return {"건수": len(df), "생산량": production, "불량수": defect,
            "불량률": defect / production * 100 if production else None}

def 집계하기(df, 기준열):
    if 기준열 not in ["라인", "제품", "날짜"]:
        raise ValueError("집계 기준은 라인·제품·날짜 중 하나입니다.")
    result = df.groupby(기준열, as_index=False).agg(
        건수=("생산량", "count"), 생산량=("생산량", "sum"), 불량수=("불량수", "sum"))
    result["불량률"] = result["불량수"] / result["생산량"].where(result["생산량"] > 0) * 100
    return result

def 엑셀바이트(df):
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    return buffer.getvalue()

def 판정하기(df, 불량기준=3.0, 주의기준=2.5):
    if not (0 <= 주의기준 < 불량기준):
        raise ValueError("0 ≤ 주의 기준 < 불량 기준이어야 합니다.")
    d = 지표추가(df)
    def judge(rate):
        if pd.isna(rate):
            return "계산불가"
        if rate > 불량기준:
            return "불량"
        if rate > 주의기준:
            return "주의"
        return "정상"
    # 반올림하기 전 값으로 판정하고, 반올림은 화면 표시에서만 합니다.
    d["판정"] = d["불량률"].apply(judge)
    return d

def 판정요약(df):
    return {name: int((df["판정"] == name).sum()) for name in ["불량", "주의", "정상", "계산불가"]}
