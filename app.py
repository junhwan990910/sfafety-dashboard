import streamlit as st
import pandas as pd
import numpy as np
import math
import re
import datetime
import io

st.set_page_config(page_title="기종별 수요량 분석 대시보드", layout="wide")

# 스타일 설정
st.markdown("""
<style>
    [data-testid="stDataFrame"] table {
        font-size: 16px;
    }
</style>
""", unsafe_allow_html=True)

st.title("🛠️ 기종별 개별 설정 및 상세 리포트 분석 대시보드")

uploaded_file = st.file_uploader("엑셀 파일 업로드", type=['xlsx'])

col1, col2 = st.columns(2)
with col1:
    global_lead_time = st.number_input("NC 리드타임:", value=6.7, step=0.1)
with col2:
    model_choice = st.selectbox(
        "수요 분석 모델:",
        options=[
            ('elastic_capa', '양방향 탄력 밴드 완충 모델 (상하방 유연 연동)'),
            ('trend_short', '단기 수주추이 중심'),
            ('trend_long', '중장기 안정세 중심')
        ],
        format_func=lambda x: x[1]
    )[0]

def safe_float(val):
    if pd.isna(val): 
        return None
    if isinstance(val, (int, float)): 
        return float(val)
    val_str = str(val).strip()
    if val_str == '' or val_str in ['-', '↓', '↑', 'N/A', 'nan', 'None']: 
        return None
    try:
        if val_str.endswith('%'):
            return float(val_str[:-1].strip()) / 100.0
        return float(val_str)
    except ValueError:
        cleaned = ''.join(c for c in val_str if c.isdigit() or c in ['.', '-'])
        try:
            return float(cleaned) if cleaned else None
        except:
            return None

def extract_op_demand_from_a23(file_bytes, sheet_name):
    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
        val = df_raw.iloc[22, 0]
        parsed = safe_float(val)
        if parsed is not None:
            return parsed
    except Exception:
        pass
    return 5.0

def extract_contract_rate_final(df):
    for r_idx in range(len(df)):
        for c_idx in range(len(df.columns)):
            cell_val = str(df.iloc[r_idx, c_idx])
            if '계약율' in cell_val or '계약률' in cell_val:
                if '※' not in cell_val:
                    match = re.search(r'([\d.]+)\s*%', cell_val)
                    if match:
                        val = float(match.group(1))
                        return val / 100.0 if val > 1.0 else val
                    for offset in range(1, 4):
                        if c_idx + offset < len(df.columns):
                            right_val = df.iloc[r_idx, c_idx + offset]
                            if pd.notna(right_val):
                                parsed = safe_float(right_val)
                                if parsed is not None:
                                    return parsed / 100.0 if parsed > 1.0 else parsed
    try:
        for r in range(12, min(16, len(df))):
            for c in range(len(df.columns)):
                val = df.iloc[r, c]
                if pd.notna(val):
                    val_str = str(val)
                    if '%' in val_str and '※' not in val_str:
                        match = re.search(r'([\d.]+)\s*%', val_str)
                        if match:
                            num = float(match.group(1))
                            if num not in [90.0, 110.0]:
                                return num / 100.0 if num > 1.0 else num
    except:
        pass
    return 0.46

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = excel_file.sheet_names

    st.markdown("---")
    st.markdown("#### ⚙️ 기종별 개별 설정 (CAPA · 현재고 · 프로모션 여부)")

    machine_configs = {}
    for s in sheet_names:
        cols = st.columns([2, 2, 2, 2])
        cols[0].markdown(f"**[{s}]**")
        capa = cols[1].number_input(f"최대 CAPA ({s})", value=10.0, step=0.5, key=f"capa_{s}")
        inv = cols[2].number_input(f"현재고 ({s})", value=19, step=1, key=f"inv_{s}")
        promo = cols[3].checkbox(f"프로모션 방어 ({s})", value=False, key=f"promo_{s}")
        machine_configs[s] = {'capa': capa, 'inventory': inv, 'promo': promo}

    st.markdown("---")

    if st.button("상세 리포트 포함 기종별 맞춤 일괄 분석 실행", type="primary"):
        summary_results = []

        with st.spinner("🚀 분석 진행 중..."):
            for sheet_name, cfg in machine_configs.items():
                try:
                    max_capa = cfg['capa']
                    current_inventory = cfg['inventory']
                    is_promo = cfg['promo']

                    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
                    prev_op_demand = extract_op_demand_from_a23(file_bytes, sheet_name)

                    years_raw = df.iloc[34, 1:].values
                    months_raw = df.iloc[35, 1:].values

                    estimates_raw = df.iloc[36, 1:len(months_raw)+1].values
                    estimates = [safe_float(x) for x in estimates_raw]
                    valid_estimates = [x for x in estimates if x is not None]

                    orders_raw = df.iloc[38, 1:len(months_raw)+1].values
                    orders = [safe_float(x) for x in orders_raw]

                    trend_raw = df.iloc[39, 1:len(months_raw)+1].values
                    trend_orders = [safe_float(x) for x in trend_raw]
                    valid_trend = [x for x in trend_orders if x is not None and x > 0]

                    excel_contract_rate = extract_contract_rate_final(df)

                    current_year = 2024
                    parsed_years = []
                    for y in years_raw:
                        if pd.notna(y):
                            digits = ''.join(filter(str.isdigit, str(y)))
                            if digits:
                                try:
                                    current_year = int(digits)
                                except:
                                    pass
                        parsed_years.append(current_year)

                    valid_data_pairs = []
                    for yr, m_val, ord_val in zip(parsed_years, months_raw, orders):
                        if pd.notna(m_val) and ord_val is not None:
                            m_str = str(m_val).strip()
                            m_num = int(m_str[:-1].strip()) if m_str.endswith('월') else int(float(m_str))
                            if 1 <= m_num <= 12:
                                valid_data_pairs.append((datetime.datetime(yr, m_num, 1), ord_val))

                    valid_data_pairs.sort(key=lambda x: x[0])

                    valid_orders = [x for _, x in valid_data_pairs] if valid_data_pairs else [x for x in orders if x is not None]

                    avg_3m = np.mean(valid_orders[-3:]) if len(valid_orders) >= 3 else (np.mean(valid_orders) if valid_orders else 0.0)
                    avg_6m = np.mean(valid_orders[-6:]) if len(valid_orders) >= 6 else (np.mean(valid_orders) if valid_orders else 0.0)
                    latest_trend = valid_trend[-1] if valid_trend else avg_3m
                    avg_est_3m = np.mean(valid_estimates[-3:]) if len(valid_estimates) >= 3 else (np.mean(valid_estimates) if valid_estimates else 0.0)

                    seasonal_arrow = "➡ 보통"
                    seasonal_desc = "YoY 계절성 보통 (연중 평균 유사)"
                    if valid_data_pairs:
                        latest_dt = valid_data_pairs[-1][0]
                        future_target_month_numbers = []
                        for i in range(round(global_lead_time) - 1, round(global_lead_time) + 2):
                            future_dt = latest_dt + datetime.timedelta(days=int(i * 30.44))
                            future_target_month_numbers.append(future_dt.month)

                        historical_matching_orders = [ord_val for dt, ord_val in valid_data_pairs if dt.month in future_target_month_numbers]
                        overall_historical_avg = np.mean([ord_val for _, ord_val in valid_data_pairs]) if valid_data_pairs else avg_6m

                        if historical_matching_orders and overall_historical_avg > 0:
                            matching_avg = np.mean(historical_matching_orders)
                            if matching_avg > overall_historical_avg * 1.10:
                                seasonal_arrow = "↑ 성수기"
                                seasonal_desc = "YoY 계절성 성수기 (상향 요인)"
                            elif matching_avg < overall_historical_avg * 0.90:
                                seasonal_arrow = "↓ 비수기"
                                seasonal_desc = "YoY 계절성 비수기 (하향 요인)"

                    # -----------------------------------------------------------------
                    # 🔍 상승 모멘텀 판단 세분화 및 사유 추적 로직 추가
                    # -----------------------------------------------------------------
                    is_boom_by_orders = (avg_3m > avg_6m * 1.15) if avg_6m > 0 else False
                    order_growth_pct = ((avg_3m - avg_6m) / avg_6m * 100) if avg_6m > 0 else 0.0

                    past_est_6m = np.mean(valid_estimates[-6:]) if len(valid_estimates) >= 6 else (np.mean(valid_estimates) if valid_estimates else 0.0)
                    is_boom_by_estimates = (avg_est_3m >= past_est_6m * 1.15) if past_est_6m > 0 else False
                    est_growth_pct = ((avg_est_3m - past_est_6m) / past_est_6m * 100) if past_est_6m > 0 else 0.0

                    is_auto_boom = is_boom_by_orders or is_boom_by_estimates

                    # 상승 모멘텀 상세 사유 조합
                    momentum_reasons = []
                    if is_promo:
                        momentum_reasons.append("프로모션 방어 설정 적용")
                    if is_boom_by_orders:
                        momentum_reasons.append(f"단기 수주 급증 (3개월 평균이 6개월 평균 대비 +{order_growth_pct:.1f}% 상승)")
                    if is_boom_by_estimates:
                        momentum_reasons.append(f"견적 추이 급증 (3개월 견적이 과거 평균 대비 +{est_growth_pct:.1f}% 상승)")

                    if momentum_reasons:
                        momentum_status = "O (활성화)"
                        momentum_arrow = "⚡ ON (↑)"
                        momentum_desc = " / ".join(momentum_reasons)
                    else:
                        momentum_status = "X (비활성)"
                        momentum_arrow = "OFF (➡)"
                        momentum_desc = "조건 미충족 (수주 및 견적 안정세)"
                    # -----------------------------------------------------------------

                    if model_choice == 'elastic_capa':
                        raw_demand = (avg_3m * 0.4) + (avg_6m * 0.3) + (latest_trend * 0.3)
                    elif model_choice == 'trend_short':
                        raw_demand = latest_trend
                    else:
                        raw_demand = avg_6m

                    if excel_contract_rate > 0.30: 
                        rate_arrow = "🔥 호조 (↑)"
                        rate_text = f"{excel_contract_rate*100:.1f}% (30% 초과)"
                    elif excel_contract_rate < 0.20: 
                        rate_arrow = "📉 부진 (↓)"
                        rate_text = f"{excel_contract_rate*100:.1f}% (20% 미만)"
                    else:
                        rate_arrow = "📊 보통"
                        rate_text = f"{excel_contract_rate*100:.1f}%"

                    if avg_3m > avg_6m: 
                        trend_arrow = "↑ 상승"
                        trend_text = f"3개월({avg_3m:.2f}) > 6개월({avg_6m:.2f})"
                    elif avg_3m < avg_6m: 
                        trend_arrow = "↓ 하락"
                        trend_text = f"3개월({avg_3m:.2f}) < 6개월({avg_6m:.2f})"
                    else:
                        trend_arrow = "  보합"
                        trend_text = f"3개월({avg_3m:.2f}) = 6개월({avg_6m:.2f})"

                    if is_auto_boom: 
                        trend_arrow += " (붐업)"

                    rate_score = 0.0 if 0.20 <= excel_contract_rate <= 0.30 else (0.03 if excel_contract_rate > 0.30 else -0.05)
                    trend_score = 0.05 if avg_3m > avg_6m else (-0.06 if avg_3m < avg_6m else 0.0)
                    if is_auto_boom: trend_score += 0.05
                    seasonal_score = 0.05 if "성수기" in seasonal_desc else (-0.05 if "비수기" in seasonal_desc else 0.0)

                    base_pct = (raw_demand - prev_op_demand) / prev_op_demand if prev_op_demand > 0 else 0.0
                    total_ai_score = base_pct + rate_score + trend_score + seasonal_score

                    step = 0.05
                    snapped_ratio = round(total_ai_score / step) * step
                    snapped_ratio = max(-0.15, min(0.15, snapped_ratio))

                    if (is_promo or is_auto_boom) and snapped_ratio < -0.05:
                        snapped_ratio = -0.05

                    final_demand_calculated = prev_op_demand * (1.0 + snapped_ratio)
                    final_demand = min(final_demand_calculated, max_capa)

                    required_qty = final_demand * global_lead_time
                    target_inv = math.ceil(required_qty)
                    net_order_qty = max(0, target_inv - current_inventory)

                    summary_results.append({
                        '기종(시트)': sheet_name,
                        '월 최대 CAPA': max_capa,
                        '기존 운영수요': round(prev_op_demand, 2),
                        '변동률': f"{snapped_ratio*100:+.0f}%",
                        '최종 운영수요': round(final_demand, 2),
                        '목표재고': target_inv,
                        '현재고': current_inventory,
                        '🚨 순발주량': f"👉 {net_order_qty} 대" if net_order_qty > 0 else "0 대"
                    })

                    with st.expander(f"📌 [{sheet_name}] 상세 진단 리포트 (최종 확정: {final_demand:.2f}대 / 변동률: {snapped_ratio*100:+.0f}%)"):
                        report_table_data = [
                            {"진단 항목": "기존 운영수요량 (A23)", "지표 내용": f"{prev_op_demand:.2f} 대", "상태 / 방향": "➡ 보합"},
                            {"진단 항목": "최근 수주량 비교", "지표 내용": trend_text, "상태 / 방향": trend_arrow},
                            {"진단 항목": "계약율 분석", "지표 내용": rate_text, "상태 / 방향": rate_arrow},
                            {"진단 항목": "상승 모멘텀 진단 (사유 및 변동치)", "지표 내용": momentum_desc, "상태 / 방향": momentum_arrow},
                            {"진단 항목": "YoY 계절성 성수기 진단", "지표 내용": seasonal_desc, "상태 / 방향": seasonal_arrow},
                            {"진단 항목": "월평균 견적 참고 (3개월)", "지표 내용": f"{avg_est_3m:.1f} 건", "상태 / 방향": "💡 참고"}
                        ]

                        df_report = pd.DataFrame(report_table_data)
                        st.dataframe(df_report, use_container_width=True, hide_index=True)

                        st.markdown(f"🎯 **최종 요약**: 기존 **{prev_op_demand:.2f}대**에서 AI 파라미터 종합 반영 결과 **{final_demand:.2f}대 (`{snapped_ratio*100:+.0f}%`)**로 최종 확정되었습니다. (CAPA 제한: {max_capa}대)")

                except Exception as e:
                    st.error(f"시트 '{sheet_name}' 분석 중 오류 발생: {e}")

            st.markdown("### 📊 기종별 개별 설정 반영 종합 수요 분석표")
            df_summary = pd.DataFrame(summary_results)
            st.dataframe(df_summary, use_container_width=True, hide_index=True)
