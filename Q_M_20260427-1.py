import streamlit as st
import pandas as pd
import os
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
import io
import numpy as np

# ==========================================
# 0. 페이지 설정
# ==========================================
st.set_page_config(page_title="M스캔 전용 서류 처리 대시보드", layout="wide", page_icon="📊")

# ==========================================
# 1. 전역 설정 & 가이드 내용
# ==========================================
EXCEL_FILE = "insurance_data.xlsx"
TARGET_FILE = "target_settings.xlsx"
APP_PASSWORD = os.environ.get("APP_PASSWORD", "incar961")
MANUAL_FILES = ["모바일동의_독려_안내.pdf", "모바일_보험가입확인서_장기_계피동일건발송절차_v2.pdf"]

GUIDANCE_DOCS = [
    ["No.", "서류명", "법적근거", "목적 및 주요내용"],
    ["1", "개인정보동의서", "개인정보보호법 15조\n대리점 표준 내부통제기준 27조", "개인정보 처리 적법 근거, 보유계약 전산 관리 과정에 따른 개인정보 처리로 신계약시 필수 징구"],
    ["2", "비교설명확인서", "보험업감독규정 별표 5-6", "유사 상품 3개 이상 비교·설명 이행 사실 고객 확인 서명"],
    ["3", "고지의무확인서", "금융소비자보호법 26조와 동법시행령 24조", "판매자 권한·책임·보상 관련 핵심 사항 고지, 소비자 오인 예방"],
    ["4", "완전판매확인서\n(대상: 종신, CI, CEO정기, 고액)", "금융소비자보호법 제17·19조\n영업지원기준안", "약관,청약서 부본 제공, 중요 상품 이해 및 자발적 가입 확인, 설명 의무 이행 증빙력 확보"]
]

MOBILE_GUIDE = {
    "title": "📱 모바일동의(M스캔) 집중 관리 안내",
    "reasons": [
        "**자동 매칭** : 실적 입력 시 서류 자동 연결, 수작업 업로드 불필요, 오류 감소",
        "**타임스탬프** : 전자서명 시점 실시간 기록, 계약 전 작성 객관적 증빙, 서명 위·변조 불가",
        "**누락 방지** : 필수 항목 입력후 다음 단계 진행, 불완전판매 리스크 최소화",
        "**비용 절감** : 비교확인서 스캔 시 5년 원본 보관 의무 해소"
    ],
    "faq": [
        ("스캔 업로드도 가능한가요?", "네, 가능합니다. 다만 자동매칭·타임스탬프·누락방지 기능으로 업무 효율과 법적 보호가 강력한 모바일동의를 권장합니다."),
        ("완전판매확인서는 모든 계약에 필수인가요?", "영업지원기준안대로 종신, CI, CEO정기보험 및 고액 계약(저축성 300만원, 비저축성 100만원 이상)은 필수이며, 이외에는 금소법 취지에 따라 모든 계약에서 활용을 권장합니다."),
        ("모바일동의 절차는 어떻게 진행되나요?", "고객 동의 본인인증 → 설계사 동의 본인인증 → 전자서명 → 타임스탬프 기록 → 실적 입력 매칭\n*(법인, 미성년 계약 외 모바일 동의 전건 가능)*")
    ],
    "do_list": [
        "✅ 계약 체결 전 필수 서류 4종 100% 완비 (미완비 시 불완전판매 및 리스크 계약 간주)",
        "✅ 모바일동의 적극 활용 (자동매칭·타임스탬프 확보)",
        "✅ 사후징구 금지 (2026년 5월부터 서류 미비 시 내부 통제 미충족 조직으로 관리)",
        "✅ 고객 정보 확인 기록 및 적합성 판단 근거 남기기",
        "✅ 대리·중개 고지사항(권한·책임·보상) 사전 안내"
    ]
}

PROCESS_FLOW = [
    {"step": "1", "title": "보험구분 및 계/피동일 설정", "desc": "손보/생보 선택 후 계약자와 피보험자가 동일할 경우 '계/피동일' 체크박스 선택. 체크 시 피보험자 입력칸 자동 생략."},
    {"step": "2", "title": "완전판매확인서 발송 기준 설정", "desc": "필수발송대상: 생보보장성(변액/종신/100만↑), 생보저축성(300만↑), 손보보장/저축성(100만↑). 조건 외 선택적 발송 가능."},
    {"step": "3", "title": "계약자 결재", "desc": "페이퍼 화면 변경 → 동의함/부동의함 선택 → 서명 입력 → 다음 클릭 → 마지막 페이지 상단 저장 클릭."},
    {"step": "4", "title": "피보험자 결재 (계피상이 시)", "desc": "계약자 결재 완료 시 자동 문자 발송 → 화면 확대 가능 → 계약자 결재와 동일한 순서로 진행."},
    {"step": "5", "title": "설계사 결재", "desc": "계약자/피보험자 결재 완료 시 자동 문자 발송 → 글씨 진하게 선택(필요시) → 다음 클릭 → 저장 완료."}
]

# ==========================================
# 2. 데이터 로딩 & 헬퍼
# ==========================================
def safe_rate(num, den):
    return (num / den.replace(0, float('nan')) * 100).round(1).fillna(0.0)

@st.cache_data(ttl=300)
def load_data():
    if not os.path.exists(EXCEL_FILE):
        st.error(f"⚠️ '{EXCEL_FILE}' 파일이 없습니다.")
        return pd.DataFrame()
    try:
        df = pd.read_excel(EXCEL_FILE)
        if df.empty: return pd.DataFrame()
        df.columns = df.columns.str.strip()
        df["보험시작일_dt"] = pd.to_datetime(df["보험시작일"], errors="coerce")
        df["월_피리어드"] = df["보험시작일_dt"].dt.to_period("M").astype(str)
        for col in ["FA고지", "비교설명", "완전판매"]:
            df[f"{col}_c"] = df[col].fillna(" ").astype(str).str.strip()
        def is_total_scanned(val): return not (pd.isna(val) or val == " ") and str(val).strip() in ["스캔", "M스캔", "보험사스캔"]
        def is_m_scanned(val): return not (pd.isna(val) or val == " ") and str(val).strip() == "M스캔"
        def is_cs_target(val): return not (pd.isna(val) or val == " ") and str(val).strip() in ["스캔", "M스캔", "미스캔"]
        df["FA_전체스캔"] = df["FA고지_c"].apply(is_total_scanned).astype(int)
        df["비교_전체스캔"] = df["비교설명_c"].apply(is_total_scanned).astype(int)
        df["완판_전체스캔"] = df["완전판매_c"].apply(is_total_scanned).astype(int)
        df["FA_M스캔"] = df["FA고지_c"].apply(is_m_scanned).astype(int)
        df["비교_M스캔"] = df["비교설명_c"].apply(is_m_scanned).astype(int)
        df["완판_M스캔"] = df["완전판매_c"].apply(is_m_scanned).astype(int)
        df["완판_대상"] = df["완전판매_c"].apply(is_cs_target).astype(int)
        df["FA_target"], df["비교_target"] = 1, 1
        df["완판_target"] = df["완판_대상"]
        df["대상건"] = df[["FA_target", "비교_target", "완판_target"]].sum(axis=1).astype(int)
        df["전체스캔건"] = df[["FA_전체스캔", "비교_전체스캔", "완판_전체스캔"]].sum(axis=1).astype(int)
        df["M스캔건"] = df[["FA_M스캔", "비교_M스캔", "완판_M스캔"]].sum(axis=1).astype(int)
        return df
    except Exception as e:
        st.error(f"❌ 엑셀 파일 읽기 오류: {e}")
        return pd.DataFrame()

def get_file_update_time():
    if os.path.exists(EXCEL_FILE):
        return datetime.fromtimestamp(os.path.getmtime(EXCEL_FILE)).strftime("%Y-%m-%d %H:%M:%S")
    return "알 수 없음"

# ==========================================
# 3. 목표 관리 & 자동배분
# ==========================================
@st.cache_data(ttl=60)
def load_targets():
    if not os.path.exists(TARGET_FILE):
        return pd.DataFrame(columns=["조직단계", "조직명", "M스캔율_목표", "배분사유", "특이사항"])
    try:
        df = pd.read_excel(TARGET_FILE)
        required_cols = ["조직단계", "조직명", "M스캔율_목표", "배분사유", "특이사항"]
        for col in required_cols:
            if col not in df.columns: df[col] = ""
        df["조직단계"] = df["조직단계"].astype(str)
        df["조직명"] = df["조직명"].astype(str)
        df["M스캔율_목표"] = pd.to_numeric(df["M스캔율_목표"], errors="coerce").fillna(0.0)
        df["배분사유"] = df["배분사유"].astype(str)
        df["특이사항"] = df["특이사항"].astype(str)
        return df[required_cols]
    except Exception as e:
        return pd.DataFrame(columns=["조직단계", "조직명", "M스캔율_목표", "배분사유", "특이사항"])

def save_targets(df_targets):
    try:
        required_cols = ["조직단계", "조직명", "M스캔율_목표", "배분사유", "특이사항"]
        for col in required_cols:
            if col not in df_targets.columns: df_targets[col] = ""
        df_targets[required_cols].to_excel(TARGET_FILE, index=False)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"목표 저장 실패: {e}")
        return False

@st.cache_data(ttl=300)
def build_org_stats(df, months=None, group_cols=["영업가족"], view_mode="누적"):
    src = df[df["월_피리어드"].isin(months)].copy() if months else df.copy()
    if src.empty: return pd.DataFrame()
    keys = group_cols.copy()
    if view_mode == "월별": keys = ["월_피리어드"] + keys
    agg_df = src.groupby(keys).agg(대상건=("대상건", "sum"), 전체스캔건=("전체스캔건", "sum"), M스캔건=("M스캔건", "sum")).reset_index()
    agg_df.columns = agg_df.columns.str.strip()
    agg_df["M스캔율_대상"] = safe_rate(agg_df["M스캔건"], agg_df["대상건"])
    agg_df["M스캔율_완료"] = safe_rate(agg_df["M스캔건"], agg_df["전체스캔건"])
    if view_mode == "월별":
        agg_df = agg_df.rename(columns={"월_피리어드": "월"})
        agg_df["월_표시"] = agg_df["월"].apply(lambda x: f"{x.replace('-', '.')[:7]}월" if pd.notna(x) else "")
        agg_df["표시명"] = agg_df["월_표시"] + " | " + agg_df[group_cols[-1]].astype(str)
    else:
        agg_df["월_표시"] = ""
        agg_df["표시명"] = agg_df[group_cols[-1]].astype(str)
    return agg_df

def auto_allocate_targets(df_actual, df_existing, increase_rate=0.10):
    if df_actual.empty or "영업가족" not in df_actual.columns:
        return pd.DataFrame(columns=["조직단계", "조직명", "M스캔율_목표", "배분사유", "특이사항"])
    def calc_targets_for_group(group_df, group_col, group_name):
        agg = group_df.groupby(group_col).agg({"대상건": "sum", "M스캔건": "sum", "전체스캔건": "sum"}).reset_index()
        agg.columns = [group_col, "대상건", "M스캔건", "전체스캔건"]
        agg["현재_실적율"] = safe_rate(agg["M스캔건"], agg["대상건"])
        p30, p70, max_vol = agg["대상건"].quantile(0.3), agg["대상건"].quantile(0.7), agg["대상건"].max()
        results = []
        for _, row in agg.iterrows():
            vol, rate = row["대상건"], row["현재_실적율"]
            if vol < p30: adj, label = 1.15 + min(0.05, (p30 - vol) / max(p30, 1) * 0.05), "소규모(성장유도)"
            elif vol < p70: adj, label = 1.10, "중규모(기준)"
            else: adj, label = 1.05 + (vol - p70) / max(max_vol - p70, 1) * 0.03, "대규모(현실유지)"
            results.append({"조직단계": group_name, "조직명": str(row[group_col]), "M스캔율_목표": round(min(max(rate * adj, 30.0), 95.0), 1), "배분사유": f"{label} | 기본+{int((adj-1)*100)}%", "특이사항": ""})
        return pd.DataFrame(results)
    all_targets = pd.concat([calc_targets_for_group(df_actual, g, g) for g in ["영업가족", "부서", "총괄", "부문"]], ignore_index=True)
    for col in ["조직단계", "조직명", "M스캔율_목표", "배분사유", "특이사항"]:
        if col not in all_targets.columns: all_targets[col] = ""
    if not df_existing.empty and "조직명" in df_existing.columns and "특이사항" in df_existing.columns:
        notes_map = dict(zip(df_existing["조직명"], df_existing["특이사항"]))
        all_targets["특이사항"] = all_targets["조직명"].map(notes_map).fillna(all_targets["특이사항"])
    return all_targets[["조직단계", "조직명", "M스캔율_목표", "배분사유", "특이사항"]]

# ==========================================
# 4. PDF 생성 함수 (✅ 한 페이지 최적화 - 여백 최소화)
# ==========================================
def fig_to_png_bytes(fig, width=600, height=300, scale=2):
    try:
        return pio.to_image(fig, format='png', width=width, height=height, scale=scale)
    except Exception:
        return b''

def generate_agent_report_pdf(title, receiver, reference, sender_dept, dispatcher_name, date_str, recipient_name, table_data, special_notes, actual_rate, target_rate, df_sel=None):
    try:
        pdfmetrics.registerFont(TTFont('NotoSansKR', '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'))
        font_name = 'NotoSansKR'
    except:
        try:
            pdfmetrics.registerFont(TTFont('NanumGothic', '/usr/share/fonts/truetype/nanum/NanumGothic.ttf'))
            font_name = 'NanumGothic'
        except:
            font_name = 'Helvetica'
    
    styles = getSampleStyleSheet()
    # ✅ spaceAfter 최소화 (한 페이지 출력용)
    if 'CustTitle' not in styles: styles.add(ParagraphStyle(name='CustTitle', fontName=font_name, fontSize=15, bold=True, alignment=TA_LEFT, spaceAfter=2*mm))
    if 'CustSubtitle' not in styles: styles.add(ParagraphStyle(name='CustSubtitle', fontName=font_name, fontSize=10, bold=True, alignment=TA_LEFT, spaceAfter=2*mm))
    if 'KoreanText' not in styles: styles.add(ParagraphStyle(name='KoreanText', fontName=font_name, fontSize=8, alignment=TA_LEFT, spaceAfter=1*mm))
    if 'SenderStyle' not in styles: styles.add(ParagraphStyle(name='SenderStyle', fontName=font_name, fontSize=11, bold=True, alignment=TA_CENTER, spaceAfter=1*mm))
        
    pdf_buffer = io.BytesIO()
    # ✅ 마진 최소화 (3mm로 축소)
    pdf_doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, rightMargin=3*mm, leftMargin=3*mm, topMargin=3*mm, bottomMargin=3*mm)
    pdf_elements = []
    
    header_table = Table([['수신: ' + receiver], ['참조: ' + reference]])
    header_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), font_name), 
        ('FONTSIZE', (0, 0), (-1, -1), 9), 
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'), 
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1*mm)
    ]))
    pdf_elements.append(header_table)
    pdf_elements.append(Spacer(1, 2*mm))
    
    pdf_elements.append(Paragraph(title, styles['CustTitle']))
    pdf_elements.append(Spacer(1, 2*mm))
    
    pdf_elements.append(Paragraph(f"【대상 조직】 {recipient_name}", styles['CustSubtitle']))
    pdf_elements.append(Spacer(1, 2*mm))

    # 1. 대시보드 차트 (✅ 높이 50mm로 축소)
    try:
        fig = go.Figure()
        fig.add_trace(go.Bar(name='현황', y=['M스캔율(%)'], x=[actual_rate], orientation='h', marker_color='#3498DB', text=[f"{actual_rate:.1f}%"], textposition='outside'))
        fig.add_trace(go.Bar(name='목표', y=['M스캔율(%)'], x=[target_rate], orientation='h', marker_color='#E74C3C', text=[f"{target_rate:.1f}%"], textposition='outside'))
        fig.add_annotation(text=f"기준: {date_str}", x=0, y=1.1, xref='paper', yref='paper', showarrow=False, font=dict(size=7, color="gray"))
        # ✅ 그래프 높이 50mm로 축소
        fig.update_layout(barmode='group', height=50, margin=dict(l=55, r=5, t=5, b=15), xaxis=dict(range=[0, max(actual_rate, target_rate)*1.2]))
        img_bytes = fig_to_png_bytes(fig, width=320, height=50, scale=2)
        if img_bytes:
            img = RLImage(io.BytesIO(img_bytes), width=100*mm, height=17*mm)
            pdf_elements.append(img)
            pdf_elements.append(Spacer(1, 1*mm))
    except Exception:
        pass
        
    # 2. 대시보드 표
    status_table = Table(table_data, colWidths=[38*mm, 35*mm, 35*mm, 35*mm])
    status_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), font_name), ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2C3E50')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#ECF0F1'), colors.white]),
        ('TOPPADDING', (0, 0), (-1, -1), 1), ('BOTTOMPADDING', (0, 0), (-1, -1), 1)
    ]))
    pdf_elements.append(status_table)
    
    # ✅ 대시보드 표와 FA 집계 사이 간격 축소
    pdf_elements.append(Spacer(1, 2*mm))

    # 3. FA별 이용현황 집계
    fa_col = None
    if df_sel is not None:
        for col in df_sel.columns:
            if "사번" in str(col) or "FA" in str(col).upper() or "AGENT" in str(col).upper():
                fa_col = col
                break
        if fa_col is None and len(df_sel.columns) >= 16:
            fa_col = df_sel.columns[15]

    if fa_col:
        pdf_elements.append(Paragraph("【FA별 M스캔 이용현황 집계】", styles['CustSubtitle']))
        
        # FA사번별 집계
        fa_stats = df_sel.groupby(fa_col).agg(
            총대상건=("대상건", "sum"),
            총M스캔건=("M스캔건", "sum")
        ).reset_index()
        fa_stats["M스캔율"] = safe_rate(fa_stats["총M스캔건"], fa_stats["총대상건"])
        
        # 통계 계산
        total_fa = len(fa_stats)
        using_fa = len(fa_stats[fa_stats["총M스캔건"] > 0])
        not_using_fa = total_fa - using_fa
        using_rate = (using_fa / total_fa * 100) if total_fa > 0 else 0
        
        # ✅ 사용자 평균 M스캔율 (실제로 이용한 FA들만의 평균)
        using_fa_stats = fa_stats[fa_stats["총M스캔건"] > 0]
        if not using_fa_stats.empty:
            avg_m_scan_rate = using_fa_stats["M스캔율"].mean()
        else:
            avg_m_scan_rate = 0
        
        # ✅ 집계 통계 표 (천단위 구분기호 적용)
        stats_data = [
            ['총 FA 수', f'{total_fa:,}명', 'M스캔 사용 FA', f'{using_fa:,}명'],
            ['미사용 FA', f'{not_using_fa:,}명', 'FA 사용률', f'{using_rate:.1f}%'],
            ['사용자 평균 M스캔율', f'{avg_m_scan_rate:.1f}%', '', '']
        ]
        stats_table = Table(stats_data, colWidths=[33*mm, 33*mm, 33*mm, 33*mm])
        stats_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name), ('FONTSIZE', (0, 0), (-1, -1), 7),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#5DADE2')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('TOPPADDING', (0, 0), (-1, -1), 1), ('BOTTOMPADDING', (0, 0), (-1, -1), 1)
        ]))
        pdf_elements.append(stats_table)

    pdf_elements.append(Spacer(1, 2*mm))
        
    # 4. 가이드 텍스트 (축약)
    pdf_elements.append(Paragraph("【모바일동의(M스캔) 안내】", styles['CustSubtitle']))
    for reason in MOBILE_GUIDE["reasons"]:
        pdf_elements.append(Paragraph(f"• {reason}", styles['KoreanText']))
    
    pdf_elements.append(Spacer(1, 1*mm))
    
    # 5. 필수 서류 4종
    pdf_elements.append(Paragraph("【필수 서류 4종 완비 원칙】", styles['CustSubtitle']))
    doc_table_data = [['No.', '서류명', '법적근거']]
    for i, doc in enumerate(GUIDANCE_DOCS[1:], 1):
        doc_table_data.append([str(i), doc[1], doc[2].replace('\n', ' ')])
    
    doc_table = Table(doc_table_data, colWidths=[9*mm, 42*mm, 82*mm])
    doc_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), font_name), ('FONTSIZE', (0, 0), (-1, -1), 6),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#34495E')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#95A5A6')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'), ('TOPPADDING', (0, 0), (-1, -1), 1), ('BOTTOMPADDING', (0, 0), (-1, -1), 1)
    ]))
    pdf_elements.append(doc_table)
    pdf_elements.append(Spacer(1, 1*mm))
    
    # 6. 요청사항
    pdf_elements.append(Paragraph("【요청사항】", styles['CustSubtitle']))
    for i in range(1, 4):
        pdf_elements.append(Paragraph(f"{i}. {['설정된 목표를 반드시 달성하여 주시기 바랍니다.', '모바일동의(M스캔)를 적극 활용하여 업무 효율성과 법적 증빙력을 확보해 주시기 바랍니다.', '2026년 5월부터는 서류 미비 시 내부 통제 미충족 조직으로 관리되오니 각별한 주의 바랍니다.'][i-1]}", styles['KoreanText']))
    
    # 7. 특이사항
    if special_notes and str(special_notes).strip() and str(special_notes).lower() != 'nan':
        pdf_elements.append(Spacer(1, 1*mm))
        special_table = Table([[f"【특이사항】 {special_notes}"]], colWidths=[138*mm])
        special_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 7),
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F9E79F')),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1)
        ]))
        pdf_elements.append(special_table)
    
    # 8. 발신 정보 (중앙 정렬)
    pdf_elements.append(Spacer(1, 2*mm))
    pdf_elements.append(Paragraph(f"{sender_dept}", styles['SenderStyle']))
    pdf_elements.append(Paragraph(f"담당자: {dispatcher_name} (직인생략)", styles['SenderStyle']))
    
    pdf_doc.build(pdf_elements)
    pdf_buffer.seek(0)
    return pdf_buffer

def generate_guide_pdf():
    try:
        pdfmetrics.registerFont(TTFont('NotoSansKR', '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'))
        font_name = 'NotoSansKR'
    except:
        try:
            pdfmetrics.registerFont(TTFont('NanumGothic', '/usr/share/fonts/truetype/nanum/NanumGothic.ttf'))
            font_name = 'NanumGothic'
        except:
            font_name = 'Helvetica'
    
    styles = getSampleStyleSheet()
    if 'GuideTitle' not in styles: styles.add(ParagraphStyle(name='GuideTitle', fontName=font_name, fontSize=14, bold=True, alignment=1, spaceAfter=3*mm))
    if 'GuideSubtitle' not in styles: styles.add(ParagraphStyle(name='GuideSubtitle', fontName=font_name, fontSize=10, bold=True, spaceAfter=1*mm))
    if 'GuideText' not in styles: styles.add(ParagraphStyle(name='GuideText', fontName=font_name, fontSize=8, spaceAfter=1*mm))
        
    pdf_buffer = io.BytesIO()
    pdf_doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    pdf_elements = []
    
    pdf_elements.append(Paragraph("📱 모바일동의(M스캔) 가이드 & 프로세스 요약", styles['GuideTitle']))
    
    doc_table_data = [['No.', '서류명', '법적근거']]
    for i, doc in enumerate(GUIDANCE_DOCS[1:], 1):
        doc_table_data.append([str(i), doc[1], doc[2].replace('\n', ' ')])
    
    doc_table = Table(doc_table_data, colWidths=[12*mm, 50*mm, 70*mm])
    doc_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), font_name), ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2C3E50')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#95A5A6')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'), ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3)
    ]))
    pdf_elements.append(doc_table)
    pdf_elements.append(Spacer(1, 2*mm))
    
    pdf_elements.append(Paragraph("🔄 모바일가입확인서 발송 및 결재 프로세스", styles['GuideSubtitle']))
    process_rows = []
    for step in PROCESS_FLOW:
        process_rows.append([Paragraph(f"Step {step['step']}", styles['GuideText']), Paragraph(f"{step['title']}: {step['desc']}", styles['GuideText'])])
        
    process_table = Table(process_rows, colWidths=[20*mm, 140*mm])
    process_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), font_name), ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2)
    ]))
    pdf_elements.append(process_table)
    pdf_elements.append(Spacer(1, 2*mm))
    
    pdf_elements.append(Paragraph("❓ 자주 묻는 질문(FAQ) 및 안내", styles['GuideSubtitle']))
    faq_items = []
    for q, a in MOBILE_GUIDE["faq"]:
        faq_items.append([Paragraph(f"Q. {q}", styles['GuideText']), Paragraph(f"A. {a}", styles['GuideText'])])
    for reason in MOBILE_GUIDE["reasons"]:
        faq_items.append([Paragraph("💡 안내", styles['GuideText']), Paragraph(reason, styles['GuideText'])])
        
    faq_table = Table(faq_items, colWidths=[30*mm, 130*mm])
    faq_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), font_name), ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [colors.white, colors.HexColor('#F4F6F7')]),
        ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2)
    ]))
    pdf_elements.append(faq_table)
    
    pdf_doc.build(pdf_elements)
    pdf_buffer.seek(0)
    return pdf_buffer

# ==========================================
# 5. UI – 로그인 & 대시보드
# ==========================================
def login_page():
    st.title("🔐 시스템 접속")
    pwd = st.text_input("접속 비밀번호를 입력하세요", type="password")
    if st.button("접속하기", use_container_width=True, type="primary"):
        if pwd == APP_PASSWORD:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")

def dashboard_page():
    st.title("📱 M스캔 전용 서류 처리 현황 대시보드")
    with st.expander("📌 책임판매 필수서류 4종 & 모바일동의 집중 관리 안내 (2026.05 시행)", expanded=False):
        st.markdown("✅ **체결 전 완비 원칙** : 4종 중 단 1개라도 미완비 시 불완전판매 및 리스크 계약 간주\n📱 **모바일동의 표준화** : 자동매칭·타임스탬프·누락방지 기능으로 업무 효율과 법적 증빙력 확보\n⚠️ **사후징구 금지** : 2026년 5월부터 서류 미비 시 내부 통제 미충족 조직으로 관리")

    df = load_data()
    if df.empty:
        st.warning("데이터가 없습니다. insurance_data.xlsx 파일을 확인해주세요.")
        return

    col1, col2 = st.columns([2, 1])
    with col1: st.success(f"총 {len(df):,}건의 데이터 로드 완료")
    with col2: st.info(f"기준: {get_file_update_time()}")

    month_col = "월_피리어드"
    all_months = sorted(df[month_col].dropna().unique())
    st.subheader("분석 기간 선택")
    sel_months = st.multiselect("월 선택", all_months, default=[all_months[-1]] if all_months else [])
    if not sel_months: st.warning("최소 1개 이상의 월을 선택해주세요."); return

    df_sel = df[df[month_col].isin(sel_months)].copy()
    if df_sel.empty: st.info("선택한 기간에 데이터가 없습니다."); return

    st.divider()

    tab_dash, tab_map, tab_target, tab_guide, tab_manual = st.tabs([
        "📊 현황 대시보드", "🗺️ M스캔 활용 현황", "🎯 목표관리 & 공문출력", 
        "📱 가이드 & 프로세스", "📚 매뉴얼 다운로드"
    ])

    # ==========================================
    # 탭 1: 현황 대시보드
    # ==========================================
    with tab_dash:
        ctrl1, ctrl2, ctrl3, ctrl4 = st.columns([1, 1, 1, 1])
        with ctrl1: agg_group = st.selectbox("집계 기준", ["부문", "총괄", "부서", "영업가족"], key="agg_group")
        with ctrl2:
            view_mode_ui = st.radio("보기 방식", ["누적 통합", "월별 비교"], horizontal=True, key="view_mode")
            view_mode = "월별" if view_mode_ui == "월별 비교" else "누적"
        with ctrl3: min_target = st.number_input("🔻 최소 대상건수 필터", min_value=0, step=10, value=10, key="min_target_dash")
        with ctrl4:
            filter_direction = st.radio("🔍 표시 방향", ["상위 N개 (높은 순)", "하위 N개 (낮은 순)"], horizontal=True, key="filter_direction")
            n_display = st.number_input("🔢 표시 개수", min_value=5, step=5, value=20, key="n_display")

        hierarchy = ["부문", "총괄", "부서", "영업가족"]
        idx = hierarchy.index(agg_group) + 1
        group_cols = hierarchy[:idx]
        agg = build_org_stats(df_sel, sel_months, group_cols, view_mode)

        rate_type = st.radio("📊 지표 선", ["M스캔율 (대상대비)", "M스캔율 (완료대비)"], horizontal=True, index=0, key="rate_type")
        compare_type = st.radio("🎯 비교 기준", ["전사 평균대비", "목표치(+10%) 대비"], horizontal=True, index=0, key="compare_type")
        is_target = "대상대비" in rate_type
        rate_col = "M스캔율_대상" if is_target else "M스캔율_완료"
        
        # ✅ 월별/누적 동적 평균 계산
        if view_mode == "월별":
            monthly_stats = df_sel.groupby("월_피리어드").agg(M스캔건=("M스캔건", "sum"), 대상건=("대상건", "sum"))
            monthly_stats["rate"] = safe_rate(monthly_stats["M스캔건"], monthly_stats["대상건"])
            avg_rate_target = monthly_stats["rate"].mean()
            month_key_map = {f"{m.replace('-', '.')[:7]}월": r for m, r in zip(monthly_stats.index.astype(str), monthly_stats["rate"])}
            agg["기준치"] = agg["월_표시"].map(month_key_map).fillna(avg_rate_target)
        else:
            avg_rate_target = safe_rate(pd.Series([int(df_sel["M스캔건"].sum())]), pd.Series([int(df_sel["대상건"].sum())])).iloc[0]
            agg["기준치"] = avg_rate_target

        target_val = round(avg_rate_target * 1.1, 1)
        baseline_val = avg_rate_target if compare_type == "전사 평균대비" else target_val
        baseline_label = "전사 평균" if compare_type == "전사 평균대비" else "목표치(+10%)"

        if min_target > 0: agg = agg[agg["대상건"] >= min_target].copy()
        ascending_sort = (filter_direction == "하위 N개 (낮은 순)")
        agg = agg.sort_values(rate_col, ascending=ascending_sort).head(n_display).reset_index(drop=True)
        
        if view_mode == "월별": agg["순위"] = agg.groupby("월").cumcount() + 1
        else: agg["순위"] = range(1, len(agg) + 1)
        
        agg["기준치"] = baseline_val
        agg["대비_격차"] = (agg[rate_col] - agg["기준치"]).round(1)

        st.markdown(f"### 📈 **{rate_type}** 현황 ({filter_direction}) | 비교 기준: **{baseline_label}** ({baseline_val:.1f}%)")
        met1, met2, met3, met4 = st.columns(4)
        with met1: st.metric("🏢 전사 평균", f"{avg_rate_target:.1f}%")
        with met2: st.metric("🎯 목표치 (+10%)", f"{target_val:.1f}%")
        with met3: st.metric("📊 선 조직 평균", f"{agg[rate_col].mean():.1f}%")
        with met4: st.metric("📐 기준 대비 격차", f"{agg[rate_col].mean() - baseline_val:+.1f}%")
        st.divider()

        display_cols = ["순위"]
        if view_mode == "월별": display_cols.append("월_표시")
        display_cols.extend(group_cols + ["대상건", "전체스캔건", "M스캔건", rate_col, "기준치", "대비_격차"])
        display_cols = [c for c in display_cols if c in agg.columns]

        st.dataframe(
            agg[display_cols].style.format({"대상건":"{:,}", "전체스캔건":"{:,}", "M스캔건":"{:,}", rate_col:"{:.1f}%", "기준치":"{:.1f}%", "대비_격차":"{:+.1f}%"}),
            use_container_width=True, hide_index=True, height=350
        )

        top = agg
        if view_mode == "월별":
            months_order = sorted(top["월_표시"].dropna().unique())
            fig = px.bar(top, x=agg_group, y=rate_col, color="월_표시", barmode="group", title=f"월별 {rate_type} 비교", text_auto=".1f%", category_orders={"월_표시": months_order})
            fig.update_layout(yaxis_range=[0, max(top[rate_col].max() * 1.2, 5)], yaxis_title="M스캔율(%)", legend_title="월", xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
            
            st.subheader("📉 조직별 월간 추이 분석")
            all_orgs = sorted(agg[agg_group].unique())
            trend_orgs = st.multiselect("추이 분석할 조직 선택", all_orgs, default=all_orgs[:5], max_selections=10, key="trend_org_select")
            show_data_labels = st.checkbox("📊 데이터 라벨 표시", value=True, key="show_trend_labels")
            
            m_stats = df_sel.groupby("월_피리어드").agg(M스캔건=("M스캔건", "sum"), 대상건=("대상건", "sum"), 전체스캔건=("전체스캔건", "sum")).reset_index()
            m_stats.columns = m_stats.columns.str.strip()
            m_stats["월_표시"] = m_stats["월_피리어드"].apply(lambda x: f"{x.replace('-', '.')[:7]}월")
            m_stats[rate_col] = safe_rate(m_stats["M스캔건"], m_stats["대상건"] if is_target else m_stats["전체스캔건"])
            m_avg = m_stats[["월_표시", rate_col]][m_stats["월_표시"].isin(months_order)]
            
            if trend_orgs:
                trend_data = agg[agg[agg_group].isin(trend_orgs)].copy()
                max_val = trend_data[rate_col].max()
                y_upper = max(max_val * 1.25, 5)
                if max_val < 10: y_upper = min(y_upper, 20)
                elif max_val < 30: y_upper = min(y_upper, 40)
                else: y_upper = 100

                fig_line = go.Figure()
                months_sorted = sorted(trend_data["월_표시"].dropna().unique())
                for org in trend_orgs:
                    org_data = trend_data[trend_data[agg_group] == org].sort_values("월")
                    txt_vals = [f"{v:.1f}%" for v in org_data[rate_col]] if show_data_labels else None
                    fig_line.add_trace(go.Scatter(x=org_data["월_표시"], y=org_data[rate_col], name=org, mode="lines+markers", text=txt_vals, textposition="top center", hoverinfo="text", hovertext=[f"{org}<br>{row['월_표시']}: {row[rate_col]:.1f}%" for _, row in org_data.iterrows()]))
                
                if not m_avg.empty:
                    txt_avg = [f"{v:.1f}%" for v in m_avg[rate_col]] if show_data_labels else None
                    fig_line.add_trace(go.Scatter(x=m_avg["월_표시"], y=m_avg[rate_col], name="전사 평균", mode="lines+markers", line=dict(color="#000000", width=2, dash="dash"), text=txt_avg, textposition="bottom center", hoverinfo="text", hovertext=[f"전사 평균: {v:.1f}%" for v in m_avg[rate_col]]))
                    
                fig_line.add_hline(y=baseline_val, line_dash="dot", line_color="red", line_width=2, annotation_text=f"{baseline_label} {baseline_val:.1f}%")
                fig_line.update_layout(title="조직별 월간 추이", yaxis_range=[0, y_upper], yaxis_title="M스캔율(%)", xaxis=dict(title="월", tickangle=-45, categoryorder="array", categoryarray=months_sorted), legend=dict(orientation="h", y=1.15))
                st.plotly_chart(fig_line, use_container_width=True)
        else:
            fig = go.Figure()
            fig.add_trace(go.Bar(x=top["표시명"], y=top[rate_col], text=[f"{v:.1f}%" for v in top[rate_col]], textposition="outside", marker_color=top[rate_col]))
            fig.add_hline(y=baseline_val, line_dash="dash", line_color="red", line_width=2, annotation_text=f"{baseline_label} {baseline_val:.1f}%")
            fig.update_layout(title=f"조직별 {rate_type} (정렬: {filter_direction})", xaxis_tickangle=-45, yaxis_range=[0, max(top[rate_col].max() * 1.2, 5)], yaxis_title="M스캔율(%)", height=420)
            st.plotly_chart(fig, use_container_width=True)

    # ==========================================
    # 탭 2: M스캔 활용 현황
    # ==========================================
    with tab_map:
        st.subheader("조직별 M스캔 활용도 분포")
        mc1, mc2, mc3, mc4, mc5 = st.columns([1, 1, 1, 1, 1])
        with mc1: map_level = st.selectbox("집계 단위", ["부문", "총괄", "부서", "영업가족"], key="map_level")
        with mc2: map_type = st.radio("차트 유형", ["가로막대", "파이 차트"], horizontal=True, key="map_type")
        with mc3: map_min_target = st.number_input("🔻 최소 대상건수", min_value=0, step=10, value=10, key="min_target_map")
        with mc4: map_top_n = st.number_input("🔝 표시 개수", min_value=5, step=5, value=20, key="top_n_map")
        with mc5: 
            map_compare = st.radio("🎯 비교 기준", ["전사 평균대비", "목표치(+10%) 대비"], horizontal=True, key="map_compare")
            map_sort = st.radio("🔀 정렬 방향", ["높은 순 (상위)", "낮은 순 (하위)"], horizontal=True, key="map_sort")

        hierarchy_cols = ['부문', '총괄', '부서', '영업가족']
        agg_dict = {col: 'first' for col in hierarchy_cols if col != map_level}
        agg_dict.update({'대상건': 'sum', '전체스캔건': 'sum', 'M스캔건': 'sum'})
        map_agg = df_sel.groupby(map_level).agg(agg_dict).reset_index()
        map_agg.columns = map_agg.columns.str.strip()
        map_agg["M스캔율_대상"] = safe_rate(map_agg["M스캔건"], map_agg["대상건"])
        
        avg_rate_map = safe_rate(pd.Series([int(df_sel["M스캔건"].sum())]), pd.Series([int(df_sel["대상건"].sum())])).iloc[0]
        map_baseline = avg_rate_map if "평균" in map_compare else round(avg_rate_map * 1.1, 1)
        map_agg["격차"] = map_agg["M스캔율_대상"] - map_baseline
        
        if map_min_target > 0: map_agg = map_agg[map_agg["대상건"] >= map_min_target].copy()
        map_agg = map_agg[map_agg["M스캔건"] > 0].sort_values("M스캔율_대상", ascending=(map_sort == "낮은 순 (하위)")).reset_index(drop=True)
        if map_top_n: map_agg = map_agg.head(map_top_n).reset_index(drop=True)
        
        if map_agg.empty:
            st.info("M스캔 활용 데이터가 없습니다.")
        else:
            if map_type == "가로막대":
                fig_bar = px.bar(map_agg, y=map_level, x="M스캔율_대상", orientation="h", color="격차", text_auto=".1f%", color_continuous_scale=["#FF4444", "#888888", "#44CC44"], title="전사 평균/목표 대비 조직별 M스캔율 분포", hover_data=["부문", "총괄", "부서", "영업가족", "대상건", "전체스캔건", "M스캔건"])
                fig_bar.update_layout(height=600, xaxis_title="M스캔율 (%)", yaxis=dict(autorange="reversed"))
                fig_bar.add_vline(x=map_baseline, line_dash="dash", line_color="black", annotation_text=f"기준선 {map_baseline:.1f}%")
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                fig_pie = px.pie(map_agg, values="M스캔건", names=map_level, title="조직별 M스캔 건수 비중", hole=0.4)
                fig_pie.update_traces(textposition="inside", textinfo="percent+label")
                st.plotly_chart(fig_pie, use_container_width=True)
            
            st.dataframe(map_agg[[map_level, "대상건", "M스캔건", "M스캔율_대상", "격차"]].style.format({"대상건":"{:,}", "M스캔건":"{:,}", "M스캔율_대상":"{:.1f}%", "격차":"{:+.1f}%"}), use_container_width=True, hide_index=True)

    # ==========================================
    # 탭 3: 목표관리 & 공문출력
    # ==========================================
    with tab_target:
        st.subheader("🎯 목표 관리 & 공문 출력 워크플로우")
        
        if "doc_preview_ready" not in st.session_state: st.session_state.doc_preview_ready = False
        
        st.markdown("### ① 목표 자동배분 및 파일 저장")
        df_targets = load_targets()
        all_families = sorted(df_sel["영업가족"].dropna().unique()) if "영업가족" in df_sel.columns else []

        if st.button("🔄 자동배분 계산 (영업가족+부서+총괄+부문)", use_container_width=True, type="primary"):
            with st.spinner("🎯 실적규모 분석 및 목표 배분 중..."):
                new_targets = auto_allocate_targets(df_sel, df_targets, increase_rate=0.10)
                if not new_targets.empty:
                    st.session_state["auto_targets"] = new_targets
                    st.success(f"✅ {len(new_targets)}개 조직 목표 자동배분 완료!")
                    st.rerun()
        
        if "auto_targets" in st.session_state:
            st.markdown("#### 📋 자동배분 결과 미리보기")
            st.dataframe(st.session_state["auto_targets"].style.format({"M스캔율_목표": "{:.1f}%"}), use_container_width=True, height=200)
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("💾 목표 파일 저장", key="save_auto_targets"):
                    if save_targets(st.session_state["auto_targets"]):
                        st.success("✅ `target_settings.xlsx` 파일이 저장되었습니다.")
                        del st.session_state["auto_targets"]
                        st.rerun()
            with col_b:
                buf = io.BytesIO()
                st.session_state["auto_targets"].to_excel(buf, index=False)
                buf.seek(0)
                st.download_button("📥 목표 파일 다운로드", data=buf, file_name="target_settings.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        
        st.divider()
        st.markdown("### ② 공문 대상 선정 → 수정 → 미리보기 → 출력")
        
        org_level = st.radio("📊 집계 기준 택", ["영업가족", "부서", "총괄", "부문"], horizontal=True, key="target_org_level")
        level_targets = df_targets[df_targets["조직단계"] == org_level].copy() if not df_targets.empty else pd.DataFrame()
        
        if level_targets.empty:
            unique_orgs = sorted(df_sel[org_level].dropna().unique())
            level_targets = pd.DataFrame({"조직단계": org_level, "조직명": unique_orgs, "M스캔율_목표": 50.0, "배분사유": "신규등록", "특이사항": ""})

        def auto_fill_receiver():
            agent = st.session_state.get("select_agent_for_doc", "")
            receiver = st.session_state.get("receiver_input", "선택 조직명")
            if receiver == "선택 조직명" or receiver == "":
                st.session_state.receiver_input = agent

        st.subheader("📝 공문 정보 입력")
        info_col1, info_col2 = st.columns(2)
        with info_col1:
            selected_agent = st.selectbox(
                "📄 공문 생성 대상 선택", 
                level_targets["조직명"].tolist(), 
                key="select_agent_for_doc",
                on_change=auto_fill_receiver
            )
            receiver_input = st.text_input("📥 수신 (수신자)", value="선택 조직명", key="receiver_input")
            reference_input = st.text_input("📎 참조 (참조자)", value="", key="reference_input")
        with info_col2:
            sender_dept_input = st.text_input("🏢 발신 (부서/조직)", value="지원센터", key="sender_dept_input")
            dispatcher_input = st.text_input("✍️ 발송인 (담당자/성명)", value="", key="dispatcher_input")

        if selected_agent:
            agent_row = level_targets[level_targets["조직명"] == selected_agent].iloc[0].copy()
            agent_stats = build_org_stats(df_sel, sel_months, [org_level], "누적")
            agent_actual = agent_stats[agent_stats[org_level] == selected_agent]
            
            if "edited_doc_df" not in st.session_state or st.session_state.get("last_selected_agent") != selected_agent:
                st.session_state.last_selected_agent = selected_agent
                st.session_state.edited_doc_df = pd.DataFrame([{
                    "조직명": selected_agent,
                    "실제_M스캔율": float(agent_actual["M스캔율_대상"].iloc[0]) if not agent_actual.empty else 0.0,
                    "실제_대상건": int(agent_actual["대상건"].iloc[0]) if not agent_actual.empty else 0,
                    "실제_M스캔건": int(agent_actual["M스캔건"].iloc[0]) if not agent_actual.empty else 0,
                    "목표_M스캔율": float(agent_row["M스캔율_목표"]),
                    "특이사항": str(agent_row["특이사항"])
                }])
            
            edited_df = st.data_editor(
                st.session_state.edited_doc_df,
                column_config={
                    "조직명": st.column_config.TextColumn("조직명", disabled=True),
                    "실제_M스캔율": st.column_config.NumberColumn("실제 M스캔율(%)", disabled=True, format="%.1f%%"),
                    "실제_대상건": st.column_config.NumberColumn("실제 대상건", disabled=True),
                    "실제_M스캔건": st.column_config.NumberColumn("실제 M스캔건", disabled=True),
                    "목표_M스캔율": st.column_config.NumberColumn("목표 M스캔율(%)", min_value=0, max_value=100, format="%.1f%%"),
                    "특이사항": st.column_config.TextColumn("특이사항 (공문 포함)", max_chars=100)
                },
                hide_index=True, use_container_width=True, key="doc_editor"
            )
            
            st.session_state.edited_doc_df = edited_df
            
            col1, col2, col3 = st.columns(3)
            with col1:
                if st.button("💾 수정사항 저장", use_container_width=True):
                    st.success("✅ 수정 내용이 반영되었습니다. 미리보기를 확인해주세요.")
            with col2:
                if st.button("👁️ 미리보기 확인", use_container_width=True):
                    st.session_state.doc_preview_ready = True
                    st.rerun()
            with col3:
                if st.session_state.doc_preview_ready:
                    if st.button("🖨️ 공문 생성 (PDF)", use_container_width=True, type="primary"):
                        with st.spinner("📄 공문 생성 중..."):
                            try:
                                d = st.session_state.edited_doc_df.iloc[0]
                                target_rate = d['목표_M스캔율']
                                actual_vol = d['실제_대상건']
                                additional_needed = int(np.ceil(actual_vol * target_rate / 100)) - d['실제_M스캔건']
                                if additional_needed < 0: additional_needed = 0
                                
                                table_data = [
                                    ['지표', '목표', '현황', '차이'],
                                    ['M스캔율', f"{d['목표_M스캔율']:.1f}%", f"{d['실제_M스캔율']:.1f}%", f"{d['실제_M스캔율']-d['목표_M스캔율']:+.1f}%"],
                                    ['대상건', '-', f"{d['실제_대상건']:,}건", '-'],
                                    ['M스캔건', '-', f"{d['실제_M스캔건']:,}건", '-'],
                                    ['목표달성필요', f"{additional_needed:,}건 추가필요", '-', '-']
                                ]
                                
                                special_notes = d['특이사항']
                                if pd.isna(special_notes) or str(special_notes).lower() == 'nan':
                                    special_notes = ""
                                
                                pdf_buf = generate_agent_report_pdf(
                                    f"M스캔 목표관리 현황 안내 ({org_level})", 
                                    st.session_state.receiver_input, 
                                    st.session_state.reference_input, 
                                    st.session_state.sender_dept_input, 
                                    st.session_state.dispatcher_input, 
                                    datetime.now().strftime("%Y년 %m월 %d일"), 
                                    d['조직명'], 
                                    table_data, 
                                    special_notes, 
                                    d['실제_M스캔율'], 
                                    d['목표_M스캔율'],
                                    df_sel=df_sel
                                )
                                st.download_button("📥 PDF 다운로드", data=pdf_buf, file_name=f"공문_{d['조직명']}_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf", use_container_width=True)
                                st.success("✅ 공문이 생성되었습니다.")
                            except Exception as e:
                                st.error(f"❌ 공문 생성 오류: {e}")
                else:
                    st.button("🖨️ 공문 생성 (PDF)", use_container_width=True, type="primary", disabled=True)

        if st.session_state.doc_preview_ready and "edited_doc_df" in st.session_state and not st.session_state.edited_doc_df.empty:
            st.divider()
            st.markdown("#### 📄 공문 미리보기")
            d = st.session_state.edited_doc_df.iloc[0]
            st.info(f"**수신**: {st.session_state.receiver_input} | **참조**: {st.session_state.reference_input} | **일자**: {datetime.now().strftime('%Y년 %m월 %d일')}")
            
            try:
                fig_preview = go.Figure()
                fig_preview.add_trace(go.Bar(name='현황', y=['M스캔율(%)'], x=[d['실제_M스캔율']], orientation='h', marker_color='#3498DB', text=[f"{d['실제_M스캔율']:.1f}%"], textposition='outside'))
                fig_preview.add_trace(go.Bar(name='목표', y=['M스캔율(%)'], x=[d['목표_M스캔율']], orientation='h', marker_color='#E74C3C', text=[f"{d['목표_M스캔율']:.1f}%"], textposition='outside'))
                fig_preview.update_layout(barmode='group', height=200, margin=dict(l=80, r=20, t=20, b=30), xaxis=dict(range=[0, max(d['실제_M스캔율'], d['목표_M스캔율'])*1.3]))
                st.plotly_chart(fig_preview, use_container_width=True)
            except Exception:
                pass
                
            preview_data = {
                "구분": ["M스캔율", "대상건", "M스캔건", "특이사항"],
                "목표": [f"{d['목표_M스캔율']:.1f}%", "-", "-", d['특이사항'] if d['특이사항'] else "-"],
                "현황": [f"{d['실제_M스캔율']:.1f}%", f"{d['실제_대상건']:,}건", f"{d['실제_M스캔건']:,}건", "-"]
            }
            st.dataframe(pd.DataFrame(preview_data).set_index("구분").T, use_container_width=True)

    # ==========================================
    # 탭 4: 가이드 & 프로세스
    # ==========================================
    with tab_guide:
        st.subheader("🔄 모바일가입확인서 발송 및 결재 프로세스")
        if st.button("📄 가이드/프로세스 PDF 저장 (1페이지)", use_container_width=True, type="primary"):
            with st.spinner("PDF 생성 중..."):
                try:
                    guide_buf = generate_guide_pdf()
                    st.download_button("📥 가이드 PDF 다운로드", data=guide_buf, file_name=f"M스캔_가이드프로세스_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf", use_container_width=True)
                    st.success("✅ 가이드 PDF가 생성되었습니다.")
                except Exception as e:
                    st.error(f"❌ 가이드 PDF 생성 오류: {e}")

        for step in PROCESS_FLOW:
            with st.expander(f"🔹 Step {step['step']}: {step['title']}", expanded=True): st.markdown(step["desc"])
        st.subheader("❓ 자주 묻는 질문(FAQ)")
        for q, a in MOBILE_GUIDE["faq"]: st.markdown(f"**Q. {q}**\n\nA. {a}")
        st.divider()
        st.subheader("📝 책임판매 필수 서류 4종")
        st.dataframe(pd.DataFrame(GUIDANCE_DOCS[1:], columns=GUIDANCE_DOCS[0]).set_index("No."), use_container_width=True, hide_index=True, height=280)
        do_cols = st.columns(2)
        for i, item in enumerate(MOBILE_GUIDE["do_list"][:4]): do_cols[i%2].markdown(item)

    # ==========================================
    # 탭 5: 매뉴얼 다운로드
    # ==========================================
    with tab_manual:
        st.subheader("📚 모바일동의 매뉴얼 다운로드")
        st.divider()
        found = False
        for mf in MANUAL_FILES:
            if os.path.exists(mf):
                found = True
                try:
                    with open(mf, "rb") as f: st.download_button(label=f"📥 {mf}", data=f.read(), file_name=mf, mime="application/pdf", key=f"dl_{mf}", use_container_width=True)
                except Exception as e: st.error(f"❌ {mf} 읽기 오류: {e}")
        if not found: st.warning("⚠️ 매뉴얼 파일이 현재 폴더에 없습니다.")

# ==========================================
# 6. Main
# ==========================================
def main():
    if not st.session_state.get("logged_in"): login_page()
    else:
        with st.sidebar:
            st.success("👋 접속 완료")
            if st.button("🚪 로그아웃", use_container_width=True): st.session_state.logged_in = False; st.rerun()
            st.divider()
            st.caption("v16.2 | 한페이지완벽최적화 | 여백초최소화 | 그래프초축소")
        dashboard_page()

if __name__ == "__main__":
    main()