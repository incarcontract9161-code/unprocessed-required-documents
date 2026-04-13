# 📊 보험 서류 스캔 관리 대시보드

GitHub 엑셀 파일 기반 실시간 대시보드 - 로그인 보안 + 자동 배포

## ✨ 주요 기능

- 🔐 단일 비밀번호 로그인
- 📊 실시간 대시보드 (GitHub 엑셀 기반)
- 📈 종합 현황 및 조직별 분석
- 🗺️ 미처리 분포 시각화
- 📋 계층 리포트 (Excel/PDF 다운로드)
- 📄 관리대장 자동 생성 (Excel/PDF)
- 🔄 자동 갱신 (5분마다)

## 🚀 배포 방법

### 1️⃣ GitHub에 파일 업로드

```bash
# 저장소 클론
git clone https://github.com/incarcontract9161-code/unprocessed-required-documents.git
cd unprocessed-required-documents

# 파일 복사:
# - app_final.py (메인 앱)
# - requirements.txt (패키지)
# - insurance_data.xlsx (실제 데이터)

# Git 커밋
git add app_final.py requirements.txt insurance_data.xlsx
git commit -m "Streamlit 앱 최종 배포"
git push
```

### 2️⃣ Streamlit Cloud 배포

1. **https://share.streamlit.io** 접속
2. **"New app"** 클릭
3. 설정:
   - Repository: `incarcontract9161-code/unprocessed-required-documents`
   - Branch: `main`
   - Main file: `app_final.py` ← **중요!**
4. **Advanced settings** 클릭
   - **Secrets** 추가 (선택사항):
     ```toml
     APP_PASSWORD = "your_custom_password"
     ```
   - 설정하지 않으면 기본값 `incar2026` 사용
5. **"Deploy!"** 클릭

### 3️⃣ 완료!

배포된 URL로 접속:
```
https://your-app-name.streamlit.app
```

**로그인:**
- 비밀번호: `incar2026` (기본값)
- 또는 Secrets에서 설정한 비밀번호

---

## 📊 데이터 업데이트 방법

### GitHub에 엑셀 파일만 수정하면 끝!

```bash
# 1. insurance_data.xlsx 파일 수정 (로컬에서)

# 2. GitHub에 커밋
git add insurance_data.xlsx
git commit -m "2024년 3월 데이터 업데이트"
git push

# 3. 완료!
# Streamlit Cloud가 자동 재배포 (1-2분)
# 모든 사용자가 자동으로 최신 데이터 확인
```

**데이터는 5분마다 자동으로 캐시 갱신됩니다.**  
수동 갱신: 화면 상단의 🔄 버튼 클릭

---

## 📁 엑셀 파일 구조

`insurance_data.xlsx` 파일에 다음 컬럼이 필요합니다:

### 필수 컬럼:
```
보험시작일 - YYYY-MM-DD 형식 (예: 2024-01-15)
부문       - 조직 구분
총괄       - 총괄 구분
부서       - 부서명
영업가족   - 영업가족명
담당자     - 담당자 이름
FA고지     - "완료" 또는 "미스캔"
비교설명   - "완료" 또는 "미스캔"
완전판매   - "완료", "미스캔", "해당없음"
```

### 선택 컬럼:
```
보종, 보험사, 상품군, 상품명, 증권번호, 계약자, 
보험료, 보험종료일, 소속, 담당자사번, 개인정보
```

**샘플 파일**: `insurance_data_sample.xlsx` 참고

---

## 🔧 로컬 실행 (테스트용)

```bash
# 패키지 설치
pip install -r requirements.txt

# 앱 실행
streamlit run app_final.py
```

브라우저에서 `http://localhost:8501` 접속

---

## 💡 사용 가이드

### 1. 로그인
- 비밀번호 입력 (기본: `incar2026`)

### 2. 기간 선택
- 왼쪽 사이드바에서 월/부문/총괄 선택
- 여러 월 선택하여 월별 비교 가능

### 3. 대시보드 탭
- **📈 현황 대시보드**: 조직별 미스캔 순위 및 차트
- **🗺️ 미처리맵**: 파이 차트 및 트리맵 시각화
- **📊 계층 리포트**: Excel/PDF 다운로드
- **📋 관리대장 출력**: 부서별 관리대장 생성

### 4. 데이터 다운로드
- 각 탭에서 Excel/PDF 다운로드 가능
- 필터 적용 후 다운로드 가능

---

## 🎯 주요 기능 상세

### 📊 종합 현황
- 전체 통계 (총 계약, 미스캔 건수, 미처리율)
- 조직별 순위 및 차트
- 검색 기능

### 🗺️ 미처리 분포
- 부문/총괄/부서/영업가족별 집계
- 파이 차트 & 트리맵
- 미처리율 시각화

### 📊 계층 리포트
- 부문 → 총괄 → 부서 → 영업가족 계층 구조
- 월별 집계
- Excel/PDF 다운로드

### 📋 관리대장
- 부서별 표지 자동 생성
- 영업가족별 상세 시트
- 서명란 포함

---

## 🔒 보안

### 비밀번호 설정
Streamlit Cloud의 **Secrets** 기능 사용:

1. Streamlit Cloud 앱 대시보드 접속
2. **Settings** → **Secrets** 클릭
3. 다음 내용 입력:
   ```toml
   APP_PASSWORD = "새로운_비밀번호_123"
   ```
4. **Save** 클릭
5. 앱 자동 재시작

### 데이터 보안
- Private 저장소 사용 시 Streamlit Cloud에서 Private 앱으로 자동 설정됨
- 민감한 개인정보는 엑셀 파일에서 제거 권장

---

## 📞 문의

문제 발생 시 GitHub Issues에 등록해주세요.

## 📝 버전

- **v5.0**: GitHub 엑셀 기반 최종 버전
  - SQLite DB 제거
  - 데이터 관리 기능 제거 (GitHub 기반)
  - 실시간 자동 갱신
  - 단일 비밀번호 로그인
  - 모든 분석/리포트 기능 유지

---

## 🛠️ 기술 스택

- **Frontend**: Streamlit
- **Data**: Pandas, Openpyxl
- **Charts**: Plotly
- **Reports**: ReportLab (PDF), Openpyxl (Excel)
- **Deploy**: Streamlit Cloud
- **Storage**: GitHub (엑셀 파일)
