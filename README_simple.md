# 📊 보험 서류 스캔 관리 대시보드

GitHub 엑셀 파일 기반 실시간 대시보드 - 로그인 없이 모든 사용자가 같은 데이터를 볼 수 있습니다.

## ✨ 주요 기능

- 📊 실시간 대시보드 (로그인 불필요)
- 📈 종합 현황 및 조직별 분석
- 📅 월별 트렌드 분석
- 📥 Excel/CSV 다운로드
- 🔄 자동 갱신 (5분마다)

## 🚀 배포 방법 (3단계만!)

### 1️⃣ GitHub에 파일 업로드

```bash
# 저장소 클론
git clone https://github.com/incarcontract9161-code/unprocessed-required-documents.git
cd unprocessed-required-documents

# 다운로드한 파일들 복사:
# - app_simple.py
# - requirements.txt
# - insurance_data.xlsx (엑셀 데이터)

# Git 커밋
git add app_simple.py requirements.txt insurance_data.xlsx
git commit -m "대시보드 배포"
git push
```

### 2️⃣ Streamlit Cloud 배포

1. **https://share.streamlit.io** 접속
2. **"New app"** 클릭
3. 설정:
   - Repository: `incarcontract9161-code/unprocessed-required-documents`
   - Branch: `main`
   - Main file: `app_simple.py`
4. **"Deploy!"** 클릭 → 완료! 🎉

### 3️⃣ 완료!

배포된 URL로 접속하면 바로 대시보드 확인 가능:
```
https://your-app-name.streamlit.app
```

**로그인 없이 누구나 바로 볼 수 있습니다!**

---

## 📊 데이터 업데이트 방법

### 엑셀 파일만 수정하면 끝!

```bash
# 1. insurance_data.xlsx 파일 수정
# 2. GitHub에 커밋
git add insurance_data.xlsx
git commit -m "2024년 3월 데이터 업데이트"
git push

# 3. 끝! Streamlit Cloud가 자동 재배포 (1-2분)
#    모든 사용자가 자동으로 최신 데이터 확인
```

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
streamlit run app_simple.py
```

브라우저에서 `http://localhost:8501` 접속

---

## 💡 사용 팁

### 1. 필터 활용
- 왼쪽 사이드바에서 월/부문/총괄 선택
- 여러 월 선택하여 월별 비교 가능

### 2. 데이터 다운로드
- "상세 데이터" 탭에서 Excel/CSV 다운로드
- 필터 적용 후 다운로드하면 필터링된 데이터만 저장

### 3. 자동 갱신
- 5분마다 자동으로 데이터 새로고침
- 수동 갱신: 화면 상단의 🔄 버튼 클릭

### 4. 기준 날짜 확인
- 화면 상단에 엑셀 파일 마지막 업데이트 시간 표시
- GitHub 커밋 시간과 일치

---

## 🎯 주요 기능

### 📊 종합 현황
- 전체 통계 (총 계약, 미스캔 건수)
- 서류별 미스캔 비율 (파이 차트)
- 월별 트렌드 (라인 차트)

### 📈 조직별 분석
- 부문/총괄/부서/영업가족별 집계
- 미스캔 TOP 10 조회
- 서류별 상세 비교

### 📋 상세 데이터
- 전체/미스캔/완료 필터
- 엑셀/CSV 다운로드
- 실시간 검색

---

## 🔒 보안

- 공개 대시보드이므로 민감한 개인정보는 엑셀 파일에서 제거하세요
- Private 저장소 사용 시 Streamlit Cloud에서 Private 앱으로 설정 가능
- 필요시 로그인 기능 추가 가능 (별도 문의)

---

## 📞 문의

문제 발생 시 GitHub Issues에 등록해주세요.

## 📝 버전

- **v5.0**: GitHub 엑셀 기반 단순 버전
  - 로그인 시스템 제거
  - 실시간 대시보드만 제공
  - 자동 갱신 기능 추가
