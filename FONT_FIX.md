# 한글 폰트 문제 해결 가이드

## 문제 증상
PDF나 Excel 파일에서 한글이 **네모 박스(□□□)**로 표시됩니다.

## 원인
Streamlit Cloud 기본 환경에 한글 폰트가 설치되어 있지 않습니다.

## 해결 방법

### ✅ 방법 1: packages.txt 사용 (권장)

1. **packages.txt 파일을 GitHub에 업로드**
   ```bash
   git add packages.txt
   git commit -m "한글 폰트 설치 설정 추가"
   git push
   ```

2. **Streamlit Cloud가 자동으로 재배포**
   - 1-2분 대기
   - 한글 폰트가 자동 설치됨

3. **확인**
   - 앱 접속 → PDF/Excel 다운로드
   - 한글이 정상적으로 표시되는지 확인

### packages.txt 내용:
```
fonts-noto-cjk
fonts-nanum
```

---

## 배포 시 체크리스트

### 필수 파일 (6개):
- [x] app_final.py
- [x] requirements.txt
- [x] packages.txt ← **한글 폰트용**
- [x] insurance_data.xlsx
- [x] .gitignore
- [x] README_final.md

### GitHub 업로드 명령어:
```bash
git add app_final.py requirements.txt packages.txt insurance_data.xlsx .gitignore README_final.md
git commit -m "한글 폰트 지원 추가"
git push
```

---

## Streamlit Cloud 배포 설정

### 1. 기본 설정
- Main file: `app_final.py`
- Python version: 3.9 이상

### 2. Advanced settings (선택사항)
- Secrets:
  ```toml
  APP_PASSWORD = "your_password"
  ```

### 3. 배포 후 확인
- [ ] 앱이 정상적으로 실행되는가?
- [ ] 로그인이 되는가?
- [ ] 데이터가 로드되는가?
- [ ] PDF 다운로드 시 한글이 정상인가?
- [ ] Excel 다운로드 시 한글이 정상인가?

---

## 문제 해결

### 여전히 한글이 깨진다면:

1. **Streamlit Cloud 로그 확인**
   - Manage app → Logs
   - "fonts-noto-cjk" 설치 확인

2. **앱 재시작**
   - Settings → Reboot app

3. **캐시 클리어**
   - 앱 내 🔄 새로고침 버튼 클릭

4. **packages.txt 확인**
   ```bash
   # GitHub에서 파일 내용 확인
   cat packages.txt
   
   # 결과:
   # fonts-noto-cjk
   # fonts-nanum
   ```

---

## 로컬 테스트 (Windows)

Windows에서는 시스템 폰트를 사용하므로 별도 설정이 필요 없습니다.

```bash
streamlit run app_final.py
```

한글이 정상적으로 표시되어야 합니다.

---

## 로컬 테스트 (Linux/Mac)

### 폰트 설치:
```bash
# Ubuntu/Debian
sudo apt-get install fonts-noto-cjk fonts-nanum

# macOS (Homebrew)
brew tap homebrew/cask-fonts
brew install font-noto-sans-cjk-kr
```

### 앱 실행:
```bash
streamlit run app_final.py
```

---

## 지원되는 폰트

코드가 자동으로 다음 폰트를 순서대로 찾습니다:

1. **Noto Sans CJK** (권장)
   - 경로: `/usr/share/fonts/opentype/noto/`
   - Google의 무료 한글 폰트

2. **나눔고딕**
   - 경로: `/usr/share/fonts/truetype/nanum/`
   - 네이버의 무료 한글 폰트

3. **맑은 고딕** (Windows)
   - 경로: `C:\Windows\Fonts\malgun.ttf`
   - Windows 기본 폰트

---

## FAQ

### Q: packages.txt를 추가했는데도 안 됩니다
A: Streamlit Cloud 재배포를 기다리세요 (1-2분). Settings → Reboot app으로 강제 재시작도 가능합니다.

### Q: 로컬에서는 되는데 배포 후 안 됩니다
A: packages.txt 파일이 GitHub에 제대로 업로드되었는지 확인하세요.

### Q: 영어는 되는데 한글만 깨집니다
A: 정상입니다. packages.txt로 한글 폰트를 설치하면 해결됩니다.

### Q: 일부 한글만 깨집니다
A: 특수 문자나 옛한글일 수 있습니다. Noto Sans CJK가 가장 많은 글자를 지원합니다.

---

## 마무리

packages.txt 파일만 추가하면 **모든 한글 폰트 문제가 자동으로 해결**됩니다!

```bash
# 최종 배포 명령어
git add packages.txt
git commit -m "한글 폰트 지원"
git push
```

배포 후 1-2분 대기하면 완료! 🎉
