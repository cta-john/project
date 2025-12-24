# 이미지 캡처 가이드 - KB증권

## 모니터 독립적 자동화를 위한 추가 이미지

코드를 **모니터 환경 변경에도 작동하도록** 수정했습니다!

이제 **3개의 추가 이미지**만 캡처하면 됩니다.

---

## 📸 캡처해야 할 이미지 (3개)

### 1. **ID 입력창** → `id_input_field.png`

**캡처 위치**: 로그인 화면의 아이디 입력창 영역

**캡처 방법**:
1. KB HTS 로그인 화면 띄우기
2. Windows 캡처 도구 실행 (`Win + Shift + S`)
3. 아이디 입력창 부분만 드래그하여 캡처
   - 입력창 테두리를 포함해서 캡처
   - 너무 크지 않게 (입력창과 약간의 주변 영역만)
4. 저장: `C:\HTS_Automation\images\KB\id_input_field.png`

**참고**:
- 입력창 자체가 특징적이면 더 잘 인식됩니다
- 빈 입력창 상태에서 캡처하세요

---

### 2. **화면번호 입력창** → `screen_number_field.png`

**캡처 위치**: 메인 화면 좌측 상단의 화면번호 입력창

**캡처 방법**:
1. KB HTS 메인 화면 진입 (로그인 후)
2. 화면 좌측 상단에 있는 화면번호 입력창 찾기
   - 보통 "화면번호" 또는 빈 입력창 형태
3. Windows 캡처 도구로 입력창 영역 캡처
   - 입력창과 "화면번호" 라벨 포함
4. 저장: `C:\HTS_Automation\images\KB\screen_number_field.png`

**참고**:
- 입력창이 작으면 주변 라벨까지 함께 캡처
- 빈 상태에서 캡처하세요 (숫자 입력 전)

---

### 3. **테이블 영역** → `table_area.png`

**캡처 위치**: 0191 화면의 데이터 테이블 일부

**캡처 방법**:
1. 화면번호 0191 입력 후 화면 진입
2. 테이블이 표시되면, **테이블 헤더 부분** 캡처
   - 예: "종목명", "증거금률", "융자가능" 등의 컬럼 헤더
3. Windows 캡처 도구로 헤더 1~2줄 캡처
   - 너무 많이 캡처하지 말고 특징적인 부분만
4. 저장: `C:\HTS_Automation\images\KB\table_area.png`

**참고**:
- 테이블 헤더는 항상 같은 위치에 있으므로 인식률이 높습니다
- 데이터 행이 아닌 **헤더 행**을 캡처하세요

---

## 💾 저장 위치 확인

모든 이미지는 다음 폴더에 저장:

```
C:\HTS_Automation\images\KB\
```

**최종 파일 목록** (총 17개):
```
C:\HTS_Automation\images\KB\
├── id_tab_active.png          ✅ (이미 있음)
├── id_tab_inactive.png        ✅ (이미 있음)
├── readonly_checked.png       ✅ (이미 있음)
├── readonly_unchecked.png     ✅ (이미 있음)
├── login_button.png           ✅ (이미 있음)
├── readonly_notice.png        ✅ (이미 있음)
├── yes_button.png             ✅ (이미 있음)
├── hts_logo.png               ✅ (이미 있음)
├── download_all_button.png    ✅ (이미 있음)
├── id_input_field.png         🆕 (새로 캡처 필요)
├── screen_number_field.png    🆕 (새로 캡처 필요)
└── table_area.png             🆕 (새로 캡처 필요)
```

---

## ⚡ 빠른 캡처 명령어 (PowerShell)

캡처한 이미지를 빠르게 이동/이름 변경하려면:

```powershell
# 다운로드 폴더에서 KB 이미지 폴더로 이동
Move-Item "$env:USERPROFILE\Downloads\*.png" "C:\HTS_Automation\images\KB\"

# 파일 이름 변경 (필요시)
Rename-Item "C:\HTS_Automation\images\KB\캡처.png" "id_input_field.png"
```

---

## ✅ 캡처 완료 후

3개 이미지 캡처가 완료되면:

1. **파일 존재 확인**:
   ```powershell
   Get-ChildItem "C:\HTS_Automation\images\KB\*.png"
   ```

2. **테스트 실행**:
   ```bash
   python test_kb.py
   ```

---

## 🎯 왜 이미지 기반으로 변경했나?

### ❌ 기존 방식 (좌표 기반)
```python
pyautogui.click(837, 173)  # 모니터 변경 시 작동 안 함
```

### ✅ 새 방식 (이미지 기반)
```python
click_at_image("id_input_field.png")  # 모니터 어디서든 작동
```

**장점**:
- 🖥️ 싱글/듀얼 모니터 자유롭게 전환 가능
- 💻 노트북 도킹/언도킹 시에도 작동
- 📏 해상도 변경 시에도 작동
- 🔧 유지보수 용이

---

## 📝 캡처 팁

1. **선명하게 캡처**: 흐릿하면 인식률 떨어짐
2. **적절한 크기**: 너무 크지도 작지도 않게
3. **빈 상태**: 입력창은 비어있는 상태로
4. **헤더 중심**: 테이블은 헤더 행 캡처
5. **PNG 형식**: 반드시 PNG로 저장

---

## ❓ 문제 해결

### Q: 이미지를 찾지 못한다고 나옴
**A**:
1. 파일명이 정확한지 확인 (영어, 소문자, .png)
2. confidence 값을 0.4로 낮춰보기
3. 이미지 크기를 조금 더 크게 캡처

### Q: 잘못 캡처했을 때
**A**:
1. 해당 이미지 파일 삭제
2. 다시 캡처하여 같은 이름으로 저장

### Q: 캡처 도구가 안 열림
**A**:
- `Win + Shift + S` 대신
- "캡처 도구" 앱 직접 실행
- 또는 "Snipping Tool" 검색

---

**준비되셨으면 3개 이미지 캡처를 시작하세요!** 🚀
