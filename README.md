# 경비신청서 자동 작성

카드사 **사용내역 엑셀** + 회사 **제경비 양식** + **영수증 사진** → 제출용 완성 `.xls` 생성 (Windows + Excel 필요).

## 권장: PowerShell 앱 (Electron 없음)

Exosphere 등에서 `electron.exe` / `node_modules` 수정이 막히는 환경에서는 아래만 사용하세요.

1. **`경비만들기.bat`** 더블클릭 (또는 `Run-ExpenseApp.bat` — 동일)
2. **카드 사용내역** (.xls) 선택
3. **회사 양식** (.xls) 선택
4. **영수증 폴더** 선택 (선택, 파일명 순으로 첨부·거래 매칭 없음)
5. 기본값(계정·거래처·사용자·업무상세) 확인 후 **완성 엑셀 만들기**
6. 생성된 `.xls`를 사내 포털에 직접 업로드

설정은 `%APPDATA%\ExpenseApp\settings.json`에 저장됩니다.

### 스크립트 구조

| 파일 | 역할 |
|------|------|
| `경비만들기.ps1` | WinForms UI |
| `scripts/read-card-statement.ps1` | 명세서 읽기 (Excel COM) |
| `scripts/fill-expense-form.ps1` | 양식 복사·채우기·영수증 삽입 |
| `scripts/ExpenseCommon.ps1` | 설정·형변환 공통 |

## (선택) Electron 개발 앱

```bash
npm install
npm run dev
```

포트 **5175** (`vite.config.ts`). 일상 사용은 PowerShell 앱을 권장합니다.

## 요구 사항

- Windows
- Microsoft Excel (COM 자동화)
- PowerShell 5.1+

## 동작

- 명세서 첫 시트 → `1. 이용내역명세서`(시트1) 14열 복사 + 노란 열(계정·거래처·사용자·상세·경비/개인금액)
- 영수증 → `1-1. 지출증빙 첨부`(시트2~3) 박스에 그림 삽입