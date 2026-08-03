# TOU Revenue Simulator v0.2

## 업로드할 파일 구조

```text
TOU-revenue-simulator/
├─ index.html
├─ app.js
├─ style.css
├─ .nojekyll
└─ data/
   └─ tou_data.xlsx
```

## 최초 업로드

1. ZIP 압축을 풉니다.
2. GitHub 저장소에서 `Add file → Upload files`를 누릅니다.
3. 압축을 푼 폴더 안의 파일과 `data` 폴더를 함께 끌어다 놓습니다.
4. `Commit changes`를 누릅니다.
5. `Settings → Pages → Deploy from a branch → main → /(root) → Save`로 설정합니다.

GitHub 공식 웹 업로드는 파일과 폴더의 드래그앤드롭을 지원합니다.

## 이후 2025·2026년 자료 갱신

1. `tou_data.xlsx`를 엽니다.
2. `산업용`, `산업용(을)` 시트의 마지막 행 아래에 날짜별 데이터를 추가합니다.
3. 파일명을 `tou_data.xlsx`로 유지합니다.
4. GitHub의 `data` 폴더에서 기존 파일을 새 파일로 교체합니다.
5. 웹페이지에서 `GitHub 데이터 다시 불러오기`를 누릅니다.

## 새 종별 추가

1. 사용량 시트를 추가합니다.
2. 머리글을 `날짜 | 연도 | 월 | 계절 | 요일 | H01 ... H24`로 통일합니다.
3. `종별목록` 시트에 행을 추가합니다.
4. `요금표`와 `시간대기준`에 해당 요금 및 시간대 행을 추가합니다.
5. `활성화`를 `Y`로 변경합니다.
6. 같은 이름의 `tou_data.xlsx`로 GitHub에 다시 올립니다.

## 보안

저장소가 Public이면 `data/tou_data.xlsx`도 공개됩니다. 내부자료는 공개 저장소에 업로드하지 마십시오.

## 외부 라이브러리

공식 SheetJS CDN의 `xlsx.full.min.js` 0.20.3을 사용합니다.
