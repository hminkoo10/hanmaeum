# 한마음 체육대회 점수 계산기

본 프로그램은 **교실, 공원강당, 운동장**에서 이루어지는 종목 점수를 입력하고, 자동으로 반별 최종 점수를 합산하여 결과를 출력합니다.

---

## 📦 설치 방법

### 1️⃣ 파이썬 설치

먼저 `install_python.bat` 파일을 실행해 파이썬이 설치되어 있는지 확인하고, 없으면 자동으로 설치 및 최신 버전으로 업데이트합니다.

> 💡 `PATH`에 자동으로 등록되며, pip도 함께 설치됩니다.

---

### 2️⃣ 필수 모듈 설치

필요한 모듈(pandas, openpyxl 등)은 아래 배치 파일을 실행하여 설치합니다:

> `requirements.txt` 파일에는 다음과 같은 모듈이 포함되어 있습니다:
> - pandas
> - openpyxl

---

## ⚙️ 설정 방법

### 3️⃣ `setting.ini` 설정

프로젝트 루트 폴더에 있는 `setting.ini` 파일에서 다음 항목을 수정해야 합니다:

(예시)
```ini
[기본설정]
class=10
```

- `class`: 반의 수

---

### 4️⃣ 종목 설정 파일 작성

`설정/교실.ini`, `설정/공원강당.ini`, `설정/운동장.ini` 등의 파일에서 종목 정보를 설정합니다.

#### 예시: `교실.ini`

```ini
[종목1]
이름=멀리뛰기
높은값우선=yes
1등=100
2등=80
3등=60
...
10등=10
```

- `이름`: 종목명
- `높은값우선`: 점수가 **높을수록 등수가 높은 경우 yes**, 반대는 no
- 각 등수에 맞는 배점을 설정

---

## ✅ 사용 방법

### 5️⃣ 장소별 점수 입력 및 최종 정산

1. `공원강당종목.py`, `교실종목.py`, `운동장종목.py` 중 하나를 실행하여 종목별 점수를 입력합니다.
2. 입력이 완료되면 자동으로 `결과/공원강당.xlsx`, `결과/교실.xlsx`, `결과/운동장.xlsx` 파일이 생성됩니다.
3. **모든 점수를 입력한 후**, `합산.py`를 실행하여 반별 최종 점수를 계산합니다.

- `결과/최종.xlsx` 파일이 생성되며,
  - 반별 종목별 점수
  - 총점
  - 순위 (`X위` 형식)

이 함께 출력됩니다.

---

## 📬 문의

- 개발자: 구현민
- 이메일: hminkoo10@isolvitm.goe.go.kr