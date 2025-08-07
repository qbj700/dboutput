# DB 산출물 생성기 v1.0

![DB 산출물 생성기](dboutput.png)

## 📋 프로젝트 소개

**DB 산출물 생성기**는 다양한 데이터베이스에서 테이블 정보를 자동으로 추출하여 Excel 형태의 명세서와 목록을 생성하는 Python GUI 애플리케이션입니다.

### ✨ 주요 특징

- 🎯 **멀티 DBMS 지원**: MySQL, MariaDB, PostgreSQL, Oracle (SID/Service Name 지원)
- 📊 **Excel 자동 생성**: 테이블 명세서와 테이블 목록을 Excel 파일로 출력
- 🔍 **스마트 테이블 선택**: 검색 필터링, 선택 상태 유지, 뷰 테이블 포함
- 🎨 **동적 UI**: Oracle 선택 시 연결 방식 선택 및 윈도우 크기 자동 조정
- 📦 **완전 독립 실행**: JRE 포함 실행파일로 Java 설치 불필요
- 🔐 **보안**: 데이터베이스 연결 정보 보호
- 🎨 **커스텀 아이콘**: 윈도우, 작업표시줄, 파일 아이콘 적용
- 📁 **스마트 폴더 열기**: 생성 완료 후 정확한 폴더 위치로 바로 이동

## 🛠️ 지원 데이터베이스

| DBMS | 버전 | 드라이버 | 연결 방식 | 상태 |
|------|------|----------|----------|------|
| **MySQL** | 5.7+ | PyMySQL | 표준 | ✅ 완전 지원 |
| **MariaDB** | 10.0+ | PyMySQL | 표준 | ✅ 완전 지원 |
| **PostgreSQL** | 9.0+ | psycopg2-binary | 표준 | ✅ 완전 지원 |
| **Oracle** | 11g+ | JDBC (ojdbc8.jar) | SID / Service Name | ✅ 완전 지원 |

## 💻 시스템 요구사항

- **운영체제**: Windows 10/11 (64-bit)
- **메모리**: 최소 4GB RAM
- **디스크 공간**: 100MB 이상
- **네트워크**: 데이터베이스 서버 접근 가능

## 📁 프로젝트 구조

```
dboutput/
├── 📁 database/              # 데이터베이스 연결 모듈
│   ├── __init__.py
│   ├── base_connection.py    # 기본 연결 클래스
│   ├── connection_factory.py # 연결 팩토리
│   ├── connection_manager.py # 연결 관리자
│   ├── data_collector.py     # 메타데이터 수집기
│   ├── mysql_connection.py   # MySQL/MariaDB 연결
│   ├── postgresql_connection.py # PostgreSQL 연결
│   ├── jdbc_oracle_connection.py # Oracle JDBC 연결
│   └── exceptions.py         # 예외 클래스
├── 📁 excel/                 # Excel 생성 모듈
│   ├── __init__.py
│   └── excel_generator.py    # Excel 파일 생성기
├── 📁 gui/                   # GUI 인터페이스
│   ├── __init__.py
│   ├── main_window.py        # 메인 윈도우
│   └── table_selector.py     # 테이블 선택 창
├── 📁 jre/                   # Eclipse Temurin JRE (포함됨)
│   ├── bin/
│   │   ├── java.exe
│   │   └── ...
│   ├── lib/
│   └── ...
├── 📁 dist/                  # 빌드된 실행파일 (배포용)
│   └── DB산출물생성기.exe
├── 📄 main.py                # 애플리케이션 진입점
├── 📄 config.py              # 설정 파일
├── 📄 utils.py               # 유틸리티 함수
├── 📄 requirements.txt       # Python 의존성
├── 📄 .gitignore            # Git 제외 파일 목록
├── 🖼️ dboutput.ico          # 애플리케이션 아이콘
├── 🖼️ dboutput.png          # 프로젝트 로고
├── 📄 ojdbc8.jar            # Oracle JDBC 드라이버
└── 📖 README.md             # 프로젝트 문서
```

## 🚀 설치 및 실행

### 🔧 사전 요구사항

#### 실행파일 사용 시 (권장)
- **없음**: 모든 것이 포함된 완전 독립 실행파일

#### 소스코드 실행 시
- **Python 3.8 이상** ([다운로드](https://www.python.org/downloads/))
- **Git** ([다운로드](https://git-scm.com/downloads))

> 💡 **JRE 및 Oracle JDBC 드라이버**: 프로젝트에 Eclipse Temurin 8과 ojdbc8.jar이 포함되어 있어 별도 설치 불필요!

### 방법 1: 실행파일 사용 (권장)

1. **저장소 클론**
   ```bash
   git clone <repository-url>
   cd dboutput
   ```

2. **실행파일 실행**
   ```bash
   # JRE가 포함된 완전 독립 실행파일
   dist\DB산출물생성기.exe
   ```

> 💡 **완전 독립 실행**: Java 설치 없이도 모든 기능 사용 가능!

### 방법 2: 소스코드에서 직접 실행

1. **저장소 클론**
   ```bash
   git clone <repository-url>
   cd dboutput
   ```

2. **가상환경 생성 및 활성화**
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. **의존성 설치**
   ```bash
   pip install -r requirements.txt
   ```

4. **애플리케이션 실행**
   ```bash
   python main.py
   ```

### 방법 3: 실행파일 재빌드 (개발자용)

#### 빌드 사전 요구사항
- Python 3.8 이상
- PyInstaller 설치: `pip install pyinstaller`
- 프로젝트에 포함된 JRE 및 ojdbc8.jar 확인
- UPX (선택사항): 실행파일 압축 도구 ([다운로드](https://upx.github.io/))
  - 파일 크기를 약 30-50% 추가 압축 가능
  - `upx` 폴더에 압축 해제 후 PATH 추가 또는 `--upx-dir=upx` 옵션 사용

#### 빌드 과정

1. **환경 설정** (방법 2의 1~3단계 동일)

2. **PyInstaller로 빌드**
   ```bash
   # Eclipse Temurin JRE 포함 빌드 (완전 최적화)
   pyinstaller --onefile --windowed --icon=dboutput.ico --add-data "dboutput.ico;." --add-data "dboutput.png;." --add-data "ojdbc8.jar;." --add-data "jre;jre" --hidden-import=pymysql --hidden-import=psycopg2 --hidden-import=jpype1 --hidden-import=cryptography.hazmat.primitives.kdf.pbkdf2 --hidden-import=cryptography.hazmat.primitives.hashes --hidden-import=cryptography.hazmat.primitives.ciphers --hidden-import=cryptography.hazmat.backends.openssl --upx-dir=upx main.py --name "DB산출물생성기"
   
   # 또는 spec 파일 사용 (권장)
   pyinstaller DB산출물생성기.spec
   ```

3. **빌드 결과 확인**
   ```bash
   # 생성된 실행파일 위치 및 크기 확인
   dir dist\DB산출물생성기.exe
   # 예상 크기: ~150MB (Eclipse Temurin JRE 8 포함)
   ```

4. **실행파일 테스트**
   ```bash
   # Java 미설치 환경에서도 실행 가능
   dist\DB산출물생성기.exe
   ```

#### 빌드 특징

**✅ JRE 포함 빌드의 장점:**
- **완전 독립 실행**: Java 설치 불필요
- **모든 기능 지원**: Oracle 포함 모든 DB 연결 가능
- **상업적 배포 허용**: Eclipse Temurin 라이선스
- **사용자 편의성**: 추가 설치 과정 없음

**❌ 고려사항:**
- **파일 크기**: 약 150MB (JRE 8 포함)
- **업데이트**: JRE 업데이트 시 재빌드 필요

### 방법 4: 개발 환경 설정

개발에 참여하거나 코드를 수정하려는 경우:

1. **개발용 의존성 추가 설치**
   ```bash
   pip install pyinstaller
   ```

2. **코드 편집기 설정**
   - **VSCode**: Python 확장 설치 권장
   - **PyCharm**: 프로젝트 인터프리터를 가상환경으로 설정

3. **JRE 확인**
   ```bash
   # 포함된 JRE 버전 확인
   jre\bin\java.exe -version
   # 출력 예시: openjdk version "1.8.0_xxx"
   ```

4. **테스트 실행**
   ```bash
   python -m pytest  # 테스트가 있는 경우
   ```

## 🔧 Oracle 연결 정보

### 연결 구조
```
DB산출물생성기.exe
├── Python 애플리케이션 (GUI/로직)
├── Eclipse Temurin JRE 8 (포함됨)
├── Oracle JDBC Driver (ojdbc8.jar, 포함됨)
└── JPype1 (Python-Java 브릿지)
     ↓
Oracle 데이터베이스
```

### 특징
- ✅ **완전 독립**: 별도 Java 설치 불필요
- ✅ **호환성**: JRE 8로 최대 호환성 보장
- ✅ **간편함**: 실행파일 하나로 모든 기능 지원



## 📖 사용 방법

### 1. 데이터베이스 연결

1. **DBMS 선택**: 드롭다운에서 사용할 데이터베이스 종류 선택
2. **Oracle 연결 방식** (Oracle 선택 시만):
   - **SID**: 시스템 식별자 방식
   - **Service Name**: 서비스명 방식 (권장)
3. **연결 정보 입력**:
   - 서버 주소: DB 서버 IP 또는 도메인
   - 포트 번호: 기본 포트 또는 사용자 지정 포트
   - 데이터베이스명/SID/서비스명: 연결할 데이터베이스 (Oracle은 선택 방식에 따라 라벨 변경)
   - 사용자 ID: DB 접속 계정
   - 비밀번호: DB 접속 암호

4. **연결 테스트**: "연결 테스트" 버튼으로 연결 확인

### 2. 테이블 명세서 생성

1. **"명세서 생성" 버튼 클릭** (연결 테스트 성공 후 활성화)
2. **스마트 테이블 선택**: 
   - **전체 선택/해제**: 기본적으로 모든 테이블 선택됨
   - **검색 기능**: 테이블명으로 실시간 필터링
   - **선택 상태 유지**: 검색 필터링 후에도 이전 선택 상태 유지
   - **뷰 테이블 포함**: 일반 테이블과 뷰 테이블 모두 표시
   - **체크박스**: 개별 테이블 선택/해제
3. **"확인" 버튼으로 생성 시작**
4. **Excel 파일 생성**: `[파일명]_명세서.xlsx` 파일 자동 생성
5. **폴더 열기**: 생성 완료 후 파일 위치로 바로 이동 옵션

### 3. 테이블 목록 생성

1. **"목록 생성" 버튼 클릭** (연결 테스트 성공 후 활성화)
2. **테이블 선택**: 명세서와 동일한 스마트 선택 방식
3. **Excel 파일 생성**: `[파일명]_목록.xlsx` 파일 자동 생성
4. **폴더 열기**: 생성 완료 후 파일 위치로 바로 이동 옵션

### 4. 추가 기능

- **입력값 초기화**: 연결 정보 초기화 (MySQL localhost:3306 기본값)
- **로그 지우기**: 화면 로그 내용 삭제
- **파일 경로 지정**: 생성할 파일의 저장 위치 선택
- **동적 윈도우 크기**: Oracle 선택 시 자동으로 윈도우 크기 확장
- **진행 상태 표시**: 프로그래스 바로 작업 진행률 표시

## 📊 출력 형식

### 테이블 명세서 (Excel)

- **시트명**: "테이블명세서"
- **형식**: 각 테이블별 상세 정보
  - 테이블 정보 (테이블명, 논리명, 테이블 설명)
  - PRIMARY KEY, FOREIGN KEY, INDEX, UNIQUE INDEX 정보
  - 컬럼별 상세 정보 (NO, 컬럼명, 타입, Default, Null 허용, Key, Extra, 설명)
- **스타일**: 맑은 고딕 폰트, 회색 헤더, 테두리 적용
- **정렬**: 컬럼 순서대로 정확한 NO 부여 (1, 2, 3...)

### 테이블 목록 (Excel)

- **시트명**: "테이블목록"  
- **컬럼**: NO, 테이블명, 논리명
- **정렬**: 테이블명 기준 오름차순
- **포함**: 일반 테이블과 뷰 테이블 모두 포함

## 🔧 개발 정보

### 기술 스택

| 구분 | 기술 |
|------|------|
| **언어** | Python 3.12 |
| **GUI** | tkinter (ttk) |
| **Excel** | openpyxl |
| **DB 드라이버** | PyMySQL, psycopg2-binary, JPype1 + JDBC |
| **Java 연동** | JPype1, ojdbc8.jar, 포함된 JRE |
| **패키징** | PyInstaller |
| **이미지 처리** | Pillow |



## 🔍 문제 해결

### 자주 묻는 질문

#### Q1. 프로그램이 실행되지 않아요
**A:** Windows Defender나 백신 프로그램에서 차단될 수 있습니다. 예외 등록하거나 관리자 권한으로 실행해보세요.

#### Q2. Excel 파일 생성이 실패해요
**A:** 동일한 이름의 Excel 파일이 열려있는지 확인하고, 다른 경로로 저장해보세요.

#### Q3. 데이터베이스 연결이 안돼요
**A:** 연결 정보(호스트, 포트, 사용자명, 비밀번호)를 다시 확인하고, 데이터베이스 서버가 실행 중인지 확인하세요.

#### Q4. Oracle Service Name과 SID의 차이는?
**A:** Service Name 방식을 권장합니다. 최신 Oracle에서 더 안정적으로 작동합니다.

## 📝 버전 기록

### v1.0.0 (2024-12-20)
- 🎉 초기 릴리스
- ✅ MySQL, MariaDB, PostgreSQL, Oracle 지원
- ✅ Oracle SID/Service Name 동적 선택
- ✅ 테이블 명세서 및 목록 Excel 생성 (`_명세서`, `_목록` 접미사)
- ✅ 스마트 테이블 선택 (검색, 상태 유지, 뷰 포함)
- ✅ 동적 GUI (Oracle 선택 시 윈도우 크기 자동 조정)
- ✅ 커스텀 아이콘 적용 (윈도우, 작업표시줄, 파일)
- ✅ 스마트 폴더 열기 (생성 완료 후 정확한 위치로 이동)
- ✅ Oracle Default 값 완벽 지원 (JDBC 메타데이터)
- ✅ **완전 독립 실행파일**: Eclipse Temurin JRE 8 + Oracle JDBC 드라이버 포함

## 🤝 기여하기

1. **Fork** 저장소
2. **Feature 브랜치** 생성 (`git checkout -b feature/AmazingFeature`)
3. **커밋** (`git commit -m 'Add some AmazingFeature'`)
4. **브랜치에 Push** (`git push origin feature/AmazingFeature`)
5. **Pull Request** 생성

## 📞 지원 및 문의

- **이슈 리포트**: GitHub Issues
- **기능 요청**: GitHub Issues
- **문서**: 이 README.md 파일

## 📄 라이선스 및 배포 정보

### 프로젝트 라이선스
이 프로젝트는 **MIT 라이선스** 하에 배포됩니다. 자세한 내용은 `LICENSE` 파일을 참조하세요.

### 포함된 구성 요소 라이선스

#### Eclipse Temurin JRE 8
- **라이선스**: GPLv2 + Classpath Exception
- **사용 권한**: 상업적 사용 및 재배포 허용 (완전 무료)
- **배포 방식**: 애플리케이션에 번들링하여 배포

#### Oracle JDBC 드라이버 (ojdbc8.jar)
- **라이선스**: Oracle Technology Network License Agreement
- **재배포 허용**: Oracle DB 연결 목적으로 애플리케이션과 번들링 허용
- **상업적 사용**: 허용 (무료 제공)

---

## 🙏 감사의 말

이 프로젝트는 다음 오픈소스 라이브러리들을 사용합니다:

- **tkinter**: Python 표준 GUI 라이브러리
- **openpyxl**: Excel 파일 생성 및 조작
- **PyMySQL**: MySQL/MariaDB Python 드라이버
- **psycopg2**: PostgreSQL Python 드라이버  
- **JPype1**: Python-Java 연동 라이브러리
- **Oracle JDBC**: ojdbc8.jar 드라이버
- **PyInstaller**: Python 애플리케이션 패키징
- **Pillow**: 이미지 처리 (아이콘 변환)

---

⭐ **이 프로젝트가 유용하다면 Star를 눌러주세요!**