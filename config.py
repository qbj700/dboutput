"""
DB 산출물 생성기 설정 파일
"""

# 애플리케이션 기본 설정
APP_CONFIG = {
    "title": "DB 산출물 생성기 v1.0",
    "version": "1.0.0",
    "author": "DB명세서 생성기",
    "window_size": "700x750",
    "min_window_size": (650, 650)
}

# 지원하는 DBMS 목록 및 기본 포트
SUPPORTED_DBMS = {
    "MySQL": {
        "name": "MySQL",
        "default_port": 3306,
        "driver": "pymysql"
    },
    "MariaDB": {
        "name": "MariaDB", 
        "default_port": 3306,
        "driver": "pymysql"
    },
    "PostgreSQL": {
        "name": "PostgreSQL",
        "default_port": 5432,
        "driver": "psycopg2"
    },
    "Oracle": {
        "name": "Oracle",
        "default_port": 1521,
        "driver": "jaydebeapi"
    }
}

# 기본 파일 설정
FILE_CONFIG = {
    "default_filename": "DB산출물.xlsx",
    "excel_extensions": [".xlsx", ".xls"],
    "encoding": "utf-8"
}

# UI 메시지
UI_MESSAGES = {
    "startup": "DB 산출물 생성기가 시작되었습니다.",
    "ready": "연결 정보를 입력하고 '연결 테스트' 버튼을 눌러 DB 연결을 확인하세요.",
    "connecting": "DB 연결 테스트를 시작합니다...",
    "connected": "✅ DB 연결 테스트가 성공했습니다!",
    "connection_failed": "❌ DB 연결 실패",
    "generating": "테이블 명세서 생성을 시작합니다...",
    "generated": "✅ 테이블 명세서가 성공적으로 생성되었습니다",
    "generation_failed": "❌ 테이블 명세서 생성 실패"
}

# 에러 메시지
ERROR_MESSAGES = {
    "empty_host": "서버 주소를 입력해주세요.",
    "empty_port": "포트 번호를 입력해주세요.", 
    "invalid_port": "올바른 포트 번호를 입력해주세요. (1-65535)",
    "empty_database": "데이터베이스 이름을 입력해주세요.",
    "empty_username": "사용자 ID를 입력해주세요.",
    "empty_filename": "파일명을 입력해주세요.",
    "invalid_filename": "올바른 파일명을 입력해주세요."
}