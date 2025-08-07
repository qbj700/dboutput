"""
데이터베이스 연결 관련 커스텀 예외 클래스들
"""


class DatabaseConnectionError(Exception):
    """데이터베이스 연결 오류"""
    def __init__(self, message, dbms=None, host=None, port=None, database=None):
        super().__init__(message)
        self.dbms = dbms
        self.host = host
        self.port = port
        self.database = database
        
    def __str__(self):
        base_msg = super().__str__()
        if self.dbms:
            return f"[{self.dbms}] {base_msg}"
        return base_msg


class DatabaseAuthenticationError(DatabaseConnectionError):
    """데이터베이스 인증 오류"""
    pass


class DatabaseNotFoundError(DatabaseConnectionError):
    """데이터베이스를 찾을 수 없음"""
    pass


class DatabaseTimeoutError(DatabaseConnectionError):
    """데이터베이스 연결 시간 초과"""
    pass


class DatabaseQueryError(Exception):
    """데이터베이스 쿼리 실행 오류"""
    def __init__(self, message, query=None):
        super().__init__(message)
        self.query = query


class UnsupportedDatabaseError(Exception):
    """지원하지 않는 데이터베이스"""
    def __init__(self, dbms):
        super().__init__(f"지원하지 않는 DBMS입니다: {dbms}")
        self.dbms = dbms