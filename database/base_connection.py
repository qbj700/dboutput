"""
데이터베이스 연결 기본 클래스
"""

from abc import ABC, abstractmethod
from contextlib import contextmanager
import time
from .exceptions import DatabaseConnectionError, DatabaseTimeoutError


class BaseConnection(ABC):
    """데이터베이스 연결 기본 클래스"""
    
    def __init__(self, host, port, database, username, password, timeout=30):
        self.host = host
        self.port = port
        self.database = database
        self.username = username
        self.password = password
        self.timeout = timeout
        self.connection = None
        self.is_connected = False
        
    @abstractmethod
    def connect(self):
        """데이터베이스 연결"""
        pass
        
    @abstractmethod
    def disconnect(self):
        """데이터베이스 연결 해제"""
        pass
        
    @abstractmethod
    def test_connection(self):
        """연결 테스트"""
        pass
        
    @abstractmethod
    def execute_query(self, query, params=None):
        """쿼리 실행"""
        pass
        
    @abstractmethod
    def get_dbms_name(self):
        """DBMS 이름 반환"""
        pass
        
    @abstractmethod
    def get_version(self):
        """DB 버전 정보 반환"""
        pass
    
    @abstractmethod
    def get_tables_info(self):
        """테이블 상세 정보 반환 (컬럼 포함)"""
        pass
    
    @abstractmethod
    def get_tables_basic_info(self):
        """테이블 기본 정보만 반환 (테이블명, 코멘트)"""
        pass
    
    @abstractmethod
    def get_foreign_keys_info(self):
        """외래키 정보 반환"""
        pass
    
    @abstractmethod
    def get_indexes_info(self):
        """인덱스 정보 반환"""
        pass
        
    @contextmanager
    def get_cursor(self):
        """커서 컨텍스트 매니저"""
        if not self.is_connected:
            raise DatabaseConnectionError("데이터베이스에 연결되지 않았습니다.")
            
        cursor = None
        try:
            cursor = self.connection.cursor()
            yield cursor
        finally:
            if cursor:
                cursor.close()
                
    def __enter__(self):
        """컨텍스트 매니저 진입"""
        self.connect()
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """컨텍스트 매니저 종료"""
        self.disconnect()
        
    def get_connection_info(self):
        """연결 정보 딕셔너리 반환"""
        return {
            'dbms': self.get_dbms_name(),
            'host': self.host,
            'port': self.port, 
            'database': self.database,
            'username': self.username,
            'is_connected': self.is_connected
        }
        
    def mask_password(self, password=None):
        """비밀번호 마스킹"""
        if password is None:
            password = self.password
        return '*' * len(password) if password else ''