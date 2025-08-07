"""
데이터베이스 연결 모듈

이 모듈은 다양한 DBMS(MySQL, MariaDB, PostgreSQL, Oracle)에 대한 
통합된 연결 인터페이스를 제공합니다.

주요 클래스:
- DatabaseConnectionFactory: DBMS별 연결 객체 생성
- ConnectionManager: 연결 관리 및 테스트
- BaseConnection: 모든 DB 연결의 기본 클래스

사용 예시:
    from database import connection_manager
    
    result = connection_manager.test_connection(
        dbms="MySQL",
        host="localhost", 
        port=3306,
        database="test_db",
        username="user",
        password="pass"
    )
"""

from .connection_factory import DatabaseConnectionFactory
from .connection_manager import ConnectionManager, connection_manager
from .data_collector import DatabaseMetadataCollector, metadata_collector
from .exceptions import (
    DatabaseConnectionError,
    DatabaseAuthenticationError, 
    DatabaseNotFoundError,
    DatabaseTimeoutError,
    DatabaseQueryError,
    UnsupportedDatabaseError
)

__all__ = [
    'DatabaseConnectionFactory',
    'ConnectionManager', 
    'connection_manager',
    'DatabaseMetadataCollector',
    'metadata_collector',
    'DatabaseConnectionError',
    'DatabaseAuthenticationError',
    'DatabaseNotFoundError', 
    'DatabaseTimeoutError',
    'DatabaseQueryError',
    'UnsupportedDatabaseError'
]