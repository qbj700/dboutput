"""
데이터베이스 연결 팩토리
"""

from .mysql_connection import MySQLConnection
from .postgresql_connection import PostgreSQLConnection
from .jdbc_oracle_connection import JdbcOracleConnection
from .exceptions import UnsupportedDatabaseError


class DatabaseConnectionFactory:
    """데이터베이스 연결 팩토리 클래스"""
    
    # 지원하는 DBMS 매핑
    CONNECTION_CLASSES = {
        'MySQL': MySQLConnection,
        'MariaDB': MySQLConnection,  # MariaDB는 MySQL 드라이버 사용
        'PostgreSQL': PostgreSQLConnection,
        'Oracle': JdbcOracleConnection  # JDBC 방식 Oracle 연결
    }
    
    @classmethod
    def create_connection(cls, dbms, host, port, database, username, password, timeout=30, oracle_type=None):
        """
        DBMS 종류에 따라 적절한 연결 객체 생성
        
        Args:
            dbms (str): DBMS 종류 (MySQL, MariaDB, PostgreSQL, Oracle)
            host (str): 서버 주소
            port (int): 포트 번호
            database (str): 데이터베이스 이름
            username (str): 사용자명
            password (str): 비밀번호
            timeout (int): 연결 시간 제한 (초)
            oracle_type (str): Oracle 연결 방식 ('service_name' 또는 'sid')
            
        Returns:
            BaseConnection: 데이터베이스 연결 객체
            
        Raises:
            UnsupportedDatabaseError: 지원하지 않는 DBMS인 경우
        """
        if dbms not in cls.CONNECTION_CLASSES:
            raise UnsupportedDatabaseError(dbms)
            
        connection_class = cls.CONNECTION_CLASSES[dbms]
        
        # Oracle DBMS인 경우 oracle_type 파라미터 추가
        if dbms == 'Oracle':
            return connection_class(
                host=host,
                port=port,
                database=database,
                username=username,
                password=password,
                timeout=timeout,
                oracle_type=oracle_type or 'service_name'  # 기본값은 service_name
            )
        else:
            return connection_class(
                host=host,
                port=port,
                database=database,
                username=username,
                password=password,
                timeout=timeout
            )
    
    @classmethod
    def get_supported_dbms(cls):
        """지원하는 DBMS 목록 반환"""
        return list(cls.CONNECTION_CLASSES.keys())
    
    @classmethod
    def is_supported(cls, dbms):
        """DBMS 지원 여부 확인"""
        return dbms in cls.CONNECTION_CLASSES