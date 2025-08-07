"""
데이터베이스 연결 관리자
"""

import threading
import time
from contextlib import contextmanager
from .connection_factory import DatabaseConnectionFactory
from .exceptions import DatabaseConnectionError


class ConnectionManager:
    """데이터베이스 연결 관리자"""
    
    def __init__(self):
        self._connections = {}
        self._lock = threading.Lock()
        
    def get_connection_key(self, dbms, host, port, database, username):
        """연결 키 생성"""
        return f"{dbms}://{username}@{host}:{port}/{database}"
        
    @contextmanager
    def get_connection(self, dbms, host, port, database, username, password, timeout=30, oracle_type=None):
        """
        연결 컨텍스트 매니저
        
        Args:
            dbms (str): DBMS 종류
            host (str): 서버 주소
            port (int): 포트 번호
            database (str): 데이터베이스 이름
            username (str): 사용자명
            password (str): 비밀번호
            timeout (int): 연결 시간 제한
            oracle_type (str): Oracle 연결 방식 ('service_name' 또는 'sid')
            
        Yields:
            BaseConnection: 데이터베이스 연결 객체
        """
        connection = None
        try:
            connection = DatabaseConnectionFactory.create_connection(
                dbms=dbms,
                host=host,
                port=port,
                database=database,
                username=username,
                password=password,
                timeout=timeout,
                oracle_type=oracle_type
            )
            
            connection.connect()
            yield connection
            
        finally:
            if connection:
                connection.disconnect()
                
    def test_connection(self, dbms, host, port, database, username, password, timeout=30, oracle_type=None):
        """
        연결 테스트
        
        Args:
            oracle_type (str): Oracle 연결 방식 ('service_name' 또는 'sid')
        
        Returns:
            dict: 테스트 결과 정보
        """
        start_time = time.time()
        
        try:
            with self.get_connection(dbms, host, port, database, username, password, timeout, oracle_type) as conn:
                # 연결 테스트 실행
                test_result = conn.test_connection()
                
                if test_result:
                    # 버전 정보 가져오기 (추가 정보)
                    try:
                        version = conn.get_version()
                    except:
                        version = "Unknown"
                        
                    end_time = time.time()
                    connection_time = round((end_time - start_time) * 1000, 2)  # 밀리초
                    
                    return {
                        'success': True,
                        'dbms': conn.get_dbms_name(),
                        'version': version,
                        'connection_time_ms': connection_time,
                        'host': host,
                        'port': port,
                        'database': database,
                        'message': f"연결 성공! ({connection_time}ms)"
                    }
                else:
                    return {
                        'success': False,
                        'error': '연결 테스트 쿼리가 실패했습니다.',
                        'dbms': dbms,
                        'host': host,
                        'port': port,
                        'database': database
                    }
                    
        except Exception as e:
            end_time = time.time()
            connection_time = round((end_time - start_time) * 1000, 2)
            
            return {
                'success': False,
                'error': str(e),
                'dbms': dbms,
                'host': host,
                'port': port,
                'database': database,
                'connection_time_ms': connection_time
            }
            
    def get_database_info(self, dbms, host, port, database, username, password, timeout=30):
        """
        데이터베이스 기본 정보 조회
        
        Returns:
            dict: 데이터베이스 정보
        """
        try:
            with self.get_connection(dbms, host, port, database, username, password, timeout) as conn:
                info = conn.get_connection_info()
                info['version'] = conn.get_version()
                return info
                
        except Exception as e:
            raise DatabaseConnectionError(f"데이터베이스 정보 조회 실패: {str(e)}")


# 싱글톤 인스턴스
connection_manager = ConnectionManager()