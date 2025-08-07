"""
MySQL/MariaDB 연결 클래스
"""

import pymysql
import pymysql.cursors
from .base_connection import BaseConnection
from .exceptions import (
    DatabaseConnectionError, DatabaseAuthenticationError,
    DatabaseNotFoundError, DatabaseTimeoutError, DatabaseQueryError
)


class MySQLConnection(BaseConnection):
    """MySQL/MariaDB 연결 클래스"""
    
    def get_dbms_name(self):
        return "MySQL/MariaDB"
        
    def connect(self):
        """MySQL/MariaDB 연결"""
        try:
            self.connection = pymysql.connect(
                host=self.host,
                port=self.port,
                user=self.username,
                password=self.password,
                database=self.database,
                charset='utf8mb4',
                connect_timeout=self.timeout,
                read_timeout=self.timeout,
                write_timeout=self.timeout,
                cursorclass=pymysql.cursors.DictCursor,
                autocommit=True
            )
            self.is_connected = True
            return True
            
        except pymysql.err.OperationalError as e:
            error_code = e.args[0] if e.args else 0
            error_msg = e.args[1] if len(e.args) > 1 else str(e)
            
            if error_code == 1045:  # Access denied
                raise DatabaseAuthenticationError(
                    f"인증 실패: 사용자명 또는 비밀번호가 올바르지 않습니다. ({error_msg})",
                    dbms=self.get_dbms_name(),
                    host=self.host,
                    port=self.port,
                    database=self.database
                )
            elif error_code == 1049:  # Unknown database
                raise DatabaseNotFoundError(
                    f"데이터베이스를 찾을 수 없습니다: '{self.database}' ({error_msg})",
                    dbms=self.get_dbms_name(),
                    host=self.host,
                    port=self.port,
                    database=self.database
                )
            elif error_code == 2003:  # Can't connect to MySQL server
                raise DatabaseConnectionError(
                    f"서버에 연결할 수 없습니다: {self.host}:{self.port} ({error_msg})",
                    dbms=self.get_dbms_name(),
                    host=self.host,
                    port=self.port,
                    database=self.database
                )
            else:
                raise DatabaseConnectionError(
                    f"연결 오류: {error_msg} (코드: {error_code})",
                    dbms=self.get_dbms_name(),
                    host=self.host,
                    port=self.port,
                    database=self.database
                )
                
        except Exception as e:
            raise DatabaseConnectionError(
                f"예상치 못한 오류: {str(e)}",
                dbms=self.get_dbms_name(),
                host=self.host,
                port=self.port,
                database=self.database
            )
            
    def disconnect(self):
        """MySQL/MariaDB 연결 해제"""
        if self.connection:
            try:
                self.connection.close()
            except:
                pass  # 연결 해제 중 오류는 무시
            finally:
                self.connection = None
                self.is_connected = False
                
    def test_connection(self):
        """연결 테스트"""
        try:
            with self.get_cursor() as cursor:
                cursor.execute("SELECT 1 as test")
                result = cursor.fetchone()
                return result is not None and result.get('test') == 1
        except Exception as e:
            raise DatabaseConnectionError(f"연결 테스트 실패: {str(e)}")
            
    def execute_query(self, query, params=None):
        """쿼리 실행"""
        try:
            with self.get_cursor() as cursor:
                cursor.execute(query, params)
                
                # SELECT 쿼리인 경우 결과 반환
                if query.strip().upper().startswith('SELECT'):
                    return cursor.fetchall()
                else:
                    return cursor.rowcount
                    
        except Exception as e:
            raise DatabaseQueryError(f"쿼리 실행 오류: {str(e)}", query=query)
            
    def get_version(self):
        """MySQL/MariaDB 버전 정보"""
        try:
            result = self.execute_query("SELECT VERSION() as version")
            if result and len(result) > 0:
                version_str = result[0]['version']
                # MariaDB인지 MySQL인지 구분
                if 'mariadb' in version_str.lower():
                    return f"MariaDB {version_str}"
                else:
                    return f"MySQL {version_str}"
            return "Unknown"
        except Exception:
            return "Unknown"
            
    def get_tables_info(self):
        """테이블 정보 조회"""
        query = """
        SELECT 
            t.TABLE_NAME,
            t.TABLE_COMMENT AS 테이블설명,
            c.ORDINAL_POSITION AS NO,
            c.COLUMN_NAME AS 컬럼명,
            c.COLUMN_TYPE AS TYPE,
            c.COLUMN_DEFAULT AS DEFAULT_VALUE,
            c.IS_NULLABLE AS NULLABLE,
            c.COLUMN_KEY AS KEY_TYPE,
            c.EXTRA AS EXTRA,
            c.COLUMN_COMMENT AS 설명
        FROM 
            INFORMATION_SCHEMA.TABLES t
        JOIN 
            INFORMATION_SCHEMA.COLUMNS c
            ON t.TABLE_NAME = c.TABLE_NAME
            AND t.TABLE_SCHEMA = c.TABLE_SCHEMA
        WHERE 
            t.TABLE_SCHEMA = %s
            AND t.TABLE_TYPE IN ('BASE TABLE', 'VIEW')
        ORDER BY 
            t.TABLE_NAME, c.ORDINAL_POSITION
        """
        return self.execute_query(query, (self.database,))
        
    def get_foreign_keys_info(self):
        """외래키 정보 조회"""
        query = """
        SELECT 
            TABLE_NAME,
            COLUMN_NAME,
            REFERENCED_TABLE_NAME,
            REFERENCED_COLUMN_NAME,
            CONSTRAINT_NAME
        FROM 
            INFORMATION_SCHEMA.KEY_COLUMN_USAGE
        WHERE 
            TABLE_SCHEMA = %s
            AND REFERENCED_TABLE_NAME IS NOT NULL
        ORDER BY TABLE_NAME, COLUMN_NAME
        """
        return self.execute_query(query, (self.database,))
    
    def get_tables_basic_info(self):
        """테이블 기본 정보만 조회 (테이블명, 코멘트)"""
        query = """
        SELECT 
            TABLE_NAME as table_name,
            TABLE_COMMENT as table_comment
        FROM 
            INFORMATION_SCHEMA.TABLES
        WHERE 
            TABLE_SCHEMA = %s
            AND TABLE_TYPE IN ('BASE TABLE', 'VIEW')
        ORDER BY 
            TABLE_NAME
        """
        return self.execute_query(query, (self.database,))
    
    def get_indexes_info(self):
        """인덱스 정보 조회"""
        query = """
        SELECT 
            TABLE_NAME,
            INDEX_NAME,
            NON_UNIQUE,
            COLUMN_NAME,
            SEQ_IN_INDEX
        FROM 
            INFORMATION_SCHEMA.STATISTICS 
        WHERE 
            TABLE_SCHEMA = %s
            AND INDEX_NAME != 'PRIMARY'
        ORDER BY 
            TABLE_NAME, INDEX_NAME, SEQ_IN_INDEX
        """
        return self.execute_query(query, (self.database,))