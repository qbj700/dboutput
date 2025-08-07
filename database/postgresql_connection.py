"""
PostgreSQL 연결 클래스
"""

import psycopg2
import psycopg2.extras
from .base_connection import BaseConnection
from .exceptions import (
    DatabaseConnectionError, DatabaseAuthenticationError,
    DatabaseNotFoundError, DatabaseTimeoutError, DatabaseQueryError
)


class PostgreSQLConnection(BaseConnection):
    """PostgreSQL 연결 클래스"""
    
    def get_dbms_name(self):
        return "PostgreSQL"
        
    def connect(self):
        """PostgreSQL 연결"""
        try:
            # PostgreSQL 연결 문자열 생성
            connection_string = (
                f"host={self.host} "
                f"port={self.port} "
                f"dbname={self.database} "
                f"user={self.username} "
                f"password={self.password} "
                f"connect_timeout={self.timeout}"
            )
            
            self.connection = psycopg2.connect(
                connection_string,
                cursor_factory=psycopg2.extras.RealDictCursor
            )
            self.connection.autocommit = True
            self.is_connected = True
            return True
            
        except psycopg2.OperationalError as e:
            error_msg = str(e).strip()
            
            if "password authentication failed" in error_msg or "authentication failed" in error_msg:
                raise DatabaseAuthenticationError(
                    f"인증 실패: 사용자명 또는 비밀번호가 올바르지 않습니다. ({error_msg})",
                    dbms=self.get_dbms_name(),
                    host=self.host,
                    port=self.port,
                    database=self.database
                )
            elif "database" in error_msg and "does not exist" in error_msg:
                raise DatabaseNotFoundError(
                    f"데이터베이스를 찾을 수 없습니다: '{self.database}' ({error_msg})",
                    dbms=self.get_dbms_name(),
                    host=self.host,
                    port=self.port,
                    database=self.database
                )
            elif "could not connect to server" in error_msg or "Connection refused" in error_msg:
                raise DatabaseConnectionError(
                    f"서버에 연결할 수 없습니다: {self.host}:{self.port} ({error_msg})",
                    dbms=self.get_dbms_name(),
                    host=self.host,
                    port=self.port,
                    database=self.database
                )
            else:
                raise DatabaseConnectionError(
                    f"연결 오류: {error_msg}",
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
        """PostgreSQL 연결 해제"""
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
                return result is not None and result['test'] == 1
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
        """PostgreSQL 버전 정보"""
        try:
            result = self.execute_query("SELECT version() as version")
            if result and len(result) > 0:
                return result[0]['version']
            return "Unknown"
        except Exception:
            return "Unknown"
            
    def get_tables_info(self):
        """테이블 정보 조회"""
        # 단순화된 쿼리 - 기본 테이블과 컬럼 정보만 조회
        query = """
        SELECT 
            t.table_name,
            COALESCE(pgd.description, '') AS 테이블설명,
            col.ordinal_position AS NO,
            col.column_name AS 컬럼명,
            CASE 
                WHEN col.data_type = 'character varying' THEN 'varchar(' || COALESCE(col.character_maximum_length::text, '') || ')'
                WHEN col.data_type = 'character' THEN 'char(' || COALESCE(col.character_maximum_length::text, '') || ')'
                WHEN col.data_type = 'numeric' THEN 'numeric(' || COALESCE(col.numeric_precision::text, '0') || ',' || COALESCE(col.numeric_scale::text, '0') || ')'
                ELSE col.data_type
            END AS TYPE,
            COALESCE(col.column_default, '') AS DEFAULT_VALUE,
            col.is_nullable AS NULLABLE,
            '' AS KEY_TYPE,  -- 별도 쿼리에서 처리
            CASE 
                WHEN col.column_default LIKE 'nextval%' THEN 'auto_increment'
                ELSE ''
            END AS EXTRA,
            '' AS 설명
        FROM 
            information_schema.tables t
        JOIN 
            information_schema.columns col ON t.table_name = col.table_name AND t.table_schema = col.table_schema
        LEFT JOIN 
            pg_class pgc ON pgc.relname = t.table_name
        LEFT JOIN 
            pg_description pgd ON pgd.objoid = pgc.oid AND pgd.objsubid = 0
        WHERE 
            t.table_schema = 'public'
            AND t.table_type IN ('BASE TABLE', 'VIEW')
        ORDER BY 
            t.table_name, col.ordinal_position
        """
        
        try:
            tables_data = self.execute_query(query)
            
            # Primary Key 정보 별도 조회
            pk_query = """
            SELECT 
                ku.table_name,
                ku.column_name
            FROM 
                information_schema.table_constraints tc
            JOIN 
                information_schema.key_column_usage ku ON tc.constraint_name = ku.constraint_name
            WHERE 
                tc.constraint_type = 'PRIMARY KEY' 
                AND tc.table_schema = 'public'
            """
            pk_data = self.execute_query(pk_query)
            pk_columns = {(row['table_name'], row['column_name']) for row in pk_data}
            
            # Foreign Key 정보 별도 조회  
            fk_query = """
            SELECT 
                ku.table_name,
                ku.column_name
            FROM 
                information_schema.table_constraints tc
            JOIN 
                information_schema.key_column_usage ku ON tc.constraint_name = ku.constraint_name
            WHERE 
                tc.constraint_type = 'FOREIGN KEY'
                AND tc.table_schema = 'public'
            """
            fk_data = self.execute_query(fk_query)
            fk_columns = {(row['table_name'], row['column_name']) for row in fk_data}
            
            # KEY_TYPE 설정
            for row in tables_data:
                table_name = row['table_name']
                column_name = row['컬럼명']
                
                if (table_name, column_name) in pk_columns:
                    row['KEY_TYPE'] = 'PRI'
                elif (table_name, column_name) in fk_columns:
                    row['KEY_TYPE'] = 'MUL'
                else:
                    row['KEY_TYPE'] = ''
                    
            return tables_data
            
        except Exception as e:
            # 오류 발생 시 최소한의 정보라도 반환
            basic_query = """
            SELECT 
                t.table_name,
                '' AS 테이블설명,
                col.ordinal_position AS NO,
                col.column_name AS 컬럼명,
                col.data_type AS TYPE,
                COALESCE(col.column_default, '') AS DEFAULT_VALUE,
                col.is_nullable AS NULLABLE,
                '' AS KEY_TYPE,
                '' AS EXTRA,
                '' AS 설명
            FROM 
                information_schema.tables t
            JOIN 
                information_schema.columns col ON t.table_name = col.table_name AND t.table_schema = col.table_schema
            WHERE 
                t.table_schema = 'public'
                AND t.table_type = 'BASE TABLE'
            ORDER BY 
                t.table_name, col.ordinal_position
            """
            return self.execute_query(basic_query)
        
    def get_foreign_keys_info(self):
        """외래키 정보 조회"""
        query = """
        SELECT 
            kcu.table_name,
            kcu.column_name,
            ccu.table_name AS referenced_table_name,
            ccu.column_name AS referenced_column_name,
            tc.constraint_name
        FROM 
            information_schema.table_constraints tc
        JOIN 
            information_schema.key_column_usage kcu ON tc.constraint_name = kcu.constraint_name
        JOIN 
            information_schema.constraint_column_usage ccu ON ccu.constraint_name = tc.constraint_name
        WHERE 
            tc.constraint_type = 'FOREIGN KEY' 
            AND tc.table_schema = 'public'
        ORDER BY 
            kcu.table_name, kcu.column_name
        """
        return self.execute_query(query)
    
    def get_tables_basic_info(self):
        """테이블 기본 정보만 조회 (테이블명, 코멘트)"""
        query = """
        SELECT 
            t.table_name,
            COALESCE(obj_description(pgc.oid), '') as table_comment
        FROM 
            information_schema.tables t
        LEFT JOIN 
            pg_class pgc ON pgc.relname = t.table_name
        LEFT JOIN 
            pg_namespace pgn ON pgn.oid = pgc.relnamespace
        WHERE 
            t.table_schema = 'public'
            AND t.table_type IN ('BASE TABLE', 'VIEW')
            AND (pgn.nspname = 'public' OR pgn.nspname IS NULL)
        ORDER BY 
            t.table_name
        """
        return self.execute_query(query)
    
    def get_indexes_info(self):
        """인덱스 정보 조회"""
        query = """
        SELECT 
            i.tablename as table_name,
            i.indexname as index_name,
            NOT ix.indisunique as non_unique,
            a.attname as column_name,
            a.attnum as seq_in_index
        FROM 
            pg_indexes i
        JOIN 
            pg_class t ON t.relname = i.tablename
        JOIN 
            pg_index ix ON ix.indrelid = t.oid
        JOIN 
            pg_class ic ON ic.oid = ix.indexrelid AND ic.relname = i.indexname  
        JOIN 
            pg_attribute a ON a.attrelid = t.oid AND a.attnum = ANY(ix.indkey)
        WHERE 
            i.schemaname = 'public'
            AND i.indexname NOT LIKE '%_pkey'
        ORDER BY 
            i.tablename, i.indexname, a.attnum
        """
        return self.execute_query(query)