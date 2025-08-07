"""
Oracle 연결 클래스
"""

import oracledb
from .base_connection import BaseConnection
from .exceptions import (
    DatabaseConnectionError, DatabaseAuthenticationError,
    DatabaseNotFoundError, DatabaseTimeoutError, DatabaseQueryError
)


class OracleConnection(BaseConnection):
    """Oracle 연결 클래스"""
    
    def __init__(self, host, port, database, username, password, timeout=30, oracle_type='service_name'):
        """
        Oracle 연결 초기화
        
        Args:
            oracle_type (str): Oracle 연결 방식 ('service_name' 또는 'sid')
        """
        super().__init__(host, port, database, username, password, timeout)
        self.oracle_type = oracle_type
    
    def get_dbms_name(self):
        return "Oracle"
        
    def connect(self):
        """Oracle 연결"""
        # Oracle DSN 생성 (연결 방식에 따라)
        if self.oracle_type == 'sid':
            # SID 방식: 단순 문자열 형식 (thin mode 호환성)
            dsn = f"{self.host}:{self.port}:{self.database}"
        else:
            # Service Name 방식: host:port/service_name (기본값)
            dsn = f"{self.host}:{self.port}/{self.database}"
        
        # 1단계: Thin Mode로 먼저 시도 (Oracle Client 설치 불필요)
        try:
            print("Oracle thin mode로 연결을 시도합니다...")
            self.connection = oracledb.connect(
                user=self.username,
                password=self.password,
                dsn=dsn,
                tcp_connect_timeout=self.timeout
            )
            self.connection.autocommit = True
            self.is_connected = True
            print("Oracle thin mode 연결 성공!")
            return True
            
        except oracledb.DatabaseError as thin_error:
            thin_error_msg = str(thin_error)
            print(f"Oracle thin mode 연결 실패: {thin_error_msg}")
            
            # DPY-3010 (버전 지원 안 됨) 또는 기타 thin mode 오류 시 thick mode 시도
            if "DPY-3010" in thin_error_msg or "thin mode" in thin_error_msg.lower():
                print("Oracle thick mode로 재시도합니다...")
                
                # 2단계: Thick Mode 초기화 및 연결 시도
                try:
                    oracledb.init_oracle_client()
                    print("Oracle thick mode 초기화 성공!")
                    
                    self.connection = oracledb.connect(
                        user=self.username,
                        password=self.password,
                        dsn=dsn,
                        tcp_connect_timeout=self.timeout
                    )
                    self.connection.autocommit = True
                    self.is_connected = True
                    print("Oracle thick mode 연결 성공!")
                    return True
                    
                except Exception as thick_error:
                    print(f"Oracle thick mode도 실패: {thick_error}")
                    # thick mode도 실패하면 원래 thin mode 오류를 던짐
                    raise thin_error
            else:
                # thin mode 오류이지만 thick mode로 해결할 수 없는 오류 (인증, 네트워크 등)
                raise thin_error
            
        except Exception as e:
            raise DatabaseConnectionError(
                f"예상치 못한 오류: {str(e)}",
                dbms=self.get_dbms_name(),
                host=self.host,
                port=self.port,
                database=self.database
            )
            
    def disconnect(self):
        """Oracle 연결 해제"""
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
                cursor.execute("SELECT 1 FROM DUAL")
                result = cursor.fetchone()
                return result is not None and result[0] == 1
        except Exception as e:
            raise DatabaseConnectionError(f"연결 테스트 실패: {str(e)}")
            
    def execute_query(self, query, params=None):
        """쿼리 실행"""
        try:
            with self.get_cursor() as cursor:
                cursor.execute(query, params or {})
                
                # SELECT 쿼리인 경우 결과 반환
                if query.strip().upper().startswith('SELECT'):
                    # Oracle은 DictCursor가 없으므로 컬럼명과 함께 딕셔너리로 변환
                    columns = [desc[0] for desc in cursor.description]
                    rows = cursor.fetchall()
                    return [dict(zip(columns, row)) for row in rows]
                else:
                    return cursor.rowcount
                    
        except Exception as e:
            raise DatabaseQueryError(f"쿼리 실행 오류: {str(e)}", query=query)
            
    def get_version(self):
        """Oracle 버전 정보"""
        try:
            result = self.execute_query("SELECT banner FROM v$version WHERE ROWNUM = 1")
            if result and len(result) > 0:
                return result[0]['BANNER']
            return "Unknown"
        except Exception:
            return "Unknown"
            
    def get_tables_info(self):
        """테이블 정보 조회"""
        # 먼저 ALL_TABLES로 시도 (접근 가능한 모든 테이블)
        query_all = """
        SELECT 
            t.table_name as "table_name",
            tc.comments AS 테이블설명,
            c.column_id AS NO,
            c.column_name AS 컬럼명,
            CASE 
                WHEN c.data_type IN ('VARCHAR2', 'CHAR', 'NVARCHAR2', 'NCHAR') 
                THEN c.data_type || '(' || c.data_length || ')'
                WHEN c.data_type = 'NUMBER' AND c.data_precision IS NOT NULL
                THEN c.data_type || '(' || c.data_precision || 
                     CASE WHEN c.data_scale > 0 THEN ',' || c.data_scale ELSE '' END || ')'
                ELSE c.data_type
            END AS TYPE,
            c.data_default AS DEFAULT_VALUE,
            c.nullable AS NULLABLE,
            CASE 
                WHEN pk.column_name IS NOT NULL THEN 'PRI'
                WHEN fk.column_name IS NOT NULL THEN 'MUL'
                ELSE ''
            END AS KEY_TYPE,
            '' AS EXTRA,
            cc.comments AS 설명
        FROM 
            (SELECT owner, table_name FROM all_tables WHERE owner = USER
             UNION ALL 
             SELECT owner, view_name as table_name FROM all_views WHERE owner = USER) t
        JOIN 
            all_tab_columns c ON t.table_name = c.table_name AND t.owner = c.owner
        LEFT JOIN 
            all_tab_comments tc ON t.table_name = tc.table_name AND t.owner = tc.owner
        LEFT JOIN 
            all_col_comments cc ON c.table_name = cc.table_name AND c.column_name = cc.column_name AND c.owner = cc.owner
        LEFT JOIN 
            (SELECT col.table_name, col.column_name
             FROM all_constraints con
             JOIN all_cons_columns col ON con.constraint_name = col.constraint_name AND con.owner = col.owner
             WHERE con.constraint_type = 'P' AND con.owner = USER) pk
            ON c.table_name = pk.table_name AND c.column_name = pk.column_name
        LEFT JOIN 
            (SELECT col.table_name, col.column_name
             FROM all_constraints con
             JOIN all_cons_columns col ON con.constraint_name = col.constraint_name AND con.owner = col.owner
             WHERE con.constraint_type = 'R' AND con.owner = USER) fk
            ON c.table_name = fk.table_name AND c.column_name = fk.column_name
        ORDER BY 
            t.table_name, c.column_id
        """
        
        try:
            return self.execute_query(query_all)
        except Exception:
            # ALL_TABLES 접근 실패 시 USER_TABLES로 fallback
            query_user = """
            SELECT 
                t.table_name as "table_name",
                tc.comments AS 테이블설명,
                c.column_id AS NO,
                c.column_name AS 컬럼명,
                CASE 
                    WHEN c.data_type IN ('VARCHAR2', 'CHAR', 'NVARCHAR2', 'NCHAR') 
                    THEN c.data_type || '(' || c.data_length || ')'
                    WHEN c.data_type = 'NUMBER' AND c.data_precision IS NOT NULL
                    THEN c.data_type || '(' || c.data_precision || 
                         CASE WHEN c.data_scale > 0 THEN ',' || c.data_scale ELSE '' END || ')'
                    ELSE c.data_type
                END AS TYPE,
                c.data_default AS DEFAULT_VALUE,
                c.nullable AS NULLABLE,
                CASE 
                    WHEN pk.column_name IS NOT NULL THEN 'PRI'
                    WHEN fk.column_name IS NOT NULL THEN 'MUL'
                    ELSE ''
                END AS KEY_TYPE,
                '' AS EXTRA,
                cc.comments AS 설명
            FROM 
                (SELECT table_name FROM user_tables 
                 UNION ALL 
                 SELECT view_name as table_name FROM user_views) t
            JOIN 
                user_tab_columns c ON t.table_name = c.table_name
            LEFT JOIN 
                user_tab_comments tc ON t.table_name = tc.table_name
            LEFT JOIN 
                user_col_comments cc ON c.table_name = cc.table_name AND c.column_name = cc.column_name
            LEFT JOIN 
                (SELECT col.table_name, col.column_name
                 FROM user_constraints con
                 JOIN user_cons_columns col ON con.constraint_name = col.constraint_name
                 WHERE con.constraint_type = 'P') pk
                ON c.table_name = pk.table_name AND c.column_name = pk.column_name
            LEFT JOIN 
                (SELECT col.table_name, col.column_name
                 FROM user_constraints con
                 JOIN user_cons_columns col ON con.constraint_name = col.constraint_name
                 WHERE con.constraint_type = 'R') fk
                ON c.table_name = fk.table_name AND c.column_name = fk.column_name
            ORDER BY 
                t.table_name, c.column_id
            """
            return self.execute_query(query_user)
        
    def get_foreign_keys_info(self):
        """외래키 정보 조회"""
        # 먼저 ALL_CONSTRAINTS로 시도
        query_all = """
        SELECT 
            col.table_name as "table_name",
            col.column_name as "column_name",
            r_col.table_name AS "referenced_table_name",
            r_col.column_name AS "referenced_column_name",
            con.constraint_name as "constraint_name"
        FROM 
            all_constraints con
        JOIN 
            all_cons_columns col ON con.constraint_name = col.constraint_name AND con.owner = col.owner
        JOIN 
            all_cons_columns r_col ON con.r_constraint_name = r_col.constraint_name AND con.r_owner = r_col.owner
        WHERE 
            con.constraint_type = 'R' AND con.owner = USER
        ORDER BY 
            col.table_name, col.column_name
        """
        
        try:
            return self.execute_query(query_all)
        except Exception:
            # ALL_CONSTRAINTS 접근 실패 시 USER_CONSTRAINTS로 fallback
            query_user = """
            SELECT 
                col.table_name as "table_name",
                col.column_name as "column_name",
                r_col.table_name AS "referenced_table_name",
                r_col.column_name AS "referenced_column_name",
                con.constraint_name as "constraint_name"
            FROM 
                user_constraints con
            JOIN 
                user_cons_columns col ON con.constraint_name = col.constraint_name
            JOIN 
                user_cons_columns r_col ON con.r_constraint_name = r_col.constraint_name
            WHERE 
                con.constraint_type = 'R'
            ORDER BY 
                col.table_name, col.column_name
            """
            return self.execute_query(query_user)
    
    def get_tables_basic_info(self):
        """테이블 기본 정보만 조회 (테이블명, 코멘트)"""
        # 먼저 ALL_TABLES로 시도 (접근 가능한 모든 테이블)
        query_all = """
        SELECT 
            t.table_name as "table_name",
            NVL(tc.comments, '') as "table_comment"
        FROM 
            (SELECT owner, table_name FROM all_tables WHERE owner = USER
             UNION ALL 
             SELECT owner, view_name as table_name FROM all_views WHERE owner = USER) t
        LEFT JOIN 
            all_tab_comments tc ON t.table_name = tc.table_name AND t.owner = tc.owner
        ORDER BY 
            t.table_name
        """
        
        try:
            return self.execute_query(query_all)
        except Exception:
            # ALL_TABLES 접근 실패 시 USER_TABLES로 fallback
            query_user = """
            SELECT 
                t.table_name as "table_name",
                NVL(tc.comments, '') as "table_comment"
            FROM 
                (SELECT table_name FROM user_tables 
                 UNION ALL 
                 SELECT view_name as table_name FROM user_views) t
            LEFT JOIN 
                user_tab_comments tc ON t.table_name = tc.table_name
            ORDER BY 
                t.table_name
            """
            return self.execute_query(query_user)
    
    def get_indexes_info(self):
        """인덱스 정보 조회"""
        # 먼저 ALL_INDEXES로 시도
        query_all = """
        SELECT 
            ind.table_name as "table_name",
            ind.index_name as "index_name",
            CASE WHEN ind.uniqueness = 'UNIQUE' THEN 0 ELSE 1 END as "non_unique",
            ic.column_name as "column_name",
            ic.column_position as "seq_in_index"
        FROM 
            all_indexes ind
        JOIN 
            all_ind_columns ic ON ind.index_name = ic.index_name AND ind.owner = ic.index_owner
        WHERE 
            ind.owner = USER
            AND ind.index_type = 'NORMAL'
            AND ind.table_name NOT LIKE 'SYS_%'
            AND NOT EXISTS (
                SELECT 1 FROM all_constraints con 
                WHERE con.constraint_type = 'P' 
                AND con.index_name = ind.index_name
                AND con.owner = ind.owner
            )
        ORDER BY 
            ind.table_name, ind.index_name, ic.column_position
        """
        
        try:
            return self.execute_query(query_all)
        except Exception:
            # ALL_INDEXES 접근 실패 시 USER_INDEXES로 fallback
            query_user = """
            SELECT 
                ind.table_name as "table_name",
                ind.index_name as "index_name",
                CASE WHEN ind.uniqueness = 'UNIQUE' THEN 0 ELSE 1 END as "non_unique",
                ic.column_name as "column_name",
                ic.column_position as "seq_in_index"
            FROM 
                user_indexes ind
            JOIN 
                user_ind_columns ic ON ind.index_name = ic.index_name
            WHERE 
                ind.index_type = 'NORMAL'
                AND ind.table_name NOT LIKE 'SYS_%'
                AND NOT EXISTS (
                    SELECT 1 FROM user_constraints con 
                    WHERE con.constraint_type = 'P' 
                    AND con.index_name = ind.index_name
                )
            ORDER BY 
                ind.table_name, ind.index_name, ic.column_position
            """
            return self.execute_query(query_user)