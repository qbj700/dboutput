#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JDBC Oracle 연결 클래스 (순수 Java JDBC + JPype 사용)
"""

import os
import sys
import jpype
from typing import List, Dict, Any, Optional

from .base_connection import BaseConnection
from .exceptions import DatabaseConnectionError, DatabaseAuthenticationError


class JdbcOracleConnection(BaseConnection):
    """JayDeBeApi를 사용한 Oracle JDBC 연결"""
    
    def __init__(self, host: str, port: int, database: str, username: str, password: str, 
                 timeout: int = 30, oracle_type: str = 'service_name'):
        super().__init__(host, port, database, username, password, timeout)
        self.oracle_type = oracle_type
        self.connection = None
        self.jdbc_url = self._build_jdbc_url()
        self._setup_jvm()
    
    def _build_jdbc_url(self) -> str:
        """JDBC URL 생성"""
        if self.oracle_type == 'sid':
            return f"jdbc:oracle:thin:@{self.host}:{self.port}:{self.database}"
        else:
            return f"jdbc:oracle:thin:@{self.host}:{self.port}/{self.database}"
    
    def _setup_jvm(self):
        """JVM 초기화"""
        # JVM이 이미 시작되었는지 확인 (매우 중요!)
        if jpype.isJVMStarted():
            print("JVM이 이미 시작되었습니다. 기존 JVM을 재사용합니다.")
            return
        
        print("새로운 JVM을 시작합니다...")
        
        try:
            # JRE 경로 설정
            if hasattr(sys, '_MEIPASS'):
                # PyInstaller 실행 파일에서
                jre_path = os.path.join(sys._MEIPASS, 'jre')
            else:
                # 개발 환경에서
                current_dir = os.path.dirname(os.path.abspath(__file__))
                jre_path = os.path.join(os.path.dirname(current_dir), 'jre')
            
            print(f"JRE 경로: {jre_path}")
            
            # JVM DLL 경로 찾기
            jvm_dll_paths = [
                os.path.join(jre_path, 'bin', 'server', 'jvm.dll'),
                os.path.join(jre_path, 'bin', 'client', 'jvm.dll'),
                os.path.join(jre_path, 'bin', 'jvm.dll')
            ]
            
            jvm_path = None
            for path in jvm_dll_paths:
                print(f"JVM DLL 확인 중: {path}")
                if os.path.exists(path):
                    jvm_path = path
                    print(f"✅ JVM DLL 발견: {jvm_path}")
                    break
            
            if not jvm_path:
                raise DatabaseConnectionError(
                    f"JVM DLL을 찾을 수 없습니다.\n확인된 경로들:\n" + 
                    "\n".join(f"- {path}" for path in jvm_dll_paths) + 
                    "\n\nJRE가 올바르게 설치되었는지 확인하세요.",
                    dbms="Oracle JDBC",
                    host=self.host,
                    port=self.port,
                    database=self.database
                )
            
            # JDBC JAR 경로 확인
            jdbc_jar = self._get_jdbc_jar_path()
            print(f"JDBC JAR: {jdbc_jar}")
            
            if not os.path.exists(jdbc_jar):
                raise DatabaseConnectionError(
                    f"JDBC JAR 파일을 찾을 수 없습니다: {jdbc_jar}",
                    dbms="Oracle JDBC",
                    host=self.host,
                    port=self.port,
                    database=self.database
                )
            
            print(f"JVM 시작 중... ({jvm_path})")
            jpype.startJVM(
                jvm_path, 
                "-ea",
                "-Xmx512m",  # 메모리 제한
                f"-Djava.class.path={jdbc_jar}"
            )
            print("✅ JVM 시작 완료!")
            
        except Exception as e:
            raise DatabaseConnectionError(
                f"JVM 초기화 실패: {str(e)}",
                dbms="Oracle JDBC",
                host=self.host,
                port=self.port,
                database=self.database
            )
    
    def _get_jdbc_jar_path(self) -> str:
        """JDBC JAR 파일 경로 반환"""
        if hasattr(sys, '_MEIPASS'):
            # PyInstaller 실행 파일에서
            return os.path.join(sys._MEIPASS, 'ojdbc8.jar')
        else:
            # 개발 환경에서
            return os.path.join(os.path.dirname(os.path.dirname(__file__)), 'ojdbc8.jar')
    
    def get_dbms_name(self) -> str:
        return "Oracle JDBC"
    
    def test_connection(self) -> bool:
        """연결 테스트"""
        try:
            result = self.connect()
            if result and self.is_connected:
                # 간단한 쿼리로 연결 확인
                test_result = self.execute_query("SELECT 'TEST' as test_col FROM dual")
                return len(test_result) > 0 and test_result[0].get('test_col') == 'TEST'
            return False
        except Exception as e:
            print(f"연결 테스트 실패: {e}")
            return False
    
    def get_version(self) -> str:
        """Oracle 버전 정보 반환"""
        try:
            if not self.is_connected:
                return "연결되지 않음"
            
            result = self.execute_query("SELECT banner FROM v$version WHERE rownum = 1")
            if result and len(result) > 0:
                return result[0].get('banner', 'Unknown')
            return "Unknown"
        except Exception as e:
            return f"버전 확인 실패: {str(e)}"
    
    def connect(self) -> bool:
        """Oracle JDBC 연결"""
        try:
            # JVM 상태 확인 및 로깅
            if jpype.isJVMStarted():
                print(f"JVM 상태: {self.get_jvm_info()}")
            else:
                print("JVM이 시작되지 않았습니다. 초기화를 다시 시도합니다.")
                self._setup_jvm()
            

            
            # 입력 검증
            if not self.username:
                raise DatabaseAuthenticationError("사용자명이 비어있습니다.")
            if not self.password:
                raise DatabaseAuthenticationError("비밀번호가 비어있습니다.")
            
            # 문자열 변환 및 정리
            username_str = str(self.username).strip()
            password_str = str(self.password).strip()
            
            # Java JDBC 연결 (Properties 방식)
            try:
                # Java 객체 생성
                JavaString = jpype.JClass("java.lang.String")
                Properties = jpype.JClass("java.util.Properties")
                DriverManager = jpype.JClass("java.sql.DriverManager")
                
                # 연결 정보 설정
                props = Properties()
                props.setProperty("user", JavaString(username_str))
                props.setProperty("password", JavaString(password_str))
                
                # JDBC 연결
                java_connection = DriverManager.getConnection(self.jdbc_url, props)
                
                # 순수 Java JDBC 사용 (JayDeBeApi 래핑 제거)
                self.java_connection = java_connection
                self.connection = self  # self를 connection으로 설정
                self.is_connected = True
                
                print("Oracle JDBC 연결 성공!")
                
            except Exception as e:
                # 계정 잠금 오류 처리
                if "ORA-28000" in str(e):
                    raise DatabaseAuthenticationError(
                        "Oracle 계정이 잠겨있습니다. DBA에게 문의하여 계정 잠금을 해제해주세요."
                    )
                
                # 기타 Oracle 오류 처리
                if "ORA-" in str(e):
                    error_code = str(e).split("ORA-")[1][:5]
                    raise DatabaseConnectionError(f"Oracle 연결 실패 (ORA-{error_code}): {e}")
                
                # 일반 연결 오류
                raise DatabaseConnectionError(f"Oracle JDBC 연결 실패: {e}")
            
            return True
            
        except Exception as e:
            # 최종 오류 처리
            error_msg = str(e).lower()
            
            if 'invalid username/password' in error_msg or 'ora-01017' in error_msg:
                raise DatabaseAuthenticationError(
                    f"인증 실패: 사용자명 또는 비밀번호가 올바르지 않습니다.",
                    dbms=self.get_dbms_name(),
                    host=self.host,
                    port=self.port,
                    database=self.database
                )
            elif 'connection refused' in error_msg or 'ora-12541' in error_msg:
                raise DatabaseConnectionError(
                    f"서버에 연결할 수 없습니다: {self.host}:{self.port}",
                    dbms=self.get_dbms_name(),
                    host=self.host,
                    port=self.port,
                    database=self.database
                )
            else:
                raise DatabaseConnectionError(
                    f"Oracle JDBC 연결 실패: {str(e)}",
                    dbms=self.get_dbms_name(),
                    host=self.host,
                    port=self.port,
                    database=self.database
                )
    
    def disconnect(self):
        """연결 종료"""
        if hasattr(self, 'java_connection') and self.java_connection:
            try:
                self.java_connection.close()
                self.java_connection = None
                self.connection = None
                self.is_connected = False
                print("Oracle JDBC 연결이 종료되었습니다.")
                # 주의: JVM은 종료하지 않음! jpype는 프로세스당 한 번만 시작 가능
            except Exception as e:
                print(f"연결 종료 중 오류: {e}")
    
    @staticmethod
    def is_jvm_started() -> bool:
        """JVM 시작 상태 확인"""
        try:
            return jpype.isJVMStarted()
        except Exception:
            return False
    
    @staticmethod
    def get_jvm_info() -> str:
        """JVM 정보 반환"""
        if jpype.isJVMStarted():
            try:
                runtime = jpype.JClass("java.lang.Runtime").getRuntime()
                total_memory = runtime.totalMemory() / (1024 * 1024)  # MB
                free_memory = runtime.freeMemory() / (1024 * 1024)   # MB
                used_memory = total_memory - free_memory
                return f"JVM 메모리: {used_memory:.1f}MB 사용 / {total_memory:.1f}MB 전체"
            except Exception as e:
                return f"JVM 실행 중 (정보 확인 실패: {e})"
        else:
            return "JVM 미실행"
    
    def execute_query(self, query: str) -> List[Dict[str, Any]]:
        """쿼리 실행 (순수 Java JDBC 사용)"""
        if not hasattr(self, 'java_connection') or not self.java_connection:
            raise DatabaseConnectionError("데이터베이스에 연결되지 않았습니다.", dbms=self.get_dbms_name())
        
        try:
            # Java Statement 생성
            statement = self.java_connection.createStatement()
            result_set = statement.executeQuery(query)
            
            # 메타데이터 가져오기
            meta_data = result_set.getMetaData()
            column_count = meta_data.getColumnCount()
            
            # 컬럼 이름 추출 (소문자로 변환)
            columns = []
            for i in range(1, column_count + 1):
                column_label = meta_data.getColumnLabel(i)
                # Java String을 Python 문자열로 변환 후 소문자 변환
                columns.append(str(column_label).lower())
            
            # 결과 데이터 추출
            result = []
            while result_set.next():
                row_dict = {}
                for i, column_name in enumerate(columns, 1):
                    value = result_set.getObject(i)
                    # Java null을 Python None으로 변환
                    if value is None:
                        row_dict[column_name] = None
                    else:
                        row_dict[column_name] = str(value)  # Java 객체를 문자열로 변환
                result.append(row_dict)
            
            # 리소스 정리
            result_set.close()
            statement.close()
            
            return result
            
        except Exception as e:
            raise DatabaseConnectionError(f"쿼리 실행 실패: {str(e)}", dbms=self.get_dbms_name())
    
    def get_tables_basic_info(self) -> List[Dict[str, Any]]:
        """테이블 기본 정보만 조회 (MySQL과 동일한 컬럼명)"""
        query = """
        SELECT
            t.table_name                             AS TABLE_NAME,
            NVL(tc.comments, '')                     AS "테이블설명"
        FROM (
            SELECT table_name  FROM user_tables  WHERE table_name NOT LIKE 'BIN$%'
            UNION ALL
            SELECT view_name AS table_name FROM user_views
        ) t
        LEFT JOIN user_tab_comments tc
            ON tc.table_name = t.table_name
        ORDER BY t.table_name
        """
        return self.execute_query(query)
    
    def _get_default_values_via_jdbc(self) -> Dict[str, str]:
        """JDBC 메타데이터로 모든 테이블의 Default 값을 한 번에 수집"""
        default_map = {}
        try:
            db_meta = self.java_connection.getMetaData()
            # Oracle에서는 username이 스키마명
            schema_name = self.username.upper()
            
            col_meta_rs = db_meta.getColumns(
                None,
                schema_name,              # TABLE_SCHEMA (username 사용)
                None,                     # TABLE_NAME (모든 테이블)
                None                      # COLUMN_NAME (모든 컬럼)
            )
            
            while col_meta_rs.next():
                table_name = col_meta_rs.getString("TABLE_NAME")
                col_name = col_meta_rs.getString("COLUMN_NAME")
                col_def = col_meta_rs.getString("COLUMN_DEF")  # DEFAULT 값
                
                # 테이블.컬럼 형태로 키 생성
                key = f"{table_name}.{col_name}"
                # Default 값 정리: None, 빈 문자열, 공백만 있는 경우 처리
                if col_def is None:
                    clean_default = ""
                else:
                    clean_default = str(col_def).strip()
                    # 특수한 Oracle Default 값들 정리
                    if clean_default in ['NULL', 'null', "''", '""', '-']:
                        clean_default = ""
                default_map[key] = clean_default
                
            col_meta_rs.close()
            return default_map
        except Exception as e:
            # JDBC 메타데이터 수집 실패 시 빈 딕셔너리 반환 (기존 동작 유지)
            return {}

    def get_tables_info(self) -> List[Dict[str, Any]]:
        """테이블+컬럼 정보 조회 (MySQL과 완전히 동일한 컬럼명·순서)"""
        # 1. JDBC 메타데이터로 Default 값들을 한 번에 수집
        default_map = self._get_default_values_via_jdbc()
        
        # 2. 기본 테이블+컬럼 정보 조회 (DEFAULT_VALUE는 여전히 빈 문자열)
        query = """
        SELECT
            t.table_name                             AS TABLE_NAME,
            NVL(tc.comments, '')                     AS "테이블설명",
            c.column_id                              AS NO,
            c.column_name                            AS "컬럼명",
            -- MySQL의 COLUMN_TYPE 과 같은 형태로 뿌려주기
            CASE
              WHEN c.data_type IN ('VARCHAR2','CHAR','NVARCHAR2','NCHAR')
                THEN c.data_type || '(' || c.char_length || ')'
              WHEN c.data_type='NUMBER'
                AND c.data_precision IS NOT NULL
                AND c.data_scale     IS NOT NULL
                THEN c.data_type || '(' || c.data_precision || ',' || c.data_scale || ')'
              WHEN c.data_type='NUMBER'
                AND c.data_precision IS NOT NULL
                THEN c.data_type || '(' || c.data_precision || ')'
              ELSE c.data_type
            END                                       AS TYPE,
            -- LONG 타입 충돌 회피를 위해 빈 문자열로만 내려줌 (나중에 JDBC 메타데이터로 교체)
            ''                                        AS DEFAULT_VALUE,
            CASE WHEN c.nullable='Y' THEN 'Y' ELSE 'N' END AS NULLABLE,
            NVL(k.key_type, '')                       AS KEY_TYPE,
            -- AUTO_INCREMENT 여부는 Oracle에서는 해당 없음
            ''                                        AS EXTRA,
            NVL(cc.comments, '')                      AS "설명"
        FROM (
            SELECT table_name  FROM user_tables  WHERE table_name NOT LIKE 'BIN$%'
            UNION ALL
            SELECT view_name AS table_name FROM user_views
        ) t
        JOIN user_tab_columns c
            ON c.table_name = t.table_name
        LEFT JOIN user_tab_comments tc
            ON tc.table_name = t.table_name
        LEFT JOIN user_col_comments cc
            ON cc.table_name  = c.table_name
           AND cc.column_name = c.column_name
        LEFT JOIN (
            -- PK/UNI/FK 를 MySQL COLUMN_KEY 값과 매핑 (PRI/UNI/MUL)
            SELECT acc.table_name,
                   acc.column_name,
                   CASE ac.constraint_type
                     WHEN 'P' THEN 'PRI'
                     WHEN 'U' THEN 'UNI'
                     WHEN 'R' THEN 'MUL'
                   END AS key_type
            FROM user_constraints ac
            JOIN user_cons_columns acc
              ON ac.constraint_name = acc.constraint_name
            WHERE ac.constraint_type IN ('P','U','R')
        ) k
          ON k.table_name  = c.table_name
         AND k.column_name = c.column_name
        ORDER BY t.table_name, c.column_id
        """
        
        # 3. 쿼리 실행
        result = self.execute_query(query)
        
        # 4. 결과에서 DEFAULT_VALUE를 JDBC 메타데이터로 교체
        for row in result:
            table_name = row.get('table_name') or row.get('TABLE_NAME')
            column_name = row.get('컬럼명') or row.get('column_name') or row.get('COLUMN_NAME')
            
            if table_name and column_name:
                key = f"{table_name}.{column_name}"
                default_value = default_map.get(key, "")
                
                # DEFAULT_VALUE 키 업데이트
                if 'default_value' in row:
                    row['default_value'] = default_value
                elif 'DEFAULT_VALUE' in row:
                    row['DEFAULT_VALUE'] = default_value
        
        return result

    
    def get_columns_info(self, table_name: str) -> List[Dict[str, Any]]:
        """컬럼 정보 조회 (기본값 포함: JDBC Metadata 사용)"""
        try:
            # 1) 공통 JDBC 메타데이터 함수 사용 (해당 테이블만 필터링)
            all_defaults = self._get_default_values_via_jdbc()
            default_map = {}
            table_prefix = f"{table_name.upper()}."
            for key, value in all_defaults.items():
                if key.startswith(table_prefix):
                    col_name = key[len(table_prefix):]  # 테이블명. 제거
                    default_map[col_name] = value

            # 2) 실제 조회 쿼리 (DATA_DEFAULT 대신 '' AS DEFAULT_VALUE)
            query = """
            SELECT 
                c.column_name                AS COLUMN_NAME,
                CASE 
                WHEN c.data_type IN ('VARCHAR2','CHAR','NVARCHAR2','NCHAR')
                    THEN c.data_type || '(' || c.char_length || ')'
                WHEN c.data_type = 'NUMBER' 
                    AND c.data_precision IS NOT NULL 
                    AND c.data_scale IS NOT NULL
                    THEN c.data_type || '(' || c.data_precision || ',' || c.data_scale || ')'
                WHEN c.data_type = 'NUMBER' 
                    AND c.data_precision IS NOT NULL
                    THEN c.data_type || '(' || c.data_precision || ')'
                ELSE c.data_type
                END                           AS TYPE,
                CASE 
                WHEN c.data_type IN ('VARCHAR2','CHAR','NVARCHAR2','NCHAR')
                    THEN c.char_length
                WHEN c.data_type = 'NUMBER' 
                    AND c.data_precision IS NOT NULL
                    THEN c.data_precision
                ELSE c.data_length
                END                           AS MAX_LENGTH,
                c.data_precision             AS PRECISION,
                c.data_scale                 AS SCALE,
                CASE WHEN c.nullable = 'Y' THEN 1 ELSE 0 END AS IS_NULLABLE,
                -- SQL에서는 빈 문자열로 채워 두고
                ''                            AS DEFAULT_VALUE,
                NVL(cc.comments, '')          AS "COMMENT",
                c.column_id                  AS ORDINAL_POSITION
            FROM user_tab_columns c
            LEFT JOIN user_col_comments cc
            ON cc.table_name  = c.table_name
            AND cc.column_name = c.column_name
            WHERE UPPER(c.table_name) = UPPER(?)
            ORDER BY c.column_id
            """
            statement = self.java_connection.prepareStatement(query)
            statement.setString(1, table_name)
            result_set = statement.executeQuery()

            # 3) 메타데이터+결과 매핑 (DEFAULT_VALUE는 default_map 사용)
            meta_data    = result_set.getMetaData()
            column_count = meta_data.getColumnCount()
            columns = [meta_data.getColumnLabel(i).lower() for i in range(1, column_count+1)]

            result = []
            while result_set.next():
                row_dict = {}
                for idx, col_key in enumerate(columns, start=1):
                    if col_key == "default_value":
                        # SQL에서는 ''지만, 실제로는 JDBC 메타데이터에서 채워 넣기
                        java_col_name = meta_data.getColumnLabel(idx)  # e.g. "DEFAULT_VALUE"
                        # 원래 컬럼 이름(COLUMN_NAME) 키 가져오기
                        actual_col_name = result_set.getString("COLUMN_NAME")
                        row_dict[col_key] = default_map.get(actual_col_name, "")
                    else:
                        val = result_set.getObject(idx)
                        row_dict[col_key] = None if val is None else str(val)
                result.append(row_dict)

            result_set.close()
            statement.close()
            return result

        except Exception as e:
            raise DatabaseConnectionError(f"컬럼 정보 조회 실패: {str(e)}")

    
    def get_foreign_keys_info(self) -> List[Dict[str, Any]]:
        """외래키 정보 조회 (MySQL 구조와 동일)"""
        query = """
        SELECT
            acc.table_name            AS TABLE_NAME,
            acc.column_name           AS COLUMN_NAME,
            r_acc.table_name          AS REFERENCED_TABLE_NAME,
            r_acc.column_name         AS REFERENCED_COLUMN_NAME,
            acc.constraint_name       AS CONSTRAINT_NAME
        FROM user_constraints ac
        JOIN user_cons_columns acc
          ON ac.constraint_name = acc.constraint_name
        JOIN user_cons_columns r_acc
          ON ac.r_constraint_name = r_acc.constraint_name
         AND acc.position = r_acc.position
        WHERE ac.constraint_type = 'R'
        ORDER BY acc.table_name, acc.column_name
        """
        return self.execute_query(query)

    def get_indexes_info(self) -> List[Dict[str, Any]]:
        """인덱스 정보 조회 (MySQL 구조와 동일)"""
        query = """
        SELECT
            ic.table_name             AS TABLE_NAME,
            ic.index_name             AS INDEX_NAME,
            CASE WHEN i.uniqueness='UNIQUE' THEN 0 ELSE 1 END AS NON_UNIQUE,
            ic.column_name            AS COLUMN_NAME,
            ic.column_position        AS SEQ_IN_INDEX
        FROM user_indexes i
        JOIN user_ind_columns ic
          ON i.index_name = ic.index_name
         AND i.table_name = ic.table_name
        WHERE i.index_type = 'NORMAL'
          AND NOT EXISTS (
            SELECT 1
              FROM user_constraints uc
             WHERE uc.constraint_name = i.index_name
               AND uc.constraint_type = 'P'
          )
        ORDER BY ic.table_name, ic.index_name, ic.column_position
        """
        return self.execute_query(query)