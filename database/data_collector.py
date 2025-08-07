"""
데이터베이스 메타데이터 수집기
"""

from .connection_manager import connection_manager
from .exceptions import DatabaseConnectionError, DatabaseQueryError
import time


class DatabaseMetadataCollector:
    """데이터베이스 메타데이터 수집 클래스"""
    
    def __init__(self):
        self.connection_info = None
        self.last_collection_time = None
        
    def collect_database_metadata(self, dbms, host, port, database, username, password, timeout=30, selected_tables=None, oracle_type=None):
        """
        데이터베이스 메타데이터 전체 수집
        
        Args:
            dbms (str): DBMS 종류
            host (str): 서버 주소  
            port (int): 포트 번호
            database (str): 데이터베이스 이름
            username (str): 사용자명
            password (str): 비밀번호
            timeout (int): 연결 시간 제한
            selected_tables (list): 선택된 테이블 목록 (None이면 전체)
            oracle_type (str): Oracle 연결 방식 ('service_name' 또는 'sid')
            
        Returns:
            dict: 수집된 메타데이터
        """
        start_time = time.time()
        
        try:
            with connection_manager.get_connection(
                dbms=dbms, host=host, port=port, database=database,
                username=username, password=password, timeout=timeout,
                oracle_type=oracle_type
            ) as conn:
                
                # 기본 정보 수집
                metadata = {
                    'connection_info': {
                        'dbms': conn.get_dbms_name(),
                        'host': host,
                        'port': port,
                        'database': database,
                        'username': username,
                        'version': conn.get_version(),
                        'collection_time': time.strftime("%Y-%m-%d %H:%M:%S")
                    },
                    'tables': [],
                    'foreign_keys': [],
                    'indexes': [],
                    'statistics': {
                        'total_tables': 0,
                        'total_columns': 0,
                        'total_foreign_keys': 0,
                        'collection_duration_ms': 0
                    }
                }
                
                # 테이블 및 컬럼 정보 수집 (모든 DBMS 동일한 방식)
                tables_data = conn.get_tables_info()
                metadata['tables'] = self._normalize_tables_data(tables_data)
                
                # 외래키 정보 수집
                fk_data = conn.get_foreign_keys_info()
                metadata['foreign_keys'] = self._normalize_foreign_keys_data(fk_data)
                
                # 인덱스 정보 수집
                indexes_data = conn.get_indexes_info()
                metadata['indexes'] = self._normalize_indexes_data(indexes_data)
                
                # 선택된 테이블들만 필터링
                if selected_tables:
                    metadata['tables'] = [table for table in metadata['tables'] 
                                        if table['table_name'] in selected_tables]
                    metadata['foreign_keys'] = [fk for fk in metadata['foreign_keys'] 
                                               if fk['table_name'] in selected_tables]
                    metadata['indexes'] = [idx for idx in metadata['indexes'] 
                                         if idx['table_name'] in selected_tables]
                
                # 통계 계산
                metadata['statistics']['total_tables'] = len(set(table['table_name'] for table in metadata['tables']))
                metadata['statistics']['total_columns'] = len(metadata['tables'])
                metadata['statistics']['total_foreign_keys'] = len(metadata['foreign_keys'])
                metadata['statistics']['collection_duration_ms'] = round((time.time() - start_time) * 1000, 2)
                
                self.connection_info = metadata['connection_info']
                self.last_collection_time = time.time()
                
                return metadata
                
        except Exception as e:
            raise DatabaseQueryError(f"메타데이터 수집 실패: {str(e)}")
            
    def _normalize_tables_data(self, tables_data):
        """테이블 데이터 정규화 (MySQL/PostgreSQL/Oracle 모두 지원)"""
        normalized = []
        
        for row in tables_data:
            # 각 DBMS별로 다른 키 이름을 통일
            normalized_row = {
                'table_name': self._get_value(row, ['table_name', 'TABLE_NAME']),
                'table_comment': self._get_value(row, ['comment', 'table_comment', 'TABLE_COMMENT', '테이블설명']) or '',
                'column_position': self._get_value(row, ['NO', 'no', 'column_position', 'ORDINAL_POSITION']),
                'column_name': self._get_value(row, ['컬럼명', 'column_name', 'COLUMN_NAME']),
                'data_type': self._get_value(row, ['TYPE', 'type', 'data_type', 'COLUMN_TYPE']),
                'default_value': self._get_value(row, ['DEFAULT_VALUE', 'default_value', 'COLUMN_DEFAULT']) or '',
                'is_nullable': self._normalize_nullable(self._get_value(row, ['NULLABLE', 'nullable', 'IS_NULLABLE'])),
                'key_type': self._normalize_key_type(self._get_value(row, ['KEY_TYPE', 'key_type', 'COLUMN_KEY'])),
                'extra': self._get_value(row, ['EXTRA', 'extra', 'AUTO_INCREMENT']) or '',
                'column_comment': self._get_value(row, ['설명', 'column_comment', 'COLUMN_COMMENT']) or ''
            }
            normalized.append(normalized_row)
            
        return normalized
        
    def _normalize_foreign_keys_data(self, fk_data):
        """외래키 데이터 정규화"""
        normalized = []
        
        for row in fk_data:
            normalized_row = {
                'table_name': self._get_value(row, ['table_name', 'TABLE_NAME']),
                'column_name': self._get_value(row, ['column_name', 'COLUMN_NAME']),
                'referenced_table_name': self._get_value(row, ['referenced_table_name', 'REFERENCED_TABLE_NAME']),
                'referenced_column_name': self._get_value(row, ['referenced_column_name', 'REFERENCED_COLUMN_NAME']),
                'constraint_name': self._get_value(row, ['constraint_name', 'CONSTRAINT_NAME']) or ''
            }
            normalized.append(normalized_row)
            
        return normalized
        
    def _get_value(self, row_dict, possible_keys):
        """딕셔너리에서 가능한 키들 중 첫 번째로 존재하는 값 반환"""
        if isinstance(row_dict, dict):
            for key in possible_keys:
                if key in row_dict and row_dict[key] is not None:
                    return row_dict[key]
        return None
        
    def _normalize_nullable(self, value):
        """nullable 값 정규화"""
        if value is None:
            return 'YES'
        
        value_str = str(value).upper()
        if value_str in ['Y', 'YES', 'TRUE', '1']:
            return 'YES'
        elif value_str in ['N', 'NO', 'FALSE', '0']:
            return 'NO'
        else:
            return 'YES'  # 기본값
            
    def _normalize_key_type(self, value):
        """키 타입 정규화"""
        if value is None or value == '':
            return ''
            
        value_str = str(value).upper()
        if value_str in ['PRI', 'PRIMARY', 'P']:
            return 'PRI'
        elif value_str in ['MUL', 'FOREIGN', 'F', 'R']:
            return 'MUL'
        elif value_str in ['UNI', 'UNIQUE', 'U']:
            return 'UNI'
        else:
            return value_str
            
    def get_tables_summary(self, metadata):
        """테이블 요약 정보 생성"""
        if not metadata or 'tables' not in metadata:
            return []
            
        tables_summary = {}
        
        for row in metadata['tables']:
            table_name = row['table_name']
            if table_name not in tables_summary:
                tables_summary[table_name] = {
                    'table_name': table_name,
                    'table_comment': row['table_comment'],
                    'column_count': 0,
                    'primary_keys': [],
                    'foreign_keys': []
                }
                
            tables_summary[table_name]['column_count'] += 1
            
            # Primary Key 수집
            if row['key_type'] == 'PRI':
                tables_summary[table_name]['primary_keys'].append(row['column_name'])
                
        # Foreign Key 정보 추가
        if 'foreign_keys' in metadata:
            for fk in metadata['foreign_keys']:
                table_name = fk['table_name']
                if table_name in tables_summary:
                    fk_info = f"{fk['column_name']} -> {fk['referenced_table_name']}.{fk['referenced_column_name']}"
                    tables_summary[table_name]['foreign_keys'].append(fk_info)
                    
        return list(tables_summary.values())
        
    def format_metadata_for_excel(self, metadata):
        """Excel 생성을 위한 메타데이터 포맷"""
        if not metadata:
            return None
            
        formatted = {
            'database_info': metadata['connection_info'],
            'tables_by_name': {},
            'foreign_keys': metadata['foreign_keys'],
            'statistics': metadata['statistics']
        }
        
        # 테이블별로 그룹화
        for row in metadata['tables']:
            table_name = row['table_name']
            if table_name not in formatted['tables_by_name']:
                formatted['tables_by_name'][table_name] = {
                    'table_info': {
                        'name': table_name,
                        'comment': row['table_comment']
                    },
                    'columns': []
                }
                
            formatted['tables_by_name'][table_name]['columns'].append({
                'position': row['column_position'],
                'name': row['column_name'],
                'type': row['data_type'],
                'default': row['default_value'],
                'nullable': row['is_nullable'],
                'key': row['key_type'],
                'extra': row['extra'],
                'comment': row['column_comment']
            })
            
        # 컬럼을 position 순으로 정렬
        for table_data in formatted['tables_by_name'].values():
            table_data['columns'].sort(key=lambda x: x['position'] or 0)
            
        return formatted
    
    def collect_table_list(self, dbms, host, port, database, username, password, timeout=30, oracle_type=None):
        """
        테이블 목록만 수집 (NO, 테이블명, 논리명)
        
        Args:
            dbms (str): DBMS 종류
            host (str): 서버 주소  
            port (int): 포트 번호
            database (str): 데이터베이스 이름
            username (str): 사용자명
            password (str): 비밀번호
            timeout (int): 연결 시간 제한
            oracle_type (str): Oracle 연결 방식 ('service_name' 또는 'sid')
            
        Returns:
            dict: 수집된 테이블 목록
        """
        start_time = time.time()
        
        try:
            with connection_manager.get_connection(
                dbms=dbms, host=host, port=port, database=database,
                username=username, password=password, timeout=timeout,
                oracle_type=oracle_type
            ) as conn:
                
                # 기본 정보 수집
                table_list_data = {
                    'connection_info': {
                        'dbms': conn.get_dbms_name(),
                        'host': host,
                        'port': port,
                        'database': database,
                        'username': username,
                        'version': conn.get_version(),
                        'collection_time': time.strftime("%Y-%m-%d %H:%M:%S")
                    }
                }
                
                # 테이블 목록만 수집 (간단한 정보)
                tables_info = conn.get_tables_basic_info()
                
                # 정규화된 테이블 목록 생성
                table_list = []
                for idx, table in enumerate(tables_info, 1):
                    table_list.append({
                        'no': idx,
                        'table_name': self._get_value(table, ['table_name', 'TABLE_NAME']),
                        'table_comment': self._get_value(table, ['table_comment', 'TABLE_COMMENT', '테이블설명']) or ''
                    })
                
                table_list_data['table_list'] = table_list
                
                # 통계 정보
                end_time = time.time()
                collection_duration = (end_time - start_time) * 1000
                
                table_list_data['statistics'] = {
                    'total_tables': len(table_list),
                    'collection_duration_ms': round(collection_duration, 2)
                }
                
                return table_list_data
                
        except Exception as e:
            raise DatabaseConnectionError(f"테이블 목록 수집 중 오류가 발생했습니다: {str(e)}")
    
    def _normalize_indexes_data(self, indexes_data):
        """인덱스 데이터 정규화"""
        normalized = []
        
        for row in indexes_data:
            # 각 DBMS별로 다른 키 이름을 통일
            normalized_row = {
                'table_name': row.get('table_name') or row.get('TABLE_NAME'),
                'index_name': row.get('index_name') or row.get('INDEX_NAME'),
                'non_unique': row.get('non_unique') or row.get('NON_UNIQUE', 0),
                'column_name': row.get('column_name') or row.get('COLUMN_NAME'),
                'seq_in_index': row.get('seq_in_index') or row.get('SEQ_IN_INDEX', 1)
            }
            
            # 필수 필드 검증
            if not normalized_row['table_name'] or not normalized_row['index_name']:
                continue
                
            normalized.append(normalized_row)
            
        return normalized


# 싱글톤 인스턴스
metadata_collector = DatabaseMetadataCollector()