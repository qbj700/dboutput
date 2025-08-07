"""
공통 유틸리티 함수들
"""

import os
import datetime
import re
from pathlib import Path


def validate_port(port_str):
    """포트 번호 유효성 검사"""
    try:
        port = int(port_str)
        return 1 <= port <= 65535
    except (ValueError, TypeError):
        return False


def validate_filename(filename):
    """파일명 유효성 검사"""
    if not filename or not filename.strip():
        return False
        
    # Windows 파일명 금지 문자 체크
    invalid_chars = r'[<>:"/\\|?*]'
    if re.search(invalid_chars, filename):
        return False
        
    # 예약된 파일명 체크
    reserved_names = [
        'CON', 'PRN', 'AUX', 'NUL',
        'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
        'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
    ]
    
    name_without_ext = os.path.splitext(filename)[0].upper()
    if name_without_ext in reserved_names:
        return False
        
    return True


def ensure_excel_extension(filename):
    """파일명에 Excel 확장자가 없으면 추가"""
    if not filename.lower().endswith(('.xlsx', '.xls')):
        return filename + '.xlsx'
    return filename


def get_timestamp():
    """현재 시간을 문자열로 반환"""
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_log_timestamp():
    """로그용 시간 스탬프 반환"""
    return datetime.datetime.now().strftime("%H:%M:%S")


def safe_filename(text, max_length=50):
    """안전한 파일명 생성"""
    # 특수문자 제거
    safe_text = re.sub(r'[<>:"/\\|?*]', '_', text)
    # 길이 제한
    if len(safe_text) > max_length:
        safe_text = safe_text[:max_length]
    return safe_text


def ensure_directory_exists(directory_path):
    """디렉토리가 존재하지 않으면 생성"""
    try:
        Path(directory_path).mkdir(parents=True, exist_ok=True)
        return True
    except Exception:
        return False


def format_file_size(size_bytes):
    """파일 크기를 읽기 쉬운 형태로 변환"""
    if size_bytes == 0:
        return "0B"
    
    size_names = ["B", "KB", "MB", "GB"]
    i = 0
    while size_bytes >= 1024.0 and i < len(size_names) - 1:
        size_bytes /= 1024.0
        i += 1
    
    return f"{size_bytes:.1f}{size_names[i]}"


def generate_spec_filename(database_name, timestamp=True):
    """명세서 파일명 자동 생성"""
    safe_db_name = safe_filename(database_name)
    base_name = f"명세서_{safe_db_name}"
    
    if timestamp:
        time_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name += f"_{time_str}"
        
    return base_name + ".xlsx"


def mask_password(password, mask_char="*"):
    """비밀번호 마스킹"""
    if not password:
        return ""
    return mask_char * len(password)


class Logger:
    """간단한 로거 클래스"""
    
    def __init__(self, log_widget=None):
        self.log_widget = log_widget
        
    def log(self, message, level="INFO"):
        """로그 메시지 출력"""
        timestamp = get_log_timestamp()
        formatted_message = f"[{timestamp}] [{level}] {message}"
        
        if self.log_widget:
            try:
                self.log_widget.insert("end", formatted_message + "\n")
                self.log_widget.see("end")
            except Exception:
                pass  # GUI가 없는 환경에서는 무시
                
        # 콘솔에도 출력 (개발 시 디버깅용)
        print(formatted_message)
        
    def info(self, message):
        """정보 로그"""
        self.log(message, "INFO")
        
    def error(self, message):
        """에러 로그"""
        self.log(message, "ERROR")
        
    def warning(self, message):
        """경고 로그"""
        self.log(message, "WARN")
        
    def success(self, message):
        """성공 로그"""
        self.log(message, "SUCCESS")