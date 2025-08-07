"""
DB 산출물 생성 GUI 메인 윈도우
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkinter import filedialog
import threading
import os
import sys
import subprocess
import time
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import APP_CONFIG, SUPPORTED_DBMS, FILE_CONFIG, UI_MESSAGES, ERROR_MESSAGES
from utils import validate_port, validate_filename, ensure_excel_extension, Logger
from database import connection_manager, metadata_collector, DatabaseConnectionError
from excel import excel_generator
from gui.table_selector import show_table_selector


class DBSpecGeneratorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.logger = None  # 로그 위젯 생성 후 초기화
        self.setup_window()
        self.create_widgets()
        self.setup_logger()
        
    def setup_window(self):
        """윈도우 기본 설정"""
        self.root.title(APP_CONFIG["title"])
        self.root.geometry(APP_CONFIG["window_size"])
        self.root.resizable(True, True)
        
        # 최소 윈도우 크기 설정
        self.root.minsize(*APP_CONFIG["min_window_size"])
        
        # 윈도우 아이콘 설정
        self._set_window_icon()
            
        # 윈도우 중앙 배치
        self.center_window()
        
        # 윈도우 종료 이벤트 처리
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def _set_window_icon(self):
        """윈도우 아이콘 설정"""
        try:
            # 현재 스크립트의 디렉토리 경로
            if getattr(sys, 'frozen', False):
                # 실행파일(exe)로 실행된 경우 - PyInstaller의 임시 폴더
                base_path = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
            else:
                # 개발 환경에서 실행된 경우
                base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            
            icon_path = os.path.join(base_path, "dboutput.ico")
            
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                if self.logger:
                    self.logger.info(f"윈도우 아이콘이 설정되었습니다: {icon_path}")
            else:
                # ico 파일이 없으면 png 파일로 시도
                png_path = os.path.join(base_path, "dboutput.png")
                if os.path.exists(png_path):
                    # PNG 파일을 PhotoImage로 로드
                    img = tk.PhotoImage(file=png_path)
                    self.root.iconphoto(False, img)
                    if self.logger:
                        self.logger.info(f"윈도우 아이콘이 설정되었습니다: {png_path}")
        except Exception as e:
            # 아이콘 설정 실패해도 프로그램은 계속 실행
            if self.logger:
                self.logger.debug(f"아이콘 설정 실패 (무시됨): {str(e)}")
            pass
        
    def center_window(self):
        """윈도우를 화면 중앙에 배치"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.root.winfo_screenheight() // 2) - (700 // 2)
        self.root.geometry(f"600x700+{x}+{y}")
        
    def create_widgets(self):
        """GUI 위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 그리드 가중치 설정
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(6, weight=1)  # 로그 영역이 확장되도록
        main_frame.grid_columnconfigure(1, weight=1)
        
        # 제목
        title_label = ttk.Label(main_frame, text="DB 산출물 생성기", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        # 초기화 버튼 (제목과 DB 연결 정보 프레임 사이 우측)
        self.reset_button = ttk.Button(main_frame, text="입력값 초기화", 
                                      command=self.reset_form, width=12)
        self.reset_button.grid(row=1, column=1, sticky=tk.E, pady=(0, 10))
        
        # DB 연결 정보 프레임
        self.create_connection_frame(main_frame)
        
        # 파일 저장 설정 프레임
        self.create_file_frame(main_frame)
        
        # 버튼 프레임
        self.create_button_frame(main_frame)
        
        # 로그/상태 메시지 프레임
        self.create_log_frame(main_frame)
        
    def create_connection_frame(self, parent):
        """DB 연결 정보 입력 프레임"""
        # 연결 정보 그룹
        conn_group = ttk.LabelFrame(parent, text="데이터베이스 연결 정보", padding="10")
        conn_group.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        conn_group.grid_columnconfigure(1, weight=1)
        
        # DBMS 종류
        ttk.Label(conn_group, text="DBMS 종류:", width=15, anchor='w').grid(row=0, column=0, sticky=tk.W, pady=5)
        self.dbms_var = tk.StringVar()
        dbms_list = list(SUPPORTED_DBMS.keys())
        self.dbms_combo = ttk.Combobox(conn_group, textvariable=self.dbms_var, 
                                      values=dbms_list,
                                      state="readonly", width=15)
        self.dbms_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        self.dbms_combo.set(dbms_list[0])  # 첫 번째 DBMS를 기본값으로
        self.dbms_combo.bind("<<ComboboxSelected>>", self.on_dbms_change)
        
        # 서버 주소
        ttk.Label(conn_group, text="서버 주소:", width=15, anchor='w').grid(row=1, column=0, sticky=tk.W, pady=5)
        self.host_var = tk.StringVar(value="localhost")
        self.host_entry = ttk.Entry(conn_group, textvariable=self.host_var)
        self.host_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
        # 포트 번호
        ttk.Label(conn_group, text="포트 번호:", width=15, anchor='w').grid(row=2, column=0, sticky=tk.W, pady=5)
        default_port = SUPPORTED_DBMS[dbms_list[0]]["default_port"]
        self.port_var = tk.StringVar(value=str(default_port))
        self.port_entry = ttk.Entry(conn_group, textvariable=self.port_var, width=15)
        self.port_entry.grid(row=2, column=1, sticky=tk.W, pady=5, padx=(10, 0))
        
        # Oracle 연결 방식 선택 (Oracle 선택 시에만 표시)
        self.oracle_label = ttk.Label(conn_group, text="Oracle 연결 방식:", width=15, anchor='w')
        self.oracle_label.grid(row=3, column=0, sticky=tk.W, pady=5)
        self.oracle_frame = ttk.Frame(conn_group)
        self.oracle_frame.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
        self.oracle_type_var = tk.StringVar(value="service_name")
        
        self.service_name_radio = ttk.Radiobutton(
            self.oracle_frame, 
            text="Service Name", 
            variable=self.oracle_type_var, 
            value="service_name",
            command=self.on_oracle_type_change
        )
        self.service_name_radio.pack(side=tk.LEFT, padx=(0, 10))
        
        self.sid_radio = ttk.Radiobutton(
            self.oracle_frame, 
            text="SID", 
            variable=self.oracle_type_var, 
            value="sid",
            command=self.on_oracle_type_change
        )
        self.sid_radio.pack(side=tk.LEFT)
        
        # 초기에는 Oracle 라벨과 프레임 숨김
        self.oracle_label.grid_remove()
        self.oracle_frame.grid_remove()
        
        # 데이터베이스 이름 (Oracle인 경우 동적으로 변경됨)
        self.database_label = ttk.Label(conn_group, text="데이터베이스 이름:", width=15, anchor='w')
        self.database_label.grid(row=4, column=0, sticky=tk.W, pady=5)
        self.database_var = tk.StringVar()
        self.database_entry = ttk.Entry(conn_group, textvariable=self.database_var)
        self.database_entry.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
        # 사용자 ID
        ttk.Label(conn_group, text="사용자 ID:", width=15, anchor='w').grid(row=5, column=0, sticky=tk.W, pady=5)
        self.username_var = tk.StringVar()
        self.username_entry = ttk.Entry(conn_group, textvariable=self.username_var)
        self.username_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
        # 비밀번호
        ttk.Label(conn_group, text="비밀번호:", width=15, anchor='w').grid(row=6, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(conn_group, textvariable=self.password_var, show="*")
        self.password_entry.grid(row=6, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
    def create_file_frame(self, parent):
        """파일 저장 설정 프레임"""
        file_group = ttk.LabelFrame(parent, text="파일 저장 설정", padding="10")
        file_group.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_group.grid_columnconfigure(1, weight=1)
        
        # 저장 경로
        ttk.Label(file_group, text="저장 경로:", width=15, anchor='w').grid(row=0, column=0, sticky=tk.W, pady=5)
        self.save_path_var = tk.StringVar(value=os.getcwd())
        self.save_path_entry = ttk.Entry(file_group, textvariable=self.save_path_var)
        self.save_path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 5))
        
        self.browse_button = ttk.Button(file_group, text="찾아보기", 
                                       command=self.browse_save_path, width=12)
        self.browse_button.grid(row=0, column=2, pady=5)
        
        # 파일명
        ttk.Label(file_group, text="파일명:", width=15, anchor='w').grid(row=1, column=0, sticky=tk.W, pady=5)
        self.filename_var = tk.StringVar(value=FILE_CONFIG["default_filename"])
        self.filename_entry = ttk.Entry(file_group, textvariable=self.filename_var)
        self.filename_entry.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
    def create_button_frame(self, parent):
        """버튼 프레임"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        # 연결 테스트 버튼
        self.test_button = ttk.Button(button_frame, text="연결 테스트", 
                                     command=self.test_connection, width=12)
        self.test_button.grid(row=0, column=0, padx=(0, 5))
        
        # 명세서 생성 버튼
        self.generate_button = ttk.Button(button_frame, text="명세서 생성", 
                                         command=self.generate_spec, width=12,
                                         style="Accent.TButton")
        self.generate_button.grid(row=0, column=1, padx=(5, 5))
        
        # 테이블 목록 생성 버튼
        self.generate_list_button = ttk.Button(button_frame, text="목록 생성", 
                                              command=self.generate_table_list, width=12)
        self.generate_list_button.grid(row=0, column=2, padx=(5, 0))
        
        # 진행 상황 표시
        self.progress_var = tk.StringVar(value="준비됨")
        self.progress_label = ttk.Label(parent, textvariable=self.progress_var, 
                                       font=("Arial", 10))
        self.progress_label.grid(row=5, column=0, columnspan=2, pady=5)
        
        # 프로그레스 바
        self.progress_bar = ttk.Progressbar(parent, mode='indeterminate')
        self.progress_bar.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
    def create_log_frame(self, parent):
        """로그/상태 메시지 프레임"""
        log_group = ttk.LabelFrame(parent, text="실행 로그", padding="5")
        log_group.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        log_group.grid_rowconfigure(0, weight=1)
        log_group.grid_columnconfigure(0, weight=1)
        
        # 로그 텍스트 영역
        self.log_text = scrolledtext.ScrolledText(log_group, height=8, width=70)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 로그 클리어 버튼
        clear_button = ttk.Button(log_group, text="로그 지우기", 
                                 command=self.clear_log, width=12)
        clear_button.grid(row=1, column=0, sticky=tk.E, pady=(5, 0))
        
    def setup_logger(self):
        """로거 설정"""
        self.logger = Logger(self.log_text)
        
    def on_dbms_change(self, event=None):
        """DBMS 변경 시 기본 포트 설정 및 Oracle 프레임 표시/숨김"""
        dbms = self.dbms_var.get()
        if dbms in SUPPORTED_DBMS:
            default_port = SUPPORTED_DBMS[dbms]["default_port"]
            self.port_var.set(str(default_port))
            
            # Oracle 선택 시 연결 방식 라벨과 프레임 표시, 다른 DBMS 선택 시 숨김
            if dbms == "Oracle":
                self.oracle_label.grid()
                self.oracle_frame.grid()
                # Oracle 기본값으로 Service Name 선택
                self.oracle_type_var.set("service_name")
                self.update_database_label()
                # Oracle 선택 시 윈도우 크기 증가
                self.adjust_window_size_for_oracle(True)
            else:
                self.oracle_label.grid_remove()
                self.oracle_frame.grid_remove()
                # 다른 DBMS는 일반적인 "데이터베이스 이름" 라벨
                self.database_label.config(text="데이터베이스 이름:")
                # 다른 DBMS 선택 시 윈도우 크기 원래대로
                self.adjust_window_size_for_oracle(False)
            
            if self.logger:
                self.logger.info(f"DBMS가 {dbms}로 변경되었습니다. 기본 포트: {default_port}")
    
    def on_oracle_type_change(self):
        """Oracle 연결 방식 변경 시 라벨 업데이트"""
        self.update_database_label()
        if self.logger:
            oracle_type = self.oracle_type_var.get()
            type_name = "Service Name" if oracle_type == "service_name" else "SID"
            self.logger.info(f"Oracle 연결 방식이 {type_name}으로 변경되었습니다.")
    
    def update_database_label(self):
        """Oracle 연결 방식에 따라 데이터베이스 라벨 업데이트"""
        oracle_type = self.oracle_type_var.get()
        if oracle_type == "service_name":
            self.database_label.config(text="서비스명:")
        else:  # sid
            self.database_label.config(text="SID:")
    
    def adjust_window_size_for_oracle(self, is_oracle: bool):
        """Oracle 선택 여부에 따라 윈도우 크기 조정"""
        if is_oracle:
            # Oracle 선택 시: 추가 필드로 인해 높이 증가 (35px 정도 추가)
            new_height = 785
        else:
            # 다른 DBMS 선택 시: 기본 높이
            new_height = 750
        
        # 현재 윈도우 크기와 위치 정보 가져오기
        current_geometry = self.root.geometry()
        # geometry 형태: "widthxheight+x+y"
        parts = current_geometry.split('+')
        size_part = parts[0]  # "widthxheight"
        width = size_part.split('x')[0]  # 현재 너비 유지
        
        if len(parts) >= 3:
            # 현재 위치 유지
            x_pos = parts[1]
            y_pos = parts[2]
            new_geometry = f"{width}x{new_height}+{x_pos}+{y_pos}"
        else:
            # 위치 정보가 없으면 크기만 조정
            new_geometry = f"{width}x{new_height}"
        
        # 윈도우 크기 조정
        self.root.geometry(new_geometry)
        
        # 최소 윈도우 크기도 함께 조정
        min_width = APP_CONFIG["min_window_size"][0]
        self.root.minsize(min_width, new_height - 100)  # 최소 높이는 조금 작게
        
    def browse_save_path(self):
        """저장 경로 선택"""
        folder = filedialog.askdirectory(initialdir=self.save_path_var.get())
        if folder:
            self.save_path_var.set(folder)
            if self.logger:
                self.logger.info(f"저장 경로가 변경되었습니다: {folder}")
        
    def clear_log(self):
        """로그 지우기"""
        self.log_text.delete(1.0, tk.END)
        
    def reset_form(self):
        """입력값 초기화"""
        # 확인 메시지
        if messagebox.askyesno("초기화 확인", "모든 입력값을 초기화하시겠습니까?\n\n이 작업은 되돌릴 수 없습니다."):
            # DB 연결 정보 초기화
            self.dbms_var.set("MySQL")  # 기본값
            self.host_var.set("localhost")  # 기본 호스트
            self.port_var.set("3306")  # MySQL 기본 포트
            self.database_var.set("")
            self.username_var.set("")
            self.password_var.set("")
            
            # Oracle 라벨과 프레임 숨김 및 라벨 초기화 (MySQL로 초기화되므로)
            self.oracle_label.grid_remove()
            self.oracle_frame.grid_remove()
            self.database_label.config(text="데이터베이스 이름:")
            # 윈도우 크기도 기본값으로 리셋
            self.adjust_window_size_for_oracle(False)
            
            # 파일 설정 초기화
            self.save_path_var.set(os.getcwd())
            self.filename_var.set(FILE_CONFIG["default_filename"])
            
            # 상태 초기화
            self.progress_var.set("준비됨")
            self.progress_bar.stop()
            
            # 로그 지우기
            self.clear_log()
            
            # 버튼 활성화
            self.test_button.config(state='normal')
            self.generate_button.config(state='normal')
            self.generate_list_button.config(state='normal')
            
            if self.logger:
                self.logger.info("✅ 입력값이 초기화되었습니다.")
                self.logger.info(UI_MESSAGES["ready"])
                
    def open_file_location(self, file_path):
        """파일 위치를 탐색기에서 열기 (파일 선택된 상태로)"""
        try:
            # 경로를 절대 경로로 변환하고 정규화
            file_path = os.path.abspath(file_path)
            
            if self.logger:
                self.logger.info(f"📂 파일 위치 열기 시도: {file_path}")
            
            # 파일이 실제로 존재하는지 확인
            if not os.path.exists(file_path):
                if self.logger:
                    self.logger.warning(f"⚠️ 파일이 존재하지 않습니다: {file_path}")
                # 파일이 없으면 폴더만 열기
                folder_path = os.path.dirname(file_path)
                if os.path.exists(folder_path):
                    result = subprocess.run(['explorer', folder_path], capture_output=True)
                    if result.returncode == 0 or result.returncode == 1:
                        if self.logger:
                            self.logger.info(f"📂 폴더를 열었습니다: {folder_path}")
                        return
                # 폴더도 없으면 에러
                raise FileNotFoundError(f"파일과 폴더가 모두 존재하지 않습니다: {file_path}")
            
            # 방법 1: explorer /select 시도 (파일 선택된 상태로)
            try:
                result = subprocess.run(['explorer', '/select,', file_path], 
                                      capture_output=True, text=True, timeout=10)
                if result.returncode == 0 or result.returncode == 1:
                    if self.logger:
                        self.logger.info(f"📂 파일 위치를 열었습니다: {os.path.dirname(file_path)}")
                    return
                else:
                    if self.logger:
                        self.logger.debug(f"explorer /select 실패 (returncode: {result.returncode})")
                    raise subprocess.CalledProcessError(result.returncode, 'explorer')
            except Exception as e:
                if self.logger:
                    self.logger.debug(f"explorer /select 명령 실패: {str(e)}")
                
                # 방법 2: 폴더만 열기
                folder_path = os.path.dirname(file_path)
                try:
                    result = subprocess.run(['explorer', folder_path], 
                                          capture_output=True, text=True, timeout=10)
                    if result.returncode == 0 or result.returncode == 1:
                        if self.logger:
                            self.logger.info(f"📂 폴더를 열었습니다: {folder_path}")
                        return
                    else:
                        raise subprocess.CalledProcessError(result.returncode, 'explorer')
                except Exception as e2:
                    if self.logger:
                        self.logger.debug(f"explorer 폴더 열기 실패: {str(e2)}")
                    
                    # 방법 3: os.startfile 시도 (최후의 수단)
                    try:
                        os.startfile(folder_path)
                        if self.logger:
                            self.logger.info(f"📂 폴더를 열었습니다 (startfile): {folder_path}")
                        return
                    except Exception as e3:
                        if self.logger:
                            self.logger.error(f"❌ 모든 폴더 열기 방법이 실패했습니다: {str(e3)}")
                        raise e3
                        
        except Exception as e:
            if self.logger:
                self.logger.error(f"❌ 폴더 열기에 실패했습니다: {str(e)}")
            messagebox.showerror("오류", f"폴더를 열 수 없습니다.\n\n파일 경로: {file_path}\n오류: {str(e)}")
        
    def validate_input(self):
        """입력값 유효성 검사"""
        if not self.host_var.get().strip():
            messagebox.showerror("입력 오류", ERROR_MESSAGES["empty_host"])
            return False
            
        if not self.port_var.get().strip():
            messagebox.showerror("입력 오류", ERROR_MESSAGES["empty_port"])
            return False
            
        if not validate_port(self.port_var.get()):
            messagebox.showerror("입력 오류", ERROR_MESSAGES["invalid_port"])
            return False
            
        if not self.database_var.get().strip():
            messagebox.showerror("입력 오류", ERROR_MESSAGES["empty_database"])
            return False
            
        if not self.username_var.get().strip():
            messagebox.showerror("입력 오류", ERROR_MESSAGES["empty_username"])
            return False
            
        return True
        
    def get_connection_info(self):
        """연결 정보 딕셔너리 반환"""
        conn_info = {
            'dbms': self.dbms_var.get(),
            'host': self.host_var.get().strip(),
            'port': int(self.port_var.get()),
            'database': self.database_var.get().strip(),
            'username': self.username_var.get().strip(),
            'password': self.password_var.get()
        }
        
        # Oracle인 경우 연결 방식 추가
        if conn_info['dbms'] == 'Oracle':
            conn_info['oracle_type'] = self.oracle_type_var.get()
            
        return conn_info
        
    def test_connection(self):
        """DB 연결 테스트"""
        if not self.validate_input():
            return
            
        if self.logger:
            self.logger.info(UI_MESSAGES["connecting"])
        self.progress_var.set("연결 테스트 중...")
        self.progress_bar.start()
        
        # 버튼 비활성화
        self.test_button.config(state='disabled')
        self.generate_button.config(state='disabled')
        self.generate_list_button.config(state='disabled')
        
        # 별도 스레드에서 연결 테스트 실행
        threading.Thread(target=self._test_connection_thread, daemon=True).start()
        
    def _test_connection_thread(self):
        """연결 테스트 스레드"""
        try:
            conn_info = self.get_connection_info()
            
            # 실제 DB 연결 테스트 실행
            result = connection_manager.test_connection(
                dbms=conn_info['dbms'],
                host=conn_info['host'],
                port=conn_info['port'],
                database=conn_info['database'],
                username=conn_info['username'],
                password=conn_info['password'],
                timeout=30,
                oracle_type=conn_info.get('oracle_type')
            )
            
            if result['success']:
                # 연결 성공 - 추가 정보와 함께 성공 처리
                self.root.after(0, lambda: self._test_connection_success(result))
            else:
                # 연결 실패
                self.root.after(0, lambda: self._test_connection_error(result['error']))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self._test_connection_error(msg))
            
    def _test_connection_success(self, result=None):
        """연결 테스트 성공 처리"""
        self.progress_bar.stop()
        self.progress_var.set("연결 성공!")
        
        if self.logger:
            if result:
                # 상세 정보와 함께 로그
                self.logger.success(f"{UI_MESSAGES['connected']} - {result['dbms']}")
                self.logger.info(f"서버: {result['host']}:{result['port']}")
                self.logger.info(f"데이터베이스: {result['database']}")
                self.logger.info(f"버전: {result.get('version', 'Unknown')}")
                self.logger.info(f"응답시간: {result.get('connection_time_ms', 0)}ms")
            else:
                self.logger.success(UI_MESSAGES["connected"])
        
        # 성공 메시지
        if result:
            detail_msg = (
                f"데이터베이스 연결이 성공했습니다!\n\n"
                f"DBMS: {result['dbms']}\n"
                f"서버: {result['host']}:{result['port']}\n"
                f"데이터베이스: {result['database']}\n"
                f"버전: {result.get('version', 'Unknown')}\n"
                f"응답시간: {result.get('connection_time_ms', 0)}ms"
            )
            # 버튼 활성화 (메시지박스 전에)
            self.test_button.config(state='normal')
            self.generate_button.config(state='normal')
            self.generate_list_button.config(state='normal')
            messagebox.showinfo("연결 성공", detail_msg)
        else:
            # 버튼 활성화 (메시지박스 전에)
            self.test_button.config(state='normal')
            self.generate_button.config(state='normal')
            self.generate_list_button.config(state='normal')
            messagebox.showinfo("연결 성공", "데이터베이스 연결이 성공했습니다.")
        
    def _test_connection_error(self, error_msg):
        """연결 테스트 실패 처리"""
        self.progress_bar.stop()
        self.progress_var.set("연결 실패")
        if self.logger:
            self.logger.error(f"{UI_MESSAGES['connection_failed']}: {error_msg}")
        
        # 버튼 활성화 (메시지박스 전에)
        self.test_button.config(state='normal')
        self.generate_button.config(state='normal')
        self.generate_list_button.config(state='normal')
        
        messagebox.showerror("연결 실패", f"데이터베이스 연결에 실패했습니다.\n\n{error_msg}")
        
    def generate_spec(self):
        """명세서 생성 (테이블 선택 포함)"""
        if not self.validate_input():
            return
            
        filename = self.filename_var.get().strip()
        if not filename:
            messagebox.showerror("입력 오류", ERROR_MESSAGES["empty_filename"])
            return
            
        if not validate_filename(filename):
            messagebox.showerror("입력 오류", ERROR_MESSAGES["invalid_filename"])
            return
        
        # 테이블 목록 가져오기 및 선택
        self._show_table_selector_for_spec()
    
    def _show_table_selector_for_spec(self):
        """명세서 생성용 테이블 선택"""
        if self.logger:
            self.logger.info("테이블 목록을 가져오는 중...")
        self.progress_var.set("테이블 목록 조회 중...")
        self.progress_bar.start()
        
        # 버튼 비활성화
        self.test_button.config(state='disabled')
        self.generate_button.config(state='disabled')
        self.generate_list_button.config(state='disabled')
        
        # 별도 스레드에서 테이블 목록 가져오기
        threading.Thread(target=self._get_tables_for_spec_selection, daemon=True).start()
    
    def _get_tables_for_spec_selection(self):
        """명세서용 테이블 목록 가져오기"""
        try:
            conn_info = self.get_connection_info()
            
            # 테이블 목록 수집
            table_list_data = metadata_collector.collect_table_list(
                dbms=conn_info['dbms'],
                host=conn_info['host'],
                port=conn_info['port'],
                database=conn_info['database'],
                username=conn_info['username'],
                password=conn_info['password'],
                timeout=30
            )
            
            # GUI 업데이트는 메인 스레드에서
            self.root.after(0, lambda: self._show_table_selector_dialog_for_spec(table_list_data['table_list']))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self._table_selection_error(msg, "spec"))
    
    def _show_table_selector_dialog_for_spec(self, table_list):
        """명세서용 테이블 선택 다이얼로그 표시"""
        self.progress_bar.stop()
        self.progress_var.set("테이블 선택 중...")
        
        if self.logger:
            self.logger.info(f"테이블 목록 조회 완료: {len(table_list)}개 테이블")
        
        # 테이블 선택 다이얼로그 표시
        selected_tables = show_table_selector(
            self.root, 
            table_list, 
            "명세서 생성 대상 테이블 선택"
        )
        
        if selected_tables:
            if self.logger:
                self.logger.info(f"선택된 테이블: {len(selected_tables)}개")
            # 선택된 테이블로 명세서 생성
            self._generate_spec_with_selected_tables(selected_tables)
        else:
            # 취소된 경우
            self.progress_var.set("준비됨")
            self.test_button.config(state='normal')
            self.generate_button.config(state='normal')
            self.generate_list_button.config(state='normal')
            if self.logger:
                self.logger.info("테이블 선택이 취소되었습니다.")
    
    def _generate_spec_with_selected_tables(self, selected_tables):
        """선택된 테이블들로 명세서 생성"""
        if self.logger:
            self.logger.info(UI_MESSAGES["generating"])
        self.progress_var.set("명세서 생성 중...")
        self.progress_bar.start()
        
        # 선택된 테이블 이름 목록
        selected_table_names = [table['table_name'] for table in selected_tables]
        
        # 별도 스레드에서 명세서 생성 실행
        threading.Thread(target=self._generate_spec_thread, args=(selected_table_names,), daemon=True).start()
    
    def _table_selection_error(self, error_msg, operation_type):
        """테이블 선택 중 오류 처리"""
        self.progress_bar.stop()
        self.progress_var.set("테이블 조회 실패")
        
        if self.logger:
            self.logger.error(f"테이블 목록 조회 실패: {error_msg}")
        
        # 버튼 활성화
        self.test_button.config(state='normal')
        self.generate_button.config(state='normal')
        self.generate_list_button.config(state='normal')
        
        operation_name = "테이블 명세서" if operation_type == "spec" else "테이블 목록"
        messagebox.showerror("테이블 조회 실패", 
                            f"{operation_name} 생성을 위한 테이블 목록 조회에 실패했습니다.\n\n{error_msg}")
        
    def _generate_spec_thread(self, selected_table_names=None):
        """명세서 생성 스레드"""
        try:
            conn_info = self.get_connection_info()
            
            # 파일명에 명세서 접미사 추가
            original_filename = self.filename_var.get()
            if '.xlsx' in original_filename:
                spec_filename = original_filename.replace('.xlsx', '_명세서.xlsx')
            else:
                spec_filename = original_filename + '_명세서.xlsx'
                
            save_path = os.path.join(self.save_path_var.get(), spec_filename)
            
            # 실제 메타데이터 수집
            if self.logger:
                if selected_table_names:
                    self.logger.info(f"선택된 {len(selected_table_names)}개 테이블의 메타데이터 수집을 시작합니다...")
                else:
                    self.logger.info("데이터베이스 메타데이터 수집을 시작합니다...")
            
            metadata = metadata_collector.collect_database_metadata(
                dbms=conn_info['dbms'],
                host=conn_info['host'],
                port=conn_info['port'],
                database=conn_info['database'],
                username=conn_info['username'],
                password=conn_info['password'],
                timeout=30,
                selected_tables=selected_table_names,
                oracle_type=conn_info.get('oracle_type')
            )
            
            if self.logger:
                self.logger.info(f"메타데이터 수집 완료: 테이블 {metadata['statistics']['total_tables']}개, 컬럼 {metadata['statistics']['total_columns']}개")
                self.logger.info(f"수집 시간: {metadata['statistics']['collection_duration_ms']}ms")
            
            # Excel 파일 생성
            if self.logger:
                self.logger.info("Excel 명세서 생성을 시작합니다...")
                
            # 파일명에 .xlsx 확장자 확인
            save_path = ensure_excel_extension(save_path)
            
            excel_path = excel_generator.generate_excel(metadata, save_path)
            
            if self.logger:
                self.logger.info(f"Excel 명세서 생성 완료: {excel_path}")
            
            # GUI 업데이트는 메인 스레드에서
            result_info = {
                'save_path': excel_path,
                'metadata': metadata
            }
            self.root.after(0, lambda: self._generate_spec_success(result_info))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self._generate_spec_error(msg))
            
    def _generate_spec_success(self, result_info):
        """명세서 생성 성공 처리"""
        self.progress_bar.stop()
        self.progress_var.set("생성 완료!")
        
        if self.logger:
            if isinstance(result_info, dict) and 'metadata' in result_info:
                metadata = result_info['metadata']
                self.logger.success(f"{UI_MESSAGES['generated']}")
                self.logger.info(f"수집 결과: 테이블 {metadata['statistics']['total_tables']}개")
                self.logger.info(f"저장 위치: {result_info['save_path']}")
            else:
                self.logger.success(f"{UI_MESSAGES['generated']}: {result_info}")
        
        # 상세 성공 메시지
        if isinstance(result_info, dict) and 'metadata' in result_info:
            metadata = result_info['metadata']
            stats = metadata['statistics']
            
            # 파일 크기 확인
            excel_size = "Unknown"
            if os.path.exists(result_info['save_path']):
                excel_file_size = os.path.getsize(result_info['save_path'])
                excel_size = f"{excel_file_size:,} bytes"
                
            detail_msg = (
                f"테이블 명세서가 성공적으로 생성되었습니다! 🎉\n\n"
                f"📊 수집 결과:\n"
                f"  • 테이블: {stats['total_tables']}개\n"
                f"  • 컬럼: {stats['total_columns']}개\n"
                f"  • 외래키: {stats['total_foreign_keys']}개\n"
                f"  • 수집 시간: {stats['collection_duration_ms']}ms\n\n"
                f"📋 생성된 Excel 명세서:\n"
                f"  • 파일 위치: {result_info['save_path']}\n"
                f"  • 파일 크기: {excel_size}\n"
                f"  • 시트 구성: 테이블명세서 (통합)\n\n"
                f"✅ 모든 테이블이 하나의 시트에 정리되었습니다."
            )
        else:
            detail_msg = f"테이블 명세서가 성공적으로 생성되었습니다.\n\n저장 위치: {result_info}"
            
        # 버튼 활성화 (메시지박스 전에)
        self.test_button.config(state='normal')
        self.generate_button.config(state='normal')
        self.generate_list_button.config(state='normal')
        
        messagebox.showinfo("생성 완료", detail_msg)
        
        # 폴더 열기 옵션 제공
        if messagebox.askyesno("폴더 열기", "생성된 Excel 파일이 있는 폴더를 여시겠습니까?"):
            self.open_file_location(result_info['save_path'])
        
    def _generate_spec_error(self, error_msg):
        """명세서 생성 실패 처리"""
        self.progress_bar.stop()
        self.progress_var.set("생성 실패")
        if self.logger:
            self.logger.error(f"{UI_MESSAGES['generation_failed']}: {error_msg}")
        # 버튼 활성화 (메시지박스 전에)
        self.test_button.config(state='normal')
        self.generate_button.config(state='normal')
        self.generate_list_button.config(state='normal')
        
        messagebox.showerror("생성 실패", f"테이블 명세서 생성에 실패했습니다.\n\n{error_msg}")
    
    def generate_table_list(self):
        """테이블 목록 생성 (테이블 선택 포함)"""
        if not self.validate_input():
            return
            
        filename = self.filename_var.get().strip()
        if not filename:
            messagebox.showerror("입력 오류", ERROR_MESSAGES["empty_filename"])
            return
            
        if not validate_filename(filename):
            messagebox.showerror("입력 오류", ERROR_MESSAGES["invalid_filename"])
            return
        
        # 테이블 목록 가져오기 및 선택
        self._show_table_selector_for_list()
    
    def _show_table_selector_for_list(self):
        """목록 생성용 테이블 선택"""
        if self.logger:
            self.logger.info("테이블 목록을 가져오는 중...")
        self.progress_var.set("테이블 목록 조회 중...")
        self.progress_bar.start()
        
        # 버튼 비활성화
        self.test_button.config(state='disabled')
        self.generate_button.config(state='disabled')
        self.generate_list_button.config(state='disabled')
        
        # 별도 스레드에서 테이블 목록 가져오기
        threading.Thread(target=self._get_tables_for_list_selection, daemon=True).start()
    
    def _get_tables_for_list_selection(self):
        """목록용 테이블 목록 가져오기"""
        try:
            conn_info = self.get_connection_info()
            
            # 테이블 목록 수집
            table_list_data = metadata_collector.collect_table_list(
                dbms=conn_info['dbms'],
                host=conn_info['host'],
                port=conn_info['port'],
                database=conn_info['database'],
                username=conn_info['username'],
                password=conn_info['password'],
                timeout=30
            )
            
            # GUI 업데이트는 메인 스레드에서
            self.root.after(0, lambda: self._show_table_selector_dialog_for_list(table_list_data['table_list']))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self._table_selection_error(msg, "list"))
    
    def _show_table_selector_dialog_for_list(self, table_list):
        """목록용 테이블 선택 다이얼로그 표시"""
        self.progress_bar.stop()
        self.progress_var.set("테이블 선택 중...")
        
        if self.logger:
            self.logger.info(f"테이블 목록 조회 완료: {len(table_list)}개 테이블")
        
        # 테이블 선택 다이얼로그 표시
        selected_tables = show_table_selector(
            self.root, 
            table_list, 
            "목록 생성 대상 테이블 선택"
        )
        
        if selected_tables:
            if self.logger:
                self.logger.info(f"선택된 테이블: {len(selected_tables)}개")
            # 선택된 테이블로 목록 생성
            self._generate_list_with_selected_tables(selected_tables)
        else:
            # 취소된 경우
            self.progress_var.set("준비됨")
            self.test_button.config(state='normal')
            self.generate_button.config(state='normal')
            self.generate_list_button.config(state='normal')
            if self.logger:
                self.logger.info("테이블 선택이 취소되었습니다.")
    
    def _generate_list_with_selected_tables(self, selected_tables):
        """선택된 테이블들로 목록 생성"""
        if self.logger:
            self.logger.info("테이블 목록 생성을 시작합니다...")
        self.progress_var.set("테이블 목록 생성 중...")
        self.progress_bar.start()
        
        # 별도 스레드에서 목록 생성 실행
        threading.Thread(target=self._generate_table_list_thread, args=(selected_tables,), daemon=True).start()
        
    def _generate_table_list_thread(self, selected_tables=None):
        """테이블 목록 생성 스레드"""
        try:
            conn_info = self.get_connection_info()
            
            # 파일명에 목록 접미사 추가
            original_filename = self.filename_var.get()
            if '.xlsx' in original_filename:
                list_filename = original_filename.replace('.xlsx', '_목록.xlsx')
            else:
                list_filename = original_filename + '_목록.xlsx'
                
            save_path = os.path.join(self.save_path_var.get(), list_filename)
            
            # 선택된 테이블이 있으면 해당 테이블만 사용, 없으면 전체 수집
            if selected_tables:
                if self.logger:
                    self.logger.info(f"선택된 {len(selected_tables)}개 테이블의 목록을 생성합니다...")
                
                # 선택된 테이블들로 목록 데이터 구성
                table_list = []
                for idx, table in enumerate(selected_tables, 1):
                    table_list.append({
                        'no': idx,
                        'table_name': table.get('table_name', ''),
                        'table_comment': table.get('table_comment', '') or ''
                    })
                
                table_list_data = {
                    'connection_info': {
                        'dbms': conn_info['dbms'],
                        'host': conn_info['host'],
                        'port': conn_info['port'],
                        'database': conn_info['database'],
                        'username': conn_info['username'],
                        'collection_time': time.strftime("%Y-%m-%d %H:%M:%S")
                    },
                    'table_list': table_list,
                    'statistics': {
                        'total_tables': len(table_list),
                        'collection_duration_ms': 0  # 이미 수집된 데이터 사용
                    }
                }
            else:
                # 실제 테이블 목록 수집
                if self.logger:
                    self.logger.info("데이터베이스 테이블 목록 수집을 시작합니다...")
                
                table_list_data = metadata_collector.collect_table_list(
                    dbms=conn_info['dbms'],
                    host=conn_info['host'],
                    port=conn_info['port'],
                    database=conn_info['database'],
                    username=conn_info['username'],
                    password=conn_info['password'],
                    timeout=30,
                    oracle_type=conn_info.get('oracle_type')
                )
            
            if self.logger:
                self.logger.info(f"테이블 목록 수집 완료: {table_list_data['statistics']['total_tables']}개")
                self.logger.info(f"수집 시간: {table_list_data['statistics']['collection_duration_ms']}ms")
            
            # Excel 파일 생성
            if self.logger:
                self.logger.info("Excel 테이블 목록을 생성합니다...")
            
            # 파일명에 .xlsx 확장자 확인
            save_path = ensure_excel_extension(save_path)
            
            excel_path = excel_generator.generate_table_list_excel(table_list_data, save_path)
            
            if self.logger:
                self.logger.info(f"Excel 테이블 목록 생성 완료: {excel_path}")
            
            # GUI 업데이트는 메인 스레드에서
            result_info = {
                'save_path': excel_path,
                'table_list_data': table_list_data
            }
            self.root.after(0, lambda: self._generate_table_list_success(result_info))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self._generate_table_list_error(msg))
    
    def _generate_table_list_success(self, result_info):
        """테이블 목록 생성 성공 처리"""
        self.progress_bar.stop()
        self.progress_var.set("목록 생성 완료!")
        
        if self.logger:
            table_list_data = result_info['table_list_data']
            self.logger.success("✅ 테이블 목록이 성공적으로 생성되었습니다")
            self.logger.info(f"수집 결과: 테이블 {table_list_data['statistics']['total_tables']}개")
            self.logger.info(f"저장 위치: {result_info['save_path']}")
        
        # 상세 성공 메시지
        table_list_data = result_info['table_list_data']
        stats = table_list_data['statistics']
        
        # 파일 크기 확인
        excel_size = "Unknown"
        if os.path.exists(result_info['save_path']):
            excel_file_size = os.path.getsize(result_info['save_path'])
            excel_size = f"{excel_file_size:,} bytes"
            
        detail_msg = (
            f"테이블 목록이 성공적으로 생성되었습니다! 🎉\n\n"
            f"📊 수집 결과:\n"
            f"  • 테이블: {stats['total_tables']}개\n"
            f"  • 수집 시간: {stats['collection_duration_ms']}ms\n\n"
            f"📋 생성된 Excel 목록:\n"
            f"  • 파일 위치: {result_info['save_path']}\n"
            f"  • 파일 크기: {excel_size}\n"
            f"  • 시트 구성: 테이블목록\n\n"
            f"✅ NO, 테이블명, 논리명 형식으로 정리되었습니다."
        )
        
        # 버튼 활성화 (메시지박스 전에)
        self.test_button.config(state='normal')
        self.generate_button.config(state='normal')
        self.generate_list_button.config(state='normal')
        
        messagebox.showinfo("목록 생성 완료", detail_msg)
        
        # 폴더 열기 옵션 제공
        if messagebox.askyesno("폴더 열기", "생성된 Excel 파일이 있는 폴더를 여시겠습니까?"):
            self.open_file_location(result_info['save_path'])
    
    def _generate_table_list_error(self, error_msg):
        """테이블 목록 생성 실패 처리"""
        self.progress_bar.stop()
        self.progress_var.set("목록 생성 실패")
        if self.logger:
            self.logger.error(f"테이블 목록 생성 실패: {error_msg}")
        
        # 버튼 활성화 (메시지박스 전에)
        self.test_button.config(state='normal')
        self.generate_button.config(state='normal')
        self.generate_list_button.config(state='normal')
        
        messagebox.showerror("목록 생성 실패", f"테이블 목록 생성에 실패했습니다.\n\n{error_msg}")
        
    def on_closing(self):
        """윈도우 종료 시 호출되는 메서드"""
        try:
            # JVM 종료 (Oracle JDBC 사용 시)
            self._shutdown_jvm()
            
            if self.logger:
                self.logger.info("프로그램이 종료됩니다.")
                
        except Exception as e:
            print(f"종료 중 오류 발생 (무시됨): {e}")
        finally:
            # 윈도우 종료
            self.root.destroy()
    
    def _shutdown_jvm(self):
        """JVM 종료 처리"""
        try:
            import jpype
            if jpype.isJVMStarted():
                jpype.shutdownJVM()
                print("JVM이 정상적으로 종료되었습니다.")
        except ImportError:
            # jpype가 없으면 무시
            pass
        except Exception as e:
            print(f"JVM 종료 중 오류: {e}")
    
    def run(self):
        """애플리케이션 실행"""
        if self.logger:
            self.logger.info(UI_MESSAGES["startup"])
            self.logger.info(UI_MESSAGES["ready"])
        self.root.mainloop()


if __name__ == "__main__":
    app = DBSpecGeneratorApp()
    app.run()