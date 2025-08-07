"""
테이블 선택 다이얼로그
"""

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import font


class TableSelectorDialog:
    """테이블 선택 다이얼로그 클래스"""
    
    def __init__(self, parent, table_list, title="테이블 선택"):
        self.parent = parent
        self.table_list = table_list
        self.title = title
        self.selected_tables = []
        self.result = None
        
        # 전체 테이블의 선택 상태를 저장하는 딕셔너리
        self.global_selections = {}
        for table in table_list:
            table_name = table.get('table_name', '')
            self.global_selections[table_name] = True  # 기본적으로 모든 테이블 선택
        
        # 다이얼로그 창 생성
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("500x600")
        self.dialog.resizable(True, True)
        
        # 다이얼로그 아이콘 설정
        self._set_dialog_icon(parent)
        
        # 모달 창으로 설정
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 중앙에 위치
        self._center_window()
        
        # UI 구성
        self.create_widgets()
        
        # 엔터키로 확인
        self.dialog.bind('<Return>', lambda e: self.confirm())
        self.dialog.bind('<Escape>', lambda e: self.cancel())
        
    def _set_dialog_icon(self, parent):
        """다이얼로그 아이콘 설정"""
        try:
            import sys
            import os
            
            # 현재 스크립트의 디렉토리 경로
            if getattr(sys, 'frozen', False):
                # 실행파일(exe)로 실행된 경우 - PyInstaller의 임시 폴더
                base_path = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
            else:
                # 개발 환경에서 실행된 경우
                base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            
            icon_path = os.path.join(base_path, "dboutput.ico")
            
            if os.path.exists(icon_path):
                self.dialog.iconbitmap(icon_path)
            else:
                # ico 파일이 없으면 부모 윈도우의 아이콘 복사 시도
                try:
                    # 부모의 아이콘을 복사
                    parent_icon = parent.iconbitmap()
                    if parent_icon:
                        self.dialog.iconbitmap(parent_icon)
                except:
                    pass
        except Exception:
            # 아이콘 설정 실패해도 다이얼로그는 정상 표시
            pass
    
    def _center_window(self):
        """창을 화면 중앙에 위치"""
        self.dialog.update_idletasks()
        
        # 부모 창의 위치와 크기
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        # 다이얼로그 크기
        dialog_width = 500
        dialog_height = 600
        
        # 중앙 계산
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
    
    def create_widgets(self):
        """위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목
        title_label = ttk.Label(main_frame, text=f"생성할 테이블을 선택하세요", 
                               font=("맑은 고딕", 12, "bold"))
        title_label.pack(pady=(0, 10))
        
        # 정보 표시
        info_text = f"총 {len(self.table_list)}개의 테이블이 발견되었습니다."
        info_label = ttk.Label(main_frame, text=info_text, 
                              font=("맑은 고딕", 9))
        info_label.pack(pady=(0, 10))
        
        # 검색 프레임
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(search_frame, text="검색:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_tables)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        
        # 전체 선택/해제 버튼
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(button_frame, text="전체 선택", 
                  command=self.select_all, width=12).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="전체 해제", 
                  command=self.deselect_all, width=12).pack(side=tk.LEFT, padx=(5, 0))
        
        # 선택된 개수 표시
        self.selection_label = ttk.Label(button_frame, text="", 
                                        font=("맑은 고딕", 9))
        self.selection_label.pack(side=tk.RIGHT)
        
        # 테이블 리스트 프레임 (스크롤바 포함)
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 리스트박스 프레임
        self.listbox_frame = ttk.Frame(list_frame)
        self.listbox_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Canvas와 Scrollbar 설정 (체크박스들을 위해)
        self.canvas = tk.Canvas(self.listbox_frame, highlightthickness=0)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.configure(command=self.canvas.yview)
        
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 마우스 휠 바인딩
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 체크박스 변수들
        self.check_vars = {}
        self.checkboxes = {}
        
        # 테이블 체크박스 생성
        self.create_table_checkboxes()
        
        # 확인/취소 버튼
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(bottom_frame, text="취소", 
                  command=self.cancel, width=12).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(bottom_frame, text="확인", 
                  command=self.confirm, width=12).pack(side=tk.RIGHT)
        
        # 선택 개수 업데이트
        self.update_selection_count()
    
    def create_table_checkboxes(self):
        """테이블 체크박스들 생성"""
        # 기존 체크박스들 제거
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.check_vars.clear()
        self.checkboxes.clear()
        
        # 검색 필터 적용
        search_text = self.search_var.get().lower()
        filtered_tables = []
        
        for table in self.table_list:
            table_name = table.get('table_name', '')
            table_comment = table.get('table_comment', '')
            
            if (search_text in table_name.lower() or 
                search_text in table_comment.lower() or 
                not search_text):
                filtered_tables.append(table)
        
        # 빈 목록 처리
        if not filtered_tables:
            # 빈 목록 메시지 표시
            if not search_text:
                # 전체 목록이 비어있는 경우
                message = "조회된 테이블이 존재하지 않습니다."
            else:
                # 검색 결과가 비어있는 경우
                message = f"'{search_text}' 검색 결과가 없습니다."
            
            empty_frame = ttk.Frame(self.scrollable_frame)
            empty_frame.pack(fill=tk.X, padx=20, pady=20)
            
            empty_label = ttk.Label(
                empty_frame, 
                text=message,
                font=("맑은 고딕", 10),
                foreground="gray"
            )
            empty_label.pack()
        
        # 체크박스 생성
        for idx, table in enumerate(filtered_tables):
            table_name = table.get('table_name', '')
            table_comment = table.get('table_comment', '')
            
            # 체크박스 변수 (전역 선택 상태에서 값 가져오기)
            var = tk.BooleanVar()
            var.set(self.global_selections.get(table_name, True))  # 전역 상태에서 복원
            self.check_vars[table_name] = var
            
            # 체크박스 프레임
            cb_frame = ttk.Frame(self.scrollable_frame)
            cb_frame.pack(fill=tk.X, padx=5, pady=2)
            
            # 체크박스 (선택 상태 변경 시 전역 상태도 업데이트)
            def on_checkbox_change(table_name=table_name):
                self.global_selections[table_name] = self.check_vars[table_name].get()
                self.update_selection_count()
                
            checkbox = ttk.Checkbutton(
                cb_frame, 
                variable=var,
                command=on_checkbox_change
            )
            checkbox.pack(side=tk.LEFT)
            self.checkboxes[table_name] = checkbox
            
            # 테이블 정보 라벨
            if table_comment:
                label_text = f"{table_name} ({table_comment})"
            else:
                label_text = table_name
                
            label = ttk.Label(cb_frame, text=label_text, 
                             font=("맑은 고딕", 9))
            label.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
            
            # 라벨 클릭시 체크박스 토글 (전역 상태도 업데이트)
            def toggle_checkbox(event, var=var, table_name=table_name):
                var.set(not var.get())
                self.global_selections[table_name] = var.get()
                self.update_selection_count()
            label.bind("<Button-1>", toggle_checkbox)
        
        # Canvas 스크롤 영역 업데이트
        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def filter_tables(self, *args):
        """테이블 필터링 (전역 선택 상태 유지)"""
        # 현재 화면의 선택 상태를 전역 상태에 저장
        for table_name, var in self.check_vars.items():
            self.global_selections[table_name] = var.get()
        
        # 체크박스 다시 생성 (create_table_checkboxes에서 전역 상태로부터 복원됨)
        self.create_table_checkboxes()
        
        self.update_selection_count()
    
    def select_all(self):
        """현재 표시된 테이블 모두 선택"""
        # 현재 필터링된 테이블들만 선택
        search_text = self.search_var.get().lower()
        for table in self.table_list:
            table_name = table.get('table_name', '')
            table_comment = table.get('table_comment', '')
            
            # 검색 조건에 맞는 테이블만 선택
            if (search_text in table_name.lower() or 
                search_text in table_comment.lower() or 
                not search_text):
                # 전역 상태와 현재 화면 상태 모두 업데이트
                self.global_selections[table_name] = True
                if table_name in self.check_vars:
                    self.check_vars[table_name].set(True)
        self.update_selection_count()
    
    def deselect_all(self):
        """현재 표시된 테이블 모두 선택 해제"""
        # 현재 필터링된 테이블들만 해제
        search_text = self.search_var.get().lower()
        for table in self.table_list:
            table_name = table.get('table_name', '')
            table_comment = table.get('table_comment', '')
            
            # 검색 조건에 맞는 테이블만 해제
            if (search_text in table_name.lower() or 
                search_text in table_comment.lower() or 
                not search_text):
                # 전역 상태와 현재 화면 상태 모두 업데이트
                self.global_selections[table_name] = False
                if table_name in self.check_vars:
                    self.check_vars[table_name].set(False)
        self.update_selection_count()
    
    def update_selection_count(self):
        """선택된 테이블 개수 업데이트"""
        # 현재 화면의 상태를 전역 상태에 반영
        for table_name, var in self.check_vars.items():
            self.global_selections[table_name] = var.get()
        
        # 전역적으로 선택된 테이블 개수
        total_selected = sum(1 for selected in self.global_selections.values() if selected)
        
        # 현재 화면에 보이는 선택된 테이블 개수
        visible_selected = sum(1 for var in self.check_vars.values() if var.get())
        visible_count = len(self.check_vars)  # 현재 화면에 보이는 테이블 수
        total_count = len(self.table_list)    # 전체 테이블 수
        
        if visible_count == total_count:
            # 검색하지 않은 상태
            self.selection_label.config(text=f"선택됨: {total_selected}/{total_count}")
        else:
            # 검색 중인 상태 - 화면의 선택 개수와 전체 선택 개수 모두 표시
            self.selection_label.config(text=f"화면: {visible_selected}/{visible_count} | 전체 선택: {total_selected}/{total_count}")
    
    def get_selected_tables(self):
        """선택된 테이블들 반환 (전역 상태 기준)"""
        # 현재 화면의 상태를 전역 상태에 반영
        for table_name, var in self.check_vars.items():
            self.global_selections[table_name] = var.get()
        
        # 전역 상태에서 선택된 테이블들 수집
        selected = []
        for table_name, is_selected in self.global_selections.items():
            if is_selected:
                # 원본 테이블 정보 찾기
                for table in self.table_list:
                    if table.get('table_name') == table_name:
                        selected.append(table)
                        break
        return selected
    
    def confirm(self):
        """확인 버튼"""
        selected_tables = self.get_selected_tables()
        
        if not selected_tables:
            messagebox.showwarning("선택 오류", "최소 1개 이상의 테이블을 선택해주세요.")
            return
        
        self.result = selected_tables
        self.dialog.destroy()
    
    def cancel(self):
        """취소 버튼"""
        self.result = None
        self.dialog.destroy()
    
    def show(self):
        """다이얼로그 표시 및 결과 반환"""
        self.dialog.wait_window()
        return self.result


def show_table_selector(parent, table_list, title="테이블 선택"):
    """테이블 선택 다이얼로그 표시"""
    dialog = TableSelectorDialog(parent, table_list, title)
    return dialog.show()