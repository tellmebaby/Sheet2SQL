import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import sqlite3
import mysql.connector
from configparser import ConfigParser
import re
import sys
from pathlib import Path

class DatabaseUpdater:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 데이터로 DB 업데이트 도구")
        self.root.geometry("700x700")
        self.root.resizable(True, True)
        
        # 설정 파일 경로
        self.config_file = Path("db_config.ini")
        self.config = ConfigParser()
        
        # 기본 설정 로드 또는 생성
        self.load_or_create_config()
        
        # UI 초기화
        self.create_widgets()
        
    def _on_mousewheel(self, event):
        """마우스 휠 스크롤 이벤트 처리"""
        main_canvas = None
        for child in self.root.winfo_children():
            if isinstance(child, tk.Canvas):
                main_canvas = child
                break
                
        if main_canvas:
            if event.num == 5 or event.delta < 0:  # 아래로 스크롤
                main_canvas.yview_scroll(1, "units")
            elif event.num == 4 or event.delta > 0:  # 위로 스크롤
                main_canvas.yview_scroll(-1, "units")
                
    def load_or_create_config(self):
        """설정 파일을 로드하거나 없으면 기본 설정으로 생성"""
        if self.config_file.exists():
            self.config.read(self.config_file)
        else:
            # 기본 설정 생성
            self.config['DATABASE'] = {
                'type': 'sqlite',
                'host': 'localhost',
                'port': '3306',
                'database': 'mydatabase.db',
                'user': 'username',
                'password': 'password',
                'table': 'users',
                'phone_column': 'phone_number',
                'update_column': 'status'
            }
            
            self.config['EXCEL'] = {
                'phone_column_index': '1',  # 0부터 시작하므로 두 번째 열은 1
                'start_row': '1',
                'has_header': 'True'
            }
            
            # 설정 파일 저장
            with open(self.config_file, 'w') as f:
                self.config.write(f)
    
    def create_widgets(self):
        # 메인 캔버스 생성 (스크롤을 위해)
        main_canvas = tk.Canvas(self.root)
        main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=main_canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 캔버스와 스크롤바 연결
        main_canvas.configure(yscrollcommand=scrollbar.set)
        main_canvas.bind('<Configure>', lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
        
        # 프레임 생성
        main_frame = ttk.Frame(main_canvas, padding="10")
        main_canvas.create_window((0, 0), window=main_frame, anchor="nw", width=680)   
       
        # 노트북 (탭) 생성
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 파일 업로드 탭
        upload_frame = ttk.Frame(notebook, padding="10")
        notebook.add(upload_frame, text="엑셀 업로드 및 업데이트")
        
        # 설정 탭
        settings_frame = ttk.Frame(notebook, padding="10")
        notebook.add(settings_frame, text="설정")
        
        # 파일 업로드 탭 내용
        self.create_upload_widgets(upload_frame)
        
        # 설정 탭 내용
        self.create_settings_widgets(settings_frame)
    
    def create_upload_widgets(self, parent):
        """엑셀 업로드 및 업데이트 탭 위젯 생성"""
        # 엑셀 파일 선택
        file_frame = ttk.LabelFrame(parent, text="엑셀 파일 선택", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50).pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="찾아보기", command=self.browse_file).pack(side=tk.LEFT, padx=5, pady=5)
        
        # 미리보기 영역
        preview_frame = ttk.LabelFrame(parent, text="엑셀 데이터 미리보기", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 트리뷰 생성 (테이블 형태로 데이터 표시)
        self.preview_tree = ttk.Treeview(preview_frame)
        self.preview_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_tree.configure(yscrollcommand=scrollbar.set)
        
        # 수정할 값 입력
        update_frame = ttk.LabelFrame(parent, text="업데이트 정보", padding="10")
        update_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(update_frame, text="업데이트할 값:").pack(side=tk.LEFT, padx=5, pady=5)
        self.update_value_var = tk.StringVar()
        ttk.Entry(update_frame, textvariable=self.update_value_var, width=30).pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        
        # 업데이트 값 설명 추가
        update_desc = ttk.Label(update_frame, text="(이 값이 설정 탭의 '업데이트할 컬럼' 필드에 지정된 컬럼에 저장됩니다)", font=("", 8))
        update_desc.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 업데이트 실행 버튼
        action_frame = ttk.Frame(parent, padding="10")
        action_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(action_frame, text="미리보기 갱신", command=self.refresh_preview).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(action_frame, text="DB 업데이트 실행", command=self.run_update).pack(side=tk.RIGHT, padx=5, pady=5)
        
        # 로그 영역
        log_frame = ttk.LabelFrame(parent, text="실행 로그", padding="10")
        log_frame.pack(fill=tk.BOTH, pady=5)
        
        self.log_text = tk.Text(log_frame, height=6, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
    
    def create_settings_widgets(self, parent):
        """설정 탭 위젯 생성"""
        # 데이터베이스 설정
        db_frame = ttk.LabelFrame(parent, text="데이터베이스 설정", padding="10")
        db_frame.pack(fill=tk.X, pady=5)
        
        # DB 유형 선택
        ttk.Label(db_frame, text="DB 유형:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.db_type_var = tk.StringVar(value=self.config.get('DATABASE', 'type'))
        db_type_combo = ttk.Combobox(db_frame, textvariable=self.db_type_var, values=["sqlite", "mysql", "mariadb"])
        db_type_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        db_type_combo.bind("<<ComboboxSelected>>", self.on_db_type_change)
        
        # 호스트
        ttk.Label(db_frame, text="호스트:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.host_var = tk.StringVar(value=self.config.get('DATABASE', 'host'))
        self.host_entry = ttk.Entry(db_frame, textvariable=self.host_var, width=30)
        self.host_entry.grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=2)
        
        # 포트
        ttk.Label(db_frame, text="포트:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.port_var = tk.StringVar(value=self.config.get('DATABASE', 'port'))
        self.port_entry = ttk.Entry(db_frame, textvariable=self.port_var, width=10)
        self.port_entry.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        # 데이터베이스명
        ttk.Label(db_frame, text="데이터베이스:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        self.database_var = tk.StringVar(value=self.config.get('DATABASE', 'database'))
        ttk.Entry(db_frame, textvariable=self.database_var, width=30).grid(row=3, column=1, sticky=tk.W+tk.E, padx=5, pady=2)
        
        # 사용자명
        ttk.Label(db_frame, text="사용자명:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=2)
        self.user_var = tk.StringVar(value=self.config.get('DATABASE', 'user'))
        self.user_entry = ttk.Entry(db_frame, textvariable=self.user_var, width=20)
        self.user_entry.grid(row=4, column=1, sticky=tk.W, padx=5, pady=2)
        
        # 비밀번호
        ttk.Label(db_frame, text="비밀번호:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=2)
        self.password_var = tk.StringVar(value=self.config.get('DATABASE', 'password'))
        self.password_entry = ttk.Entry(db_frame, textvariable=self.password_var, width=20, show="*")
        self.password_entry.grid(row=5, column=1, sticky=tk.W, padx=5, pady=2)
        
        # 테이블 설정
        table_frame = ttk.LabelFrame(parent, text="테이블 설정", padding="10")
        table_frame.pack(fill=tk.X, pady=5)
        
        # 테이블명
        ttk.Label(table_frame, text="테이블명:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.table_var = tk.StringVar(value=self.config.get('DATABASE', 'table'))
        ttk.Entry(table_frame, textvariable=self.table_var, width=30).grid(row=0, column=1, sticky=tk.W+tk.E, padx=5, pady=2)
        
        # 전화번호 컬럼명
        ttk.Label(table_frame, text="전화번호 컬럼:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.phone_column_var = tk.StringVar(value=self.config.get('DATABASE', 'phone_column'))
        ttk.Entry(table_frame, textvariable=self.phone_column_var, width=30).grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=2)
        
        # 업데이트할 컬럼명
        ttk.Label(table_frame, text="업데이트할 컬럼:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.update_column_var = tk.StringVar(value=self.config.get('DATABASE', 'update_column'))
        ttk.Entry(table_frame, textvariable=self.update_column_var, width=30).grid(row=2, column=1, sticky=tk.W+tk.E, padx=5, pady=2)
        
        # 엑셀 설정
        excel_frame = ttk.LabelFrame(parent, text="엑셀 파일 설정", padding="10")
        excel_frame.pack(fill=tk.X, pady=5)
        
        # 전화번호 열 인덱스
        ttk.Label(excel_frame, text="전화번호 열 인덱스 (0부터 시작):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.phone_col_idx_var = tk.StringVar(value=self.config.get('EXCEL', 'phone_column_index'))
        ttk.Entry(excel_frame, textvariable=self.phone_col_idx_var, width=5).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        # 시작 행
        ttk.Label(excel_frame, text="데이터 시작 행 (0부터 시작):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.start_row_var = tk.StringVar(value=self.config.get('EXCEL', 'start_row'))
        ttk.Entry(excel_frame, textvariable=self.start_row_var, width=5).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # 헤더 여부
        ttk.Label(excel_frame, text="헤더 포함 여부:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.has_header_var = tk.BooleanVar(value=self.config.getboolean('EXCEL', 'has_header'))
        ttk.Checkbutton(excel_frame, variable=self.has_header_var).grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        # 저장 버튼
        ttk.Button(parent, text="설정 저장", command=self.save_config).pack(pady=10)
        
        # SQLite 또는 MySQL에 따라 UI 변경
        self.on_db_type_change(None)
    
    def on_db_type_change(self, event):
        """DB 유형에 따라 UI 요소 활성화/비활성화"""
        db_type = self.db_type_var.get()
        
        if db_type == "sqlite":
            # SQLite는 호스트, 포트, 사용자명, 비밀번호가 필요 없음
            self.host_entry.config(state="disabled")
            self.port_entry.config(state="disabled")
            self.user_entry.config(state="disabled")
            self.password_entry.config(state="disabled")
        else:
            # MySQL/MariaDB는 모든 필드 필요
            self.host_entry.config(state="normal")
            self.port_entry.config(state="normal")
            self.user_entry.config(state="normal")
            self.password_entry.config(state="normal")
    
    def browse_file(self):
        """엑셀 파일 선택 다이얼로그"""
        filetypes = (("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*"))
        filename = filedialog.askopenfilename(title="엑셀 파일 선택", filetypes=filetypes)
        
        if filename:
            self.file_path_var.set(filename)
            self.refresh_preview()
    
    def refresh_preview(self):
        """엑셀 데이터 미리보기 갱신"""
        file_path = self.file_path_var.get()
        
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("오류", "유효한 엑셀 파일을 선택해주세요.")
            return
        
        try:
            # 엑셀 파일 로드
            start_row = int(self.start_row_var.get())
            has_header = self.has_header_var.get()
            
            if has_header:
                df = pd.read_excel(file_path, header=start_row)
            else:
                df = pd.read_excel(file_path, header=None, skiprows=start_row)
            
            # 트리뷰 초기화
            for child in self.preview_tree.get_children():
                self.preview_tree.delete(child)
            
            # 열 헤더 설정
            columns = list(df.columns)
            self.preview_tree["columns"] = columns
            
            # 첫 번째 열 (인덱스) 설정
            self.preview_tree.column("#0", width=50, stretch=tk.NO)
            self.preview_tree.heading("#0", text="No.")
            
            # 나머지 열 설정
            for col in columns:
                self.preview_tree.column(col, width=100, stretch=tk.YES)
                self.preview_tree.heading(col, text=str(col))
            
            # 데이터 추가 (최대 100행까지만 표시)
            for i, row in df.head(100).iterrows():
                values = [row[col] for col in columns]
                self.preview_tree.insert("", "end", text=str(i+1), values=values)
            
            self.log_message(f"{len(df)} 행의 데이터를 로드했습니다.")
            
            # 전화번호 열 인덱스를 기준으로 특정 열의 데이터를 추출
            phone_col_idx = int(self.phone_col_idx_var.get())
            
            if phone_col_idx >= len(df.columns):
                self.log_message(f"오류: 전화번호 열 인덱스({phone_col_idx})가 범위를 벗어났습니다. 열 개수: {len(df.columns)}")
                return
            
            phone_col_name = df.columns[phone_col_idx]
            self.log_message(f"전화번호 열: {phone_col_name}")
            
            # 전화번호 데이터 샘플 표시
            sample_phones = df.iloc[:5, phone_col_idx].tolist()
            self.log_message(f"전화번호 샘플: {sample_phones}")
            
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일 로드 중 오류 발생: {str(e)}")
            self.log_message(f"오류: {str(e)}")
    
    def save_config(self):
        """설정 저장"""
        try:
            # 데이터베이스 설정
            self.config['DATABASE'] = {
                'type': self.db_type_var.get(),
                'host': self.host_var.get(),
                'port': self.port_var.get(),
                'database': self.database_var.get(),
                'user': self.user_var.get(),
                'password': self.password_var.get(),
                'table': self.table_var.get(),
                'phone_column': self.phone_column_var.get(),
                'update_column': self.update_column_var.get()
            }
            
            # 엑셀 설정
            self.config['EXCEL'] = {
                'phone_column_index': self.phone_col_idx_var.get(),
                'start_row': self.start_row_var.get(),
                'has_header': str(self.has_header_var.get())
            }
            
            # 설정 파일 저장
            with open(self.config_file, 'w') as f:
                self.config.write(f)
            
            messagebox.showinfo("알림", "설정이 저장되었습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"설정 저장 중 오류 발생: {str(e)}")
    
    def normalize_phone_number(self, phone):
        """전화번호 형식 정규화"""
        if phone is None:
            return None
            
        # 숫자만 추출
        digits = re.sub(r'\D', '', str(phone))
        
        # 빈 문자열이면 None 반환
        if not digits:
            return None
            
        # 국가 코드 처리 (한국의 경우 +82 또는 82로 시작하는 경우)
        if digits.startswith('82'):
            digits = '0' + digits[2:]
        
        return digits
    
    def get_db_connection(self):
        """데이터베이스 연결 객체 반환"""
        db_type = self.config.get('DATABASE', 'type')
        
        if db_type == 'sqlite':
            # SQLite 연결
            db_path = self.config.get('DATABASE', 'database')
            conn = sqlite3.connect(db_path)
            conn.row_factory = sqlite3.Row
            return conn
        
        elif db_type in ['mysql', 'mariadb']:
            # MySQL/MariaDB 연결
            return mysql.connector.connect(
                host=self.config.get('DATABASE', 'host'),
                port=self.config.getint('DATABASE', 'port'),
                database=self.config.get('DATABASE', 'database'),
                user=self.config.get('DATABASE', 'user'),
                password=self.config.get('DATABASE', 'password'),
                charset='utf8mb4',
                collation='utf8mb4_unicode_ci'
            )
        
        else:
            raise ValueError(f"지원하지 않는 데이터베이스 유형: {db_type}")
    
    def execute_query(self, conn, query, params=None):
        """쿼리 실행"""
        db_type = self.config.get('DATABASE', 'type')
        cursor = conn.cursor()
        
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
                
            if db_type in ['mysql', 'mariadb']:
                result = cursor.fetchall()
            else:  # sqlite
                result = [dict(row) for row in cursor.fetchall()]
                
            return result
        finally:
            cursor.close()
    
    def run_update(self):
        """데이터베이스 업데이트 실행"""
        file_path = self.file_path_var.get()
        update_value = self.update_value_var.get()
        
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("오류", "유효한 엑셀 파일을 선택해주세요.")
            return
        
        if not update_value:
            messagebox.showerror("오류", "업데이트할 값을 입력해주세요.")
            return
        
        try:
            # 엑셀 파일 로드
            start_row = int(self.start_row_var.get())
            has_header = self.has_header_var.get()
            
            if has_header:
                df = pd.read_excel(file_path, header=start_row)
            else:
                df = pd.read_excel(file_path, header=None, skiprows=start_row)
            
            # 전화번호 열 추출
            phone_col_idx = int(self.phone_col_idx_var.get())
            
            if phone_col_idx >= len(df.columns):
                messagebox.showerror("오류", f"전화번호 열 인덱스({phone_col_idx})가 범위를 벗어났습니다.")
                return
            
            # 전화번호 정규화
            phones = df.iloc[:, phone_col_idx].apply(self.normalize_phone_number).dropna().tolist()
            
            if not phones:
                messagebox.showerror("오류", "유효한 전화번호를 찾을 수 없습니다.")
                return
            
            self.log_message(f"{len(phones)}개의 전화번호를 처리합니다.")
            
            # 데이터베이스 연결
            conn = self.get_db_connection()
            
            try:
              table = self.config.get('DATABASE', 'table')
              phone_column = self.config.get('DATABASE', 'phone_column')
              update_column = self.config.get('DATABASE', 'update_column')
              db_type = self.config.get('DATABASE', 'type')
              
              # 로그에 데이터베이스 테이블 구조 출력
              self.log_message(f"테이블: {table}, 전화번호 컬럼: {phone_column}, 업데이트 컬럼: {update_column}")
              
              # 전화번호 샘플 확인
              self.log_message(f"데이터베이스 연결 성공. 전화번호 형식 확인 중...")
              
              try:
                  # 데이터베이스의 전화번호 샘플 확인
                  sample_query = f"SELECT {phone_column} FROM {table} LIMIT 5"
                  cursor = conn.cursor()
                  cursor.execute(sample_query)
                  
                  if db_type in ['mysql', 'mariadb']:
                      db_samples = [row[0] for row in cursor.fetchall()]
                  else:  # sqlite
                      db_samples = [dict(row)[phone_column] for row in cursor.fetchall()]
                  
                  cursor.close()
                  
                  self.log_message(f"DB 전화번호 샘플: {db_samples}")
              except Exception as e:
                  self.log_message(f"샘플 쿼리 실행 중 오류: {str(e)}")
              
              # 플레이스홀더 형식 (SQLite는 ?, MySQL은 %s)
              placeholder = "?" if db_type == 'sqlite' else "%s"
              
              # 전화번호 정규화를 위한 SQL 조건 생성
              if db_type == 'sqlite':
                  # SQLite의 경우, 정규식을 사용할 수 없으므로 OR 조건으로 처리
                  phone_conditions = []
                  
                  for phone in phones:
                      # 여러 형식의 전화번호 매칭을 위한 조건 추가
                      phone_conditions.append(f"{phone_column} = '{phone}'")
                      
                      # 지역번호가 있는 경우 (예: 010-1234-5678 -> 01012345678)
                      if len(phone) >= 10:  # 최소 10자리 번호 (예: 0101234567)
                          # 하이픈 추가된 형식
                          if len(phone) == 10:  # 10자리
                              formatted = f"{phone[:3]}-{phone[3:6]}-{phone[6:]}"
                          elif len(phone) == 11:  # 11자리
                              formatted = f"{phone[:3]}-{phone[3:7]}-{phone[7:]}"
                          else:
                              formatted = phone
                          
                          phone_conditions.append(f"{phone_column} = '{formatted}'")
                          
                          # 하이픈이 있는 경우, 하이픈 없는 형식도 추가
                          if '-' in phone:
                              cleaned = phone.replace('-', '')
                              phone_conditions.append(f"{phone_column} = '{cleaned}'")
                  
                  phone_condition_sql = " OR ".join(phone_conditions)
                  
              else:  # MySQL/MariaDB
                  # LIKE 연산자를 사용하여 전화번호 매칭을 더 유연하게 처리
                  like_conditions = []
                  params = [update_value]  # 첫 번째 파라미터는 업데이트 값
                  
                  for phone in phones:
                      # 전화번호에서 하이픈 제거한 버전
                      cleaned = re.sub(r'[-\s]', '', phone)
                      
                      # 원본 전화번호와 정규화된 전화번호 모두 찾기
                      like_conditions.append(f"{phone_column} = %s")
                      params.append(phone)  # 원본 형식 (하이픈 포함)
                      
                      # 정규화된 번호가 원본과 다르면 이것도 추가
                      if cleaned != phone:
                          like_conditions.append(f"{phone_column} = %s")
                          params.append(cleaned)  # 하이픈 제거 형식
                      
                      # 하이픈 추가 형식도 시도 (10-11자리 번호인 경우)
                      cleaned_len = len(cleaned)
                      if 10 <= cleaned_len <= 11:
                          if cleaned_len == 10:
                              formatted = f"{cleaned[:3]}-{cleaned[3:6]}-{cleaned[6:]}"
                          else:  # 11자리
                              formatted = f"{cleaned[:3]}-{cleaned[3:7]}-{cleaned[7:]}"
                          
                          if formatted != phone:  # 이미 추가한 형식이 아니면
                              like_conditions.append(f"{phone_column} = %s")
                              params.append(formatted)
                  
                  # 모든 조건을 OR로 연결
                  where_clause = " OR ".join(like_conditions)
                  update_sql = f"UPDATE {table} SET {update_column} = %s WHERE {where_clause}"
                  
                  self.log_message(f"수정된 쿼리: {update_sql}")
                  self.log_message(f"파라미터: {params}")
                  
                  cursor = conn.cursor()
                  cursor.execute(update_sql, params)
                  affected_rows = cursor.rowcount
                  cursor.close()
                  
                  conn.commit()
                  
                  self.log_message(f"{affected_rows}개 행이 업데이트되었습니다.")
                  messagebox.showinfo("완료", f"{affected_rows}개 행이 성공적으로 업데이트되었습니다.")
                  
              # SQLite 처리 및 커밋
              if db_type == 'sqlite':
                  update_sql = f"UPDATE {table} SET {update_column} = ? WHERE {phone_condition_sql}"
                  self.log_message(f"실행 쿼리: {update_sql}")
                  
                  cursor = conn.cursor()
                  cursor.execute(f"UPDATE {table} SET {update_column} = ? WHERE {phone_condition_sql}", (update_value,))
                  affected_rows = cursor.rowcount
                  cursor.close()
                  
                  conn.commit()
                  
                  self.log_message(f"{affected_rows}개 행이 업데이트되었습니다.")
                  messagebox.showinfo("완료", f"{affected_rows}개 행이 성공적으로 업데이트되었습니다.")
                  
            finally:
                conn.close()
            
        except Exception as e:
            messagebox.showerror("오류", f"업데이트 과정에서 오류 발생: {str(e)}")
            self.log_message(f"오류: {str(e)}")
    
    def log_message(self, message):
        """로그 메시지 추가"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)  # 스크롤 맨 아래로

def main():
    root = tk.Tk()
    app = DatabaseUpdater(root)
    root.mainloop()

if __name__ == "__main__":
    main()