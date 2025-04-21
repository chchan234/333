import pygetwindow as gw
import pywinctl
import pyautogui
import cv2
import numpy as np
import pandas as pd
import pyperclip
import time
import os
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog
import json
import traceback

# 현재 스크립트 위치 기준으로 상대 경로 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')
EXCEL_DIR = os.path.join(BASE_DIR, 'excel_data')

# 치트 엑셀 파일 경로
CHEAT_FILE = os.path.join(BASE_DIR, 'cheat.xlsx')

class GameCheaterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("게임 치트 자동화 프로그램")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        self.window = None
        self.cheat_data = None
        self.active_windows = []
        self.window_titles = []
        self.threshold = 0.6
        self.current_category = None
        self.cheat_categories = {}  # 치트 카테고리 저장할 딕셔너리
        self.original_cheat_list = []  # 원본 치트 목록 저장 (검색/필터용)
        self.filtered_cheat_list = []  # 필터링된 치트 목록
        
        # 필터 하위 카테고리 옵션 - 드롭다운에 표시될 항목들
        self.filter_categories = ["아스터", "아바타", "아이템", "정령", "탈것", "무기소울"]
        
        # 치트 카테고리 메뉴 옵션 - 항상 이 세 가지만 표시
        self.category_menu_options = ["기타", "필터", "검색"]
        
        self.create_gui()
        self.load_cheat_categories()
        self.get_window_list()
        
    def create_gui(self):
        # 메인 프레임 설정
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 탭 컨트롤 생성
        self.tab_control = ttk.Notebook(self.main_frame)
        
        # 세 개의 탭 생성
        self.window_tab = ttk.Frame(self.tab_control)
        self.cheat_tab = ttk.Frame(self.tab_control)
        self.log_tab = ttk.Frame(self.tab_control)
        
        # 탭 추가
        self.tab_control.add(self.window_tab, text="윈도우 선택")
        self.tab_control.add(self.cheat_tab, text="치트 카테고리")
        self.tab_control.add(self.log_tab, text="로그")
        
        self.tab_control.pack(expand=1, fill=tk.BOTH)
        
        # 각 탭 설정
        self.setup_window_tab()
        self.setup_cheat_tab()
        self.setup_log_tab()
    
    def setup_window_tab(self):
        # 윈도우 선택 탭 설정
        window_frame = ttk.LabelFrame(self.window_tab, text="게임 윈도우 선택", padding="10")
        window_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 윈도우 리스트 설명
        ttk.Label(window_frame, text="아래 목록에서 게임 윈도우를 선택해주세요:").pack(anchor=tk.W, padx=10, pady=10)
        
        # 윈도우 목록 프레임
        list_frame = ttk.Frame(window_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 윈도우 목록 리스트박스
        self.window_listbox = tk.Listbox(list_frame, width=70, height=15)
        self.window_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.window_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.window_listbox.config(yscrollcommand=scrollbar.set)
        
        # 버튼 프레임
        button_frame = ttk.Frame(window_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        refresh_btn = ttk.Button(button_frame, text="새로고침", command=self.get_window_list)
        refresh_btn.pack(side=tk.LEFT, padx=5)
        
        apply_btn = ttk.Button(button_frame, text="선택 적용", command=self.apply_selected_window_and_switch_tab)
        apply_btn.pack(side=tk.RIGHT, padx=5)
        
        # 임계값 설정 프레임
        threshold_frame = ttk.LabelFrame(window_frame, text="이미지 인식 설정", padding="10")
        threshold_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(threshold_frame, text="인식 임계값:").pack(side=tk.LEFT, padx=5)
        self.threshold_var = tk.DoubleVar(value=self.threshold)
        threshold_scale = ttk.Scale(threshold_frame, from_=0.1, to=1.0, orient=tk.HORIZONTAL, 
                                  length=200, variable=self.threshold_var, command=self.update_threshold)
        threshold_scale.pack(side=tk.LEFT, padx=5)
        
        self.threshold_label = ttk.Label(threshold_frame, text=f"{self.threshold:.1f}")
        self.threshold_label.pack(side=tk.LEFT, padx=5)
        
        # 디버그 버튼
        debug_btn = ttk.Button(threshold_frame, text="템플릿 디버그", command=self.debug_templates)
        debug_btn.pack(side=tk.LEFT, padx=20)
        
    def setup_cheat_tab(self):
        # 치트 카테고리 탭 설정
        cheat_frame = ttk.Frame(self.cheat_tab, padding="10")
        cheat_frame.pack(fill=tk.BOTH, expand=True)
        
        # 선택된 윈도우 표시
        self.window_info_label = ttk.Label(cheat_frame, text="선택된 윈도우: 없음", font=("Arial", 10, "bold"))
        self.window_info_label.pack(anchor=tk.W, padx=10, pady=5)
        
        # 카테고리 선택 영역
        category_frame = ttk.LabelFrame(cheat_frame, text="치트 카테고리 선택", padding="10")
        category_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 카테고리 드롭다운
        ttk.Label(category_frame, text="카테고리:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        self.category_var = tk.StringVar()
        # 카테고리 옵션: 기타, 필터, 검색 (self.category_menu_options에서 가져옴)
        self.category_combo = ttk.Combobox(category_frame, textvariable=self.category_var, 
                                      values=self.category_menu_options, width=15, state="readonly")
        self.category_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 카테고리 선택 이벤트 바인딩
        self.category_combo.bind("<<ComboboxSelected>>", self.on_category_selected)
        
        # 하위 카테고리 프레임 (필터, 검색 선택 시 표시됨)
        self.subcategory_frame = ttk.Frame(category_frame)
        self.subcategory_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # 하위 카테고리 및 검색 관련 변수 초기화
        self.subcategory_var = tk.StringVar()
        self.subcategory_combo = None
        self.search_var = tk.StringVar()
        self.search_entry = None
        self.grade_var = tk.StringVar()
        
        # 기본 카테고리 선택
        self.category_combo.current(0)  # "기타" 선택
        
        # 치트 선택/결과 영역 - 다양한 카테고리에 따라 용도가 달라짐
        self.cheat_display_frame = ttk.Frame(cheat_frame)
        self.cheat_display_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 기본 카테고리("기타")의 치트 선택 영역 생성
        self.create_cheat_selection_ui()
        
        # 실행 버튼 영역
        self.button_frame = ttk.Frame(cheat_frame)
        self.button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.execute_btn = ttk.Button(self.button_frame, text="치트 실행", command=self.execute_selected_cheat)
        self.execute_btn.pack(side=tk.RIGHT, padx=5)
    
        # 설명 표시 영역 (카테고리에 따라 표시 여부 결정)
        self.description_frame = ttk.LabelFrame(cheat_frame, text="치트 설명", padding="10")
        self.description_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.description_text = tk.Text(self.description_frame, wrap=tk.WORD, width=70, height=10, 
                                  font=("Courier", 10), state=tk.DISABLED)
        self.description_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
    def create_cheat_selection_ui(self):
        """기본 치트 선택 UI 생성 (드롭다운 방식)"""
        # 기존 위젯 제거
        for widget in self.cheat_display_frame.winfo_children():
            widget.destroy()
            
        # 치트 선택 영역 생성
        cheat_select_frame = ttk.LabelFrame(self.cheat_display_frame, text="치트 선택", padding="10")
        cheat_select_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(cheat_select_frame, text="치트:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 치트 드롭다운
        self.cheat_var = tk.StringVar()
        self.cheat_combo = ttk.Combobox(cheat_select_frame, textvariable=self.cheat_var, 
                                    width=60, state="readonly")
        self.cheat_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 치트 선택 이벤트 바인딩
        self.cheat_combo.bind("<<ComboboxSelected>>", self.on_cheat_selected)
        
        # 파라미터 입력 프레임 (중괄호 포함 치트 코드용)
        self.param_frame = ttk.Frame(cheat_select_frame)
        self.param_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W+tk.E)
        self.param_entries = {}  # 파라미터 입력 필드를 저장할 딕셔너리
        
    def create_results_list_ui(self):
        """결과 목록 UI 생성 (리스트박스 방식)"""
        # 기존 위젯 제거
        for widget in self.cheat_display_frame.winfo_children():
            widget.destroy()
            
        # 결과 목록 영역 생성
        results_frame = ttk.LabelFrame(self.cheat_display_frame, text="검색 결과", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        # 결과 목록 리스트박스
        list_frame = ttk.Frame(results_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.results_listbox = tk.Listbox(list_frame, width=70, height=12, font=("Arial", 10))
        self.results_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.results_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_listbox.config(yscrollcommand=scrollbar.set)
        
        # 결과 선택 이벤트 바인딩
        self.results_listbox.bind('<<ListboxSelect>>', self.on_result_selected)
        
        # 치트 변수 초기화 (내부적으로 사용)
        self.cheat_var = tk.StringVar()
    
    def on_category_selected(self, event):
        """카테고리가 선택되었을 때 호출되는 함수"""
        category = self.category_var.get()
        if not category:
            return
            
        self.log(f"카테고리 선택: {category}")
        
        # 하위 카테고리 프레임 초기화
        for widget in self.subcategory_frame.winfo_children():
            widget.destroy()
        
        # 설명 영역 표시/숨김 처리
        if category in ["필터", "검색"]:
            # 필터, 검색 카테고리에서는 설명 영역 숨김
            self.description_frame.pack_forget()
        else:
            # 기타 카테고리에서는 설명 영역 표시
            self.description_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        if category == "필터":
            # UI를 결과 목록 방식으로 변경
            self.create_results_list_ui()
            
            # 필터 하위 카테고리 프레임 생성
            filter_frame = ttk.Frame(self.subcategory_frame)
            filter_frame.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W+tk.E)
            
            # 필터 하위 카테고리 드롭다운 생성
            ttk.Label(filter_frame, text="필터 항목:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
            
            self.subcategory_var = tk.StringVar()
            self.subcategory_combo = ttk.Combobox(filter_frame, textvariable=self.subcategory_var, 
                                               values=self.filter_categories, width=15, state="readonly")
            self.subcategory_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
            
            # 등급 필터 생성 (동시에 표시)
            ttk.Label(filter_frame, text="등급:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
            
            # 등급 옵션 (한글로 표시, 영어로 필터링)
            grade_options = [
                "전체",
                "일반 (Common)",
                "고급 (Advance)",
                "희귀 (Rare)",
                "에픽 (Epic)",
                "전설 (Legend)",
                "신화 (Myth)"
            ]
            
            self.grade_var = tk.StringVar()
            self.grade_var.set(grade_options[0])  # 기본값 "전체"
            
            grade_combo = ttk.Combobox(filter_frame, textvariable=self.grade_var, 
                                      values=grade_options, width=15, state="readonly")
            grade_combo.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
            
            # 필터 적용 버튼 (둘 다 적용)
            apply_btn = ttk.Button(filter_frame, text="적용", command=self.load_filtered_data)
            apply_btn.grid(row=0, column=4, padx=5, pady=5)
            
            # 첫 번째 항목 선택
            if self.filter_categories:
                self.subcategory_combo.current(0)
            
        elif category == "검색":
            # UI를 결과 목록 방식으로 변경
            self.create_results_list_ui()
            
            # 검색 입력 필드 생성
            ttk.Label(self.subcategory_frame, text="검색어:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
            
            self.search_var = tk.StringVar()
            self.search_entry = ttk.Entry(self.subcategory_frame, textvariable=self.search_var, width=30)
            self.search_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
            
            # 검색 버튼
            search_btn = ttk.Button(self.subcategory_frame, text="검색", command=self.apply_search)
            search_btn.grid(row=0, column=2, padx=5, pady=5)
            
        else:  # 기타 카테고리 선택 시
            # UI를 드롭다운 방식으로 변경
            self.create_cheat_selection_ui()
            self.select_category(category)
            
    def on_result_selected(self, event):
        """결과 목록에서 항목이 선택되었을 때 호출되는 함수"""
        if not hasattr(self, 'results_listbox'):
            return
            
        selection = self.results_listbox.curselection()
        if not selection:
            return
            
        index = selection[0]
        selected_item = self.results_listbox.get(index)
        
        self.log(f"결과 항목 선택됨: '{selected_item}'")
        
        # 치트 변수에 선택된 항목 설정 (내부용)
        self.cheat_var.set(selected_item)
        
        # 선택된 항목에 대한 설명 업데이트
        self.on_cheat_selected(None)
            
    def apply_filter(self):
        """필터 하위 카테고리 선택 적용"""
        selected_filter = self.subcategory_var.get()
        if not selected_filter:
            return
            
        self.log(f"필터 적용: {selected_filter}")
        
        # 선택된 필터 저장
        self.selected_filter_category = selected_filter
        
        # 바로 필터링된 데이터 로드 (필터 항목과 등급 함께 적용)
        self.load_filtered_data()
        
    
    def load_filtered_data(self):
        """선택된 필터와 등급에 따라 데이터 로드"""
        if not hasattr(self, 'subcategory_var') or not self.subcategory_var.get():
            self.log("필터 카테고리가 선택되지 않았습니다.")
            return
            
        category = self.subcategory_var.get()
        self.selected_filter_category = category  # 선택된 필터 저장
        
        grade_text = self.grade_var.get()
        
        # 등급 한글에서 영어로 변환
        grade_map = {
            "전체": None,  # 모든 등급
            "일반 (Common)": "common",
            "고급 (Advance)": "advance",
            "희귀 (Rare)": "rare",
            "에픽 (Epic)": "epic",
            "전설 (Legend)": "legend",
            "신화 (Myth)": "myth"
        }
        
        grade = grade_map.get(grade_text)
        
        # 카테고리에 따른 엑셀 파일 매핑
        excel_map = {
            "아스터": "asters.xlsx",
            "아바타": "avatars.xlsx",
            "아이템": "Items.xlsx",
            "정령": "spirits.xlsx",
            "탈것": "vehicles.xlsx",
            "무기소울": "weapon_souls.xlsx"
        }
        
        excel_file = excel_map.get(category)
        if not excel_file:
            self.log(f"오류: 카테고리 '{category}'에 해당하는 엑셀 파일을 찾을 수 없습니다.")
            return
            
        # 엑셀 파일 로드
        excel_path = os.path.join(EXCEL_DIR, excel_file)
        if not os.path.exists(excel_path):
            self.log(f"오류: 엑셀 파일을 찾을 수 없습니다: {excel_path}")
            return
            
        try:
            self.log(f"'{excel_file}' 파일 로드 중...")
            df = pd.read_excel(excel_path)
            
            # 자동으로 열 이름 탐지 (대소문자 구분 없이)
            column_map = {}
            for col in df.columns:
                if col.lower() == 'name':
                    column_map['name'] = col
                elif col.lower() == 'id':
                    column_map['id'] = col
                elif col.lower() == 'grade':
                    column_map['grade'] = col
            
            # Grade 열이 있는지 확인 (등급 필터링 용)
            if grade and 'grade' not in column_map:
                self.log(f"경고: '{excel_file}'에 'Grade' 열이 없습니다. 등급 필터링을 할 수 없습니다.")
                # 리스트박스 초기화하고 메시지 표시
                if hasattr(self, 'results_listbox'):
                    self.results_listbox.delete(0, tk.END)
                    self.results_listbox.insert(tk.END, "Grade 컬럼이 없어 등급 필터링을 할 수 없습니다")
                return
            
            # 데이터 필터링 (등급이 '전체'가 아닌 경우)
            if grade and 'grade' in column_map:
                grade_column = column_map['grade']
                # 영어 등급명으로 필터링 (소문자로 비교)
                df = df[df[grade_column].str.lower() == grade.lower()]
                self.log(f"등급 '{grade_text}' 기준으로 필터링되었습니다. {len(df)}개 항목 발견.")
            
            # 결과 표시
            if len(df) == 0:
                self.log(f"필터링 결과가 없습니다.")
                # 리스트박스 초기화하고 메시지 표시
                if hasattr(self, 'results_listbox'):
                    self.results_listbox.delete(0, tk.END)
                    self.results_listbox.insert(tk.END, "검색 결과 없음")
                return
                
            # 이제 치트 생성을 위해 name과 id 열이 필요한지 확인
            if 'name' not in column_map or 'id' not in column_map:
                missing_columns = []
                if 'name' not in column_map:
                    missing_columns.append('name')
                if 'id' not in column_map:
                    missing_columns.append('id')
                
                self.log(f"경고: '{excel_file}'에 치트 생성에 필요한 열이 없습니다: {', '.join(missing_columns)}")
                
                # 필터링 결과만 보여주기
                if hasattr(self, 'results_listbox'):
                    self.results_listbox.delete(0, tk.END)
                    
                    # Grade 열만으로 결과 표시 (치트 코드 생성은 안 함)
                    for idx, row in df.iterrows():
                        grade_value = row[column_map['grade']] if 'grade' in column_map else "N/A"
                        item_name = row[column_map['name']] if 'name' in column_map else f"항목 #{idx+1}"
                        
                        # 표시할 텍스트 구성
                        display_text = f"{item_name} (등급: {grade_value})"
                        self.results_listbox.insert(tk.END, display_text)
                    
                    # 첫 번째 항목 선택
                    if self.results_listbox.size() > 0:
                        self.results_listbox.selection_set(0)
                        self.results_listbox.see(0)
                
                self.log(f"'{category}' 카테고리에서 {len(df)}개 항목이 필터링되었습니다 (치트 코드 생성 불가)")
                return
                
            # 데이터를 치트 형식으로 변환
            cheat_list = []
            
            for _, row in df.iterrows():
                name = str(row[column_map['name']])
                id_value = str(row[column_map['id']])
                
                # 필터링 엑셀에는 ID와 이름만 필요함
                # 치트키는 나중에 실행할 때 cheat.xlsx에서 가져올 것임
                # 지금은 임시 치트 코드 형식만 저장 (category 정보)
                cheat_code = f"[{category}] {id_value}"
                
                # 치트 형식으로 추가
                full_cheat = f"{name} — {cheat_code}"
                cheat_list.append(full_cheat)
                    
            # 치트 목록 업데이트
            if cheat_list:
                # 해당 카테고리에 치트 목록 저장
                self.cheat_categories[category] = cheat_list
                
                # 치트 콤보박스 업데이트
                cheat_display_names = []
                self.full_cheat_data = {}
                
                for cheat in cheat_list:
                    if " — " in cheat:
                        display_name = cheat.split(" — ")[0]
                        cheat_display_names.append(display_name)
                        self.full_cheat_data[display_name] = cheat
                    else:
                        cheat_display_names.append(cheat)
                        self.full_cheat_data[cheat] = cheat
                
                # 결과를 리스트박스에 표시
                if hasattr(self, 'results_listbox'):
                    self.results_listbox.delete(0, tk.END)
                    for item in cheat_display_names:
                        self.results_listbox.insert(tk.END, item)
                    
                    # 첫 번째 항목 선택
                    if len(cheat_display_names) > 0:
                        self.results_listbox.selection_set(0)
                        self.results_listbox.see(0)
                        self.on_result_selected(None)
                
                self.log(f"'{category}' 카테고리에 {len(cheat_list)}개 치트 로드됨")
            else:
                # "전체" 등급 선택 시 별도 메시지
                if grade_text == "전체":
                    self.log(f"'{category}' 카테고리에 항목이 없습니다.")
                else:
                    self.log("치트 코드 생성에 필요한 데이터가 부족합니다.")
                
                # 리스트박스 초기화하고 메시지 표시
                if hasattr(self, 'results_listbox'):
                    self.results_listbox.delete(0, tk.END)
                    self.results_listbox.insert(tk.END, "해당 필터 조건에 맞는 항목이 없습니다")
                
        except Exception as e:
            self.log(f"엑셀 파일 처리 중 오류 발생: {e}")
            self.log(traceback.format_exc())
        
    def apply_search(self):
        """검색 적용"""
        search_text = self.search_var.get()
        if not search_text:
            self.log("검색어를 입력해주세요.")
            return
            
        self.log(f"검색 적용: '{search_text}'")
        
        # 모든 카테고리에서 검색
        filtered_cheats = []
        
        # 각 카테고리 순회
        for category_name, cheat_list in self.cheat_categories.items():
            for cheat in cheat_list:
                if search_text.lower() in cheat.lower():
                    # 카테고리 정보와 함께 치트 추가
                    display_name = cheat.split(" — ")[0] if " — " in cheat else cheat
                    filtered_cheats.append(f"[{category_name}] {display_name}")
                    
                    # 치트 데이터 저장
                    if " — " in cheat:
                        self.full_cheat_data[f"[{category_name}] {display_name}"] = cheat
        
        # 검색 결과 표시 (리스트박스 방식)
        if hasattr(self, 'results_listbox'):
            self.results_listbox.delete(0, tk.END)
            
            if filtered_cheats:
                for item in filtered_cheats:
                    self.results_listbox.insert(tk.END, item)
                # 첫 번째 항목 선택
                self.results_listbox.selection_set(0)
                self.results_listbox.see(0)
                self.on_result_selected(None)
                self.log(f"검색 결과: {len(filtered_cheats)}개 항목 발견")
            else:
                self.results_listbox.insert(tk.END, "검색 결과 없음")
                self.log("검색 결과가 없습니다.")
        else:
            # 드롭다운 방식 (예전 방식 호환성 유지)
            if filtered_cheats:
                self.cheat_combo['values'] = filtered_cheats
                self.cheat_combo.current(0)
                self.on_cheat_selected(None)
                self.log(f"검색 결과: {len(filtered_cheats)}개 항목 발견")
            else:
                self.cheat_combo['values'] = ["검색 결과 없음"]
                self.cheat_combo.current(0)
                self.log("검색 결과가 없습니다.")
    
    def select_category(self, category):
        """카테고리 선택 시 해당 카테고리의 치트만 표시"""
        self.current_category = category
        
        # 치트 콤보박스 업데이트 - 코드 부분을 제외한 이름만 표시
        cheat_display_names = []
        self.full_cheat_data = {}  # 치트 이름을 키로, 전체 치트 문자열을 값으로 저장
        
        self.log(f"카테고리 '{category}' 치트 목록 처리 시작")
        
        # 치트 목록 가져오기
        cheat_list = self.cheat_categories.get(category, [])
        self.log(f"카테고리 '{category}'에서 {len(cheat_list)}개 치트 로드됨")
        
        # 원본 치트 목록 저장 (검색/필터용)
        self.original_cheat_list = cheat_list.copy()
        
        for cheat in cheat_list:
            if " — " in cheat:
                display_name = cheat.split(" — ")[0]  # "HP,MP 전체 회복" 부분만 추출
                cheat_display_names.append(display_name)
                self.full_cheat_data[display_name] = cheat
                self.log(f"치트 등록: '{display_name}' -> '{cheat}'")
            else:
                cheat_display_names.append(cheat)
                self.full_cheat_data[cheat] = cheat
                self.log(f"치트 등록(코드 없음): '{cheat}'")
        
        # 필터링된 목록도 초기화
        self.filtered_cheat_list = cheat_display_names.copy()
                
        self.cheat_combo['values'] = cheat_display_names
        if len(cheat_display_names) > 0:
            self.cheat_combo.current(0)  # 첫 번째 항목 선택
            self.on_cheat_selected(None)  # 처음 선택된 치트에 대한 파라미터 필드 생성
        
        # 검색 필드 초기화
        if hasattr(self, 'search_var'):
            self.search_var.set("")
        
        # 설명 텍스트 업데이트
        self.update_description()
        
        self.log(f"카테고리 '{category}' 선택됨, {len(cheat_list)}개 치트 표시")
    
    def update_description(self):
        """현재 선택된 치트에 대한 설명 업데이트"""
        selected_cheat_display = self.cheat_var.get()
        
        # 설명 텍스트 업데이트
        self.description_text.config(state=tk.NORMAL)
        self.description_text.delete(1.0, tk.END)
        
        if selected_cheat_display:
            # 전체 치트 문자열 가져오기 (코드 포함)
            full_cheat = self.full_cheat_data.get(selected_cheat_display, selected_cheat_display)
            
            # 설명 텍스트 생성
            description = f"선택된 치트: {selected_cheat_display}\n\n"
            
            # 코드 부분 추출 (GT.로 시작하는 부분)
            if " — GT." in full_cheat:
                cheat_code = full_cheat.split(" — ")[-1]
                description += f"실행될 코드: {cheat_code}\n"
            
            self.description_text.insert(tk.END, description)
        
        self.description_text.config(state=tk.DISABLED)
    
    def load_cheat_categories(self):
        """치트 카테고리 데이터 로드 - 엑셀에서만 로드"""
        try:
            # 초기화 - 빈 카테고리 딕셔너리
            self.cheat_categories = {}
            self.use_excel_data = False
            
            # 기타 카테고리만 기본으로 추가 (필터, 검색은 드롭다운 메뉴 옵션으로만 존재)
            default_categories = {"기타": []}
            
            # 엑셀 파일 로드
            if os.path.exists(CHEAT_FILE):
                try:
                    # 위치 기반 접근을 위해 헤더 없이 로드
                    self.cheat_data = pd.read_excel(CHEAT_FILE, header=None)
                    self.log(f"치트 데이터 로드 완료: {len(self.cheat_data)} 개 항목")
                    
                    # 엑셀 파일의 내용 일부 출력 (디버깅)
                    for i in range(min(5, len(self.cheat_data))):
                        row_data = [str(x) if not pd.isna(x) else "NaN" for x in self.cheat_data.iloc[i]]
                        self.log(f"행 {i}: {row_data}")
                    
                    # 엑셀 데이터 처리 - 직접 행과 열 지정 (엑셀 파일 구조 기반)
                    current_category = None
                    
                    # 파일 구조 파악 (컬럼 헤더가 있는지)
                    header_row = -1
                    for i in range(min(10, len(self.cheat_data))):
                        row = self.cheat_data.iloc[i]
                        row_data = [str(x).strip() for x in row if not pd.isna(x)]
                        if any(header in row_data for header in ['치트명', '치트키', '이름', '코드']):
                            header_row = i
                            self.log(f"컬럼 헤더 행 발견: {header_row}")
                            break
                    
                    # 기본 카테고리 설정
                    self.cheat_categories = default_categories.copy()
                    
                    # 각 행 처리
                    for i in range(len(self.cheat_data)):
                        row = self.cheat_data.iloc[i]
                        
                        # 빈 행이면 건너뛰기
                        if all(pd.isna(x) for x in row):
                            continue
                            
                        # 헤더 행이면 건너뛰기    
                        if i == header_row:
                            continue
                            
                        # 카테고리 행 확인 (첫 번째 열에 값이 있고 나머지는 대부분 비어있음)
                        if not pd.isna(row[0]) and len(str(row[0]).strip()) > 0:
                            # 다른 열이 대부분 비어있으면 카테고리로 간주
                            non_empty_cells = sum(1 for x in row if not pd.isna(x))
                            if non_empty_cells <= 2:  # 카테고리 이름과 설명 정도만 있을 수 있음
                                current_category = str(row[0]).strip()
                                if current_category not in self.cheat_categories:
                                    self.cheat_categories[current_category] = []
                                self.log(f"카테고리 발견: '{current_category}'")
                                continue
                        
                        # 치트 항목 처리 (두 번째, 세 번째 열에 이름과 코드가 있음)
                        if not pd.isna(row[1]) and not pd.isna(row[2]):
                            # 현재 카테고리가 없으면 "기타" 카테고리에 추가
                            if current_category is None:
                                current_category = "기타"
                                
                            cheat_name = str(row[1]).strip()
                            cheat_code = str(row[2]).strip()
                            
                            # 사용 예시가 있으면 포함
                            example = ""
                            if len(row) > 3 and not pd.isna(row[3]):
                                example = str(row[3]).strip()
                            
                            # 치트 정보 구성
                            full_cheat = f"{cheat_name} — {cheat_code}"
                            if example:
                                full_cheat += f" — {example}"
                                
                            # 치트 데이터 추가
                            self.cheat_categories[current_category].append(full_cheat)
                            self.log(f"치트 추가: '{cheat_name}' -> '{cheat_code}'")
                    
                    # 엑셀 데이터 로드 성공
                    if self.cheat_categories and sum(len(cheats) for cheats in self.cheat_categories.values()) > 0:
                        self.log(f"엑셀에서 {len(self.cheat_categories)} 개의 카테고리와 {sum(len(cheats) for cheats in self.cheat_categories.values())} 개의 치트 로드됨")
                        self.use_excel_data = True
                        
                        # 카테고리 콤보박스 업데이트 - 항상 기타, 필터, 검색만 표시
                        self.category_combo['values'] = self.category_menu_options
                        
                        # 기본 카테고리 선택
                        self.category_combo.set("기타")
                        self.select_category("기타")
                        return
                    else:
                        self.log("엑셀 파일에서 유효한 치트 데이터를 찾을 수 없습니다.")
                        raise ValueError("치트 데이터가 없습니다.")
                        
                except Exception as e:
                    self.log(f"엑셀 파일 처리 중 오류 발생: {e}")
                    import traceback
                    self.log(traceback.format_exc())
                    raise
            else:
                self.log(f"오류: 치트 엑셀 파일을 찾을 수 없습니다: {CHEAT_FILE}")
                self.log("기본 카테고리만 사용합니다.")
                self.cheat_categories = default_categories.copy()
                
                # 치트 예시 추가 (모두 기타 카테고리에 추가)
                # 참고: 엑셀 파일에 정의된 대로 치트키 형식 사용
                self.cheat_categories["기타"].append("기본 아바타 — GT.AVATAR Basic")
                self.cheat_categories["기타"].append("치유 물약 — GT.ITEM Potion")
                self.cheat_categories["기타"].append("체력 회복 — GT.HEAL({HP})")
                
                # 카테고리 콤보박스 업데이트 - 기타, 필터, 검색만 표시
                self.category_combo['values'] = self.category_menu_options
                self.category_combo.set("기타")  # 기본 카테고리 선택
                self.select_category("기타")
                return
                
        except Exception as e:
            self.log(f"치트 데이터 로드 실패: {e}")
            import traceback
            self.log(traceback.format_exc())
            
            # 기본 카테고리 생성 (필터 항목과 동일)
            default_categories = {}
            for category in self.filter_categories:
                default_categories[category] = []
            default_categories["기타"] = []
            
            # 치트 예시 추가 (모두 기타 카테고리에 추가)
            # 참고: 엑셀 파일에 정의된 대로 치트키 형식 사용
            default_categories["기타"].append("기본 아바타 — GT.AVATAR Basic")
            default_categories["기타"].append("치유 물약 — GT.ITEM Potion")
            default_categories["기타"].append("체력 회복 — GT.HEAL({HP})")
            
            self.log("기본 카테고리 생성 중...")
            self.cheat_categories = default_categories
            # 카테고리 콤보박스 업데이트 - 기타, 필터, 검색만 표시
            self.category_combo['values'] = self.category_menu_options
            self.category_combo.set("기타")
            self.select_category("기타")
    
    def execute_selected_cheat(self):
        """선택된 치트 실행 버튼 핸들러"""
        if not self.window:
            self.log("경고: 윈도우를 먼저 선택하고 적용해주세요.")
            return
            
        # 현재 카테고리 확인
        current_category = self.category_var.get()
        
        # 치트 선택 확인 - 카테고리에 따라 다른 방식으로 가져옴
        selected_cheat_display = None
        
        if current_category in ["필터", "검색"]:
            # 리스트박스에서 선택 가져오기
            if hasattr(self, 'results_listbox'):
                selection = self.results_listbox.curselection()
                if selection:
                    selected_cheat_display = self.results_listbox.get(selection[0])
                else:
                    self.log("경고: 실행할 항목을 선택해주세요.")
                    return
        else:
            # 콤보박스에서 선택 가져오기
            selected_cheat_display = self.cheat_var.get()
            
        if not selected_cheat_display:
            self.log("경고: 실행할 치트를 선택해주세요.")
            return
        
        # 전체 치트 문자열 가져오기 (코드 포함)
        full_cheat = self.full_cheat_data.get(selected_cheat_display, selected_cheat_display)
        
        # 먼저 치트 메뉴 열기
        self.log("치트 메뉴 열기 시도 중...")
        if not self.open_cheat_menu():
            self.log("경고: 치트 메뉴를 열지 못했습니다.")
            return
        
        # 필터/검색 결과에서 선택한 경우와 일반 치트 선택의 경우를 분리 처리
        if current_category in ["필터", "검색"] and "[" in full_cheat and "]" in full_cheat:
            # [카테고리] ID 형식에서 카테고리와 ID 추출
            try:
                parts = full_cheat.split(" — ")[1].strip()  # "name — [category] id" 형식에서 "[category] id" 부분 추출
                category_match = parts.split("]")[0] + "]"  # "[category]" 부분 추출
                category = category_match.strip("[]")  # "category" 부분 추출
                id_value = parts.split("]")[1].strip()  # "id" 부분 추출
                
                # cheat.xlsx에서 해당 카테고리의 치트키 형식 찾기
                cheat_code = self.get_cheat_format_from_cheat_xlsx(category, id_value)
                
                if not cheat_code:
                    # 치트키 형식을 찾지 못한 경우, 기본 형식 사용
                    self.log(f"경고: '{category}' 카테고리의 치트키 형식을 찾을 수 없습니다. 기본 형식을 사용합니다.")
                    cheat_code = f"GT.{category.upper()} {id_value}"
            except Exception as e:
                self.log(f"치트 코드 추출 중 오류 발생: {e}")
                # 기본 형식 사용
                cheat_code = full_cheat.split(" — ")[1] if " — " in full_cheat else full_cheat
        else:
            # 일반 치트 선택의 경우 - "— GT." 문자열을 기준으로 코드 추출
            if " — " in full_cheat:
                cheat_code = full_cheat.split(" — ")[1]
            else:
                cheat_code = full_cheat
        
        self.log(f"실행할 치트 코드: {cheat_code}")
        
        # 중괄호가 있는지 확인하고 파라미터 값 적용
        if '{' in cheat_code and '}' in cheat_code and self.param_entries:
            import re
            
            # 각 파라미터에 대해 입력된 값 적용
            for param, entry_var in self.param_entries.items():
                value = entry_var.get()
                if not value:  # 값이 비어있으면 알림
                    self.log(f"경고: '{param}' 값이 입력되지 않았습니다.")
                    if not messagebox.askyesno("파라미터 없음", f"'{param}' 값이 입력되지 않았습니다. 계속 진행하시겠습니까?"):
                        self.log("치트 실행이 취소되었습니다.")
                        return
                
                # 중괄호와 함께 파라미터를 사용자 입력으로 교체
                cheat_code = cheat_code.replace(f"{{{param}}}", value)
                self.log(f"파라미터 '{param}'에 '{value}' 값이 적용되었습니다.")
        
        # 치트 실행
        self.execute_cheat(cheat_code)
        
    def get_cheat_format_from_cheat_xlsx(self, category, id_value):
        """cheat.xlsx에서 해당 카테고리의 치트키 형식을 찾아 반환"""
        try:
            if not os.path.exists(CHEAT_FILE):
                self.log(f"경고: 치트 엑셀 파일을 찾을 수 없습니다: {CHEAT_FILE}")
                return None
                
            # 치트 엑셀 파일 로드
            cheat_df = pd.read_excel(CHEAT_FILE, header=None)
            
            # 카테고리와 치트키 컬럼 찾기
            category_found = False
            for i in range(len(cheat_df)):
                row = cheat_df.iloc[i]
                
                # 빈 행이면 건너뛰기
                if all(pd.isna(x) for x in row):
                    continue
                
                # 카테고리 행 확인
                if not pd.isna(row[0]) and category.lower() in str(row[0]).lower():
                    category_found = True
                    continue
                
                # 해당 카테고리의 첫 번째 치트키 찾기 (카테고리 이후)
                if category_found and not pd.isna(row[2]) and "GT." in str(row[2]):
                    cheat_key = str(row[2]).strip()
                    
                    # 치트키에서 ID 부분 제거하고 기본 형식만 추출
                    if " " in cheat_key:
                        cheat_format = cheat_key.split(" ")[0]  # 공백 앞부분만 가져오기
                        self.log(f"'{category}' 카테고리의 치트키 형식 찾음: {cheat_format}")
                        return f"{cheat_format} {id_value}"
                    else:
                        self.log(f"'{category}' 카테고리의 치트키 형식 찾음: {cheat_key}")
                        return f"{cheat_key} {id_value}"
            
            self.log(f"경고: '{category}' 카테고리의 치트키 형식을 cheat.xlsx에서 찾을 수 없습니다.")
            return None
            
        except Exception as e:
            self.log(f"cheat.xlsx 파일 처리 중 오류 발생: {e}")
            import traceback
            self.log(traceback.format_exc())
            return None
    
    def process_cheat_code_with_params(self, cheat_code):
        """중괄호({})가 포함된 치트 코드에서 사용자 입력을 받아 처리"""
        import re
        
        # 중괄호 내부의 파라미터 추출 (예: {RATE}, {VALUE} 등)
        params = re.findall(r'{([^}]+)}', cheat_code)
        
        if not params:
            return cheat_code
            
        self.log(f"치트 코드에 {len(params)}개의 파라미터가 필요합니다.")
        
        # 각 파라미터에 대해 사용자 입력 받기
        for param in params:
            param_prompt = f"'{param}' 값을 입력하세요:"
            user_input = simpledialog.askstring("파라미터 입력", param_prompt)
            
            if user_input is None:  # 사용자가 취소한 경우
                self.log(f"'{param}' 입력이 취소되었습니다.")
                return None
                
            # 중괄호와 함께 파라미터를 사용자 입력으로 교체
            cheat_code = cheat_code.replace(f"{{{param}}}", user_input)
            self.log(f"파라미터 '{param}'에 '{user_input}' 값이 입력되었습니다.")
        
        return cheat_code
    
    def open_cheat_menu(self):
        """치트 메뉴 열기 버튼 핸들러"""
        if not self.window:
            self.log("경고: 윈도우를 먼저 선택하고 적용해주세요.")
            return False
        
        # 템플릿 매칭으로 메뉴 접근 시도
        self.log("템플릿 매칭으로 메뉴 접근 시도")
        
        # menu2의 존재 여부 확인
        menu2_result = self.find_image_on_screen('menu2.png')
        
        if menu2_result:
            self.log("menu2 발견: 치트 메뉴가 이미 열려있습니다.")
            # menu3 클릭
            menu3_result = self.find_image_on_screen('menu3.png')
            if menu3_result:
                pyautogui.click(menu3_result[0], menu3_result[1])
                self.log("menu3 클릭 완료 (이미지 매칭)")
                time.sleep(0.5)
                return True
            else:
                self.log("menu3를 찾을 수 없습니다.")
                return False
        else:
            # menu 클릭
            menu_result = self.find_image_on_screen('menu.png', report_max_val=True)
            if menu_result:
                pyautogui.click(menu_result[0], menu_result[1])
                self.log("menu 클릭 완료 (이미지 매칭)")
                time.sleep(0.5)
                
                # menu3 클릭
                menu3_result = self.find_image_on_screen('menu3.png')
                if menu3_result:
                    pyautogui.click(menu3_result[0], menu3_result[1])
                    self.log("menu3 클릭 완료 (이미지 매칭)")
                    time.sleep(0.5)
                    return True
                else:
                    self.log("menu3를 찾을 수 없습니다.")
                    return False
            else:
                self.log("menu를 찾을 수 없습니다.")
                return False
    
    def execute_cheat(self, cheat_code):
        """치트 실행 - 코드 입력 및 실행"""
        # 코드를 클립보드에 복사
        pyperclip.copy(cheat_code)
        self.log(f"치트 코드 '{cheat_code}' 복사됨")
        
        # 먼저 code2가 화면에 있는지 확인
        code2_result = self.find_image_on_screen('code2.png', report_max_val=True)
        if code2_result:
            self.log("code2 이미 존재함, code 버튼 클릭 단계 건너뜀")
            # code2 바로 클릭
            pyautogui.click(code2_result[0], code2_result[1])
            self.log(f"code2 클릭 완료 (이미지 매칭 위치: {code2_result[0]}, {code2_result[1]})")
        else:
            # code2가 없으면 일반적인 방법으로 진행
            if not self.click_button('code'):
                self.log("경고: code 버튼을 찾을 수 없습니다.")
                return False
            time.sleep(0.2)  # 대기 시간 변경
            
            if not self.click_button('code2'):
                self.log("경고: code2 버튼을 찾을 수 없습니다.")
                return False
        
        time.sleep(0.2)  # 대기 시간 변경
        
        # 코드 붙여넣기
        pyautogui.hotkey('ctrl', 'v')
        self.log("코드 붙여넣기 완료")
        time.sleep(0.2)  # 대기 시간 변경
        
        self.log("code3 클릭 시도...")
        if not self.click_button('code3'):
            self.log("경고: code3 버튼을 찾을 수 없습니다.")
            return False
        time.sleep(0.2)  # 대기 시간 변경
        
        # code5 먼저 클릭
        self.log("code5 클릭 시도...")
        code5_result = self.find_image_on_screen('code5.png', report_max_val=True)
        if code5_result:
            self.log(f"code5 발견! 위치: ({code5_result[0]}, {code5_result[1]})")
            pyautogui.click(code5_result[0], code5_result[1])
            self.log("code5 클릭 완료")
        else:
            self.log("code5 버튼을 찾을 수 없어 건너뜁니다.")
        
        time.sleep(0.2)  # 대기 시간 변경
        
        # 클릭 전 조금 더 대기
        self.log("code4 클릭 준비 중...")
        time.sleep(0.2)  # 대기 시간 변경
        
        # code4 버튼에 대해서는 이미지 검색과 클릭을 직접 처리
        code4_result = self.find_image_on_screen('code4.png', report_max_val=True)
        if code4_result:
            self.log(f"code4 발견! 위치: ({code4_result[0]}, {code4_result[1]})")
            # 두 번 클릭 시도
            pyautogui.click(code4_result[0], code4_result[1])
            time.sleep(0.2)  # 대기 시간 변경
            pyautogui.click(code4_result[0], code4_result[1])
            self.log("code4 클릭 완료 (두 번 시도)")
        else:
            self.log("경고: code4 버튼을 찾을 수 없습니다.")
            return False
        
        time.sleep(0.2)  # 대기 시간 변경
        self.log("치트 실행 완료")
        self.log("성공: 치트가 성공적으로 실행되었습니다.")
        return True

    def setup_log_tab(self):
        # 로그 탭 설정
        log_frame = ttk.Frame(self.log_tab, padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # 로그 영역
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=30)
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.log_area.config(state=tk.DISABLED)
        
        # 로그 제어 버튼
        button_frame = ttk.Frame(log_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        clear_log_btn = ttk.Button(button_frame, text="로그 지우기", command=self.clear_log)
        clear_log_btn.pack(side=tk.RIGHT, padx=5)
    
    def clear_log(self):
        """로그 영역 지우기"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        self.log("로그가 지워졌습니다.")
    
    def update_threshold(self, value):
        """임계값 업데이트"""
        self.threshold = float(value)
        self.threshold_label.config(text=f"{self.threshold:.1f}")
    
    def debug_templates(self):
        """템플릿 이미지 디버그"""
        if not self.window:
            self.log("경고: 윈도우를 먼저 선택하고 적용해주세요.")
            return
            
        self.log("=== 템플릿 디버그 시작 ===")
        
        # 각 템플릿 파일 테스트
        template_files = ['menu.png', 'menu2.png', 'menu3.png', 
                         'code.png', 'code2.png', 'code3.png', 'code4.png', 'code5.png']
        
        for template_file in template_files:
            # 템플릿 파일 확인
            template_path = os.path.join(TEMPLATES_DIR, template_file)
            if not os.path.exists(template_path):
                self.log(f"템플릿 파일 없음: {template_file}")
                continue
                
            # 이미지 매칭 시도
            result = self.find_image_on_screen(template_file, report_max_val=True)
            if result:
                self.log(f"템플릿 '{template_file}' 매칭 성공: 위치 ({result[0]}, {result[1]}), 정확도: {result[2]:.2f}")
            else:
                self.log(f"템플릿 '{template_file}' 매칭 실패")
        
        self.log("=== 템플릿 디버그 완료 ===")
    
    def log(self, message):
        """로그 영역에 메시지 추가"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
        print(message)  # 콘솔에도 출력
    
    def get_window_list(self):
        """활성화된 윈도우 목록 가져오기"""
        try:
            # 모든 윈도우 목록 가져오기
            all_windows = pywinctl.getAllWindows()
            self.log(f"총 {len(all_windows)}개 윈도우 감지됨")
            
            # 보이는 창만 필터링 (타이틀이 있고 visible 속성이 True인 창)
            visible_windows = []
            for window in all_windows:
                title = window.title if hasattr(window, 'title') else ""
                is_visible = hasattr(window, 'visible') and window.visible
                
                # 타이틀이 있고 보이는 창만 추가
                if title and title.strip() and is_visible:
                    visible_windows.append(window)
            
            self.active_windows = visible_windows
            self.log(f"보이는 창 {len(self.active_windows)}개 필터링됨")
            
            # 리스트박스 업데이트
            self.window_titles = []
            self.window_listbox.delete(0, tk.END)
            
            for window in self.active_windows:
                title = window.title if hasattr(window, 'title') else str(window)
                if title.strip():  # 빈 타이틀은 제외
                    self.window_titles.append(title)
                    self.window_listbox.insert(tk.END, title)
            
            if not self.window_titles:
                self.log("활성화된 윈도우가 없습니다.")
                
        except Exception as e:
            self.log(f"오류: 윈도우 목록 가져오기 실패: {e}")
    
    def apply_selected_window_and_switch_tab(self):
        """윈도우를 선택하고 치트 탭으로 전환"""
        result = self.select_window()
        if result:
            # 선택된 윈도우 정보 업데이트
            self.window_info_label.config(text=f"선택된 윈도우: {self.window_titles[self.window_listbox.curselection()[0]]}")
            
            # 치트 카테고리 탭으로 전환
            self.tab_control.select(1)  # 두 번째 탭(index 1)으로 이동
            
            self.log("성공: 선택한 윈도우가 적용되었습니다.")
        else:
            self.log("경고: 윈도우 적용에 실패했습니다.")
    
    def select_window(self):
        """리스트박스에서 선택된 윈도우 활성화"""
        if not self.active_windows:
            self.log("활성화된 윈도우가 없습니다.")
            return False
        
        try:
            selected_indices = self.window_listbox.curselection()
            if not selected_indices:
                self.log("윈도우를 선택해주세요.")
                return False
                
            selected_index = selected_indices[0]
            if selected_index >= 0 and selected_index < len(self.window_titles):
                selected_title = self.window_titles[selected_index]
                
                # 타이틀로 윈도우 찾기
                for window in self.active_windows:
                    title = window.title if hasattr(window, 'title') else str(window)
                    if title == selected_title:
                        self.window = window
                        
                        # 윈도우 활성화
                        if hasattr(self.window, 'activate'):
                            self.window.activate()
                        else:
                            self.window.focus()  # pywinctl에서는 focus() 메서드를 사용
                        
                        self.log(f"'{selected_title}' 윈도우 선택됨")
                        time.sleep(0.2)  # 안정성을 위한 대기 (시간 변경)
                        return True
                
                self.log(f"선택한 윈도우를 찾을 수 없습니다: {selected_title}")
                return False
            else:
                self.log("윈도우를 선택해주세요.")
                return False
                
        except Exception as e:
            self.log(f"윈도우 선택 실패: {e}")
            return False
    
    def find_image_on_screen(self, template_name, threshold=None, report_max_val=False):
        """화면에서 이미지 찾기"""
        if threshold is None:
            threshold = self.threshold
            
        template_path = os.path.join(TEMPLATES_DIR, template_name)
        
        # 화면 캡처
        screenshot = pyautogui.screenshot()
        screenshot = np.array(screenshot)
        screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)
        
        # 템플릿 이미지 로드
        template = cv2.imread(template_path)
        
        if template is None:
            self.log(f"템플릿 이미지를 찾을 수 없습니다: {template_path}")
            return None
        
        # 이미지 매칭
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        
        # 매칭 결과 로깅
        if report_max_val:
            self.log(f"템플릿 '{template_name}' 매칭 값: {max_val:.2f}, 임계값: {threshold:.2f}")
        
        if max_val >= threshold:
            # 탐지된 위치의 중앙점 계산
            h, w = template.shape[:2]
            center_x = max_loc[0] + w // 2
            center_y = max_loc[1] + h // 2
            return (center_x, center_y, max_val)
        
        return None
    
    def click_button(self, button_name):
        """버튼 클릭 - 이미지 매칭 방법만 사용"""
        # 템플릿 매칭으로 시도
        result = self.find_image_on_screen(f'{button_name}.png', report_max_val=True)
        if result:
            self.log(f"{button_name} 클릭 시도 (위치: {result[0]}, {result[1]})")
            pyautogui.click(result[0], result[1])
            time.sleep(0.2)  # 클릭 후 잠시 대기
            self.log(f"{button_name} 클릭 완료 (이미지 매칭 위치: {result[0]}, {result[1]})")
            return True
        
        self.log(f"{button_name} 버튼을 찾을 수 없습니다.")
        return False
    
    def on_cheat_selected(self, event):
        """치트가 선택되었을 때 호출되는 함수"""
        selected_cheat_display = self.cheat_var.get()
        if not selected_cheat_display:
            return
            
        self.log(f"치트 선택됨: '{selected_cheat_display}'")
        
        # 검색 결과에서 카테고리 태그가 포함된 경우 처리 ([카테고리] 치트명)
        if selected_cheat_display.startswith("[") and "] " in selected_cheat_display:
            parts = selected_cheat_display.split("] ", 1)
            if len(parts) == 2:
                category = parts[0][1:]  # '[아바타]' -> '아바타'
                cheat_name = parts[1]
                self.log(f"검색 결과에서 선택: 카테고리 '{category}', 치트 '{cheat_name}'")
                
                # 원래 치트 데이터 찾기
                found = False
                for cheat in self.cheat_categories.get(category, []):
                    if cheat.startswith(cheat_name + " —") or cheat == cheat_name:
                        self.full_cheat_data[cheat_name] = cheat
                        found = True
                        break
                
                if not found:
                    self.log(f"경고: 원본 치트 데이터를 찾을 수 없습니다: {cheat_name}")
            
        # 설명 업데이트
        self.update_description()
        
        # 파라미터 입력 필드 업데이트
        self.update_parameter_fields()
        
    def on_search_filter_change(self, *args):
        """검색 또는 필터 변경 시 치트 목록 업데이트"""
        if not hasattr(self, 'search_var') or not hasattr(self, 'filter_var'):
            return
            
        search_text = self.search_var.get().lower()
        filter_option = self.filter_var.get()
        
        # 현재 선택된 카테고리가 없으면 종료
        if not self.current_category:
            return
            
        self.log(f"검색 필터링: '{search_text}', 옵션: '{filter_option}'")
        
        # 원본 치트 목록에서 시작
        filtered_cheats = []
        
        for cheat in self.original_cheat_list:
            if " — " in cheat:
                # 이름과 코드로 분리
                parts = cheat.split(" — ")
                name = parts[0].lower()
                code = parts[1].lower() if len(parts) > 1 else ""
                
                # 필터 옵션에 따라 검색
                if filter_option == "전체":
                    if search_text in name or search_text in code:
                        display_name = parts[0]
                        filtered_cheats.append(display_name)
                elif filter_option == "이름만":
                    if search_text in name:
                        display_name = parts[0]
                        filtered_cheats.append(display_name)
                elif filter_option == "코드만":
                    if search_text in code:
                        display_name = parts[0]
                        filtered_cheats.append(display_name)
            else:
                # 코드가 없는 경우
                if search_text in cheat.lower():
                    filtered_cheats.append(cheat)
        
        # 콤보박스 업데이트
        self.filtered_cheat_list = filtered_cheats
        self.cheat_combo['values'] = filtered_cheats
        
        # 결과가 있으면 첫 번째 항목 선택
        if filtered_cheats:
            self.cheat_combo.current(0)
            self.on_cheat_selected(None)
        else:
            self.cheat_var.set("")
            # 파라미터 입력 필드 초기화
            for widget in self.param_frame.winfo_children():
                widget.destroy()
            self.param_entries.clear()
            # 설명 업데이트
            self.update_description()
            
        self.log(f"검색 결과: {len(filtered_cheats)}개 치트 표시")
        
    def clear_search_filter(self):
        """검색 및 필터 초기화"""
        if hasattr(self, 'search_var') and hasattr(self, 'filter_var'):
            self.search_var.set("")
            self.filter_var.set(self.filter_options[0])  # 기본값으로 설정
            
            # 현재 선택된 카테고리의 모든 치트 표시
            if self.current_category:
                self.select_category(self.current_category)
    
    def update_parameter_fields(self):
        """선택된 치트에 필요한 파라미터 입력 필드 생성"""
        # 기존 파라미터 입력 필드 삭제
        for widget in self.param_frame.winfo_children():
            widget.destroy()
        self.param_entries.clear()
        
        selected_cheat_display = self.cheat_var.get()
        if not selected_cheat_display:
            return
            
        # 전체 치트 문자열 가져오기 (코드 포함)
        full_cheat = self.full_cheat_data.get(selected_cheat_display, selected_cheat_display)
        self.log(f"전체 치트 문자열: '{full_cheat}'")
        
        # 중괄호 안의 파라미터 추출
        import re
        params = []
        param_options = {}  # 파라미터별 옵션을 저장할 딕셔너리
        
        # 치트 코드 부분 (GT.로 시작하는 부분) 추출
        if " — GT." in full_cheat:
            # GT. 뒤의 모든 코드 부분 추출
            cheat_code_parts = full_cheat.split(" — GT.")
            for part in cheat_code_parts[1:]:  # 첫 번째는 이름이므로 건너뛰기
                # 각 코드 부분에서 중괄호 파라미터 검색
                param_matches = re.finditer(r'{([^}]+)}', part)
                for match in param_matches:
                    param_text = match.group(1)
                    
                    # 파이프(|)가 있는지 확인 (예: ON|OFF)
                    if '|' in param_text:
                        param_name = param_text.split('|')[0].split(':')[0].strip()
                        options = [opt.strip() for opt in param_text.split('|')]
                        # 파라미터 이름이 옵션에 포함되어 있으면 제거
                        if ':' in param_name:
                            param_name = param_name.split(':')[0].strip()
                            options[0] = options[0].split(':')[1].strip()
                        
                        params.append(param_name)
                        param_options[param_name] = options
                    else:
                        params.append(param_text)
        
        self.log(f"찾은 파라미터: {params}")
        
        if not params:
            self.log("중괄호 파라미터가 발견되지 않았습니다.")
            return
            
        # 중복 제거
        unique_params = []
        for p in params:
            if p not in unique_params:
                unique_params.append(p)
        params = unique_params
        self.log(f"중복 제거 후 파라미터: {params}")
            
        # 파라미터 라벨 추가
        ttk.Label(self.param_frame, text="파라미터:", font=("Arial", 9, "bold")).grid(
            row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 각 파라미터에 대한 입력 필드 생성
        for i, param in enumerate(params):
            ttk.Label(self.param_frame, text=f"{param}:").grid(
                row=i+1, column=0, padx=5, pady=2, sticky=tk.W)
            
            # 토글/선택 옵션이 있는 경우 콤보박스 사용
            if param in param_options:
                self.log(f"파라미터 '{param}'에 옵션 선택 필드 생성: {param_options[param]}")
                combo_var = tk.StringVar()
                combo = ttk.Combobox(self.param_frame, textvariable=combo_var, 
                                  width=15, state="readonly")
                combo['values'] = param_options[param]
                combo.current(0)  # 첫 번째 옵션 선택
                combo.grid(row=i+1, column=1, padx=5, pady=2, sticky=tk.W)
                self.param_entries[param] = combo_var
            else:
                # 일반 입력 필드
                entry_var = tk.StringVar()
                entry = ttk.Entry(self.param_frame, textvariable=entry_var, width=30)
                entry.grid(row=i+1, column=1, padx=5, pady=2, sticky=tk.W)
                self.param_entries[param] = entry_var
            
        self.log(f"{len(params)}개의 파라미터 입력 필드가 생성되었습니다.")

# 메인 실행
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = GameCheaterGUI(root)
        root.mainloop()
    except KeyboardInterrupt:
        print("\n프로그램이 종료되었습니다.")
    except Exception as e:
        print(f"오류 발생: {e}")
        messagebox.showerror("치명적 오류", f"프로그램 실행 중 오류 발생: {e}")
