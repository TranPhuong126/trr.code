import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Canvas
import pandas as pd
import numpy as np
from collections import defaultdict
import heapq
import os
from datetime import datetime

# Cố gắng import để vẽ đồ thị
try:
    import networkx as nx
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    HAS_GRAPH = True
except:
    HAS_GRAPH = False


class ExamSchedulerPro:
    def __init__(self, root):
        self.root = root
        self.root.title("🎓 Xếp Lịch Thi Thông Minh - DSatur Pro v3.0")
        self.root.geometry("1600x950")
        self.root.configure(bg='#f8f9fa')
        self.root.state('zoomed')

        # Dữ liệu
        self.data = None
        self.subjects = []
        self.student_subjects = defaultdict(set)
        self.subject_students = defaultdict(set)
        self.conflict_graph = defaultdict(set)
        self.schedule = {}

        # Màu sắc hiện đại & dễ thương
        self.colors = {
            'bg': '#f8f9fa',
            'card': '#ffffff',
            'primary': '#7c3aed',      # Tím đậm
            'primary_light': '#a78bfa', # Tím nhạt
            'success': '#10b981',       # Xanh lá
            'success_light': '#6ee7b7', # Xanh lá nhạt
            'warning': '#f59e0b',       # Cam
            'danger': '#ef4444',        # Đỏ
            'info': '#3b82f6',          # Xanh dương
            'pink': '#ec4899',          # Hồng
            'dark': '#1f2937',
            'light': '#f3f4f6',
            'border': '#e5e7eb'
        }

        self.setup_styles()
        self.create_ui()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Button style
        style.configure("TButton", 
            padding=12, 
            font=('Segoe UI', 10, 'bold'),
            borderwidth=0,
            relief='flat'
        )
        
        # Treeview style
        style.configure("Treeview",
            background="white",
            fieldbackground="white",
            rowheight=35,
            borderwidth=0,
            font=('Segoe UI', 10)
        )
        style.configure("Treeview.Heading",
            font=('Segoe UI', 11, 'bold'),
            background=self.colors['primary_light'],
            foreground='white',
            borderwidth=0
        )
        style.map('Treeview',
            background=[('selected', self.colors['primary'])],
            foreground=[('selected', 'white')]
        )

    def create_ui(self):
        # Header với gradient effect
        header = tk.Frame(self.root, bg=self.colors['primary'], height=90)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        header_content = tk.Frame(header, bg=self.colors['primary'])
        header_content.pack(expand=True)
        
        tk.Label(header_content, 
                text="🎓 XẾP LỊCH THI THÔNG MINH",
                font=('Segoe UI', 32, 'bold'),
                fg='white', 
                bg=self.colors['primary']).pack()
        tk.Label(header_content,
                text="DSatur Algorithm - Tối ưu số ca thi",
                font=('Segoe UI', 12),
                fg=self.colors['primary_light'],
                bg=self.colors['primary']).pack()

        main_container = tk.Frame(self.root, bg=self.colors['bg'])
        main_container.pack(fill='both', expand=True, padx=20, pady=20)

        main = tk.PanedWindow(main_container, orient=tk.HORIZONTAL, 
                             sashrelief=tk.FLAT, bg=self.colors['bg'],
                             sashwidth=10)
        main.pack(fill='both', expand=True)

        # === SIDEBAR TRÁI ===
        left = tk.Frame(main, bg=self.colors['card'], width=400, relief='flat', bd=0)
        left.pack_propagate(False)
        main.add(left)

        # Thêm shadow effect bằng frame
        self._add_shadow(left)

        # Upload section
        upload_frame = self._create_card_frame(left, "📁 NHẬP DỮ LIỆU")
        upload_frame.pack(fill='x', padx=20, pady=(20, 10))
        
        upload_btn = tk.Button(upload_frame, 
                              text="📂 CHỌN FILE EXCEL",
                              command=self.load_file,
                              bg=self.colors['primary'],
                              fg='white',
                              font=('Segoe UI', 11, 'bold'),
                              relief='flat',
                              padx=25,
                              pady=12,
                              cursor='hand2',
                              activebackground=self.colors['primary_light'])
        upload_btn.pack(pady=15)
        self._add_hover_effect(upload_btn, self.colors['primary'], self.colors['primary_light'])
        
        self.file_label = tk.Label(upload_frame,
                                   text="Chưa chọn file...",
                                   bg=self.colors['card'],
                                   fg='#9ca3af',
                                   font=('Segoe UI', 10),
                                   wraplength=350)
        self.file_label.pack(pady=(0, 15))

        # Run button - to hơn
        run_frame = tk.Frame(left, bg=self.colors['card'])
        run_frame.pack(pady=30, padx=30, fill='x')
        
        run_btn = tk.Button(run_frame,
                           text="🚀 CHẠY DSATUR",
                           command=self.run_dsatur,
                           bg=self.colors['success'],
                           fg='white',
                           font=('Segoe UI', 16, 'bold'),
                           relief='flat',
                           padx=30,
                           pady=18,
                           cursor='hand2',
                           activebackground=self.colors['success_light'])
        run_btn.pack(fill='x')
        self._add_hover_effect(run_btn, self.colors['success'], self.colors['success_light'])

        # Thống kê
        stats_frame = self._create_card_frame(left, "📊 THỐNG KÊ")
        stats_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        self.stats_text = tk.Text(stats_frame,
                                 height=12,
                                 bg=self.colors['light'],
                                 relief='flat',
                                 font=('Consolas', 10),
                                 borderwidth=0,
                                 padx=15,
                                 pady=15)
        self.stats_text.pack(fill='both', padx=15, pady=15)

        # === PHẦN PHẢI ===
        right = tk.Frame(main, bg=self.colors['card'], relief='flat', bd=0)
        main.add(right)
        self._add_shadow(right)

        notebook = ttk.Notebook(right)
        notebook.pack(fill='both', expand=True, padx=20, pady=20)

        # Tab 1: Lịch thi
        tab1 = self._create_tab(notebook, '📅 Lịch Thi Theo Ca')
        
        tree_frame = tk.Frame(tab1, bg='white')
        tree_frame.pack(fill='both', expand=True, padx=15, pady=15)
        
        self.tree_schedule = ttk.Treeview(tree_frame,
                                         columns=('Ca', 'Môn', 'SV'),
                                         show='headings',
                                         height=25)
        self.tree_schedule.heading('Ca', text='🎯 Ca Thi')
        self.tree_schedule.heading('Môn', text='📚 Môn Học')
        self.tree_schedule.heading('SV', text='👥 Số SV')
        self.tree_schedule.column('Ca', width=100, anchor='center')
        self.tree_schedule.column('Môn', width=550)
        self.tree_schedule.column('SV', width=120, anchor='center')
        
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', 
                                 command=self.tree_schedule.yview)
        self.tree_schedule.configure(yscrollcommand=scrollbar.set)
        
        self.tree_schedule.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Tab 2: Lịch SV
        tab2 = self._create_tab(notebook, '👨‍🎓 Lịch Sinh Viên')
        
        search_frame = tk.Frame(tab2, bg='white')
        search_frame.pack(fill='x', padx=15, pady=15)
        
        tk.Label(search_frame,
                text="🔍 Tìm kiếm:",
                bg='white',
                font=('Segoe UI', 11, 'bold')).pack(side='left', padx=(0, 10))
        
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame,
                               textvariable=self.search_var,
                               width=50,
                               font=('Segoe UI', 11),
                               relief='solid',
                               borderwidth=1)
        search_entry.pack(side='left', padx=5, ipady=8)
        self.search_var.trace('w', self.filter_students)
        
        tree_frame2 = tk.Frame(tab2, bg='white')
        tree_frame2.pack(fill='both', expand=True, padx=15, pady=(0, 15))
        
        self.tree_student = ttk.Treeview(tree_frame2,
                                        columns=('MSSV', 'Tên', 'Ca', 'Môn'),
                                        show='headings')
        self.tree_student.heading('MSSV', text='🆔 MSSV')
        self.tree_student.heading('Tên', text='👤 Họ Tên')
        self.tree_student.heading('Ca', text='🎯 Ca Thi')
        self.tree_student.heading('Môn', text='📚 Môn Học')
        self.tree_student.column('MSSV', width=120, anchor='center')
        self.tree_student.column('Tên', width=250)
        self.tree_student.column('Ca', width=100, anchor='center')
        self.tree_student.column('Môn', width=300)
        
        scrollbar2 = ttk.Scrollbar(tree_frame2, orient='vertical',
                                  command=self.tree_student.yview)
        self.tree_student.configure(yscrollcommand=scrollbar2.set)
        
        self.tree_student.pack(side='left', fill='both', expand=True)
        scrollbar2.pack(side='right', fill='y')

        # Tab 3: Đồ thị
        tab3 = self._create_tab(notebook, '🎨 Đồ Thị Xung Đột')
        self.graph_canvas = tk.Canvas(tab3, bg='white')
        self.graph_canvas.pack(fill='both', expand=True, padx=20, pady=20)
        if not HAS_GRAPH:
            tk.Label(tab3,
                    text="⚠️ Cài đặt networkx + matplotlib để xem đồ thị!",
                    fg=self.colors['danger'],
                    font=('Segoe UI', 12, 'bold')).pack(pady=50)

        # Tab 4: Export & Kiểm tra
        tab4 = self._create_tab(notebook, '💾 Export & Kiểm Tra')
        
        export_btn = tk.Button(tab4,
                              text="💾 XUẤT LỊCH THI EXCEL",
                              command=self.export_all,
                              bg=self.colors['warning'],
                              fg='white',
                              font=('Segoe UI', 13, 'bold'),
                              relief='flat',
                              pady=15,
                              cursor='hand2',
                              activebackground='#fbbf24')
        export_btn.pack(pady=30, padx=30, fill='x')
        self._add_hover_effect(export_btn, self.colors['warning'], '#fbbf24')
        
        warning_frame = tk.Frame(tab4, bg='#fef2f2', relief='flat', bd=1)
        warning_frame.pack(fill='both', expand=True, padx=30, pady=(0, 30))
        
        self.warning_text = tk.Text(warning_frame,
                                   height=20,
                                   bg='#fef2f2',
                                   fg=self.colors['danger'],
                                   font=('Segoe UI', 11),
                                   relief='flat',
                                   borderwidth=0,
                                   padx=20,
                                   pady=20)
        self.warning_text.pack(fill='both', expand=True)

    def _create_card_frame(self, parent, title):
        """Tạo frame card đẹp với title"""
        frame = tk.LabelFrame(parent,
                             text=f"  {title}  ",
                             bg=self.colors['card'],
                             fg=self.colors['dark'],
                             font=('Segoe UI', 11, 'bold'),
                             relief='flat',
                             borderwidth=1,
                             highlightthickness=0)
        return frame

    def _create_tab(self, notebook, title):
        """Tạo tab với background trắng"""
        tab = tk.Frame(notebook, bg='white')
        notebook.add(tab, text=title)
        return tab

    def _add_shadow(self, widget):
        """Thêm hiệu ứng shadow cho widget"""
        widget.configure(relief='flat', borderwidth=1, 
                        highlightbackground=self.colors['border'],
                        highlightthickness=1)

    def _add_hover_effect(self, button, normal_color, hover_color):
        """Thêm hiệu ứng hover cho button"""
        def on_enter(e):
            button['background'] = hover_color
        def on_leave(e):
            button['background'] = normal_color
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return

        try:
            all_dfs = []
            excel = pd.ExcelFile(path, engine='openpyxl')
            
            for sheet in excel.sheet_names:
                try:
                    df = pd.read_excel(excel, sheet_name=sheet, header=None, dtype=str, engine='openpyxl')
                    df = df.fillna('')
                    
                    header_row = None
                    for idx in range(min(5, len(df))):
                        row_text = ' '.join(df.iloc[idx].astype(str).str.lower().tolist())
                        if 'mã sv' in row_text or 'mssv' in row_text or 'ma sv' in row_text:
                            header_row = idx
                            break
                    
                    if header_row is None:
                        continue
                    
                    subject_name = sheet
                    if header_row > 0:
                        first_cell = str(df.iloc[0, 0]).strip()
                        if first_cell and len(first_cell) > 0:
                            subject_name = first_cell
                    
                    df.columns = df.iloc[header_row]
                    df = df.iloc[header_row + 1:].reset_index(drop=True)
                    
                    masv_col = None
                    hoten_col = None
                    
                    for col in df.columns:
                        col_str = str(col).lower().strip()
                        if 'mã sv' in col_str or 'mssv' in col_str or 'ma sv' in col_str:
                            masv_col = col
                        if 'họ' in col_str and 'tên' in col_str:
                            hoten_col = col
                        elif 'tên' in col_str and hoten_col is None:
                            hoten_col = col
                    
                    if masv_col is None:
                        continue
                    
                    if hoten_col:
                        df_clean = df[[masv_col, hoten_col]].copy()
                        df_clean.columns = ['MaSV', 'HoTen']
                    else:
                        df_clean = df[[masv_col]].copy()
                        df_clean.columns = ['MaSV']
                        df_clean['HoTen'] = 'N/A'
                    
                    df_clean['MaSV'] = df_clean['MaSV'].astype(str).str.strip()
                    df_clean = df_clean.loc[df_clean['MaSV'].str.len() > 0].copy()
                    
                    mask_numeric = df_clean['MaSV'].str.match(r'^\d+$', na=False)
                    df_clean = df_clean.loc[mask_numeric].copy()
                    
                    if len(df_clean) > 0:
                        df_clean['ChuongTrinh'] = subject_name
                        all_dfs.append(df_clean)
                        print(f"✓ Đọc thành công sheet '{sheet}': {len(df_clean)} sinh viên")
                
                except Exception as e:
                    print(f"Lỗi đọc sheet {sheet}: {str(e)}")
                    continue
            
            if not all_dfs:
                messagebox.showerror("Lỗi", "Không tìm thấy dữ liệu hợp lệ!")
                return
            
            self.data = pd.concat(all_dfs, ignore_index=True)
            self.data.drop_duplicates(subset=['MaSV', 'ChuongTrinh'], inplace=True)
            
            self.file_label.config(
                text=f"✅ ĐÃ TẢI: {os.path.basename(path)}\n📊 {len(self.data)} dòng • 📚 {self.data['ChuongTrinh'].nunique()} môn",
                fg=self.colors['success'],
                font=('Segoe UI', 10, 'bold')
            )
            
            messagebox.showinfo("🎉 Thành công!", 
                f"Đã tải thành công!\n\n"
                f"📄 {len(self.data):,} bản ghi\n"
                f"📑 {len(excel.sheet_names)} sheet\n"
                f"👥 {self.data['MaSV'].nunique()} sinh viên\n"
                f"📚 {self.data['ChuongTrinh'].nunique()} môn học")
            
            self.process_data()

        except Exception as e:
            messagebox.showerror("Lỗi đọc file", f"Chi tiết lỗi:\n{str(e)}")

    def process_data(self):
        """Xử lý dữ liệu và xây dựng đồ thị xung đột"""
        self.subjects = self.data['ChuongTrinh'].unique().tolist()
        self.student_subjects.clear()
        self.subject_students.clear()
        self.conflict_graph.clear()

        for _, row in self.data.iterrows():
            sid = str(row['MaSV']).strip()
            subj = row['ChuongTrinh']
            self.student_subjects[sid].add(subj)
            self.subject_students[subj].add(sid)

        # Xây đồ thị xung đột - 2 môn xung đột nếu có sinh viên chung
        for subs in self.student_subjects.values():
            subs = list(subs)
            for i in range(len(subs)):
                for j in range(i+1, len(subs)):
                    self.conflict_graph[subs[i]].add(subs[j])
                    self.conflict_graph[subs[j]].add(subs[i])

        self.update_stats()

    def update_stats(self):
        """Cập nhật thống kê"""
        text = "📊 TỔNG QUAN DỮ LIỆU\n"
        text += "═" * 40 + "\n\n"
        text += f"👥 Sinh viên:      {len(self.student_subjects):,}\n"
        text += f"📚 Môn học:        {len(self.subjects):,}\n"
        text += f"⚠️ Xung đột:      {sum(len(v) for v in self.conflict_graph.values())//2:,} cạnh\n"
        
        if self.schedule:
            text += f"\n🎯 Số ca thi:      {max(self.schedule.values())}\n"
            text += f"✅ Đã xếp lịch:    {len(self.schedule)} môn"
        
        self.stats_text.delete(1.0, 'end')
        self.stats_text.insert('end', text)
        self.stats_text.tag_configure('bold', font=('Consolas', 10, 'bold'))

    def run_dsatur(self):
        """Chạy thuật toán DSatur để tô màu đồ thị - Đảm bảo không trùng ca"""
        if self.data is None or self.data.empty or len(self.subjects) == 0:
            messagebox.showwarning("⚠️ Cảnh báo", "Chưa tải dữ liệu!")
            return

        self.schedule.clear()

        # DSatur Algorithm - Tối ưu bằng heap
        degree = {s: len(self.conflict_graph[s]) for s in self.subjects}
        saturation = {s: 0 for s in self.subjects}
        color_of = {}
        
        # Heap: (-saturation, -degree, subject)
        heap = [(-0, -degree[s], s) for s in self.subjects]
        heapq.heapify(heap)
        colored = set()

        for _ in range(len(self.subjects)):
            # Lấy node có saturation cao nhất (nếu bằng nhau thì lấy degree cao nhất)
            while heap:
                _, _, subj = heapq.heappop(heap)
                if subj not in colored:
                    break
            
            # Tìm màu nhỏ nhất chưa được dùng bởi các hàng xóm
            used_colors = {color_of.get(neighbor) 
                          for neighbor in self.conflict_graph[subj] 
                          if neighbor in color_of}
            
            color = 1
            while color in used_colors:
                color += 1
            
            color_of[subj] = color
            colored.add(subj)

            # Cập nhật saturation của các hàng xóm chưa được tô
            for neighbor in self.conflict_graph[subj]:
                if neighbor not in colored:
                    saturation[neighbor] += 1
                    heapq.heappush(heap, (-saturation[neighbor], -degree[neighbor], neighbor))

        self.schedule = color_of
        
        # Kiểm tra xung đột
        conflict_found = self.check_conflicts()
        
        self.display_results()
        self.draw_graph()
        self.update_stats()
        
        if not conflict_found:
            messagebox.showinfo("🎉 HOÀN THÀNH!", 
                              f"✅ Đã xếp lịch thành công!\n\n"
                              f"🎯 Tổng số ca thi: {max(color_of.values())}\n"
                              f"📚 Tổng số môn: {len(color_of)}\n"
                              f"✨ Không có xung đột!")
        else:
            messagebox.showwarning("⚠️ Cảnh báo",
                                 "Có xung đột! Kiểm tra tab 'Export & Kiểm Tra'")

    def display_results(self):
        """Hiển thị kết quả lên giao diện"""
        # Xóa dữ liệu cũ
        for i in self.tree_schedule.get_children():
            self.tree_schedule.delete(i)
        for i in self.tree_student.get_children():
            self.tree_student.delete(i)

        # Hiển thị lịch thi theo ca - ĐÃ SẮP XẾP
        ca_dict = defaultdict(list)
        for subj, ca in self.schedule.items():
            ca_dict[ca].append((subj, len(self.subject_students[subj])))
        
        # Sắp xếp theo số ca
        for ca in sorted(ca_dict.keys()):
            for subj, count in sorted(ca_dict[ca], key=lambda x: -x[1]):
                self.tree_schedule.insert('', 'end', 
                                         values=(f'Ca {ca}', subj, count),
                                         tags=('evenrow' if ca % 2 == 0 else 'oddrow',))
        
        # Thêm màu cho các hàng
        self.tree_schedule.tag_configure('evenrow', background='#f9fafb')
        self.tree_schedule.tag_configure('oddrow', background='white')

        # Hiển thị lịch sinh viên
        for sid, subs in self.student_subjects.items():
            name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
            name = name_df.iloc[0] if len(name_df) > 0 else "N/A"
            
            for sub in sorted(subs):
                ca = self.schedule.get(sub, "?")
                self.tree_student.insert('', 'end', 
                                        values=(sid, name, f'Ca {ca}', sub))

    def filter_students(self, *args):
        """Lọc sinh viên theo từ khóa"""
        search = self.search_var.get().lower()
        for i in self.tree_student.get_children():
            self.tree_student.delete(i)
        
        for sid, subs in self.student_subjects.items():
            name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
            name = name_df.iloc[0] if len(name_df) > 0 else "N/A"
            
            if search in sid.lower() or search in name.lower():
                for sub in sorted(subs):
                    ca = self.schedule.get(sub, "?")
                    self.tree_student.insert('', 'end', 
                                            values=(sid, name, f'Ca {ca}', sub))

    def check_conflicts(self):
        """Kiểm tra xung đột - sinh viên có bị trùng ca không"""
        self.warning_text.delete(1.0, 'end')
        conflicts = []
        
        for sid, subs in self.student_subjects.items():
            cas = [self.schedule.get(s) for s in subs]
            # Kiểm tra trùng ca
            if len(cas) != len(set(cas)):
                name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
                name = name_df.iloc[0] if len(name_df) > 0 else "N/A"
                
                # Tìm các môn bị trùng
                ca_dict = defaultdict(list)
                for sub in subs:
                    ca = self.schedule.get(sub)
                    ca_dict[ca].append(sub)
                
                duplicate_info = []
                for ca, subj_list in ca_dict.items():
                    if len(subj_list) > 1:
                        duplicate_info.append(f"  Ca {ca}: {', '.join(subj_list)}")
                
                conflicts.append(f"⚠️ {sid} - {name}\n" + "\n".join(duplicate_info))

        if conflicts:
            self.warning_text.insert('1.0', 
                "🚨 PHÁT HIỆN XUNG ĐỘT!\n" + 
                "═" * 50 + "\n\n" +
                "\n\n".join(conflicts[:50]))
            self.warning_text.config(fg=self.colors['danger'])
            return True
        else:
            self.warning_text.insert('1.0', 
                "✅ HOÀN HẢO!\n" +
                "═" * 50 + "\n\n" +
                "🎉 Không có sinh viên nào bị trùng ca thi!\n\n"
                "📊 Thống kê:\n" +
                f"   • Tổng số ca: {max(self.schedule.values()) if self.schedule else 0}\n"
                f"   • Tổng số môn: {len(self.schedule)}\n"
                f"   • Tổng sinh viên: {len(self.student_subjects)}")
            self.warning_text.config(fg=self.colors['success'])
            return False

    def draw_graph(self):
        """Vẽ đồ thị xung đột với màu sắc đẹp"""
        if not HAS_GRAPH or not self.schedule:
            return
        
        for widget in self.graph_canvas.winfo_children():
            widget.destroy()

        G = nx.Graph()
        G.add_nodes_from(self.subjects)
        for u, vs in self.conflict_graph.items():
            G.add_edges_from((u, v) for v in vs)

        # Tạo figure với background đẹp
        fig = plt.figure(figsize=(12, 9), facecolor='white')
        ax = plt.gca()
        ax.set_facecolor('#fafafa')

        # Layout đẹp hơn
        if len(G.nodes()) < 20:
            pos = nx.spring_layout(G, k=3, iterations=100, seed=42)
        else:
            pos = nx.spring_layout(G, k=2.5, iterations=80, seed=42)

        # Màu sắc tươi sáng và đẹp mắt
        colors_palette = [
            '#7c3aed', '#ec4899', '#10b981', '#f59e0b', 
            '#3b82f6', '#ef4444', '#8b5cf6', '#14b8a6',
            '#f97316', '#06b6d4', '#84cc16', '#f43f5e',
            '#6366f1', '#a855f7', '#22c55e', '#eab308'
        ]
        
        node_colors = [colors_palette[(self.schedule.get(n, 1) - 1) % len(colors_palette)] 
                      for n in G.nodes()]

        # Vẽ edges với màu nhạt
        nx.draw_networkx_edges(G, pos, 
                              edge_color="#343453D5",
                              width=2,
                              alpha=0.4)

        # Vẽ nodes với màu sắc đẹp
        nx.draw_networkx_nodes(G, pos,
                              node_color=node_colors,
                              node_size=2000,
                              alpha=1,
                              edgecolors='gray',
                              linewidths=1)

        # Labels đẹp hơn
        nx.draw_networkx_labels(G, pos,
                               font_size=8,
                               font_weight='bold',
                               font_color='black',
                               font_family='sans-serif')

        plt.title("🎨 ĐỒ THỊ XUNG ĐỘT - MỖI MÀU = 1 CA THI",
                 fontsize=18,
                 fontweight='bold',
                 pad=20,
                 color="#000000")
        
        plt.axis('off')
        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.graph_canvas)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)

    def export_all(self):
        """Xuất file Excel với lịch thi đã sắp xếp"""
        if not self.schedule:
            messagebox.showwarning("⚠️ Cảnh báo", "Chưa chạy thuật toán!")
            return
        
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"LichThi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not path:
            return
        
        try:
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                # Sheet 1: Lịch thi theo ca - ĐÃ SẮP XẾP
                rows = []
                for subj, ca in sorted(self.schedule.items(), key=lambda x: (x[1], x[0])):
                    rows.append({
                        "Ca thi": ca,
                        "Môn học": subj,
                        "Số SV": len(self.subject_students[subj])
                    })
                df_schedule = pd.DataFrame(rows)
                df_schedule.to_excel(writer, sheet_name="LichThi", index=False)

                # Sheet 2: Lịch sinh viên
                sv_rows = []
                for sid in sorted(self.student_subjects.keys()):
                    subs = self.student_subjects[sid]
                    name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
                    name = name_df.iloc[0] if len(name_df) > 0 else "N/A"
                    
                    for s in sorted(subs):
                        sv_rows.append({
                            "MSSV": sid,
                            "Họ tên": name,
                            "Ca thi": self.schedule.get(s, ""),
                            "Môn học": s
                        })
                df_student = pd.DataFrame(sv_rows)
                df_student = df_student.sort_values(['MSSV', 'Ca thi'])
                df_student.to_excel(writer, sheet_name="LichSinhVien", index=False)

                # Sheet 3: Thống kê
                stats_rows = [
                    {"Chỉ số": "Tổng số môn học", "Giá trị": len(self.subjects)},
                    {"Chỉ số": "Tổng số sinh viên", "Giá trị": len(self.student_subjects)},
                    {"Chỉ số": "Số ca thi", "Giá trị": max(self.schedule.values())},
                    {"Chỉ số": "Số xung đột", "Giá trị": sum(len(v) for v in self.conflict_graph.values())//2},
                ]
                pd.DataFrame(stats_rows).to_excel(writer, sheet_name="ThongKe", index=False)

            messagebox.showinfo("🎉 Thành công!",
                              f"✅ Đã xuất file thành công!\n\n"
                              f"📂 Vị trí: {os.path.basename(path)}\n"
                              f"📊 3 sheet: LichThi, LichSinhVien, ThongKe")
        
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExamSchedulerPro(root)
    root.mainloop()
