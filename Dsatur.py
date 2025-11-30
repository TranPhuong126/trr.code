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
        self.root.title("Xếp Lịch Thi Thông Minh - DSatur Pro v2.0")
        self.root.geometry("1500x900")
        self.root.configure(bg='#f0f2f5')
        self.root.state('zoomed')

        # Dữ liệu
        self.data = None
        self.subjects = []
        self.student_subjects = defaultdict(set)
        self.subject_students = defaultdict(set)
        self.conflict_graph = defaultdict(set)
        self.schedule = {}
        self.max_exams_per_day = 3

        # Style hiện đại
        self.colors = {
            'bg': '#f0f2f5',
            'card': '#ffffff',
            'primary': '#4361ee',
            'success': '#4cc9f0',
            'warning': '#f72585',
            'danger': '#d90429',
            'dark': '#2b2d42',
            'light': '#edf2f4'
        }

        self.setup_styles()
        self.create_ui()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TButton", padding=10, font=('Segoe UI', 10, 'bold'))
        style.configure("Treeview", background="white", fieldbackground="white", rowheight=30)
        style.map('Treeview', background=[('selected', self.colors['primary'])])

    def create_ui(self):
        # Header
        header = tk.Frame(self.root, bg=self.colors['primary'], height=80)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text="XẾP LỊCH THI THÔNG MINH - DSATUR PRO", font=('Segoe UI', 28, 'bold'),
                 fg='white', bg=self.colors['primary']).pack(pady=15)

        main = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, bg=self.colors['bg'])
        main.pack(fill='both', expand=True, padx=15, pady=15)

        # === SIDEBAR TRÁI ===
        left = tk.Frame(main, bg=self.colors['card'], width=380, relief='flat')
        main.add(left)

        # Upload
        upload_frame = tk.LabelFrame(left, text="NHẬP DỮ LIỆU", bg=self.colors['card'], fg=self.colors['dark'], font=('Segoe UI', 11, 'bold'))
        upload_frame.pack(fill='x', padx=15, pady=10)
        tk.Button(upload_frame, text="CHỌN FILE EXCEL", command=self.load_file,
                  bg=self.colors['primary'], fg='white', font=('Segoe UI', 10, 'bold'),
                  relief='flat', padx=20, pady=10, cursor='hand2').pack(pady=10)
        self.file_label = tk.Label(upload_frame, text="Chưa chọn file...", bg=self.colors['card'], fg='gray', wraplength=340)
        self.file_label.pack(pady=5)

        # Cài đặt
        setting_frame = tk.LabelFrame(left, text="CÀI ĐẶT", bg=self.colors['card'], fg=self.colors['dark'], font=('Segoe UI', 11, 'bold'))
        setting_frame.pack(fill='x', padx=15, pady=10)
        tk.Label(setting_frame, text="Số môn tối đa mỗi ngày:", bg=self.colors['card']).pack(anchor='w', padx=10, pady=5)
        self.max_var = tk.IntVar(value=3)
        tk.Spinbox(setting_frame, from_=1, to=5, textvariable=self.max_var, width=5).pack(anchor='w', padx=10)

        # Nút chạy
        tk.Button(left, text="CHẠY DSATUR", command=self.run_dsatur,
                  bg=self.colors['success'], fg='white', font=('Segoe UI', 14, 'bold'),
                  relief='flat', padx=20, pady=15, cursor='hand2').pack(pady=30, padx=30, fill='x')

        # Thống kê
        stats_frame = tk.LabelFrame(left, text="THỐNG KÊ", bg=self.colors['card'], fg=self.colors['dark'], font=('Segoe UI', 11, 'bold'))
        stats_frame.pack(fill='both', expand=True, padx=15, pady=10)
        self.stats_text = tk.Text(stats_frame, height=12, bg=self.colors['light'], relief='flat', font=('Consolas', 10))
        self.stats_text.pack(fill='both', padx=10, pady=10)

        # === PHẦN PHẢI ===
        right = tk.Frame(main, bg=self.colors['card'])
        main.add(right)

        notebook = ttk.Notebook(right)
        notebook.pack(fill='both', expand=True, padx=15, pady=15)

        # Tab 1: Lịch thi
        tab1 = tk.Frame(notebook, bg='white')
        notebook.add(tab1, text='Lịch Thi Theo Ca')
        self.tree_schedule = ttk.Treeview(tab1, columns=('Ca', 'Môn', 'SV'), show='headings', height=25)
        self.tree_schedule.heading('Ca', text='Ca Thi')
        self.tree_schedule.heading('Môn', text='Môn Học')
        self.tree_schedule.heading('SV', text='Số SV')
        self.tree_schedule.column('Ca', width=80, anchor='center')
        self.tree_schedule.column('Môn', width=500)
        self.tree_schedule.column('SV', width=100, anchor='center')
        self.tree_schedule.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        ttk.Scrollbar(tab1, orient='vertical', command=self.tree_schedule.yview).pack(side='right', fill='y')
        self.tree_schedule.configure(yscrollcommand=lambda f, l: None)

        # Tab 2: Lịch SV
        tab2 = tk.Frame(notebook, bg='white')
        notebook.add(tab2, text='Lịch Thi Sinh Viên')
        search_frame = tk.Frame(tab2, bg='white')
        search_frame.pack(fill='x', padx=10, pady=5)
        tk.Label(search_frame, text="Tìm:", bg='white').pack(side='left')
        self.search_var = tk.StringVar()
        tk.Entry(search_frame, textvariable=self.search_var, width=40).pack(side='left', padx=5)
        self.search_var.trace('w', self.filter_students)
        self.tree_student = ttk.Treeview(tab2, columns=('MSSV', 'Tên', 'Ca', 'Môn'), show='headings')
        self.tree_student.heading('MSSV', text='MSSV')
        self.tree_student.heading('Tên', text='Họ Tên')
        self.tree_student.heading('Ca', text='Ca Thi')
        self.tree_student.heading('Môn', text='Môn Học')
        self.tree_student.pack(fill='both', expand=True, padx=10, pady=10)

        # Tab 3: Đồ thị
        tab3 = tk.Frame(notebook, bg='white')
        notebook.add(tab3, text='Đồ Thị Xung Đột')
        self.graph_canvas = tk.Canvas(tab3, bg='white')
        self.graph_canvas.pack(fill='both', expand=True, padx=20, pady=20)
        if not HAS_GRAPH:
            tk.Label(tab3, text="Cài networkx + matplotlib để xem đồ thị!", fg='red').pack(pady=50)

        # Tab 4: Export & Cảnh báo
        tab4 = tk.Frame(notebook, bg='white')
        notebook.add(tab4, text='Export & Kiểm Tra')
        tk.Button(tab4, text="XUẤT LỊCH THI EXCEL", command=self.export_all,
                  bg=self.colors['warning'], fg='white', font=('Segoe UI', 12, 'bold'), pady=10).pack(pady=20)
        self.warning_text = tk.Text(tab4, height=20, bg='#fff5f5', fg='red', font=('Segoe UI', 10))
        self.warning_text.pack(fill='both', expand=True, padx=20, pady=10)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return

        try:
            all_dfs = []
            excel = pd.ExcelFile(path, engine='openpyxl')
            
            for sheet in excel.sheet_names:
                try:
                    # Đọc toàn bộ sheet như string
                    df = pd.read_excel(excel, sheet_name=sheet, header=None, dtype=str, engine='openpyxl')
                    df = df.fillna('')
                    
                    # Tìm dòng header (chứa "Mã SV" hoặc "MSSV")
                    header_row = None
                    for idx in range(min(5, len(df))):  # Chỉ tìm trong 5 dòng đầu
                        row_text = ' '.join(df.iloc[idx].astype(str).str.lower().tolist())
                        if 'mã sv' in row_text or 'mssv' in row_text or 'ma sv' in row_text:
                            header_row = idx
                            break
                    
                    if header_row is None:
                        print(f"Không tìm thấy header trong sheet {sheet}")
                        continue
                    
                    # Lấy tên môn học (dòng đầu tiên hoặc tên sheet)
                    subject_name = sheet
                    if header_row > 0:
                        first_cell = str(df.iloc[0, 0]).strip()
                        if 'VJU' in first_cell or 'AET' in first_cell:
                            subject_name = first_cell
                    
                    # Đặt header
                    df.columns = df.iloc[header_row]
                    df = df.iloc[header_row + 1:].reset_index(drop=True)
                    
                    # Tìm cột Mã SV và Họ Tên
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
                        print(f"Không tìm thấy cột Mã SV trong sheet {sheet}")
                        continue
                    
                    # Lọc dữ liệu - FIX LỖI BOOLEAN INDEXING
                    if hoten_col:
                        df_clean = df[[masv_col, hoten_col]].copy()
                        df_clean.columns = ['MaSV', 'HoTen']
                    else:
                        df_clean = df[[masv_col]].copy()
                        df_clean.columns = ['MaSV']
                        df_clean['HoTen'] = 'N/A'
                    
                    # Làm sạch dữ liệu
                    df_clean['MaSV'] = df_clean['MaSV'].astype(str).str.strip()
                    
                    # FIX: Sử dụng .loc để tránh lỗi boolean indexing
                    df_clean = df_clean.loc[df_clean['MaSV'].str.len() > 0].copy()
                    
                    # Chỉ giữ những dòng có MSSV là số
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
                messagebox.showerror("Lỗi", "Không tìm thấy dữ liệu hợp lệ!\n\nKiểm tra:\n- File có cột 'Mã SV'\n- Có ít nhất 1 sinh viên")
                return
            
            self.data = pd.concat(all_dfs, ignore_index=True)
            self.data.drop_duplicates(subset=['MaSV', 'ChuongTrinh'], inplace=True)
            
            self.file_label.config(
                text=f"ĐÃ TẢI: {os.path.basename(path)}\n{len(self.data)} dòng • {self.data['ChuongTrinh'].nunique()} môn",
                fg='green'
            )
            
            messagebox.showinfo("Thành công", 
                f"Đã tải thành công!\n\n"
                f"• {len(self.data):,} bản ghi\n"
                f"• {len(excel.sheet_names)} sheet\n"
                f"• {self.data['MaSV'].nunique()} sinh viên\n"
                f"• {self.data['ChuongTrinh'].nunique()} môn học")
            
            self.process_data()

        except Exception as e:
            messagebox.showerror("Lỗi đọc file", f"Chi tiết lỗi:\n{str(e)}")
            import traceback
            traceback.print_exc()

    def process_data(self):
        self.subjects = self.data['ChuongTrinh'].unique().tolist()
        self.student_subjects.clear()
        self.subject_students.clear()
        self.conflict_graph.clear()

        for _, row in self.data.iterrows():
            sid = str(row['MaSV']).strip()
            subj = row['ChuongTrinh']
            self.student_subjects[sid].add(subj)
            self.subject_students[subj].add(sid)

        # Xây đồ thị xung đột
        for subs in self.student_subjects.values():
            subs = list(subs)
            for i in range(len(subs)):
                for j in range(i+1, len(subs)):
                    self.conflict_graph[subs[i]].add(subs[j])
                    self.conflict_graph[subs[j]].add(subs[i])

        self.update_stats()

    def update_stats(self):
        text = f"TỔNG QUAN DỮ LIỆU\n"
        text += f"{'='*40}\n"
        text += f"Sinh viên: {len(self.student_subjects):,}\n"
        text += f"Môn học: {len(self.subjects):,}\n"
        text += f"Xung đột cạnh: {sum(len(v) for v in self.conflict_graph.values())//2:,}\n"
        self.stats_text.delete(1.0, 'end')
        self.stats_text.insert('end', text)

    def run_dsatur(self):
        # FIX: Thay đổi cách kiểm tra DataFrame
        if self.data is None or self.data.empty or len(self.subjects) == 0:
            messagebox.showwarning("Cảnh báo", "Chưa tải dữ liệu!")
            return

        self.max_exams_per_day = self.max_var.get()
        self.schedule.clear()

        # DSatur tối ưu bằng heap
        degree = {s: len(self.conflict_graph[s]) for s in self.subjects}
        saturation = {s: 0 for s in self.subjects}
        color_of = {}
        heap = [(-0, -degree[s], s) for s in self.subjects]
        heapq.heapify(heap)
        colored = set()

        for _ in range(len(self.subjects)):
            while heap:
                _, _, subj = heapq.heappop(heap)
                if subj not in colored:
                    break
            used = {color_of.get(n) for n in self.conflict_graph[subj] if n in color_of}
            color = 1
            while color in used:
                color += 1
            color_of[subj] = color
            colored.add(subj)

            for nei in self.conflict_graph[subj]:
                if nei not in colored:
                    saturation[nei] += 1
                    heapq.heappush(heap, (-saturation[nei], -degree[nei], nei))

        self.schedule = color_of
        self.display_results()
        self.check_conflicts()
        self.draw_graph()
        messagebox.showinfo("HOÀN THÀNH", f"Đã xếp lịch thành công!\nTổng: {max(color_of.values())} ca thi")

    def display_results(self):
        for i in self.tree_schedule.get_children():
            self.tree_schedule.delete(i)
        for i in self.tree_student.get_children():
            self.tree_student.delete(i)

        # Lịch thi
        ca_dict = defaultdict(list)
        for subj, ca in self.schedule.items():
            ca_dict[ca].append((subj, len(self.subject_students[subj])))
        for ca in sorted(ca_dict):
            for subj, count in sorted(ca_dict[ca], key=lambda x: -x[1]):
                self.tree_schedule.insert('', 'end', values=(f'Ca {ca}', subj, count))

        # Lịch SV
        for sid, subs in self.student_subjects.items():
            # FIX: Sử dụng .loc để tránh lỗi
            name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
            name = name_df.iloc[0] if len(name_df) > 0 else "N/A"
            
            for sub in subs:
                ca = self.schedule.get(sub, "?")
                self.tree_student.insert('', 'end', values=(sid, name, ca, sub))

    def filter_students(self, *args):
        search = self.search_var.get().lower()
        for i in self.tree_student.get_children():
            self.tree_student.delete(i)
        
        for sid, subs in self.student_subjects.items():
            # FIX: Sử dụng .loc
            name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
            name = name_df.iloc[0] if len(name_df) > 0 else "N/A"
            
            if search in sid.lower() or search in name.lower():
                for sub in subs:
                    ca = self.schedule.get(sub, "?")
                    self.tree_student.insert('', 'end', values=(sid, name, ca, sub))

    def check_conflicts(self):
        self.warning_text.delete(1.0, 'end')
        conflicts = []
        
        for sid, subs in self.student_subjects.items():
            cas = [self.schedule.get(s) for s in subs]
            if len(cas) != len(set(cas)):
                # FIX: Sử dụng .loc
                name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
                name = name_df.iloc[0] if len(name_df) > 0 else "N/A"
                conflicts.append(f"TRÙNG: {sid} - {name}")

        if conflicts:
            self.warning_text.insert('1.0', "CÓ LỖI TRÙNG CA!\n" + "\n".join(conflicts[:50]))
            self.warning_text.config(fg='red')
        else:
            self.warning_text.insert('1.0', "HOÀN HẢO! Không có sinh viên nào bị trùng ca thi")
            self.warning_text.config(fg='green')

    def draw_graph(self):
        if not HAS_GRAPH or not self.schedule:
            return
        for widget in self.graph_canvas.winfo_children():
            widget.destroy()

        G = nx.Graph()
        G.add_nodes_from(self.subjects)
        for u, vs in self.conflict_graph.items():
            G.add_edges_from((u, v) for v in vs)

        plt.figure(figsize=(10, 8))
        pos = nx.spring_layout(G, k=2, iterations=50)
        colors = [self.schedule.get(n, 0) for n in G.nodes()]
        nx.draw(G, pos, node_color=colors, node_size=1500, font_size=9,
                with_labels=True, cmap=plt.cm.tab10, edge_color='#95a5a6')

        plt.title("ĐỒ THỊ XUNG ĐỘT - MÀU = CA THI", fontsize=10, pad=15)
        fig = plt.gcf()
        canvas = FigureCanvasTkAgg(fig, master=self.graph_canvas)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)

    def export_all(self):
        if not self.schedule:
            messagebox.showwarning("Cảnh báo", "Chưa chạy thuật toán!")
            return
        
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            # Sheet lịch thi
            rows = []
            for subj, ca in self.schedule.items():
                rows.append({"Ca thi": ca, "Môn học": subj, "Số SV": len(self.subject_students[subj])})
            pd.DataFrame(rows).to_excel(writer, sheet_name="LichThi", index=False)

            # Sheet lịch SV
            sv_rows = []
            for sid, subs in self.student_subjects.items():
                # FIX: Sử dụng .loc
                name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
                name = name_df.iloc[0] if len(name_df) > 0 else "N/A"
                
                for s in subs:
                    sv_rows.append({"MSSV": sid, "Họ tên": name, "Ca thi": self.schedule.get(s,""), "Môn": s})
            pd.DataFrame(sv_rows).to_excel(writer, sheet_name="LichSinhVien", index=False)

        messagebox.showinfo("Thành công", f"Đã xuất file:\n{os.path.basename(path)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExamSchedulerPro(root)
    root.mainloop()