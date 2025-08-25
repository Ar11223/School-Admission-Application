import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
from datetime import datetime
from tkcalendar import DateEntry
import json
from ttkthemes import ThemedTk

# --- 关键修复点 1：重新引入精确格式的CSV文件头部信息 ---
CSV_PREAMBLE = """访问形式*：可填值：公务拜访或入校参观,,,,,,,,,,,
访客姓名*：访客姓名必填,,,,,,,,,,,
手机号*：手机号必填，以#号结尾,,,,,,,,,,,
证件类型*：证件类型必填 填写:身份证或护照,,,,,,,,,,,
证件号码*：证件号码必填，以#号结尾,,,,,,,,,,,
车辆号码：车辆号码选填,,,,,,,,,,,
审批人学工号：审批人学工号 公务拜访必填 /入校参观不填，以#号结尾,,,,,,,,,,,
审批人姓名：审批人姓名 公务拜访选填 /入校参观不填,,,,,,,,,,,
场所名称*：场所名称必填 公务拜访为拜访场所/入校参观为参观场所，多个用@号隔开，最小层级为校区，填写场所名称如下:东区@西区@北区@梅山校区,,,,,,,,,,,
访问开始时间*：访问开始时间必填，时间格式如下:2023-06-27 00:00#，以#结尾,,,,,,,,,,,
访问结束时间*：访问结束时间必填，时间格式如下:2023-06-30 23:00#，以#结尾,,,,,,,,,,,
拜访人及事由：拜访人及事由 公务拜访选填 /入校参观不填,,,,,,,,,,,
"""

# --- 手动添加访客的弹出窗口类 (无变化) ---
class AddVisitorWindow(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.parent = parent
        self.callback = callback
        self.title("手动添加访客")
        self.geometry("400x250")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self, padding="15")
        frame.pack(fill="both", expand=True)
        fields = ["访客姓名", "手机号", "证件号码", "车辆号码"]
        self.entries = {}
        for i, field in enumerate(fields):
            ttk.Label(frame, text=f"{field}:").grid(row=i, column=0, sticky="w", pady=5)
            entry = ttk.Entry(frame, width=30)
            entry.grid(row=i, column=1, sticky="ew", pady=5)
            self.entries[field] = entry
        self.entries["访客姓名"].focus_set()
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=20)
        add_btn = ttk.Button(btn_frame, text="添加", command=self.submit_data)
        add_btn.pack(side="left", padx=10)
        cancel_btn = ttk.Button(btn_frame, text="取消", command=self.destroy)
        cancel_btn.pack(side="left", padx=10)

    def submit_data(self):
        name = self.entries["访客姓名"].get().strip()
        phone_raw = self.entries["手机号"].get().strip()
        id_raw = self.entries["证件号码"].get().strip()
        plate_raw = self.entries["车辆号码"].get().strip()
        if not name or not phone_raw or not id_raw:
            messagebox.showwarning("输入错误", "访客姓名、手机号和证件号码不能为空！", parent=self)
            return
        phone = phone_raw + '#'
        id_card = id_raw + '#'
        plate_upper = plate_raw.upper()
        plate_cleaned = re.sub(r'[^A-Z0-9\u4e00-\u9fa5]', '', plate_upper)
        processed_data = {
            '访客姓名*': name,
            '手机号*': phone,
            '证件号码*': id_card,
            '车辆号码': plate_cleaned
        }
        self.callback(processed_data)
        self.destroy()

# --- 数据处理核心功能 (无变化) ---
def process_excel_data(file_path):
    try:
        df = pd.read_excel(file_path, dtype=str)
        required_columns = ['访客姓名', '手机号', '证件号码', '车辆号码']
        if not all(col in df.columns for col in required_columns):
            missing = [col for col in required_columns if col not in df.columns]
            messagebox.showerror("错误", f"Excel文件中缺少以下列: {', '.join(missing)}")
            return None
        df.fillna('', inplace=True)
        processed_data = []
        for index, row in df.iterrows():
            phone = str(row['手机号']) + '#' if row['手机号'] else ''
            id_card = str(row['证件号码']) + '#' if row['证件号码'] else ''
            plate_raw = str(row['车辆号码']).upper()
            plate_cleaned = re.sub(r'[^A-Z0-9\u4e00-\u9fa5]', '', plate_raw)
            processed_data.append({
                '访客姓名*': str(row['访客姓名']),
                '手机号*': phone,
                '证件号码*': id_card,
                '车辆号码': plate_cleaned
            })
        return processed_data
    except Exception as e:
        messagebox.showerror("文件读取错误", f"处理Excel文件时出错: {e}")
        return None

# --- 主应用窗口类 (使用美化后的界面) ---
class VisitorApp(ThemedTk):
    def __init__(self):
        super().__init__()
        self.set_theme("arc")
        self.title("访客信息生成器 (v4.6 - 最终格式版)")
        self.geometry("850x800")

        style = ttk.Style(self)
        style.configure("TLabel", font=("微软雅黑", 10))
        style.configure("TButton", font=("微软雅黑", 10))
        style.configure("TEntry", font=("微软雅黑", 10))
        style.configure("TCombobox", font=("微软雅黑", 10))
        style.configure("TLabelframe.Label", font=("微软雅黑", 11, "bold"))
        style.configure("Treeview.Heading", font=("微软雅黑", 10, "bold"))
        style.configure("Treeview", rowheight=25, font=("微软雅黑", 10))
        style.map("Treeview", background=[("selected", "#0078D7")])
        style.configure("Accent.TButton", font=("微软雅黑", 11, "bold"))

        self.visitor_data = []
        self.approver_history = []
        self.start_time_widgets = {}
        self.end_time_widgets = {}
        self._load_history()
        self.create_widgets()

    def _load_history(self):
        try:
            with open("history.json", 'r', encoding='utf-8') as f:
                self.approver_history = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self.approver_history = []

    def _save_history(self, new_id, new_name):
        new_entry = {"id": new_id, "name": new_name}
        if new_entry in self.approver_history: return
        self.approver_history.insert(0, new_entry)
        self.approver_history = self.approver_history[:10]
        try:
            with open("history.json", 'w', encoding='utf-8') as f:
                json.dump(self.approver_history, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showwarning("警告", f"无法保存审批人历史记录: {e}")

    def _create_datetime_picker(self, parent, label_text, row):
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky="w", pady=5)
        dt_frame = ttk.Frame(parent)
        dt_frame.grid(row=row, column=1, sticky="ew")
        date_entry = DateEntry(dt_frame, width=12, borderwidth=2, date_pattern='yyyy-mm-dd', mindate=datetime.now())
        date_entry.pack(side="left")
        ttk.Label(dt_frame, text=" H:").pack(side="left", padx=(10, 0))
        hour_combo = ttk.Combobox(dt_frame, width=4, state="readonly", values=[f"{h:02d}" for h in range(24)])
        hour_combo.set("08")
        hour_combo.pack(side="left")
        ttk.Label(dt_frame, text=" M:").pack(side="left", padx=(5, 0))
        minute_combo = ttk.Combobox(dt_frame, width=4, state="readonly", values=[f"{m:02d}" for m in range(60)])
        minute_combo.set("00")
        minute_combo.pack(side="left")
        return {"date": date_entry, "hour": hour_combo, "minute": minute_combo}

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill="both", expand=True)
        main_frame.columnconfigure(1, weight=1)

        left_panel = ttk.Frame(main_frame)
        left_panel.grid(row=0, column=0, sticky="ns", padx=(0, 10))

        right_panel = ttk.Frame(main_frame)
        right_panel.grid(row=0, column=1, sticky="nsew")
        right_panel.rowconfigure(0, weight=1)
        right_panel.columnconfigure(0, weight=1)

        input_frame = ttk.LabelFrame(left_panel, text="第一步：添加访客数据", padding="15")
        input_frame.pack(fill="x", pady=(5, 10))
        self.upload_btn = ttk.Button(input_frame, text="从Excel批量导入", command=self.upload_file)
        self.upload_btn.pack(fill="x", padx=5, pady=5)
        self.file_label = ttk.Label(input_frame, text="尚未选择文件", anchor="center")
        self.file_label.pack(fill="x", padx=5, pady=5)
        self.manual_add_btn = ttk.Button(input_frame, text="手动添加单个访客", command=self.open_add_visitor_window)
        self.manual_add_btn.pack(fill="x", padx=5, pady=5)

        form_frame = ttk.LabelFrame(left_panel, text="第二步：填写公共信息", padding="15")
        form_frame.pack(fill="x", pady=10)
        form_frame.columnconfigure(1, weight=1)
        ttk.Label(form_frame, text="审批人学工号:").grid(row=0, column=0, sticky="w", pady=2)
        id_history = [item['id'] for item in self.approver_history]
        self.approver_id_combo = ttk.Combobox(form_frame, values=id_history)
        self.approver_id_combo.grid(row=0, column=1, sticky="ew", pady=2)
        ttk.Label(form_frame, text="审批人姓名:").grid(row=1, column=0, sticky="w", pady=2)
        name_history = [item['name'] for item in self.approver_history]
        self.approver_name_combo = ttk.Combobox(form_frame, values=name_history)
        self.approver_name_combo.grid(row=1, column=1, sticky="ew", pady=2)
        self.approver_id_combo.bind("<<ComboboxSelected>>", self.on_approver_selected)
        self.approver_name_combo.bind("<<ComboboxSelected>>", self.on_approver_selected)
        ttk.Label(form_frame, text="拜访人及事由:").grid(row=2, column=0, sticky="w", pady=2)
        self.reason_entry = ttk.Entry(form_frame)
        self.reason_entry.grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Label(form_frame, text="访问形式:").grid(row=3, column=0, sticky="w", pady=2)
        self.visit_type_combo = ttk.Combobox(form_frame, values=["公务拜访", "入校参观"], state="readonly")
        self.visit_type_combo.grid(row=3, column=1, sticky="ew", pady=2)
        self.visit_type_combo.set("公务拜访")
        ttk.Label(form_frame, text="证件类型:").grid(row=4, column=0, sticky="w", pady=2)
        self.id_type_combo = ttk.Combobox(form_frame, values=["身份证", "护照"], state="readonly")
        self.id_type_combo.grid(row=4, column=1, sticky="ew", pady=2)
        self.id_type_combo.set("身份证")
        ttk.Label(form_frame, text="场所名称 (可多选):").grid(row=5, column=0, sticky="w", pady=5)
        places_frame = ttk.Frame(form_frame)
        places_frame.grid(row=5, column=1, sticky="w")
        self.places = ["东区", "西区", "北区", "梅山校区"]
        self.place_vars = [tk.BooleanVar() for _ in self.places]
        for i, place in enumerate(self.places):
            cb = ttk.Checkbutton(places_frame, text=place, variable=self.place_vars[i])
            cb.pack(side="left", padx=5)
        self.start_time_widgets = self._create_datetime_picker(form_frame, "访问开始时间:", 6)
        self.end_time_widgets = self._create_datetime_picker(form_frame, "访问结束时间:", 7)

        export_frame = ttk.LabelFrame(left_panel, text="第三步：生成并导出", padding="15")
        export_frame.pack(fill="x", pady=(10, 5))
        self.generate_btn = ttk.Button(export_frame, text="生成 CSV 文件", command=self.generate_csv, style="Accent.TButton")
        self.generate_btn.pack(fill="x", padx=5, pady=10, ipady=5)

        preview_frame = ttk.LabelFrame(right_panel, text="数据预览", padding="15")
        preview_frame.grid(row=0, column=0, sticky="nsew", pady=(5, 10))
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        
        self.tree = ttk.Treeview(preview_frame, columns=('姓名', '手机', '证件', '车牌'), show='headings')
        self.tree.heading('姓名', text='访客姓名')
        self.tree.heading('手机', text='手机号')
        self.tree.heading('证件', text='证件号码')
        self.tree.heading('车牌', text='车辆号码')
        self.tree.column('姓名', width=100, anchor='center')
        self.tree.column('手机', width=150, anchor='center')
        self.tree.column('证件', width=180, anchor='center')
        self.tree.column('车牌', width=120, anchor='center')
        self.tree.grid(row=0, column=0, sticky="nsew")

        v_scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        v_scrollbar.grid(row=0, column=1, sticky='ns')

        h_scrollbar = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=h_scrollbar.set)
        h_scrollbar.grid(row=1, column=0, sticky='ew')

    def on_approver_selected(self, event):
        widget = event.widget
        selected_index = widget.current()
        if selected_index != -1:
            selected_entry = self.approver_history[selected_index]
            self.approver_id_combo.set(selected_entry['id'])
            self.approver_name_combo.set(selected_entry['name'])

    def update_preview_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        self.tree.tag_configure('oddrow', background='#F0F0F0')
        self.tree.tag_configure('evenrow', background='white')

        if self.visitor_data:
            for i, visitor in enumerate(self.visitor_data):
                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                self.tree.insert('', tk.END, values=(
                    visitor['访客姓名*'],
                    visitor['手机号*'],
                    visitor['证件号码*'],
                    visitor['车辆号码']
                ), tags=(tag,))

    def upload_file(self):
        file_path = filedialog.askopenfilename(title="请选择一个Excel文件", filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path: return
        
        self.file_label.config(text=file_path.split('/')[-1])
        new_data = process_excel_data(file_path)
        
        if new_data:
            self.visitor_data.clear()
            self.visitor_data.extend(new_data)
            messagebox.showinfo("成功", f"成功导入并处理了 {len(self.visitor_data)} 条访客数据。")
            self.update_preview_table()

    def open_add_visitor_window(self):
        AddVisitorWindow(self, self.add_visitor_from_manual_entry)

    def add_visitor_from_manual_entry(self, data):
        self.visitor_data.append(data)
        self.update_preview_table()
        messagebox.showinfo("成功", f"访客 {data['访客姓名*']} 已添加。")

    def get_selected_places(self):
        selected = [place for place, var in zip(self.places, self.place_vars) if var.get()]
        return "@".join(selected)

    # --- 这里是关键的修复点 ---
    def generate_csv(self):
        if not self.visitor_data:
            messagebox.showwarning("警告", "访客列表为空，请先导入或添加访客！")
            return
        
        approver_id = self.approver_id_combo.get()
        approver_name = self.approver_name_combo.get()
        reason = self.reason_entry.get()
        visit_type = self.visit_type_combo.get()
        id_type = self.id_type_combo.get()
        places = self.get_selected_places()
        
        if not all([approver_id, approver_name, reason, visit_type, id_type, places]):
            messagebox.showwarning("警告", "请填写所有必填信息，并至少选择一个场所！")
            return
        
        try:
            start_datetime = datetime.combine(self.start_time_widgets['date'].get_date(), datetime.min.time()).replace(hour=int(self.start_time_widgets['hour'].get()), minute=int(self.start_time_widgets['minute'].get()))
            end_datetime = datetime.combine(self.end_time_widgets['date'].get_date(), datetime.min.time()).replace(hour=int(self.end_time_widgets['hour'].get()), minute=int(self.end_time_widgets['minute'].get()))
        except ValueError:
            messagebox.showerror("错误", "请选择有效的小时和分钟！")
            return
        
        if start_datetime < datetime.now():
            messagebox.showwarning("警告", "访问开始时间不能早于当前时间！")
            return
        if end_datetime <= start_datetime:
            messagebox.showwarning("警告", "结束时间必须晚于开始时间！")
            return
        
        start_time_str = start_datetime.strftime('%Y-%m-%d %H:%M') + "#"
        end_time_str = end_datetime.strftime('%Y-%m-%d %H:%M') + "#"

        # 定义最终的列顺序，这必须和第13行的表头完全一致
        final_columns = [
            '访问形式*', '访客姓名*', '手机号*', '证件类型*', '证件号码*', '车辆号码',
            '审批人学工号', '审批人姓名', '场所名称*', '访问开始时间*', '访问结束时间*', '拜访人及事由'
        ]
        
        final_data = []
        for visitor in self.visitor_data:
            # 修复了这里的字典键，确保与最终列名匹配
            row = {
                '访问形式*': visit_type,
                '访客姓名*': visitor['访客姓名*'],
                '手机号*': visitor['手机号*'],
                '证件类型*': id_type,
                '证件号码*': visitor['证件号码*'],
                '车辆号码': visitor['车辆号码'],
                '审批人学工号': approver_id + '#',
                '审批人姓名': approver_name,
                '场所名称*': places,
                '访问开始时间*': start_time_str,
                '访问结束时间*': end_time_str,
                '拜访人及事由': reason
            }
            final_data.append(row)

        # 创建DataFrame，并指定列的顺序
        output_df = pd.DataFrame(final_data, columns=final_columns)

        save_path = filedialog.asksaveasfilename(
            title="保存CSV文件",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not save_path: return

        try:
            # 使用 'w' 模式打开文件，并指定GBK编码
            with open(save_path, 'w', encoding='gbk', newline='') as f:
                # 步骤一：写入第1-12行的固定头部信息
                f.write(CSV_PREAMBLE)
                
                # 步骤二：将DataFrame写入文件，使用默认的逗号作为分隔符
                # 这会先写入表头（成为第13行），然后写入数据（从第14行开始）
                output_df.to_csv(f, index=False)

            messagebox.showinfo("成功", f"文件已成功导出到: {save_path}")
            self._save_history(approver_id, approver_name)
        except PermissionError:
            messagebox.showerror("导出失败", "权限错误：无法写入所选位置。\n请尝试选择其他文件夹，例如“文档”或“桌面”。")
        except Exception as e:
            messagebox.showerror("导出失败", f"保存文件时发生未知错误: {e}")


if __name__ == "__main__":
    app = VisitorApp()
    app.mainloop()