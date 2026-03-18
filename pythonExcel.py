import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import qrcode
import re
from PIL import ImageTk, Image
from datetime import datetime

# 全局路径锁定
if getattr(sys, 'frozen', False):
    APP_PATH = os.path.dirname(sys.executable)
else:
    APP_PATH = os.path.dirname(os.path.abspath(__file__))


class MedicalAppV3:
    def __init__(self, root):
        self.root = root
        self.root.title("医疗费报销系统 V3.8 (紧凑布局版)")
        self.root.geometry("850x750")

        # 1. 核心协议与映射
        self.out_order = ["医事服务费", "检查费", "治疗费", "西药", "中药", "卫生材料费", "其他费"]
        self.in_order = ["医事服务费", "检查费", "治疗费", "西药", "中药", "卫生材料费", "床位费", "其他费"]
        self.mapping = {
            "检查费": ["检查费", "化验费"], "治疗费": ["治疗费"], "西药": ["西药费"],
            "中药": ["中药饮片", "中草药", "中成药"], "卫生材料费": ["材料费", "卫生材料费"],
            "床位费": ["床位费", "空调费", "住院费", "住院"]
        }

        # 2. 变量定义
        self.data_out = self.init_struct(self.out_order)
        self.data_in = self.init_struct(self.in_order)
        self.out_totals = {k: tk.StringVar(value="0.00") for k in ["amt", "self", "refund"]}
        self.in_totals = {k: tk.StringVar(value="0.00") for k in ["amt", "self", "refund"]}

        self.in_days_var = tk.StringVar(value="0")
        self.base_dir = tk.StringVar(value=APP_PATH)

        self.active_serial_var = tk.StringVar(value="未开始")
        self.next_serial_var = tk.StringVar()
        self.current_seq = 1
        self._repeat_job = None

        self.out_calc_entries = []
        self.in_calc_entries = []

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 3. 布局
        self.setup_ui()
        self.bind_traces()

        # 4. 初始化序号
        self.refresh_next_serial()

    def init_struct(self, order_list):
        return {cat: {"amt": tk.StringVar(value="0.00"), "self": tk.StringVar(value="0.00"),
                      "refund": tk.StringVar(value="0.00"), "calc": tk.StringVar(value="")} for cat in order_list}

    def on_closing(self):
        if messagebox.askokcancel("退出确认", "确定要退出吗？"):
            self.root.destroy()

    def setup_ui(self):
        # 顶部管理区 - 根目录 (稍作收紧)
        top_frame = tk.Frame(self.root, pady=5, bg="#f8f9fa")
        top_frame.pack(fill='x')

        tk.Label(top_frame, text=" 存档根目录:", bg="#f8f9fa").pack(side='left', padx=(20, 5))
        tk.Entry(top_frame, textvariable=self.base_dir, width=60, state='readonly').pack(side='left', padx=5)
        tk.Button(top_frame, text="更改目录", command=self.browse_base_dir, font=("微软雅黑", 8)).pack(side='left',
                                                                                                       padx=5)

        # 核心控制区 - 蓝色背景 (高度压缩版)
        ctrl_frame = tk.Frame(self.root, pady=5, bg="#e3f2fd")  # pady 从 12 降至 5
        ctrl_frame.pack(fill='x')

        tk.Label(ctrl_frame, text="下一报销单序号:", bg="#e3f2fd", font=("微软雅黑", 9)).pack(side='left', padx=(20, 5))
        tk.Entry(ctrl_frame, textvariable=self.next_serial_var, width=12, state='readonly',
                 fg="#1565c0", font=("Consolas", 10, "bold"), justify='center').pack(side='left', padx=2)

        # 增减按钮组 (微型化)
        btn_f = tk.Frame(ctrl_frame, bg="#e3f2fd")
        btn_f.pack(side='left')
        btn_up = tk.Button(btn_f, text="▲", font=("Arial", 6), width=2, height=0)
        btn_up.pack(side='top')
        btn_dn = tk.Button(btn_f, text="▼", font=("Arial", 6), width=2, height=0)
        btn_dn.pack(side='bottom')

        btn_up.bind("<ButtonPress-1>", lambda e: self.start_adjust(1))
        btn_up.bind("<ButtonRelease-1>", self.stop_adjust)
        btn_dn.bind("<ButtonPress-1>", lambda e: self.start_adjust(-1))
        btn_dn.bind("<ButtonRelease-1>", self.stop_adjust)

        # 新建按钮 (字体减小，内边距减小)
        tk.Button(ctrl_frame, text="新建报销单", command=self.execute_create_serial,
                  bg="#1976D2", fg="white", font=("微软雅黑", 9, "bold"), pady=1).pack(side='left', padx=15)

        # 当前执行序号
        tk.Label(ctrl_frame, text="| 当前录入:", bg="#e3f2fd", font=("微软雅黑", 9)).pack(side='left', padx=(5, 2))
        tk.Label(ctrl_frame, textvariable=self.active_serial_var, fg="#2e7d32", font=("Consolas", 10, "bold"),
                 bg="#ffffff", width=12, relief="sunken", bd=1).pack(side='left', padx=5)

        # 工具栏 (上传/扫码)
        tool_frame = tk.Frame(self.root, pady=8, bg="#f5f5f5")
        tool_frame.pack(fill='x')
        tk.Button(tool_frame, text=" 1. 上传数据表 (Excel) ", command=self.load_excel,
                  bg="#2196F3", fg="white", font=("微软雅黑", 10, "bold"), width=25).pack(side='left', padx=20)
        tk.Button(tool_frame, text=" 2. 生成二维码 (32项) ", command=self.generate_qr,
                  bg="#FF9800", fg="white", font=("微软雅黑", 10, "bold"), width=25).pack(side='left', padx=10)

        # 滚动区域 (保持不变)
        main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        self.scroll_frame = tk.Frame(main_canvas)
        self.scroll_frame.bind("<Configure>", lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
        main_canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        main_canvas.pack(side="left", fill="both", expand=True);
        scrollbar.pack(side="right", fill="y")

        self.col_widths = [18, 15, 15, 15, 35]
        self.create_section(self.scroll_frame, "【 门 诊 收 据 明 细 】", self.data_out, self.out_order,
                            self.out_calc_entries, self.out_totals)

        day_f = tk.Frame(self.scroll_frame, pady=5);
        day_f.pack(fill='x', padx=20)
        tk.Label(day_f, text="住院天数：", font=("微软雅黑", 10, "bold")).pack(side='left')
        tk.Entry(day_f, textvariable=self.in_days_var, width=10, bg="#fffde7", justify='center').pack(side='left')

        self.create_section(self.scroll_frame, "【 住 院 收 据 明 细 】", self.data_in, self.in_order,
                            self.in_calc_entries, self.in_totals)

    # --- 后台逻辑 (保持一致) ---
    def start_adjust(self, delta):
        self.adjust_seq(delta)
        self._repeat_job = self.root.after(500, lambda: self.repeat_adjust(delta))

    def repeat_adjust(self, delta):
        self.adjust_seq(delta)
        self._repeat_job = self.root.after(100, lambda: self.repeat_adjust(delta))

    def stop_adjust(self, event):
        if self._repeat_job:
            self.root.after_cancel(self._repeat_job)
            self._repeat_job = None

    def adjust_seq(self, delta):
        self.current_seq = max(1, self.current_seq + delta)
        self.update_serial_display()

    def update_serial_display(self):
        today = datetime.now().strftime("%Y%m%d")
        self.next_serial_var.set(f"{today}{self.current_seq:03d}")

    def refresh_next_serial(self):
        today = datetime.now().strftime("%Y%m%d")
        base = self.base_dir.get()
        if not os.path.exists(base): return
        existing = [d for d in os.listdir(base) if d.startswith(today) and os.path.isdir(os.path.join(base, d))]
        if not existing:
            self.current_seq = 1
        else:
            try:
                seqs = [int(d[-3:]) for d in existing if d[-3:].isdigit()]
                self.current_seq = max(seqs) + 1 if seqs else 1
            except:
                self.current_seq = 1
        self.update_serial_display()

    def browse_base_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.base_dir.set(path)
            self.refresh_next_serial()

    def execute_create_serial(self):
        new_name = self.next_serial_var.get()
        target_path = os.path.join(self.base_dir.get(), new_name)
        if os.path.exists(target_path):
            if not messagebox.askyesno("提示", f"目录 {new_name} 已存在，是否使用？"):
                return
        try:
            if not os.path.exists(target_path):
                os.makedirs(target_path)
            self.active_serial_var.set(new_name)
            self.reset_all_data()
            self.current_seq += 1
            self.update_serial_display()
        except Exception as e:
            messagebox.showerror("系统错误", f"无法创建目录: {e}")

    def reset_all_data(self):
        self.in_days_var.set("0")
        for d in [self.data_out, self.data_in]:
            for cat in d:
                d[cat]["amt"].set("0.00");
                d[cat]["self"].set("0.00");
                d[cat]["calc"].set("")

    def create_section(self, parent, title, data_dict, order, entries_list, totals_var):
        frame = tk.LabelFrame(parent, text=title, padx=10, pady=10, font=("微软雅黑", 10, "bold"))
        frame.pack(fill='x', padx=15, pady=5)
        headers = ["诊疗项目", "票面金额", "自付金额", "实报金额", "辅助计算"]
        for c, text in enumerate(headers):
            tk.Label(frame, text=text, width=self.col_widths[c], relief="ridge", bg="#e0e0e0").grid(row=0, column=c,
                                                                                                    sticky='nsew')
        for i, cat in enumerate(order):
            r = i + 1
            tk.Label(frame, text=cat, width=self.col_widths[0], relief="groove", anchor='w', padx=5).grid(row=r,
                                                                                                          column=0,
                                                                                                          sticky='nsew')
            tk.Entry(frame, textvariable=data_dict[cat]["amt"], state='readonly', justify='right').grid(row=r, column=1,
                                                                                                        sticky='nsew')
            tk.Entry(frame, textvariable=data_dict[cat]["self"], justify='right', bg="#e8f5e9").grid(row=r, column=2,
                                                                                                     sticky='nsew')
            tk.Label(frame, textvariable=data_dict[cat]["refund"], relief="groove", anchor='e', padx=5,
                     fg="green").grid(row=r, column=3, sticky='nsew')
            ent = tk.Entry(frame, textvariable=data_dict[cat]["calc"], bg="#fffde7");
            ent.grid(row=r, column=4, sticky='nsew')
            entries_list.append(ent)
            ent.bind("<FocusOut>", lambda e, c=cat, d=data_dict: self.perform_single_calc(c, d))
            ent.bind("<Return>",
                     lambda e, c=cat, d=data_dict, l=entries_list, idx=i: self.handle_enter(e, c, d, l, idx))

        row_sum = len(order) + 1
        tk.Label(frame, text="该表合计", width=self.col_widths[0], relief="ridge", bg="#f5f5f5",
                 font=("微软雅黑", 9, "bold")).grid(row=row_sum, column=0, sticky='nsew')
        tk.Label(frame, textvariable=totals_var["amt"], relief="ridge", bg="#f5f5f5", anchor='e', padx=5).grid(
            row=row_sum, column=1, sticky='nsew')
        tk.Label(frame, textvariable=totals_var["self"], relief="ridge", bg="#f5f5f5", anchor='e', padx=5).grid(
            row=row_sum, column=2, sticky='nsew')
        tk.Label(frame, textvariable=totals_var["refund"], relief="ridge", bg="#f5f5f5", anchor='e', padx=5,
                 font=("微软雅黑", 9, "bold"), fg="blue").grid(row=row_sum, column=3, sticky='nsew')

    def handle_enter(self, event, cat, data_dict, entries_list, current_idx):
        self.perform_single_calc(cat, data_dict)
        next_idx = (current_idx + 1) % len(entries_list)
        entries_list[next_idx].focus_set()

    def perform_single_calc(self, cat, data_dict):
        calc_str = data_dict[cat]["calc"].get().strip()
        all_nums = re.findall(r'\+(\d+(?:\.\d+)?)', calc_str)
        total_sum = sum(float(n) for n in all_nums) if all_nums else 0.0
        data_dict[cat]["self"].set(f"{total_sum:.2f}")

    def bind_traces(self):
        for cat in self.data_out:
            self.data_out[cat]["amt"].trace_add("write", lambda *a: self.refresh(self.data_out, self.out_totals))
            self.data_out[cat]["self"].trace_add("write", lambda *a: self.refresh(self.data_out, self.out_totals))
        for cat in self.data_in:
            self.data_in[cat]["amt"].trace_add("write", lambda *a: self.refresh(self.data_in, self.in_totals))
            self.data_in[cat]["self"].trace_add("write", lambda *a: self.refresh(self.data_in, self.in_totals))

    def refresh(self, d, totals_var):
        s_a, s_s, s_r = 0.0, 0.0, 0.0
        for cat in d:
            try:
                a, s = float(d[cat]["amt"].get() or 0), float(d[cat]["self"].get() or 0)
                r = a - s
                d[cat]["refund"].set(f"{r:.2f}")
                s_a += a;
                s_s += s;
                s_r += r
            except:
                pass
        totals_var["amt"].set(f"{s_a:.2f}");
        totals_var["self"].set(f"{s_s:.2f}");
        totals_var["refund"].set(f"{s_r:.2f}")

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if not path: return
        try:
            df = pd.read_excel(path)
            res_out = {c: 0.0 for c in self.out_order};
            res_in = {c: 0.0 for c in self.in_order}
            df['f_key'] = df['发票代码'].fillna('A').astype(str) + df['发票号码'].fillna('B').astype(str)
            for _, group in df.groupby('f_key'):
                m_type = str(group['医疗类型'].iloc[0]).strip()
                is_in = (m_type == "住院");
                t_res = res_in if is_in else res_out
                all_names = group['货物或应税劳务名称'].astype(str).tolist()
                has_zhencha = any("诊察费" == n.strip() for n in all_names)
                inv_total = float(group['票面金额'].iloc[0]);
                known = 0.0
                for _, row in group.iterrows():
                    name, amt = str(row['货物或应税劳务名称']).strip(), float(row['金额'] or 0)
                    if amt <= 0: continue
                    matched = False
                    if has_zhencha:
                        if name == "诊察费": t_res["医事服务费"] += amt; known += amt; matched = True
                    else:
                        if name in ["医事服务费", "急诊诊察费"]: t_res[
                            "医事服务费"] += amt; known += amt; matched = True
                    if not matched:
                        for cat, keywords in self.mapping.items():
                            if any(k in name for k in keywords):
                                t_res[cat] += amt;
                                known += amt;
                                matched = True;
                                break
                t_res["其他费"] += (inv_total - known)
            for c in self.out_order: self.data_out[c]["amt"].set(f"{max(0, res_out[c]):.2f}")
            for c in self.in_order: self.data_in[c]["amt"].set(f"{max(0, res_in[c]):.2f}")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def generate_qr(self):
        self.root.focus_set()
        cur_serial = self.active_serial_var.get()
        if cur_serial == "未开始":
            messagebox.showwarning("提醒", "请先新建报销单。")
            return
        data = [cur_serial]
        for cat in self.out_order: data.extend([self.data_out[cat]["amt"].get(), self.data_out[cat]["self"].get()])
        data.append(self.in_days_var.get())
        for cat in self.in_order: data.extend([self.data_in[cat]["amt"].get(), self.data_in[cat]["self"].get()])
        qr_str = "\t".join(data);
        qr_win = tk.Toplevel(self.root)
        qr_gen = qrcode.QRCode(box_size=10, border=2);
        qr_gen.add_data(qr_str);
        qr_gen.make(fit=True)
        img = qr_gen.make_image(fill_color="black", back_color="white")
        self.tk_img = ImageTk.PhotoImage(img)
        lbl = tk.Label(qr_win, image=self.tk_img, padx=20, pady=20);
        lbl.pack()


if __name__ == "__main__":
    root = tk.Tk();
    app = MedicalAppV3(root);
    root.mainloop()