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
        self.root.title("医疗费报销系统 V3.6 (带明细汇总版)")
        self.root.geometry("850x700")

        # 1. 扫码数据结构定义 (32项协议，不含汇总行)
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

        # 新增：用于存储两个表格底部的合计变量
        self.out_totals = {k: tk.StringVar(value="0.00") for k in ["amt", "self", "refund"]}
        self.in_totals = {k: tk.StringVar(value="0.00") for k in ["amt", "self", "refund"]}

        self.in_days_var = tk.StringVar(value="0")
        self.base_dir = tk.StringVar(value=APP_PATH)
        self.serial_folder = tk.StringVar(value="")

        self.out_calc_entries = []
        self.in_calc_entries = []

        # 3. UI布局
        self.setup_ui()
        self.bind_traces()

    def init_struct(self, order_list):
        return {cat: {"amt": tk.StringVar(value="0.00"), "self": tk.StringVar(value="0.00"),
                      "refund": tk.StringVar(value="0.00"), "calc": tk.StringVar(value="")} for cat in order_list}

    def setup_ui(self):
        # 顶部管理区
        manage_frame = tk.Frame(self.root, pady=10, bg="#e3f2fd")
        manage_frame.pack(fill='x')
        tk.Label(manage_frame, text="存档根目录:", bg="#e3f2fd").pack(side='left', padx=(20, 5))
        tk.Entry(manage_frame, textvariable=self.base_dir, width=40).pack(side='left', padx=5)
        tk.Button(manage_frame, text="选择路径", command=self.browse_base_dir).pack(side='left', padx=5)
        tk.Button(manage_frame, text=" 新建报销单 (重置数据) ", command=self.create_new_serial,
                  bg="#1976D2", fg="white", font=("微软雅黑", 9, "bold")).pack(side='left', padx=20)
        tk.Label(manage_frame, textvariable=self.serial_folder, fg="red", font=("微软雅黑", 10, "bold"),
                 bg="#e3f2fd").pack(side='left')

        # 工具栏
        tool_frame = tk.Frame(self.root, pady=10, bg="#f5f5f5")
        tool_frame.pack(fill='x')
        tk.Button(tool_frame, text=" 1. 上传 Excel 数据表 ", command=self.load_excel,
                  bg="#2196F3", fg="white", font=("微软雅黑", 10, "bold"), width=25).pack(side='left', padx=20)
        tk.Button(tool_frame, text=" 2. 生成录入二维码 ", command=self.generate_qr,
                  bg="#FF9800", fg="white", font=("微软雅黑", 10, "bold"), width=25).pack(side='left', padx=10)

        # 滚动区
        main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        self.scroll_frame = tk.Frame(main_canvas)
        self.scroll_frame.bind("<Configure>", lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
        main_canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        main_canvas.pack(side="left", fill="both", expand=True);
        scrollbar.pack(side="right", fill="y")

        self.col_widths = [18, 15, 15, 15, 35]

        # 门诊部分
        self.create_section(self.scroll_frame, "【 门 诊 收 据 明 细 】", self.data_out, self.out_order,
                            self.out_calc_entries, self.out_totals)

        # 住院天数
        day_f = tk.Frame(self.scroll_frame, pady=10);
        day_f.pack(fill='x', padx=20)
        tk.Label(day_f, text="住院天数：", font=("微软雅黑", 10, "bold")).pack(side='left')
        tk.Entry(day_f, textvariable=self.in_days_var, width=10, bg="#fffde7", justify='center').pack(side='left')

        # 住院部分
        self.create_section(self.scroll_frame, "【 住 院 收 据 明 细 】", self.data_in, self.in_order,
                            self.in_calc_entries, self.in_totals)

    def browse_base_dir(self):
        path = filedialog.askdirectory()
        if path: self.base_dir.set(path)

    def create_new_serial(self):
        today = datetime.now().strftime("%Y%m%d")
        base = self.base_dir.get()
        if not os.path.exists(base): return
        existing = [d for d in os.listdir(base) if d.startswith(today) and os.path.isdir(os.path.join(base, d))]
        new_name = f"{today}{len(existing) + 1:03d}"
        try:
            os.makedirs(os.path.join(base, new_name))
            self.serial_folder.set(new_name);
            self.reset_all_data()
            messagebox.showinfo("成功", f"新单已创建：{new_name}")
        except Exception as e:
            messagebox.showerror("失败", str(e))

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

        # 表头
        headers = ["诊疗项目", "票面金额", "自付金额", "实报金额", "辅助计算"]
        for c, text in enumerate(headers):
            tk.Label(frame, text=text, width=self.col_widths[c], relief="ridge", bg="#e0e0e0").grid(row=0, column=c,
                                                                                                    sticky='nsew')

        # 数据行
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

        # --- 新增：表格底部合计行 ---
        row_sum = len(order) + 1
        tk.Label(frame, text="该表合计", width=self.col_widths[0], relief="ridge", bg="#f5f5f5",
                 font=("微软雅黑", 9, "bold")).grid(row=row_sum, column=0, sticky='nsew')
        tk.Label(frame, textvariable=totals_var["amt"], width=self.col_widths[1], relief="ridge", bg="#f5f5f5",
                 anchor='e', padx=5).grid(row=row_sum, column=1, sticky='nsew')
        tk.Label(frame, textvariable=totals_var["self"], width=self.col_widths[2], relief="ridge", bg="#f5f5f5",
                 anchor='e', padx=5).grid(row=row_sum, column=2, sticky='nsew')
        tk.Label(frame, textvariable=totals_var["refund"], width=self.col_widths[3], relief="ridge", bg="#f5f5f5",
                 anchor='e', padx=5, font=("微软雅黑", 9, "bold"), fg="blue").grid(row=row_sum, column=3, sticky='nsew')

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
        # 监听门诊数据
        for cat in self.data_out:
            self.data_out[cat]["amt"].trace_add("write", lambda *a: self.refresh(self.data_out, self.out_totals))
            self.data_out[cat]["self"].trace_add("write", lambda *a: self.refresh(self.data_out, self.out_totals))
        # 监听住院数据
        for cat in self.data_in:
            self.data_in[cat]["amt"].trace_add("write", lambda *a: self.refresh(self.data_in, self.in_totals))
            self.data_in[cat]["self"].trace_add("write", lambda *a: self.refresh(self.data_in, self.in_totals))

    def refresh(self, d, totals_var):
        """刷新单行实报金额并更新该表合计值"""
        s_a, s_s, s_r = 0.0, 0.0, 0.0
        for cat in d:
            try:
                a = float(d[cat]["amt"].get() or 0)
                s = float(d[cat]["self"].get() or 0)
                r = a - s
                d[cat]["refund"].set(f"{r:.2f}")
                s_a += a;
                s_s += s;
                s_r += r
            except:
                pass
        totals_var["amt"].set(f"{s_a:.2f}")
        totals_var["self"].set(f"{s_s:.2f}")
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
                is_in = (m_type == "住院")
                t_res = res_in if is_in else res_out

                all_names = group['货物或应税劳务名称'].astype(str).tolist()
                has_zhencha = any("诊察费" == n.strip() for n in all_names)
                inv_total = float(group['票面金额'].iloc[0])
                known = 0.0

                for _, row in group.iterrows():
                    name = str(row['货物或应税劳务名称']).strip()
                    amt = float(row['金额'] or 0)
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
            messagebox.showinfo("成功", "数据已按新规则加载")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def generate_qr(self):
        self.root.focus_set()
        # 严格按照 32 项协议构造数据字符串
        data = [self.serial_folder.get() or "无流水号"]
        # 门诊 14 项 (7项目 * 2 [票面, 自付])
        for cat in self.out_order: data.extend([self.data_out[cat]["amt"].get(), self.data_out[cat]["self"].get()])
        # 住院天数 1 项
        data.append(self.in_days_var.get())
        # 住院 16 项 (8项目 * 2 [票面, 自付])
        for cat in self.in_order: data.extend([self.data_in[cat]["amt"].get(), self.data_in[cat]["self"].get()])

        qr_str = "\t".join(data)
        qr_win = tk.Toplevel(self.root);
        qr_win.title("扫码录入 (32项)")
        qr_gen = qrcode.QRCode(box_size=10, border=2);
        qr_gen.add_data(qr_str);
        qr_gen.make(fit=True)
        img = qr_gen.make_image(fill_color="black", back_color="white")
        self.tk_img = ImageTk.PhotoImage(img)
        lbl = tk.Label(qr_win, image=self.tk_img, padx=20, pady=20);
        lbl.pack()
        tk.Label(qr_win, text=f"流水号：{data[0]}", font=("微软雅黑", 10, "bold"), fg="blue").pack()


if __name__ == "__main__":
    root = tk.Tk();
    app = MedicalAppV3(root);
    root.mainloop()