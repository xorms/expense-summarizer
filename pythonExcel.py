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
        self.root.title("医疗费报销系统 V3.4")
        self.root.geometry("850x700")

        # 1. 扫码数据结构定义
        self.out_order = ["医事服务费", "检查费", "治疗费", "西药", "中药", "卫生材料费", "其他费"]
        self.in_order = ["医事服务费", "检查费", "治疗费", "西药", "中药", "卫生材料费", "床位费", "其他费"]

        self.mapping = {
            "医事服务费": ["医事服务费", "诊察费"], "检查费": ["检查费", "化验费"],
            "治疗费": ["治疗费"], "西药": ["西药费"], "中药": ["中药饮片", "中草药", "中成药"],
            "卫生材料费": ["材料费", "卫生材料费"], "床位费": ["床位费", "空调费", "住院费", "住院"]
        }

        # 2. 变量定义
        self.data_out = self.init_struct(self.out_order)
        self.data_in = self.init_struct(self.in_order)
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
        self.create_section(self.scroll_frame, "【 门 诊 收 据 明 细 】", self.data_out, self.out_order,
                            self.out_calc_entries)

        day_f = tk.Frame(self.scroll_frame, pady=10);
        day_f.pack(fill='x', padx=20)
        tk.Label(day_f, text="住院天数：", font=("微软雅黑", 10, "bold")).pack(side='left')
        tk.Entry(day_f, textvariable=self.in_days_var, width=10, bg="#fffde7", justify='center').pack(side='left')

        self.create_section(self.scroll_frame, "【 住 院 收 据 明 细 】", self.data_in, self.in_order,
                            self.in_calc_entries)

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

    def create_section(self, parent, title, data_dict, order, entries_list):
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
        for d in [self.data_out, self.data_in]:
            for cat in d:
                d[cat]["amt"].trace_add("write", lambda *a, x=d: self.refresh(x))
                d[cat]["self"].trace_add("write", lambda *a, x=d: self.refresh(x))

    def refresh(self, d):
        for cat in d:
            try:
                a, s = float(d[cat]["amt"].get() or 0), float(d[cat]["self"].get() or 0)
                d[cat]["refund"].set(f"{a - s:.2f}")
            except:
                pass

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
                t_res, t_order = (res_in, self.in_order) if is_in else (res_out, self.out_order)
                inv_total = float(group['票面金额'].iloc[0])
                known = 0.0
                for _, row in group.iterrows():
                    name, amt = str(row['货物或应税劳务名称']), float(row['金额'] or 0)
                    for cat in t_order:
                        if cat != "其他费" and any(k in name for k in self.mapping.get(cat, [])):
                            t_res[cat] += amt;
                            known += amt;
                            break
                t_res["其他费"] += (inv_total - known)
            for c in self.out_order: self.data_out[c]["amt"].set(f"{max(0, res_out[c]):.2f}")
            for c in self.in_order: self.data_in[c]["amt"].set(f"{max(0, res_in[c]):.2f}")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def generate_qr(self):
        # 强制同步所有输入框数据
        self.root.focus_set()

        # 准备数据字符串
        data = [self.serial_folder.get() or "无流水号"]
        for cat in self.out_order: data.extend([self.data_out[cat]["amt"].get(), self.data_out[cat]["self"].get()])
        data.append(self.in_days_var.get())
        for cat in self.in_order: data.extend([self.data_in[cat]["amt"].get(), self.data_in[cat]["self"].get()])
        qr_str = "\t".join(data)

        # 创建二维码窗口
        qr_win = tk.Toplevel(self.root)
        qr_win.title("扫码录入")

        # 生成二维码图片
        qr_gen = qrcode.QRCode(box_size=10, border=2)
        qr_gen.add_data(qr_str)
        qr_gen.make(fit=True)
        img = qr_gen.make_image(fill_color="black", back_color="white")

        # 将 PIL Image 转换为 Tkinter 可用的 PhotoImage
        # 关键修正：必须显式保留这个对象的引用，防止被垃圾回收
        self.tk_img = ImageTk.PhotoImage(img)

        # 使用 Label 显示
        lbl = tk.Label(qr_win, image=self.tk_img, padx=20, pady=20)
        lbl.pack()

        # 提示文字
        tk.Label(qr_win, text=f"流水号：{data[0]}", font=("微软雅黑", 10, "bold"), fg="blue").pack(pady=5)
        tk.Label(qr_win, text="使用扫码枪扫描上方二维码进行录入", fg="gray").pack(pady=5)


if __name__ == "__main__":
    root = tk.Tk();
    app = MedicalAppV3(root);
    root.mainloop()