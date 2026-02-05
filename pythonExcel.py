import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import qrcode
import re
from PIL import ImageTk, Image

# 获取程序运行目录
if getattr(sys, 'frozen', False):
    APP_PATH = os.path.dirname(sys.executable)
else:
    APP_PATH = os.path.dirname(os.path.abspath(__file__))


class MedicalAppV3:
    def __init__(self, root):
        self.root = root
        self.root.title("医疗费报销系统 V3.1 (32项数据+文件名显示版)")
        self.root.geometry("800x630")

        # 1. 顺序定义 (扫码枪协议)
        self.out_order = ["医事服务费", "检查费", "治疗费", "西药", "中药", "卫生材料费", "其他费"]
        self.in_order = ["医事服务费", "检查费", "治疗费", "西药", "中药", "卫生材料费", "床位费", "其他费"]

        self.mapping = {
            "医事服务费": ["医事服务费", "诊察费"],
            "检查费": ["检查费", "化验费"],
            "治疗费": ["治疗费"],
            "西药": ["西药费"],
            "中药": ["中药饮片", "中草药", "中成药"],
            "卫生材料费": ["材料费", "卫生材料费"],
            "床位费": ["床位费", "空调费", "住院费", "住院"]
        }

        # 2. 变量与控件
        self.data_out = self.init_struct(self.out_order)
        self.data_in = self.init_struct(self.in_order)
        self.in_days_var = tk.StringVar(value="0")
        self.file_name_var = tk.StringVar(value="未加载文件")  # 存储文件名

        self.out_calc_entries = []
        self.in_calc_entries = []

        # 3. 布局
        self.setup_ui()
        self.bind_traces()

    def init_struct(self, order_list):
        return {cat: {"amt": tk.StringVar(value="0.00"),
                      "self": tk.StringVar(value="0.00"),
                      "refund": tk.StringVar(value="0.00"),
                      "calc": tk.StringVar(value="")} for cat in order_list}

    def setup_ui(self):
        # 顶部工具栏
        tool_frame = tk.Frame(self.root, pady=10, bg="#f5f5f5")
        tool_frame.pack(fill='x')

        tk.Button(tool_frame, text=" 1. 上传 Excel 数据表 ", command=self.load_excel,
                  bg="#2196F3", fg="white", font=("微软雅黑", 10, "bold"), width=25).pack(side='left', padx=20)

        tk.Button(tool_frame, text=" 2. 生成录入二维码 ", command=self.generate_qr,
                  bg="#FF9800", fg="white", font=("微软雅黑", 10, "bold"), width=25).pack(side='left', padx=10)

        # 底部状态栏（显示文件名）
        status_bar = tk.Frame(self.root, bd=1, relief=tk.SUNKEN, bg="#eeeeee")
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Label(status_bar, text="当前文件：", bg="#eeeeee", font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=5)
        tk.Label(status_bar, textvariable=self.file_name_var, bg="#eeeeee", fg="#d32f2f",
                 font=("微软雅黑", 9, "bold")).pack(side=tk.LEFT)

        # 滚动区域
        main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        self.scroll_frame = tk.Frame(main_canvas)
        self.scroll_frame.bind("<Configure>", lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
        main_canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.col_widths = [18, 15, 15, 15, 35]

        # 门诊部分
        self.create_section(self.scroll_frame, "【 门 诊 收 据 明 细 】", self.data_out, self.out_order,
                            self.out_calc_entries)

        # 中间住院天数
        day_f = tk.Frame(self.scroll_frame, pady=10)
        day_f.pack(fill='x', padx=20)
        tk.Label(day_f, text="住院天数：", font=("微软雅黑", 10, "bold")).pack(side='left')
        tk.Entry(day_f, textvariable=self.in_days_var, width=10, bg="#fffde7", justify='center').pack(side='left')

        # 住院部分
        self.create_section(self.scroll_frame, "【 住 院 收 据 明 细 】", self.data_in, self.in_order,
                            self.in_calc_entries)

    def create_section(self, parent, title, data_dict, order, entries_list):
        frame = tk.LabelFrame(parent, text=title, padx=10, pady=10, font=("微软雅黑", 10, "bold"))
        frame.pack(fill='x', padx=15, pady=5)

        headers = ["诊疗项目", "票面金额", "自付金额", "实报金额", "辅助计算 (输入+10+5后回车)"]
        for c, text in enumerate(headers):
            tk.Label(frame, text=text, width=self.col_widths[c], relief="ridge", bg="#e0e0e0",
                     font=("微软雅黑", 9)).grid(row=0, column=c, sticky='nsew')

        for i, cat in enumerate(order):
            r = i + 1
            tk.Label(frame, text=cat, width=self.col_widths[0], relief="groove", anchor='w', padx=5,
                     font=("微软雅黑", 9)).grid(row=r, column=0, sticky='nsew')
            tk.Entry(frame, textvariable=data_dict[cat]["amt"], width=self.col_widths[1], state='readonly',
                     justify='right', font=("微软雅黑", 9)).grid(row=r, column=1, sticky='nsew')
            tk.Entry(frame, textvariable=data_dict[cat]["self"], width=self.col_widths[2], justify='right',
                     bg="#e8f5e9", font=("微软雅黑", 9)).grid(row=r, column=2, sticky='nsew')
            tk.Label(frame, textvariable=data_dict[cat]["refund"], width=self.col_widths[3], relief="groove",
                     anchor='e', padx=5, fg="green", font=("微软雅黑", 9)).grid(row=r, column=3, sticky='nsew')

            ent = tk.Entry(frame, textvariable=data_dict[cat]["calc"], width=self.col_widths[4], bg="#fffde7",
                           font=("Consolas", 10))
            ent.grid(row=r, column=4, sticky='nsew')
            entries_list.append(ent)

            ent.bind("<Return>",
                     lambda e, c=cat, d=data_dict, l=entries_list, idx=i: self.handle_calc_enter(e, c, d, l, idx))

    def handle_calc_enter(self, event, cat, data_dict, entries_list, current_idx):
        calc_str = data_dict[cat]["calc"].get().strip()
        all_nums = re.findall(r'\+(\d+(?:\.\d+)?)', calc_str)
        if all_nums:
            total_sum = sum(float(n) for n in all_nums)
            data_dict[cat]["self"].set(f"{total_sum:.2f}")
        elif not calc_str:
            data_dict[cat]["self"].set("0.00")

        next_idx = (current_idx + 1) % len(entries_list)
        entries_list[next_idx].focus_set()
        entries_list[next_idx].icursor(tk.END)

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
            # 提取并显示文件名
            fname = os.path.basename(path)
            self.file_name_var.set(fname)

            df = pd.read_excel(path)
            res_out = {c: 0.0 for c in self.out_order}
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
            messagebox.showinfo("成功", f"文件 [{fname}] 已分类加载")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def generate_qr(self):
        # 1-14: 门诊金额与自付
        data = []
        for cat in self.out_order:
            data.extend([self.data_out[cat]["amt"].get(), self.data_out[cat]["self"].get()])

        # 15: 住院天数
        data.append(self.in_days_var.get())

        # 16-31: 住院金额与自付
        for cat in self.in_order:
            data.extend([self.data_in[cat]["amt"].get(), self.data_in[cat]["self"].get()])

        # 32: 增加文件名项
        data.append(self.file_name_var.get())

        qr_str = "\t".join(data)
        qr_win = tk.Toplevel(self.root);
        qr_win.title("扫码录入 (含文件名)")
        qr_gen = qrcode.QRCode(box_size=10, border=2)
        qr_gen.add_data(qr_str);
        qr_gen.make(fit=True)
        img = qr_gen.make_image(fill_color="black", back_color="white")
        path = os.path.join(APP_PATH, "temp_qr.png");
        img.save(path)
        photo = ImageTk.PhotoImage(Image.open(path))
        lbl = tk.Label(qr_win, image=photo, padx=20, pady=20);
        lbl.image = photo;
        lbl.pack()

        total_refund = sum(float(self.data_out[c]['refund'].get()) for c in self.out_order) + sum(
            float(self.data_in[c]['refund'].get()) for c in self.in_order)
        tk.Label(qr_win, text=f"包含 32 项数据 (第32项为文件名)\n实报合计：{total_refund:.2f}", fg="red",
                 font=("微软雅黑", 10, "bold")).pack()
        tk.Label(qr_win, text=f"文件名: {self.file_name_var.get()}", fg="gray").pack()


if __name__ == "__main__":
    root = tk.Tk();
    app = MedicalAppV3(root);
    root.mainloop()