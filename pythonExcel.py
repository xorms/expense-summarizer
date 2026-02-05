import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import qrcode  # 用于生成二维码
from PIL import ImageTk, Image  # 用于在界面显示二维码
from datetime import datetime

# 获取程序运行目录
if getattr(sys, 'frozen', False):
    APP_PATH = os.path.dirname(sys.executable)
else:
    APP_PATH = os.path.dirname(os.path.abspath(__file__))


class MedicalAppV3:
    def __init__(self, root):
        self.root = root
        self.root.title("医疗费报销系统 V3.0 (扫码录入版)")
        self.root.geometry("900x800")

        # 1. 诊疗项目顺序定义 (严格按照扫码需求排序)
        self.out_order = ["医事服务费", "检查费", "治疗费", "西药", "中药", "卫生材料费", "其他费"]
        self.in_order = ["医事服务费", "检查费", "治疗费", "西药", "中药", "卫生材料费", "床位费", "其他费"]

        # 关键词映射
        self.mapping = {
            "医事服务费": ["医事服务费", "诊察费"],
            "检查费": ["检查费", "化验费"],
            "治疗费": ["治疗费"],
            "西药": ["西药费"],
            "中药": ["中药饮片", "中草药", "中成药"],
            "卫生材料费": ["材料费", "卫生材料费"],
            "床位费": ["床位费", "空调费", "住院费", "住院"]
        }

        # 2. 变量初始化
        self.data_out = self.init_struct(self.out_order)
        self.data_in = self.init_struct(self.in_order)
        self.in_days_var = tk.StringVar(value="0")

        # 3. 界面布局
        self.setup_ui()
        self.bind_traces()

    def init_struct(self, order_list):
        return {cat: {"amt": tk.StringVar(value="0.00"),
                      "self": tk.StringVar(value="0.00"),
                      "refund": tk.StringVar(value="0.00")} for cat in order_list}

    def setup_ui(self):
        # 顶部工具栏
        tool_frame = tk.Frame(self.root, pady=10, bg="#f0f0f0")
        tool_frame.pack(fill='x')

        tk.Button(tool_frame, text=" 1. 上传 Excel (自动分类) ", command=self.load_excel,
                  bg="#2196F3", fg="white", font=("微软雅黑", 10, "bold"), width=25).pack(side='left', padx=20)

        tk.Button(tool_frame, text=" 2. 生成录入二维码 ", command=self.generate_qr,
                  bg="#FF9800", fg="white", font=("微软雅黑", 10, "bold"), width=25).pack(side='left', padx=10)

        # 主滚动窗口
        main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        self.scroll_frame = tk.Frame(main_canvas)
        self.scroll_frame.bind("<Configure>", lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
        main_canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 门诊表
        self.create_section(self.scroll_frame, "【 门 诊 收 据 】", self.data_out, self.out_order)

        # 住院天数中间条
        day_frame = tk.Frame(self.scroll_frame, pady=10)
        day_frame.pack(fill='x', padx=20)
        tk.Label(day_frame, text="住院天数：", font=("微软雅黑", 10, "bold")).pack(side='left')
        tk.Entry(day_frame, textvariable=self.in_days_var, width=10, bg="#fffde7", justify='center').pack(side='left')

        # 住院表
        self.create_section(self.scroll_frame, "【 住 院 收 据 】", self.data_in, self.in_order)

    def create_section(self, parent, title, data_dict, order):
        frame = tk.LabelFrame(parent, text=title, padx=10, pady=10, font=("微软雅黑", 10, "bold"))
        frame.pack(fill='x', padx=15, pady=5)

        headers = ["诊疗项目", "票面金额", "自付金额", "实报金额"]
        for c, text in enumerate(headers):
            tk.Label(frame, text=text, width=20, relief="ridge", bg="#e0e0e0").grid(row=0, column=c)

        for i, cat in enumerate(order):
            tk.Label(frame, text=cat, width=20, relief="groove", anchor='w').grid(row=i + 1, column=0, sticky='nsew')
            tk.Entry(frame, textvariable=data_dict[cat]["amt"], state='readonly', justify='right').grid(row=i + 1,
                                                                                                        column=1,
                                                                                                        sticky='nsew')
            tk.Entry(frame, textvariable=data_dict[cat]["self"], justify='right', bg="#fffde7").grid(row=i + 1,
                                                                                                     column=2,
                                                                                                     sticky='nsew')
            tk.Label(frame, textvariable=data_dict[cat]["refund"], relief="groove", anchor='e', fg="green").grid(
                row=i + 1, column=3, sticky='nsew')

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
            res_out = {c: 0.0 for c in self.out_order}
            res_in = {c: 0.0 for c in self.in_order}

            df['f_key'] = df['发票代码'].fillna('A').astype(str) + df['发票号码'].fillna('B').astype(str)

            for _, group in df.groupby('f_key'):
                m_type = str(group['医疗类型'].iloc[0]).strip()
                is_in = (m_type == "住院")
                target_res = res_in if is_in else res_out
                target_order = self.in_order if is_in else self.out_order

                inv_total = float(group['票面金额'].iloc[0])
                known = 0.0
                for _, row in group.iterrows():
                    name, amt = str(row['货物或应税劳务名称']), float(row['金额'] or 0)
                    for cat in target_order:
                        if cat != "其他费" and any(k in name for k in self.mapping.get(cat, [])):
                            target_res[cat] += amt
                            known += amt;
                            break
                target_res["其他费"] += (inv_total - known)

            for c in self.out_order: self.data_out[c]["amt"].set(f"{max(0, res_out[c]):.2f}")
            for c in self.in_order: self.data_in[c]["amt"].set(f"{max(0, res_in[c]):.2f}")
            messagebox.showinfo("成功", "数据已分类加载")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def generate_qr(self):
        """核心功能：按 31 项顺序生成 Tab 间隔的字符串二维码"""
        data_list = []

        # 1-14: 门诊 (7项目 * 2)
        for cat in self.out_order:
            data_list.append(self.data_out[cat]["amt"].get())
            data_list.append(self.data_out[cat]["self"].get())

        # 15: 住院天数
        data_list.append(self.in_days_var.get())

        # 16-31: 住院 (8项目 * 2)
        for cat in self.in_order:
            data_list.append(self.data_in[cat]["amt"].get())
            data_list.append(self.data_in[cat]["self"].get())

        # 用 Tab 拼接
        qr_string = "\t".join(data_list)

        # 生成二维码窗口
        qr_win = tk.Toplevel(self.root)
        qr_win.title("扫码录入 (请确保光标点在目标电脑首格)")

        qr = qrcode.QRCode(box_size=10, border=2)
        qr.add_data(qr_string)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")

        # 显示图片
        img_path = os.path.join(APP_PATH, "temp_qr.png")
        img.save(img_path)
        photo = ImageTk.PhotoImage(Image.open(img_path))

        lbl = tk.Label(qr_win, image=photo, padx=20, pady=20)
        lbl.image = photo
        lbl.pack()

        tk.Label(qr_win,
                 text="[扫码提示]\n1. 请在目标电脑打开报销系统，点中第一个输入框。\n2. 使用扫码枪扫描上方二维码。\n3. 数据将自动按 Tab 键顺序填入 31 个格子。",
                 fg="red", justify='left', pady=10).pack()


if __name__ == "__main__":
    root = tk.Tk();
    app = MedicalAppV3(root);
    root.mainloop()