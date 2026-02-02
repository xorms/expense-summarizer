import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys

# 获取程序运行目录
if getattr(sys, 'frozen', False):
    APP_PATH = os.path.dirname(sys.executable)
else:
    APP_PATH = os.path.dirname(os.path.abspath(__file__))

class MedicalAppV3:
    def __init__(self, root):
        self.root = root
        self.root.title("医疗费报销处理系统 V3.0")
        self.root.geometry("700x650")

        # 1. 诊疗项目关键词映射
        self.outpatient_mapping = {
            "医事服务费": ["医事服务费", "诊察费"],
            "检查费": ["检查费", "化验费"],
            "治疗费": ["治疗费"],
            "西药": ["西药费"],
            "中药": ["中药饮片", "中草药", "中成药"],
            "卫生材料费": ["材料费", "卫生材料费"]
        }
        self.inpatient_mapping = self.outpatient_mapping.copy()
        self.inpatient_mapping["床位费"] = ["床位费", "空调费", "住院费", "住院"]

        # 2. 数据变量
        self.data_out = self.init_struct(self.outpatient_mapping) # 门诊数据
        self.data_in = self.init_struct(self.inpatient_mapping)   # 住院数据
        
        # 3. 界面布局
        self.setup_ui()
        
        # 4. 绑定实时计算
        self.bind_traces()

    def init_struct(self, m):
        """ 初始化数据字典结构：项目名称 -> {变量} """
        return {c: {"amt": tk.StringVar(value="0.00"), 
                      "self": tk.StringVar(value="0.00"), 
                      "refund": tk.StringVar(value="0.00")} 
                for c in list(m.keys()) + ["其他费"]}

    def setup_ui(self):
        # 顶部上传区域
        top_frame = tk.Frame(self.root, pady=10)
        top_frame.pack(fill='x')
        tk.Button(top_frame, text=" 1. 上传 Excel 数据表 (自动分流门诊/住院) ", 
                  command=self.load_excel, bg="#2196F3", fg="white", font=("微软雅黑", 10, "bold")).pack()
        self.lbl_status = tk.Label(top_frame, text="等待加载文件...", fg="gray")
        self.lbl_status.pack()

        # 主滚动容器（防止屏幕高度不够）
        main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        self.scrollable_frame = tk.Frame(main_canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

        main_canvas.pack(side="left", fill="both", expand=True, padx=10)
        scrollbar.pack(side="right", fill="y")

        # --- 门诊收据区域 ---
        self.create_table_section(self.scrollable_frame, "【 门 诊 收 据 明 细 】", self.data_out)
        
        # 分隔空隙
        tk.Label(self.scrollable_frame, text="").pack()

        # --- 住院收据区域 ---
        self.create_table_section(self.scrollable_frame, "【 住 院 收 据 明 细 】", self.data_in)

    def create_table_section(self, parent, title, data_dict):
        """ 创建表格区域的通用函数 """
        frame = tk.LabelFrame(parent, text=title, padx=10, pady=10, font=("微软雅黑", 10, "bold"), fg="#2E7D32")
        frame.pack(fill='x', padx=5, pady=5)

        header_frame = tk.Frame(frame)
        header_frame.pack(fill='x')

        headers = ["诊疗项目", "票面金额汇总", "自付金额 (手工录入)", "实报金额 (计算)"]
        widths = [25, 20, 20, 20]
        
        for c, text in enumerate(headers):
            tk.Label(header_frame, text=text, width=widths[c], relief="ridge", bg="#e0e0e0").grid(row=0, column=c)

        for i, cat in enumerate(data_dict.keys()):
            row_idx = i + 1
            tk.Label(header_frame, text=cat, width=widths[0], relief="groove", anchor='w', padx=5).grid(row=row_idx, column=0, sticky='nsew')
            tk.Entry(header_frame, textvariable=data_dict[cat]["amt"], width=widths[1], state='readonly', justify='right').grid(row=row_idx, column=1, sticky='nsew')
            tk.Entry(header_frame, textvariable=data_dict[cat]["self"], width=widths[2], justify='right', bg="#fffde7").grid(row=row_idx, column=2, sticky='nsew')
            tk.Label(header_frame, textvariable=data_dict[cat]["refund"], width=widths[3], relief="groove", anchor='e', padx=5, fg="green").grid(row=row_idx, column=3, sticky='nsew')

    def bind_traces(self):
        """ 绑定所有自付金额输入框的实时计算监听 """
        for d in [self.data_out, self.data_in]:
            for cat in d:
                d[cat]["amt"].trace_add("write", lambda *a, x=d: self.refresh_calculations(x))
                d[cat]["self"].trace_add("write", lambda *a, x=d: self.refresh_calculations(x))

    def refresh_calculations(self, data_dict):
        """ 计算实报金额：实报 = 票面 - 自付 """
        for cat in data_dict:
            try:
                a = float(data_dict[cat]["amt"].get() or 0)
                s = float(data_dict[cat]["self"].get() or 0)
                data_dict[cat]["refund"].set(f"{a - s:.2f}")
            except:
                data_dict[cat]["refund"].set("0.00")

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path: return
        try:
            df = pd.read_excel(path)
            
            # 初始化临时汇总字典
            res_out = {c: 0.0 for c in self.data_out.keys()}
            res_in = {c: 0.0 for c in self.data_in.keys()}

            # 核心逻辑：按发票唯一标识分组
            df['code_key'] = df['发票代码'].fillna('N/A').astype(str)
            df['num_key'] = df['发票号码'].fillna('N/A').astype(str)
            
            for _, group in df.groupby(['code_key', 'num_key']):
                # 获取该发票的医疗类型（取第一行）
                m_type = str(group['医疗类型'].iloc[0]).strip()
                is_inpatient = (m_type == "住院")
                
                target_res = res_in if is_inpatient else res_out
                target_mapping = self.inpatient_mapping if is_inpatient else self.outpatient_mapping
                
                # 计算该票总额及已知明细
                inv_total = float(group['票面金额'].iloc[0])
                current_known = 0.0
                
                for _, row in group.iterrows():
                    item_name = str(row['货物或应税劳务名称'])
                    item_amt = float(row['金额'] if '金额' in row else 0)
                    
                    if item_amt > 0:
                        matched = False
                        for cat, keywords in target_mapping.items():
                            if any(k in item_name for k in keywords):
                                target_res[cat] += item_amt
                                current_known += item_amt
                                matched = True
                                break
                
                # 差额进入其他费
                target_res["其他费"] += (inv_total - current_known)

            # 更新到界面变量
            for cat in self.data_out: self.data_out[cat]["amt"].set(f"{max(0, res_out[cat]):.2f}")
            for cat in self.data_in: self.data_in[cat]["amt"].set(f"{max(0, res_in[cat]):.2f}")
            
            self.lbl_status.config(text=f"成功加载: {os.path.basename(path)}", fg="green")
            # messagebox.showinfo("完成", "数据已自动分流至门诊和住院列表。")
            
        except Exception as e:
            messagebox.showerror("读取失败", f"请检查Excel列名是否包含：\n医疗类型、票面金额、货物或应税劳务名称、发票代码、发票号码\n\n错误详情: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MedicalAppV3(root)
    root.mainloop()