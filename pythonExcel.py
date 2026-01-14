import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
from fpdf import FPDF


class MedicalReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("医疗费用报销处理系统 V3.0 (CSV+PDF版)")
        self.root.geometry("850x920")

        # 1. 诊疗项目关键词定义
        self.project_mapping = {
            "医事服务费": ["医事服务费", "诊察费"],
            "检查费": ["检查费"],
            "治疗费": ["治疗费"],
            "药费": ["西药费", "中成药费", "中草药费"],
            "手术费": ["手术费"],
            "卫生材料费": ["材料费", "卫生材料费"]
        }
        self.display_cats = list(self.project_mapping.keys()) + ["其他项目"]

        # 变量存储
        self.amt_vars = {cat: tk.StringVar(value="0.00") for cat in self.display_cats}
        self.self_vars = {cat: tk.StringVar(value="0.00") for cat in self.display_cats}
        self.refund_labels = {}

        # --- 第一部分：个人信息录入区 ---
        info_frame = tk.LabelFrame(root, text="报销单基本信息", padx=10, pady=10, font=("微软雅黑", 10, "bold"))
        info_frame.pack(fill="x", padx=20, pady=10)

        fields = [
            ("姓名:", "name", 0, 0), ("报销日期:", "date", 0, 2), ("单位:", "unit", 0, 4),
            ("身份证号:", "id", 1, 0), ("手机号:", "phone", 1, 2), ("银行卡号:", "bank", 1, 4)
        ]
        self.info_entries = {}
        for label, key, r, c in fields:
            tk.Label(info_frame, text=label).grid(row=r, column=c, sticky="e", pady=5)
            ent = tk.Entry(info_frame, width=20)
            ent.grid(row=r, column=c + 1, padx=5, sticky="w")
            self.info_entries[key] = ent
        self.info_entries["date"].insert(0, datetime.now().strftime("%Y-%m-%d"))

        # --- 第二部分：功能按钮 ---
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text=" 1. 读取Excel并自动计算 ", command=self.load_excel, bg="#2196F3", fg="white",
                  font=("微软雅黑", 10, "bold")).pack(side="left", padx=10)
        self.lbl_file = tk.Label(root, text="请加载Excel文件", fg="gray")
        self.lbl_file.pack()

        # --- 第三部分：明细表格 ---
        table_frame = tk.Frame(root, padx=20)
        table_frame.pack(fill="both", expand=True)
        headers = ["诊疗项目", "票面金额合计", "自付金额(输入)", "实报金额"]
        for c, text in enumerate(headers):
            tk.Label(table_frame, text=text, relief="ridge", bg="#e0e0e0", width=22, font=("微软雅黑", 9, "bold")).grid(
                row=0, column=c, sticky="nsew")

        for r, cat in enumerate(self.display_cats):
            tk.Label(table_frame, text=cat, relief="groove", anchor="w", padx=10).grid(row=r + 1, column=0,
                                                                                       sticky="nsew")
            tk.Entry(table_frame, textvariable=self.amt_vars[cat], state="readonly", justify="right").grid(row=r + 1,
                                                                                                           column=1,
                                                                                                           sticky="nsew")
            ent_self = tk.Entry(table_frame, textvariable=self.self_vars[cat], justify="right", bg="#fffde7")
            ent_self.grid(row=r + 1, column=2, sticky="nsew")
            ent_self.bind("<KeyRelease>", self.update_all_totals)
            lbl_ref = tk.Label(table_frame, text="0.00", relief="groove", anchor="e", padx=10, fg="blue")
            lbl_ref.grid(row=r + 1, column=3, sticky="nsew")
            self.refund_labels[cat] = lbl_ref

        # --- 第四部分：合计行 ---
        self.row_idx_total = len(self.display_cats) + 1
        tk.Label(table_frame, text="总 计", relief="ridge", bg="#f5f5f5", font=("微软雅黑", 9, "bold")).grid(
            row=self.row_idx_total, column=0, sticky="nsew")
        self.lbl_total_amt = tk.Label(table_frame, text="0.00", relief="ridge", bg="#f5f5f5", anchor="e", padx=10)
        self.lbl_total_amt.grid(row=self.row_idx_total, column=1, sticky="nsew")
        self.lbl_total_self = tk.Label(table_frame, text="0.00", relief="ridge", bg="#f5f5f5", anchor="e", padx=10)
        self.lbl_total_self.grid(row=self.row_idx_total, column=2, sticky="nsew")
        self.lbl_total_refund = tk.Label(table_frame, text="0.00", relief="ridge", bg="#f5f5f5", anchor="e", padx=10,
                                         font=("微软雅黑", 9, "bold"))
        self.lbl_total_refund.grid(row=self.row_idx_total, column=3, sticky="nsew")

        # --- 第五部分：输出 ---
        tk.Button(root, text=" 2. 保存并导出 PDF + CSV ", command=self.output_files, bg="#4CAF50", fg="white",
                  font=("微软雅黑", 12, "bold"), height=2, width=30).pack(pady=30)

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path: return
        try:
            df = pd.read_excel(file_path)
            # 提取姓名
            for col in ['购方名称', '交款人', '姓名']:
                if col in df.columns:
                    name = str(df[col].dropna().iloc[0]).strip()
                    if name:
                        self.info_entries["name"].delete(0, tk.END);
                        self.info_entries["name"].insert(0, name)
                        break

            # 按发票逻辑计算 (其他项目 = 票面总额 - 识别出的明细项)
            df['code_filler'] = df['发票代码'].fillna('N/A').astype(str)
            df['num_filler'] = df['发票号码'].fillna('N/A').astype(str)
            inv_groups = df.groupby(['code_filler', 'num_filler'])

            grand_sums = {cat: 0.0 for cat in self.display_cats}
            for _, group in inv_groups:
                inv_total = float(group['票面金额'].iloc[0])
                inv_cat_sums = {cat: 0.0 for cat in self.project_mapping}
                for _, row in group.iterrows():
                    item_name = str(row['货物或应税劳务名称'])
                    item_amt = float(row['金额'] if '金额' in row else 0)
                    if item_amt > 0:
                        for cat, keywords in self.project_mapping.items():
                            if any(k in item_name for k in keywords):
                                inv_cat_sums[cat] += item_amt
                                break
                known_sum = sum(inv_cat_sums.values())
                for cat in self.project_mapping: grand_sums[cat] += inv_cat_sums[cat]
                grand_sums["其他项目"] += (inv_total - known_sum)

            for cat in self.display_cats: self.amt_vars[cat].set(f"{max(0, grand_sums[cat]):.2f}")
            self.lbl_file.config(text=f"已加载: {os.path.basename(file_path)}", fg="green")
            self.update_all_totals()
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def update_all_totals(self, event=None):
        s_amt, s_self, s_ref = 0.0, 0.0, 0.0
        for cat in self.display_cats:
            try:
                a = float(self.amt_vars[cat].get() or 0)
                s = float(self.self_vars[cat].get() or 0)
                self.refund_labels[cat].config(text=f"{a - s:.2f}")
                s_amt += a;
                s_self += s;
                s_ref += (a - s)
            except:
                pass
        self.lbl_total_amt.config(text=f"{s_amt:.2f}");
        self.lbl_total_self.config(text=f"{s_self:.2f}");
        self.lbl_total_refund.config(text=f"{s_ref:.2f}")

    def output_files(self):
        name = self.info_entries["name"].get() or "未命名"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")

        # 1. 保存 CSV
        try:
            csv_data = {
                "项目": self.display_cats + ["总计"],
                "票面合计": [self.amt_vars[c].get() for c in self.display_cats] + [self.lbl_total_amt.cget("text")],
                "自付金额": [self.self_vars[c].get() or "0.00" for c in self.display_cats] + [
                    self.lbl_total_self.cget("text")],
                "实报金额": [self.refund_labels[c].cget("text") for c in self.display_cats] + [
                    self.lbl_total_refund.cget("text")]
            }
            # 记录个人信息
            info_log = {k: v.get() for k, v in self.info_entries.items()}
            df_out = pd.DataFrame(csv_data)
            csv_name = f"报销单_{name}_{timestamp}.csv"
            df_out.to_csv(csv_name, index=False, encoding="utf-8-sig")
        except Exception as e:
            messagebox.showerror("CSV保存失败", str(e))

        # 2. 生成 PDF
        try:
            pdf = FPDF()
            pdf.add_page()
            # 设置字体 (Windows 系统自带宋体路径)
            font_path = r"C:\Windows\Fonts\simsun.ttc"
            pdf.add_font("SimSun", style="", fname=font_path)
            pdf.set_font("SimSun", size=16)

            pdf.cell(190, 10, "医药费报销明细表", ln=True, align="C")
            pdf.set_font("SimSun", size=10)
            pdf.ln(5)

            # 个人信息栏
            info_text = f"姓名: {self.info_entries['name'].get():<15} 日期: {self.info_entries['date'].get():<15} 单位: {self.info_entries['unit'].get()}\n"
            info_text += f"身份证: {self.info_entries['id'].get():<25} 手机: {self.info_entries['phone'].get()}\n"
            info_text += f"银行卡: {self.info_entries['bank'].get()}"
            pdf.multi_cell(190, 8, info_text, border=0)
            pdf.ln(5)

            # 表格
            pdf.set_fill_color(220, 220, 220)
            cols = ["诊疗项目", "票面合计", "自付金额", "实报金额"]
            widths = [60, 40, 40, 50]
            for i, head in enumerate(cols): pdf.cell(widths[i], 10, head, border=1, align="C", fill=True)
            pdf.ln()

            for cat in self.display_cats:
                pdf.cell(widths[0], 10, cat, border=1)
                pdf.cell(widths[1], 10, self.amt_vars[cat].get(), border=1, align="R")
                pdf.cell(widths[2], 10, self.self_vars[cat].get() or "0.00", border=1, align="R")
                pdf.cell(widths[3], 10, self.refund_labels[cat].cget("text"), border=1, align="R")
                pdf.ln()

            # 合计行
            pdf.set_font("SimSun", style="", size=10)
            pdf.cell(widths[0], 10, "总 计", border=1, fill=True)
            pdf.cell(widths[1], 10, self.lbl_total_amt.cget("text"), border=1, align="R", fill=True)
            pdf.cell(widths[2], 10, self.lbl_total_self.cget("text"), border=1, align="R", fill=True)
            pdf.cell(widths[3], 10, self.lbl_total_refund.cget("text"), border=1, align="R", fill=True)

            pdf_name = f"报销单_{name}_{timestamp}.pdf"
            pdf.output(pdf_name)
            os.startfile(pdf_name)
            messagebox.showinfo("成功", f"文件已保存：\nCSV: {csv_name}\nPDF: {pdf_name}")
        except Exception as e:
            messagebox.showerror("PDF生成失败", f"提示：请确保电脑有宋体字体文件\n错误：{str(e)}")


if __name__ == "__main__":
    root = tk.Tk();
    app = MedicalReportApp(root);
    root.mainloop()