# 导入必要的库
import pandas as pd  # 用于处理Excel/CSV数据
import tkinter as tk  # 用于创建图形用户界面
from tkinter import filedialog, messagebox  # 用于文件选择对话框和消息提示框
import os  # 用于操作系统功能，如文件路径
from datetime import datetime  # 用于处理日期和时间
from fpdf import FPDF  # 用于生成PDF文件


class MedicalReportApp:
    def __init__(self, root):
        """
        初始化应用程序界面和变量
        :param root: tkinter的主窗口对象
        """
        self.root = root
        self.root.title("医疗费用报销处理系统 V3.0 (CSV+PDF版)")  # 设置窗口标题
        self.root.geometry("850x920")  # 设置窗口大小

        # 1. 诊疗项目关键词定义：将Excel中的项目名称映射到六大类
        self.project_mapping = {
            "医事服务费": ["医事服务费", "诊察费"],  # 包含的关键词
            "检查费": ["检查费"],
            "治疗费": ["治疗费"],
            "药费": ["西药费", "中成药费", "中草药费"],
            "手术费": ["手术费"],
            "卫生材料费": ["材料费", "卫生材料费"]
        }
        # 显示的分类列表（增加"其他项目"用于未分类项）
        self.display_cats = list(self.project_mapping.keys()) + ["其他项目"]

        # 变量存储：创建tkinter字符串变量用于存储金额数据
        self.amt_vars = {cat: tk.StringVar(value="0.00") for cat in self.display_cats}  # 票面金额
        self.self_vars = {cat: tk.StringVar(value="0.00") for cat in self.display_cats}  # 自付金额
        self.refund_labels = {}  # 用于显示实报金额的标签字典

        # --- 第一部分：个人信息录入区 ---
        # 创建一个带标签的框架，用于个人信息输入
        info_frame = tk.LabelFrame(root, text="报销单基本信息", padx=10, pady=10, font=("微软雅黑", 10, "bold"))
        info_frame.pack(fill="x", padx=20, pady=10)  # 填充X方向，留出边距

        # 定义个人信息字段：标签名、变量键、行、列
        fields = [
            ("姓名:", "name", 0, 0), ("报销日期:", "date", 0, 2), ("单位:", "unit", 0, 4),
            ("身份证号:", "id", 1, 0), ("手机号:", "phone", 1, 2), ("银行卡号:", "bank", 1, 4)
        ]
        self.info_entries = {}  # 存储所有输入框的字典
        for label, key, r, c in fields:
            tk.Label(info_frame, text=label).grid(row=r, column=c, sticky="e", pady=5)  # 创建标签
            ent = tk.Entry(info_frame, width=20)  # 创建输入框
            ent.grid(row=r, column=c + 1, padx=5, sticky="w")  # 放置输入框
            self.info_entries[key] = ent  # 保存输入框引用
        # 设置报销日期为当前日期
        self.info_entries["date"].insert(0, datetime.now().strftime("%Y-%m-%d"))

        # --- 第二部分：功能按钮 ---
        btn_frame = tk.Frame(root)  # 创建按钮框架
        btn_frame.pack(pady=10)
        # 创建"读取Excel"按钮
        tk.Button(btn_frame, text=" 1. 读取Excel并自动计算 ", command=self.load_excel, bg="#2196F3", fg="white",
                  font=("微软雅黑", 10, "bold")).pack(side="left", padx=10)
        # 创建文件加载状态标签
        self.lbl_file = tk.Label(root, text="请加载Excel文件", fg="gray")
        self.lbl_file.pack()

        # --- 第三部分：明细表格 ---
        table_frame = tk.Frame(root, padx=20)  # 创建表格框架
        table_frame.pack(fill="both", expand=True)  # 填充并扩展
        # 创建表头
        headers = ["诊疗项目", "票面金额合计", "自付金额(输入)", "实报金额"]
        for c, text in enumerate(headers):
            tk.Label(table_frame, text=text, relief="ridge", bg="#e0e0e0", width=22, font=("微软雅黑", 9, "bold")).grid(
                row=0, column=c, sticky="nsew")

        # 创建表格内容行（每一行对应一个诊疗项目分类）
        for r, cat in enumerate(self.display_cats):
            # 诊疗项目名称列
            tk.Label(table_frame, text=cat, relief="groove", anchor="w", padx=10).grid(row=r + 1, column=0,
                                                                                       sticky="nsew")
            # 票面金额列（只读）
            tk.Entry(table_frame, textvariable=self.amt_vars[cat], state="readonly", justify="right").grid(row=r + 1,
                                                                                                           column=1,
                                                                                                           sticky="nsew")
            # 自付金额列（可编辑，黄色背景）
            ent_self = tk.Entry(table_frame, textvariable=self.self_vars[cat], justify="right", bg="#fffde7")
            ent_self.grid(row=r + 1, column=2, sticky="nsew")
            ent_self.bind("<KeyRelease>", self.update_all_totals)  # 绑定键盘释放事件，实时更新合计
            # 实报金额列（自动计算，蓝色文字）
            lbl_ref = tk.Label(table_frame, text="0.00", relief="groove", anchor="e", padx=10, fg="blue")
            lbl_ref.grid(row=r + 1, column=3, sticky="nsew")
            self.refund_labels[cat] = lbl_ref  # 保存标签引用

        # --- 第四部分：合计行 ---
        # 计算合计行的行索引（在最后一行）
        self.row_idx_total = len(self.display_cats) + 1
        # "总计"标签
        tk.Label(table_frame, text="总 计", relief="ridge", bg="#f5f5f5", font=("微软雅黑", 9, "bold")).grid(
            row=self.row_idx_total, column=0, sticky="nsew")
        # 票面金额合计
        self.lbl_total_amt = tk.Label(table_frame, text="0.00", relief="ridge", bg="#f5f5f5", anchor="e", padx=10)
        self.lbl_total_amt.grid(row=self.row_idx_total, column=1, sticky="nsew")
        # 自付金额合计
        self.lbl_total_self = tk.Label(table_frame, text="0.00", relief="ridge", bg="#f5f5f5", anchor="e", padx=10)
        self.lbl_total_self.grid(row=self.row_idx_total, column=2, sticky="nsew")
        # 实报金额合计（加粗显示）
        self.lbl_total_refund = tk.Label(table_frame, text="0.00", relief="ridge", bg="#f5f5f5", anchor="e", padx=10,
                                         font=("微软雅黑", 9, "bold"))
        self.lbl_total_refund.grid(row=self.row_idx_total, column=3, sticky="nsew")

        # --- 第五部分：输出按钮 ---
        tk.Button(root, text=" 2. 保存并导出 PDF + CSV ", command=self.output_files, bg="#4CAF50", fg="white",
                  font=("微软雅黑", 12, "bold"), height=2, width=30).pack(pady=30)

    def load_excel(self):
        """加载Excel文件并自动解析数据"""
        # 打开文件选择对话框，仅显示Excel文件
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:  # 如果用户取消了选择，直接返回
            return

        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)

            # 提取姓名：尝试从不同列名中查找姓名信息
            for col in ['购方名称', '交款人', '姓名']:
                if col in df.columns:
                    # 获取第一个非空值作为姓名
                    name = str(df[col].dropna().iloc[0]).strip()
                    if name:
                        # 清空姓名输入框并填入找到的姓名
                        self.info_entries["name"].delete(0, tk.END)
                        self.info_entries["name"].insert(0, name)
                        break

            # 按发票逻辑计算分类金额
            # 处理发票代码和发票号码中的空值
            df['code_filler'] = df['发票代码'].fillna('N/A').astype(str)
            df['num_filler'] = df['发票号码'].fillna('N/A').astype(str)

            # 按发票代码和发票号码分组（每张发票一个组）
            inv_groups = df.groupby(['code_filler', 'num_filler'])

            # 初始化各类别总金额字典
            grand_sums = {cat: 0.0 for cat in self.display_cats}

            # 遍历每张发票
            for _, group in inv_groups:
                # 获取该发票的总金额
                inv_total = float(group['票面金额'].iloc[0])
                # 初始化该发票的分类金额
                inv_cat_sums = {cat: 0.0 for cat in self.project_mapping}

                # 遍历发票中的每个明细项目
                for _, row in group.iterrows():
                    item_name = str(row['货物或应税劳务名称'])  # 项目名称
                    item_amt = float(row['金额'] if '金额' in row else 0)  # 项目金额

                    if item_amt > 0:
                        # 根据项目名称关键词分类
                        for cat, keywords in self.project_mapping.items():
                            if any(k in item_name for k in keywords):
                                inv_cat_sums[cat] += item_amt
                                break

                # 计算该发票中已识别的金额总和
                known_sum = sum(inv_cat_sums.values())

                # 累加到总金额字典
                for cat in self.project_mapping:
                    grand_sums[cat] += inv_cat_sums[cat]
                # 未识别的金额归入"其他项目"
                grand_sums["其他项目"] += (inv_total - known_sum)

            # 更新界面中的票面金额显示
            for cat in self.display_cats:
                self.amt_vars[cat].set(f"{max(0, grand_sums[cat]):.2f}")

            # 更新文件加载状态
            self.lbl_file.config(text=f"已加载: {os.path.basename(file_path)}", fg="green")

            # 更新所有合计金额
            self.update_all_totals()
        except Exception as e:
            # 异常处理：显示错误信息
            messagebox.showerror("错误", str(e))

    def update_all_totals(self, event=None):
        """
        更新所有合计金额（票面、自付、实报）
        :param event: 事件对象（可选，用于事件绑定）
        """
        s_amt, s_self, s_ref = 0.0, 0.0, 0.0  # 初始化总和

        # 遍历所有分类
        for cat in self.display_cats:
            try:
                # 获取票面金额和自付金额
                a = float(self.amt_vars[cat].get() or 0)
                s = float(self.self_vars[cat].get() or 0)

                # 计算实报金额 = 票面金额 - 自付金额
                self.refund_labels[cat].config(text=f"{a - s:.2f}")

                # 累加到总和
                s_amt += a
                s_self += s
                s_ref += (a - s)
            except:
                pass  # 忽略转换错误

        # 更新合计行显示
        self.lbl_total_amt.config(text=f"{s_amt:.2f}")
        self.lbl_total_self.config(text=f"{s_self:.2f}")
        self.lbl_total_refund.config(text=f"{s_ref:.2f}")

    def output_files(self):
        """保存CSV文件并生成PDF文件"""
        # 获取姓名（用于文件名）
        name = self.info_entries["name"].get() or "未命名"
        # 生成时间戳（用于文件名去重）
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")

        # 1. 保存 CSV 文件
        try:
            # 准备CSV数据
            csv_data = {
                "项目": self.display_cats + ["总计"],  # 第一列：项目名称
                "票面合计": [self.amt_vars[c].get() for c in self.display_cats] + [self.lbl_total_amt.cget("text")],
                "自付金额": [self.self_vars[c].get() or "0.00" for c in self.display_cats] + [
                    self.lbl_total_self.cget("text")],
                "实报金额": [self.refund_labels[c].cget("text") for c in self.display_cats] + [
                    self.lbl_total_refund.cget("text")]
            }

            # 记录个人信息（当前未使用）
            info_log = {k: v.get() for k, v in self.info_entries.items()}

            # 创建DataFrame并保存为CSV
            df_out = pd.DataFrame(csv_data)
            csv_name = f"报销单_{name}_{timestamp}.csv"  # 生成CSV文件名
            df_out.to_csv(csv_name, index=False, encoding="utf-8-sig")  # 保存CSV，使用UTF-8-BOM编码支持Excel中文
        except Exception as e:
            messagebox.showerror("CSV保存失败", str(e))
            return  # CSV保存失败时不再继续生成PDF

        # 2. 生成 PDF 文件
        try:
            pdf = FPDF()  # 创建PDF对象
            pdf.add_page()  # 添加页面

            # 设置中文字体（使用Windows系统自带的宋体）
            font_path = r"C:\Windows\Fonts\simsun.ttc"  # 宋体字体路径
            pdf.add_font("SimSun", style="", fname=font_path)  # 添加字体
            pdf.set_font("SimSun", size=16)  # 设置字体和大小

            # 添加标题
            pdf.cell(190, 10, "医药费报销明细表", ln=True, align="C")
            pdf.set_font("SimSun", size=10)
            pdf.ln(5)  # 换行，留出间距

            # 添加个人信息
            info_text = f"姓名: {self.info_entries['name'].get():<15} 日期: {self.info_entries['date'].get():<15} 单位: {self.info_entries['unit'].get()}\n"
            info_text += f"身份证: {self.info_entries['id'].get():<25} 手机: {self.info_entries['phone'].get()}\n"
            info_text += f"银行卡: {self.info_entries['bank'].get()}"
            pdf.multi_cell(190, 8, info_text, border=0)  # 多行文本单元格
            pdf.ln(5)  # 换行

            # 创建表格
            pdf.set_fill_color(220, 220, 220)  # 设置表头背景色（灰色）

            # 定义列标题和宽度
            cols = ["诊疗项目", "票面合计", "自付金额", "实报金额"]
            widths = [60, 40, 40, 50]

            # 添加表头
            for i, head in enumerate(cols):
                pdf.cell(widths[i], 10, head, border=1, align="C", fill=True)
            pdf.ln()  # 换行

            # 添加表格数据行
            for cat in self.display_cats:
                pdf.cell(widths[0], 10, cat, border=1)  # 诊疗项目
                pdf.cell(widths[1], 10, self.amt_vars[cat].get(), border=1, align="R")  # 票面合计，右对齐
                pdf.cell(widths[2], 10, self.self_vars[cat].get() or "0.00", border=1, align="R")  # 自付金额，右对齐
                pdf.cell(widths[3], 10, self.refund_labels[cat].cget("text"), border=1, align="R")  # 实报金额，右对齐
                pdf.ln()  # 换行

            # 添加合计行（灰色背景）
            pdf.set_font("SimSun", style="", size=10)
            pdf.cell(widths[0], 10, "总 计", border=1, fill=True)  # 填充灰色背景
            pdf.cell(widths[1], 10, self.lbl_total_amt.cget("text"), border=1, align="R", fill=True)
            pdf.cell(widths[2], 10, self.lbl_total_self.cget("text"), border=1, align="R", fill=True)
            pdf.cell(widths[3], 10, self.lbl_total_refund.cget("text"), border=1, align="R", fill=True)

            # 保存PDF文件
            pdf_name = f"报销单_{name}_{timestamp}.pdf"
            pdf.output(pdf_name)

            # 自动打开生成的PDF文件
            os.startfile(pdf_name)

            # 显示成功消息
            messagebox.showinfo("成功", f"文件已保存：\nCSV: {csv_name}\nPDF: {pdf_name}")
        except Exception as e:
            # PDF生成失败处理
            messagebox.showerror("PDF生成失败", f"提示：请确保电脑有宋体字体文件\n错误：{str(e)}")


# 程序入口
if __name__ == "__main__":
    root = tk.Tk()  # 创建主窗口
    app = MedicalReportApp(root)  # 创建应用程序实例
    root.mainloop()  # 启动事件循环