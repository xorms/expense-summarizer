# 导入必要的库
import pandas as pd  # 用于处理Excel和CSV数据
import tkinter as tk  # 用于创建图形用户界面
from tkinter import filedialog, messagebox  # 用于文件选择对话框和消息提示框
import os  # 用于操作系统功能，如文件路径和打开文件
from datetime import datetime  # 用于处理日期和时间
from fpdf import FPDF  # 用于生成PDF文件


class MedicalReportApp:
    def __init__(self, root):
        """
        初始化医疗报销应用程序
        :param root: tkinter主窗口对象
        """
        self.root = root
        self.root.title("医疗费用报销处理系统 - V3.0")  # 设置窗口标题
        self.root.geometry("850x600")  # 设置窗口大小（宽x高）

        # 1. 诊疗项目关键词定义：将Excel中的明细项目名称映射到标准分类
        self.project_mapping = {
            "医事服务费": ["医事服务费", "诊察费"],  # 关键词列表，用于识别分类
            "检查费": ["检查费"],
            "治疗费": ["治疗费"],
            "药费": ["西药费", "中成药费", "中草药费"],
            "手术费": ["手术费"],
            "卫生材料费": ["材料费", "卫生材料费"]
        }
        # 显示的类别列表：映射中的6个分类 + "其他项目"（用于未识别项目）
        self.display_cats = list(self.project_mapping.keys()) + ["其他项目"]

        # 创建tkinter变量用于存储金额数据
        self.amt_vars = {cat: tk.StringVar(value="0.00") for cat in self.display_cats}  # 票面金额变量
        self.self_vars = {cat: tk.StringVar(value="0.00") for cat in self.display_cats}  # 自付金额变量
        self.refund_labels = {}  # 存储实报金额标签的字典

        # --- 界面部分: 报销单基本信息录入区 ---
        # 创建一个带标签的框架，用于个人信息输入
        info_frame = tk.LabelFrame(root, text="报销单基本信息", padx=10, pady=10, font=("微软雅黑", 10, "bold"))
        info_frame.pack(fill="x", padx=20, pady=10)  # 填充X方向，设置内边距

        # 定义个人信息字段: (标签文本, 变量键名, 行, 列)
        fields = [
            ("姓名:", "name", 0, 0), ("报销日期:", "date", 0, 2), ("单位:", "unit", 0, 4),
            ("身份证号:", "id", 1, 0), ("手机号:", "phone", 1, 2), ("银行卡号:", "bank", 1, 4)
        ]
        self.info_entries = {}  # 存储所有输入框的字典

        # 创建标签和输入框
        for label, key, r, c in fields:
            tk.Label(info_frame, text=label).grid(row=r, column=c, sticky="e", pady=5)  # 创建标签，右对齐
            ent = tk.Entry(info_frame, width=20)  # 创建输入框
            ent.grid(row=r, column=c + 1, padx=5, sticky="w")  # 放置输入框，左对齐
            self.info_entries[key] = ent  # 保存输入框引用到字典

        # 设置报销日期为当前日期
        self.info_entries["date"].insert(0, datetime.now().strftime("%Y-%m-%d"))

        # --- 功能按钮区域 ---
        btn_frame = tk.Frame(root)  # 创建框架容器
        btn_frame.pack(pady=10)  # 放置框架
        # 创建"读取Excel"按钮
        tk.Button(btn_frame, text=" 1. 读取Excel并自动计算 ", command=self.load_excel, bg="#2196F3", fg="white",
                  font=("微软雅黑", 10, "bold")).pack(side="left", padx=10)
        # 文件加载状态标签
        self.lbl_file = tk.Label(root, text="请加载Excel文件", fg="gray")
        self.lbl_file.pack()

        # --- 明细表格区域 ---
        table_frame = tk.Frame(root, padx=20)  # 创建表格框架
        table_frame.pack(fill="both", expand=True)  # 填充并扩展

        # 创建表头
        headers = ["诊疗项目", "票面金额合计", "自付金额(输入)", "实报金额"]
        for c, text in enumerate(headers):
            tk.Label(table_frame, text=text, relief="ridge", bg="#e0e0e0", width=22, font=("微软雅黑", 9, "bold")).grid(
                row=0, column=c, sticky="nsew")  # 创建表头标签，带背景色和边框

        # 创建表格内容行（每个分类一行）
        for r, cat in enumerate(self.display_cats):
            # 诊疗项目名称列（只读）
            tk.Label(table_frame, text=cat, relief="groove", anchor="w", padx=10).grid(row=r + 1, column=0,
                                                                                       sticky="nsew")
            # 票面金额列（只读）
            tk.Entry(table_frame, textvariable=self.amt_vars[cat], state="readonly", justify="right").grid(row=r + 1,
                                                                                                           column=1,
                                                                                                           sticky="nsew")
            # 自付金额列（可编辑，淡黄色背景）
            ent_self = tk.Entry(table_frame, textvariable=self.self_vars[cat], justify="right", bg="#fffde7")
            ent_self.grid(row=r + 1, column=2, sticky="nsew")
            ent_self.bind("<KeyRelease>", self.update_all_totals)  # 绑定键盘释放事件，实时更新合计
            # 实报金额列（自动计算，蓝色文字）
            lbl_ref = tk.Label(table_frame, text="0.00", relief="groove", anchor="e", padx=10, fg="blue")
            lbl_ref.grid(row=r + 1, column=3, sticky="nsew")
            self.refund_labels[cat] = lbl_ref  # 保存标签引用到字典

        # --- 合计行 ---
        # 计算合计行的行索引（在最后一行）
        self.row_idx_total = len(self.display_cats) + 1

        # "总计"标签
        tk.Label(table_frame, text="总 计", relief="ridge", bg="#f5f5f5", font=("微软雅黑", 9, "bold")).grid(
            row=self.row_idx_total, column=0, sticky="nsew")
        # 票面金额合计显示
        self.lbl_total_amt = tk.Label(table_frame, text="0.00", relief="ridge", bg="#f5f5f5", anchor="e", padx=10)
        self.lbl_total_amt.grid(row=self.row_idx_total, column=1, sticky="nsew")
        # 自付金额合计显示
        self.lbl_total_self = tk.Label(table_frame, text="0.00", relief="ridge", bg="#f5f5f5", anchor="e", padx=10)
        self.lbl_total_self.grid(row=self.row_idx_total, column=2, sticky="nsew")
        # 实报金额合计显示（加粗）
        self.lbl_total_refund = tk.Label(table_frame, text="0.00", relief="ridge", bg="#f5f5f5", anchor="e", padx=10,
                                         font=("微软雅黑", 9, "bold"))
        self.lbl_total_refund.grid(row=self.row_idx_total, column=3, sticky="nsew")

        # --- 导出按钮 ---
        tk.Button(root, text=" 2. 保存并导出 PDF + CSV ", command=self.output_files, bg="#4CAF50", fg="white",
                  font=("微软雅黑", 12, "bold"), height=2, width=30).pack(pady=30)

    def load_excel(self):
        """
        加载Excel文件并自动解析和分类数据
        从Excel发票明细中提取信息，按项目分类汇总
        """
        # 打开文件选择对话框，限制只能选择Excel文件
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:  # 用户取消了选择
            return

        try:
            # 使用pandas读取Excel文件
            df = pd.read_excel(file_path)

            # 提取姓名：尝试从不同列名中查找姓名信息
            for col in ['购方名称', '交款人', '姓名']:
                if col in df.columns:
                    name = str(df[col].dropna().iloc[0]).strip()  # 获取第一个非空值
                    if name:  # 如果找到姓名
                        self.info_entries["name"].delete(0, tk.END)  # 清空姓名输入框
                        self.info_entries["name"].insert(0, name)  # 填入找到的姓名
                        break  # 找到姓名后跳出循环

            # 处理发票数据：按发票代码和发票号码分组（每张发票一个组）
            df['code_filler'] = df['发票代码'].fillna('N/A').astype(str)  # 填充空值，转换为字符串
            df['num_filler'] = df['发票号码'].fillna('N/A').astype(str)  # 填充空值，转换为字符串
            inv_groups = df.groupby(['code_filler', 'num_filler'])  # 按发票分组

            # 初始化各类别总金额字典
            grand_sums = {cat: 0.0 for cat in self.display_cats}

            # 遍历每张发票
            for _, group in inv_groups:
                # 获取该发票的总金额（票面金额）
                inv_total = float(group['票面金额'].iloc[0])
                # 初始化该发票的分类金额字典
                inv_cat_sums = {cat: 0.0 for cat in self.project_mapping}

                # 遍历发票中的每个明细项目
                for _, row in group.iterrows():
                    item_name = str(row['货物或应税劳务名称'])  # 项目名称
                    # 获取项目金额，如果'金额'列不存在则使用0
                    item_amt = float(row['金额'] if '金额' in row else 0)

                    if item_amt > 0:  # 只处理金额大于0的项目
                        # 根据项目名称关键词分类
                        for cat, keywords in self.project_mapping.items():
                            if any(k in item_name for k in keywords):  # 如果名称包含关键词
                                inv_cat_sums[cat] += item_amt  # 累加到对应分类
                                break  # 找到分类后跳出循环

                # 计算该发票中已识别的金额总和
                known_sum = sum(inv_cat_sums.values())

                # 累加到总金额字典
                for cat in self.project_mapping:
                    grand_sums[cat] += inv_cat_sums[cat]
                # 未识别的金额归入"其他项目"
                grand_sums["其他项目"] += (inv_total - known_sum)

            # 更新界面中的票面金额显示
            for cat in self.display_cats:
                self.amt_vars[cat].set(f"{max(0, grand_sums[cat]):.2f}")  # 确保非负数

            # 更新文件加载状态标签
            self.lbl_file.config(text=f"已加载: {os.path.basename(file_path)}", fg="green")

            # 更新所有合计金额
            self.update_all_totals()

        except Exception as e:
            # 异常处理：显示错误信息对话框
            messagebox.showerror("错误", str(e))

    def update_all_totals(self, event=None):
        """
        更新所有合计金额（票面金额合计、自付金额合计、实报金额合计）
        实时计算并更新表格底部的总计行
        :param event: 事件对象（可选，用于事件绑定）
        """
        s_amt, s_self, s_ref = 0.0, 0.0, 0.0  # 初始化总和变量

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
                pass  # 忽略转换错误（如果输入的不是有效数字）

        # 更新合计行显示
        self.lbl_total_amt.config(text=f"{s_amt:.2f}")
        self.lbl_total_self.config(text=f"{s_self:.2f}")
        self.lbl_total_refund.config(text=f"{s_ref:.2f}")

    def output_files(self):
        """
        保存CSV文件并生成PDF文件
        将当前数据导出为CSV和PDF格式
        """
        # 获取姓名（用于文件名），如果为空则使用"未命名"
        name = self.info_entries["name"].get() or "未命名"
        # 生成时间戳（用于文件名去重），格式：年月日_时分
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")

        # 1. 保存 CSV 文件
        try:
            # 准备CSV数据
            csv_data = {
                "项目": self.display_cats + ["总计"],  # 项目列：6个分类 + 总计行
                "票面合计": [self.amt_vars[c].get() for c in self.display_cats] + [self.lbl_total_amt.cget("text")],
                "自付金额": [self.self_vars[c].get() or "0.00" for c in self.display_cats] + [
                    self.lbl_total_self.cget("text")],
                "实报金额": [self.refund_labels[c].cget("text") for c in self.display_cats] + [
                    self.lbl_total_refund.cget("text")]
            }

            # 创建DataFrame并保存为CSV
            df_out = pd.DataFrame(csv_data)
            csv_name = f"报销单_{name}_{timestamp}.csv"  # 生成CSV文件名
            # 保存CSV，使用UTF-8-BOM编码（支持Excel中文显示）
            df_out.to_csv(csv_name, index=False, encoding="utf-8-sig")
        except Exception as e:
            messagebox.showerror("CSV保存失败", str(e))
            return  # CSV保存失败时不再继续生成PDF

        # 2. 生成 PDF 文件
        try:
            pdf = FPDF()  # 创建PDF对象
            pdf.add_page()  # 添加一个新页面

            # 设置中文字体（使用Windows系统自带的宋体）
            font_path = r"C:\Windows\Fonts\simsun.ttc"  # 宋体字体文件路径
            pdf.add_font("SimSun", style="", fname=font_path)  # 添加字体
            pdf.set_font("SimSun", size=14)  # 设置字体和大小

            # 标题
            pdf.cell(190, 8, "医药费报销明细表", ln=True, align="C")
            pdf.ln(2)  # 换行，留出间距

            # 个人信息第一行：姓名左对齐，日期右对齐
            pdf.set_font("SimSun", size=9)
            pdf.cell(95, 6, f"姓名: {self.info_entries['name'].get()}", ln=0, align="L")  # 左对齐，不换行
            pdf.cell(95, 6, f"报销日期: {self.info_entries['date'].get()}", ln=1, align="R")  # 右对齐，换行

            # 个人信息第二行：单位、身份证、手机、银行卡 紧凑排列
            info_line2 = (f"单位: {self.info_entries['unit'].get()} | "
                          f"身份证: {self.info_entries['id'].get()} | "
                          f"手机: {self.info_entries['phone'].get()} | "
                          f"银行卡: {self.info_entries['bank'].get()}")
            pdf.cell(190, 6, info_line2, ln=1, align="L")
            pdf.ln(2)  # 换行

            # 表格参数设置（高度由10缩减为7，控制总高）
            widths = [60, 40, 40, 50]  # 各列宽度
            header = ["诊疗项目", "票面合计", "自付金额", "实报金额"]  # 表头

            # 表头 - 无底色
            for i, head in enumerate(header):
                pdf.cell(widths[i], 7, head, border=1, align="C")  # 居中对齐
            pdf.ln()  # 换行

            # 表格内容行
            for cat in self.display_cats:
                pdf.cell(widths[0], 7, cat, border=1)  # 诊疗项目
                pdf.cell(widths[1], 7, self.amt_vars[cat].get(), border=1, align="R")  # 票面合计，右对齐
                pdf.cell(widths[2], 7, self.self_vars[cat].get() or "0.00", border=1, align="R")  # 自付金额，右对齐
                pdf.cell(widths[3], 7, self.refund_labels[cat].cget("text"), border=1, align="R")  # 实报金额，右对齐
                pdf.ln()  # 换行

            # 合计行 - 无底色
            pdf.set_font("SimSun", style="", size=9)
            pdf.cell(widths[0], 7, "总 计", border=1, align="C")  # 居中对齐
            pdf.cell(widths[1], 7, self.lbl_total_amt.cget("text"), border=1, align="R")
            pdf.cell(widths[2], 7, self.lbl_total_self.cget("text"), border=1, align="R")
            pdf.cell(widths[3], 7, self.lbl_total_refund.cget("text"), border=1, align="R")

            # 保存PDF文件
            pdf_name = f"报销单_{name}_{timestamp}.pdf"
            pdf.output(pdf_name)

            # 尝试自动打开生成的PDF文件
            try:
                os.startfile(pdf_name)  # 使用系统默认程序打开PDF
            except:
                # 如果os.startfile不可用（如在某些Linux系统），显示文件保存位置
                messagebox.showinfo("PDF已保存", f"PDF文件已保存到：\n{os.path.abspath(pdf_name)}")

            # 显示成功消息
            messagebox.showinfo("成功", f"文件已保存：\nCSV: {csv_name}\nPDF: {pdf_name}")
        except Exception as e:
            # PDF生成失败处理
            messagebox.showerror("PDF生成失败", f"错误：{str(e)}")


# 程序入口点
if __name__ == "__main__":
    root = tk.Tk()  # 创建tkinter主窗口
    app = MedicalReportApp(root)  # 创建应用程序实例
    root.mainloop()  # 启动GUI事件循环