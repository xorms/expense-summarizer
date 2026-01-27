import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import re  # 导入正则表达式模块，用于校验身份证格式
from datetime import datetime
from fpdf import FPDF

# --- 全局路径配置 / Global Path Configuration ---
# 确保在 Win11 或打包成 EXE 后，文件始终保存在程序所在目录
if getattr(sys, 'frozen', False):
    # 打包后的环境 / Bundled EXE environment
    APP_PATH = os.path.dirname(sys.executable)
else:
    # 脚本运行环境 / Script environment
    APP_PATH = os.path.dirname(os.path.abspath(__file__))


def cn_currency(value):
    """
    将阿拉伯数字金额转换为中文大写 (Financial Chinese Currency Formatter)
    例如: 123.45 -> 壹佰贰拾叁元肆角伍分
    """
    units = ["", "拾", "佰", "仟", "万", "拾", "佰", "仟", "亿"]
    digits = "零壹贰叁肆伍陆柒捌玖"
    try:
        s_val = f"{float(value):.2f}".replace(".", "")
        if float(value) <= 0: return "零元整"
        res = ""
        for i, d in enumerate(s_val[::-1]):
            if i == 0:  # 分位 / Fen
                if d != '0':
                    res = f"{digits[int(d)]}分" + res
                else:
                    res = "整"
            elif i == 1:  # 角位 / Jiao
                if d != '0':
                    res = f"{digits[int(d)]}角" + res
                elif res != "整":
                    res = "零" + res
            elif i == 2:  # 元位 / Yuan
                res = "元" + res
                res = digits[int(d)] + res
            else:  # 拾佰仟万位 / Ten, Hundred, Thousand...
                if d != '0':
                    res = digits[int(d)] + units[i - 2] + res
                elif not res.startswith("零"):
                    res = "零" + res
        return res.replace("零元", "元").replace("零零", "零").strip("零")
    except:
        return "零元整"


class MedicalApp:
    def __init__(self, root):
        """
        初始化主界面及数据变量 (Initialization of GUI and Variables)
        """
        self.root = root
        self.root.title("医疗费报销系统 V2.0")
        self.root.geometry("850x400")  # 窗口默认大小

        # 1. 定义分类关键词映射 (Keyword Mapping for Excel Categorization)
        self.outpatient_mapping = {
            "医事服务费": ["医事服务费", "诊察费"],
            "检查费": ["检查费", "化验费"],
            "治疗费": ["治疗费"],
            "西药": ["西药费"],
            "中药": ["中药饮片", "中草药", "中成药"],
            "卫生材料费": ["材料费", "卫生材料费"]
        }
        # 住院类别继承门诊并添加床位费 / Inpatient inherits outpatient and adds bed fee
        self.inpatient_mapping = self.outpatient_mapping.copy()
        self.inpatient_mapping["床位费"] = ["床位费", "空调费", "住院费", "住院"]

        # 汇总单显示的固定顺序 / Display order in Summary Tab
        self.summary_cats = ["医事服务费", "检查费", "治疗费", "西药", "中药", "卫生材料费", "床位费", "其他费"]

        # 2. 定义界面变量 (GUI Variables Binding)
        self.info_vars = {k: tk.StringVar() for k in ["name", "id", "bank", "age", "unit", "type", "phone", "date"]}
        self.info_vars["date"].set(datetime.now().strftime("%Y-%m-%d"))  # 默认当天日期
        self.in_days_var = tk.StringVar(value="0")  # 住院天数

        # 初始化门诊、住院及汇总的数值存储结构
        self.data_out = self.init_struct(self.outpatient_mapping)
        self.data_in = self.init_struct(self.inpatient_mapping)
        self.out_totals = {k: tk.StringVar(value="0.00") for k in ["amt", "self", "refund"]}
        self.in_totals = {k: tk.StringVar(value="0.00") for k in ["amt", "self", "refund"]}
        self.sum_amt_vars = {cat: tk.StringVar(value="0.00") for cat in self.summary_cats}

        # 3. 构建多标签页界面 (Notebook Tabs Setup)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)

        self.setup_tab1()  # 汇总单 Tab
        self.setup_detail_tab("门诊收据", self.data_out, self.outpatient_mapping, self.out_totals)
        self.setup_detail_tab("住院收据", self.data_in, self.inpatient_mapping, self.in_totals, True)

        # 4. 绑定变量监听 (Trace variables to update totals automatically)
        self.bind_traces()

    def init_struct(self, m):
        """ 初始化每个项目的 票面、自付、实报 变量 """
        return {c: {"amt": tk.StringVar(value="0.00"), "self": tk.StringVar(value="0.00"),
                    "refund": tk.StringVar(value="0.00")}
                for c in list(m.keys()) + ["其他费"]}

    def setup_tab1(self):
        """ 构建 Tab1: 报销单汇总界面 """
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=" 报销单汇总 ")

        # 个人信息区域 / Personal Info Section
        f_info = tk.LabelFrame(tab, text="基本信息", padx=5, pady=5)
        f_info.pack(fill='x', padx=5)

        tk.Label(f_info, text="姓名:").grid(row=0, column=0, sticky='e')
        tk.Entry(f_info, textvariable=self.info_vars["name"], width=10).grid(row=0, column=1, padx=2)

        # --- 核心修改：保存身份证输入框的对象引用，以便改色 ---
        tk.Label(f_info, text="身份证:").grid(row=0, column=2, sticky='e')
        self.ent_id = tk.Entry(f_info, textvariable=self.info_vars["id"], width=22)
        self.ent_id.grid(row=0, column=3, padx=2)

        tk.Label(f_info, text="银行卡:").grid(row=0, column=4, sticky='e')
        tk.Entry(f_info, textvariable=self.info_vars["bank"], width=20).grid(row=0, column=5, padx=2)
        tk.Label(f_info, text="日期:").grid(row=0, column=6, sticky='e')
        tk.Entry(f_info, textvariable=self.info_vars["date"], width=11).grid(row=0, column=7, padx=2)
        tk.Label(f_info, text="年龄:").grid(row=1, column=0, sticky='e')
        tk.Entry(f_info, textvariable=self.info_vars["age"], width=10).grid(row=1, column=1, padx=2)
        tk.Label(f_info, text="单位:").grid(row=1, column=2, sticky='e')
        tk.Entry(f_info, textvariable=self.info_vars["unit"], width=22).grid(row=1, column=3, padx=2)
        tk.Label(f_info, text="类型:").grid(row=1, column=4, sticky='e')
        tk.Entry(f_info, textvariable=self.info_vars["type"], width=20).grid(row=1, column=5, padx=2)
        tk.Label(f_info, text="手机:").grid(row=1, column=6, sticky='e')
        tk.Entry(f_info, textvariable=self.info_vars["phone"], width=15).grid(row=1, column=7, padx=2)

        # 汇总金额网格区域 / Summary Grid Section
        f_table = tk.LabelFrame(tab, text="项目汇总", padx=5, pady=5)
        f_table.pack(fill='both', expand=True, padx=5)
        grid = [("医事服务费", 0, 0), ("西药", 0, 2), ("床位费", 0, 4), ("检查费", 1, 0), ("中药", 1, 2),
                ("其他费", 1, 4), ("治疗费", 2, 0), ("卫生材料费", 2, 2)]
        for cat, r, c in grid:
            tk.Label(f_table, text=cat + ":").grid(row=r, column=c, sticky='e', pady=8)
            tk.Entry(f_table, textvariable=self.sum_amt_vars[cat], state='readonly', width=12, justify='right').grid(
                row=r, column=c + 1, padx=5)

        # 底部操作按钮及状态 / Footer with Calc Status and Button
        f_footer = tk.Frame(tab, pady=10)
        f_footer.pack(fill='x')
        self.lbl_final = tk.Label(f_footer, text="票面总金额: 0.00   自费自负: 0.00   实报数合计: 0.00",
                                  font=("微软雅黑", 10, "bold"), fg="blue")
        self.lbl_final.pack()
        tk.Button(f_footer, text=" 生成结算 CSV 并打印 PDF ", command=self.generate_output, bg="#4CAF50", fg="white",
                  font=("微软雅黑", 10, "bold")).pack(pady=5)

    def setup_detail_tab(self, title, d, m, t, days=False):
        """ 构建 Tab2 & Tab3: 明细收据界面 """
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=f" {title} ")

        top = tk.Frame(tab, pady=5);
        top.pack(fill='x')
        tk.Button(top, text="上传Excel", command=lambda: self.load_excel(d, m)).pack(side='left', padx=10)
        if days: # 住院标签页特有天数输入
            tk.Label(top, text="住院天数:").pack(side='left')
            tk.Entry(top, textvariable=self.in_days_var, width=5).pack(side='left')
        f_g = tk.Frame(tab, padx=5);
        f_g.pack(fill='both')
        # 详情页表格标题 / Column Headers
        for c, text in enumerate(["诊疗项目", "票面金额", "自付金额", "实报金额"]):
            tk.Label(f_g, text=text, width=22, relief="ridge", bg="#e0e0e0").grid(row=0, column=c)

        # 动态创建行 / Dynamic row generation
        for i, cat in enumerate(d.keys()):
            tk.Label(f_g, text=cat, relief="groove", anchor='w', padx=5).grid(row=i + 1, column=0, sticky='nsew')
            tk.Entry(f_g, textvariable=d[cat]["amt"], state='readonly', justify='right').grid(row=i + 1, column=1,
                                                                                              sticky='nsew')
            tk.Entry(f_g, textvariable=d[cat]["self"], justify='right', bg="#fffde7").grid(row=i + 1, column=2,
                                                                                           sticky='nsew')
            tk.Label(f_g, textvariable=d[cat]["refund"], relief="groove", anchor='e', padx=5, fg="green").grid(
                row=i + 1, column=3, sticky='nsew')

        # 详情页合计行 / Detail Tab Total Row
        row_t = len(d) + 1
        tk.Label(f_g, text="该表合计", relief="ridge", bg="#f5f5f5", font=("", 9, "bold")).grid(row=row_t, column=0,
                                                                                                sticky='nsew')
        tk.Label(f_g, textvariable=t["amt"], relief="ridge", bg="#f5f5f5", anchor='e', padx=5).grid(row=row_t, column=1,
                                                                                                    sticky='nsew')
        tk.Label(f_g, textvariable=t["self"], relief="ridge", bg="#f5f5f5", anchor='e', padx=5).grid(row=row_t,
                                                                                                     column=2,
                                                                                                     sticky='nsew')
        tk.Label(f_g, textvariable=t["refund"], relief="ridge", bg="#f5f5f5", anchor='e', padx=5,
                 font=("", 9, "bold")).grid(row=row_t, column=3, sticky='nsew')

    def bind_traces(self):
        # 绑定身份证输入监听
        self.info_vars["id"].trace_add("write", self.handle_id_input)
        """ 绑定变量变化监听事件 """
        for d, t in [(self.data_out, self.out_totals), (self.data_in, self.in_totals)]:
            for cat in d:
                # 当金额或自付变化时，触发计算 / Trigger calculation on amount/self-pay change
                d[cat]["amt"].trace_add("write", lambda *a, x=d, y=t: self.update_calc(x, y))
                d[cat]["self"].trace_add("write", lambda *a, x=d, y=t: self.update_calc(x, y))

    def handle_id_input(self, *args):
        """处理身份证号输入：校验格式 + 自动计算年龄"""
        id_str = self.info_vars["id"].get().strip()

        # 18位身份证简易正则表达式
        # 前6位地址 + 中间8位生日(18/19/20xx年) + 后3位顺序 + 1位校验码(0-9/X/x)
        pattern = r"^[1-9]\d{5}(18|19|20)\d{2}((0[1-9])|(1[0-2]))(([0-2][1-9])|10|20|30|31)\d{3}[0-9Xx]$"

        if len(id_str) == 0:
            self.ent_id.config(bg="white")
        elif len(id_str) < 18:
            # 正在输入中，如果包含非法字符（除最后一位），可以即时变红
            if not re.match(r"^\d*$", id_str[:17]):
                self.ent_id.config(bg="#FFCCCC")
            else:
                self.ent_id.config(bg="white")
        elif len(id_str) == 18:
            if re.match(pattern, id_str):
                # 格式正确
                self.ent_id.config(bg="white")
                # 自动计算年龄
                try:
                    birth_year = int(id_str[6:10])
                    current_year = datetime.now().year
                    self.info_vars["age"].set(str(current_year - birth_year))
                except:
                    pass
            else:
                # 18位但格式不正确
                self.ent_id.config(bg="#FFCCCC")  # 浅红色
        else:
            # 超过18位
            self.ent_id.config(bg="#FFCCCC")

    def update_calc(self, d, t):
        """ 核心计算逻辑：计算单行实报和详情页合计 """
        s_a, s_s, s_r = 0.0, 0.0, 0.0
        for c in d:
            try:
                a, s = float(d[c]["amt"].get() or 0), float(d[c]["self"].get() or 0)
                d[c]["refund"].set(f"{a - s:.2f}")  # 实报 = 票面 - 自付
                s_a += a;
                s_s += s;
                s_r += (a - s)
            except:
                pass
        t["amt"].set(f"{s_a:.2f}")
        t["self"].set(f"{s_s:.2f}")
        t["refund"].set(f"{s_r:.2f}")
        self.update_summary()  # 同步更新 Tab1 汇总

    def update_summary(self):
        """ 同步更新 Tab1 汇总界面和最终实报数 """
        t_p, t_s = 0.0, 0.0
        for c in self.summary_cats:
            # 汇总 = 门诊对应项 + 住院对应项
            v1 = float(self.data_out.get(c, {}).get("amt", tk.StringVar(value="0")).get() or 0)
            v2 = float(self.data_in.get(c, {}).get("amt", tk.StringVar(value="0")).get() or 0)
            s1 = float(self.data_out.get(c, {}).get("self", tk.StringVar(value="0")).get() or 0)
            s2 = float(self.data_in.get(c, {}).get("self", tk.StringVar(value="0")).get() or 0)
            self.sum_amt_vars[c].set(f"{v1 + v2:.2f}");
            t_p += (v1 + v2);
            t_s += (s1 + s2)
        self.lbl_final.config(text=f"票面总金额: {t_p:.2f}   自费自负: {t_s:.2f}   实报合计: {t_p - t_s:.2f}")

    def load_excel(self, d, m):
        """ 处理 Excel 文件读取与智能分类逻辑 """
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path: return
        try:
            df = pd.read_excel(path)
            # 1. 自动填姓名 / Auto-fill name
            if not self.info_vars["name"].get():
                for c in ['购方名称', '交款人', '姓名']:
                    if c in df.columns: self.info_vars["name"].set(str(df[c].dropna().iloc[0]).strip()); break

            # 2. 发票分组逻辑 (根据代码+号码确定唯一发票)
            df['code'] = df['发票代码'].fillna('N/A').astype(str)
            df['num'] = df['发票号码'].fillna('N/A').astype(str)
            tmp_sums = {c: 0.0 for c in d.keys()}
            for _, g in df.groupby(['code', 'num']):
                total = float(g['票面金额'].iloc[0])  # 发票总额
                known = 0.0
                for _, r in g.iterrows():
                    name, amt = str(r['货物或应税劳务名称']), float(r['金额'] if '金额' in r else 0)
                    if amt > 0:
                        # 3. 根据关键词匹配分类 / Categorization by keywords
                        for cat, keys in m.items():
                            if any(k in name for k in keys): tmp_sums[cat] += amt; known += amt; break
                # 剩下的差额归入“其他费” / Remaining amount to "Others"
                tmp_sums["其他费"] += (total - known)

            for c in d: d[c]["amt"].set(f"{max(0, tmp_sums[c]):.2f}")
        except Exception as e:
            messagebox.showerror("读取错误", f"Excel处理失败: {e}")

    def generate_output(self):
        """ 核心功能：生成 CSV 数据报表并按 A4 1/3 比例渲染 PDF """
        name = self.info_vars["name"].get() or "未命名"
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        pdf_path = os.path.join(APP_PATH, f"报销单_{name}_{ts}.pdf")
        csv_path = os.path.join(APP_PATH, f"报销单_{name}_{ts}.csv")

        # 计算核心数值 / Calculation of key totals
        p_t = sum(float(self.sum_amt_vars[c].get()) for c in self.summary_cats)
        s_t = float(self.out_totals["self"].get()) + float(self.in_totals["self"].get())
        f_a = p_t - s_t

        try:
            # 1. 生成并保存全数据 CSV 报表
            csv_rows = [["【个人信息】"]]
            for k in ["name", "id", "bank", "date", "age", "unit", "type", "phone"]:
                csv_rows.append([k, self.info_vars[k].get()])
            csv_rows.append(["in_days", self.in_days_var.get()]);
            csv_rows.append([])
            csv_rows.append(["【门诊明细】"]);
            csv_rows.append(["项目", "票面", "自费", "实报"])
            for c in self.data_out: csv_rows.append(
                [c, self.data_out[c]["amt"].get(), self.data_out[c]["self"].get(), self.data_out[c]["refund"].get()])
            csv_rows.append([]);
            csv_rows.append(["【住院明细】"]);
            csv_rows.append(["项目", "票面", "自费", "实报"])
            for c in self.data_in: csv_rows.append(
                [c, self.data_in[c]["amt"].get(), self.data_in[c]["self"].get(), self.data_in[c]["refund"].get()])
            csv_rows.append([]);
            csv_rows.append(["【汇总结果】"])
            csv_rows.append(["票面总计", p_t]);
            csv_rows.append(["自费总计", s_t]);
            csv_rows.append(["实报数", f_a])
            pd.DataFrame(csv_rows).to_csv(csv_path, index=False, header=False, encoding="utf-8-sig")

            # 2. 渲染 PDF (PDF Rendering Setup)
            pdf = FPDF();
            pdf.add_page()
            # 加载中文字体 / Load Chinese font
            pdf.add_font("SimSun", "", r"C:\Windows\Fonts\simsun.ttc")

            # 标题与基本信息 / Header & Personal Info
            pdf.set_font("SimSun", size=18);
            pdf.cell(190, 12, "医 药 费 报 销 单", ln=True, align="C")
            pdf.set_font("SimSun", size=11);
            pdf.cell(190, 8, f"日期：{self.info_vars['date'].get()}", ln=True, align="R")

            h = 8  # 统一行高 / Row height
            # 个人信息格 / Info Grid
            pdf.cell(45, h, f"姓名：{self.info_vars['name'].get()}", border=1)
            pdf.cell(70, h, f"身份证号：{self.info_vars['id'].get()}", border=1)
            pdf.cell(75, h, f"银行卡号：{self.info_vars['bank'].get()}", border=1, ln=True)
            pdf.cell(45, h, f"年龄：{self.info_vars['age'].get()}", border=1)
            pdf.cell(70, h, f"单位：{self.info_vars['unit'].get()}", border=1)
            pdf.cell(35, h, f"人员类型：{self.info_vars['type'].get()}", border=1)
            pdf.cell(40, h, f"手机号：{self.info_vars['phone'].get()}", border=1, ln=True)

            # 表格标题行 (三列并排格式) / Table Header (3-column layout)
            pdf.ln(2);
            pdf.set_font("SimSun", size=10)
            w1, w2 = 33, 30
            for _ in range(3):
                pdf.cell(w1, h, "项目", border=1, align="C")
                pdf.cell(w2, h, "金额", border=1, align="C")
            pdf.ln()

            # 表格内容 / Table Body
            pdf.set_font("SimSun", size=11)
            grid = [("医事服务费", "西药", "床位费"), ("检查费", "中药", "其他费"), ("治疗费", "卫生材料费", "")]
            for row in grid:
                for cat in row:
                    pdf.cell(w1, h, cat, border=1)
                    val = self.sum_amt_vars[cat].get() if cat else ""
                    pdf.cell(w2, h, val, border=1, align="R")
                pdf.ln()

            # 结算公式展示区域 / Settlement Formula Section
            # 使用整体边框并移除内部横线 / Single border without internal lines
            s_x, s_y = pdf.get_x(), pdf.get_y()
            pdf.rect(s_x, s_y, 189, h * 3)
            pdf.cell(189, h, f"票面总金额：{p_t:.2f}    -自费自负：{s_t:.2f}", ln=True, border=0)
            pdf.cell(189, h, f" =总合计：            -个人负担：            =实报数：{f_a:.2f}", ln=True, border=0)
            pdf.cell(189, h, f"实报数(大写)：{cn_currency(f_a)}", ln=True, border=1)

            # 签字岗位区 / Audit Signature Section
            pdf.ln(2)
            pdf.set_font("SimSun", size=14)
            pdf.cell(63, h, "复核：", ln=0);
            pdf.cell(63, h, "制表：", ln=0);
            pdf.cell(64, h, "初审：", ln=1)

            # 法律承诺及右对齐签字 / Commitment & Signature Right Aligned
            pdf.ln(6)
            pdf.cell(190, 7, "本人承诺所提交票据（含电子票据）真实有效，无重复报销。", ln=True, align='R')
            pdf.cell(190, 7, "承诺并确认签字：____________________", ln=True, align='R')

            # 保存并自动打开预览 / Save and auto-preview
            pdf.output(pdf_path);
            os.startfile(pdf_path)
            messagebox.showinfo("成功", f"文件已保存至：\n{pdf_path}")
        except Exception as e:
            messagebox.showerror("导出错误", f"文件生成失败: {e}")


# --- 程序启动入口 ---
if __name__ == "__main__":
    root = tk.Tk()
    app = MedicalApp(root)
    root.mainloop()