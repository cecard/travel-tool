import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Alignment
import shutil

# --- 配置文件路径 ---
CONFIG_FILE = "config.json"

# --- 默认配置 ---
DEFAULT_CONFIG = {
    "users": [],
    "current_user_index": -1,
    "station_info": {
        "name": "龙潭供电所",
        "county": "桃源县",
        "city": "常德市"
    },
    "rules": {
        "local": {"traffic": 0, "food": 40, "stay": 0, "misc": 0},
        "county": {"traffic": 0, "food": 0, "stay": 0, "misc_one_way": 15, "misc_round_trip": 30},
        "city": {"traffic": 0, "food": 0, "stay": 0, "misc_one_way": 25, "misc_round_trip": 50}
    },
    "template_paths": {
        "expense": "差旅费报销单模板.xlsx",
        "audit": "报销审核单模板.xlsx",
        "no_car": "未派车证明模板.xlsx"
    }
}

# --- 数字转大写金额函数 ---
def num_to_cn_amount(num):
    if num == 0: return "零元整"
    units = ["", "拾", "佰", "仟"]
    big_units = ["", "万", "亿"]
    num_str = str(int(num))
    fraction = str(round(num - int(num), 2))[2:]
    
    result = ""
    length = len(num_str)
    for i, digit in enumerate(num_str):
        n = int(digit)
        if n != 0:
            result += "零壹贰叁肆伍陆柒捌玖"[n] + units[(length - 1 - i) % 4]
        if (length - 1 - i) % 4 == 0:
            result += big_units[(length - 1 - i) // 4]
            
    result = result.replace("零零", "零").strip("零")
    result += "元"
    
    if len(fraction) > 0:
        jiao = int(fraction[0])
        fen = int(fraction[1]) if len(fraction) > 1 else 0
        if jiao > 0: result += "零壹贰叁肆伍陆柒捌玖"[jiao] + "角"
        if fen > 0: result += "零壹贰叁肆伍陆柒捌玖"[fen] + "分"
    else:
        result += "整"
    return result

class TravelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("供电所差旅费自动生成工具")
        self.root.geometry("700x650")
        
        self.config = self.load_config()
        self.setup_ui()

    def load_config(self):
        if not os.path.exists(CONFIG_FILE):
            return DEFAULT_CONFIG
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return DEFAULT_CONFIG

    def save_config(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)

    def setup_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(expand=True, fill='both')

        self.frame_gen = ttk.Frame(notebook)
        notebook.add(self.frame_gen, text="生成报销单")
        self.setup_gen_tab()

        self.frame_user = ttk.Frame(notebook)
        notebook.add(self.frame_user, text="人员管理")
        self.setup_user_tab()

        self.frame_rules = ttk.Frame(notebook)
        notebook.add(self.frame_rules, text="规则设置")
        self.setup_rules_tab()

    def setup_gen_tab(self):
        p = ttk.Frame(self.frame_gen, padding=15)
        p.pack(fill='both', expand=True)

        # 报销人
        row = 0
        ttk.Label(p, text="报销人:").grid(row=row, column=0, sticky='w', pady=5)
        self.cb_users = ttk.Combobox(p, state="readonly")
        self.cb_users.grid(row=row, column=1, sticky='ew')
        self.update_user_combobox()

        # 出发日期
        row += 1
        ttk.Label(p, text="出发日期 (YYYY-MM-DD):").grid(row=row, column=0, sticky='w', pady=5)
        self.entry_start_date = ttk.Entry(p)
        self.entry_start_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.entry_start_date.grid(row=row, column=1, sticky='ew')

        # 填报日期
        row += 1
        ttk.Label(p, text="填报日期 (YYYY-MM-DD):").grid(row=row, column=0, sticky='w', pady=5)
        self.entry_fill_date = ttk.Entry(p)
        self.entry_fill_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.entry_fill_date.grid(row=row, column=1, sticky='ew')

        # 起点
        row += 1
        ttk.Label(p, text="起点:").grid(row=row, column=0, sticky='w', pady=5)
        self.cb_start = ttk.Combobox(p, values=["本所", self.config['station_info']['county'], self.config['station_info']['city']])
        self.cb_start.current(0)
        self.cb_start.grid(row=row, column=1, sticky='ew')

        # 终点
        row += 1
        ttk.Label(p, text="终点:").grid(row=row, column=0, sticky='w', pady=5)
        self.cb_end = ttk.Combobox(p, values=["辖区线路", self.config['station_info']['county'], self.config['station_info']['city']])
        self.cb_end.bind("<<ComboboxSelected>>", self.on_end_point_change)
        self.cb_end.grid(row=row, column=1, sticky='ew')

        # 当天往返
        row += 1
        self.var_same_day = tk.BooleanVar(value=True)
        self.chk_same_day = ttk.Checkbutton(p, text="是否当天往返", variable=self.var_same_day, command=self.on_sameday_change)
        self.chk_same_day.grid(row=row, column=1, sticky='w', pady=5)

        # 返回日期
        row += 1
        ttk.Label(p, text="返回日期 (若非当天):").grid(row=row, column=0, sticky='w', pady=5)
        self.entry_end_date = ttk.Entry(p)
        self.entry_end_date.grid(row=row, column=1, sticky='ew')
        self.entry_end_date.config(state='disabled')

        # 未派车证明
        row += 1
        self.var_need_nocar = tk.BooleanVar(value=False)
        self.chk_nocar = ttk.Checkbutton(p, text="生成《未派车证明》", variable=self.var_need_nocar, command=self.on_nocar_change)
        self.chk_nocar.grid(row=row, column=0, sticky='w', pady=5)
        
        ttk.Label(p, text="出差事由:").grid(row=row, column=1, sticky='w')
        self.entry_reason = ttk.Entry(p)
        self.entry_reason.insert(0, "差旅")
        self.entry_reason.grid(row=row+1, column=1, sticky='ew')

        # 按钮
        row += 2
        btn_gen = ttk.Button(p, text="生成 Excel 表格", command=self.generate_excel)
        btn_gen.grid(row=row, column=0, columnspan=2, pady=20)

    def on_end_point_change(self, event):
        val = self.cb_end.get()
        if val == "辖区线路":
            self.var_same_day.set(True)
            self.chk_same_day.config(state='disabled')
            self.entry_end_date.config(state='disabled')
            self.entry_end_date.delete(0, tk.END)
        else:
            self.chk_same_day.config(state='normal')
            self.on_sameday_change()

    def on_sameday_change(self):
        if self.var_same_day.get():
            self.entry_end_date.delete(0, tk.END)
            self.entry_end_date.config(state='disabled')
        else:
            self.entry_end_date.config(state='normal')
            start = self.entry_start_date.get()
            try:
                d = datetime.strptime(start, "%Y-%m-%d")
                self.entry_end_date.insert(0, (d + timedelta(days=1)).strftime("%Y-%m-%d"))
            except:
                pass

    def on_nocar_change(self):
        pass # UI上已经处理

    def setup_user_tab(self):
        p = ttk.Frame(self.frame_user, padding=10)
        p.pack(fill='both', expand=True)

        cols = ("姓名", "电话", "银行", "卡号")
        self.tree = ttk.Treeview(p, columns=cols, show='headings', height=10)
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        self.tree.pack(fill='x')

        frame_input = ttk.Frame(p)
        frame_input.pack(pady=10)
        
        self.entries_user = {}
        for i, col in enumerate(cols):
            ttk.Label(frame_input, text=col).grid(row=0, column=i, padx=5)
            e = ttk.Entry(frame_input, width=15)
            e.grid(row=1, column=i, padx=5)
            self.entries_user[col] = e

        btn_box = ttk.Frame(p)
        btn_box.pack(pady=5)
        ttk.Button(btn_box, text="添加用户", command=self.add_user).pack(side='left', padx=5)
        ttk.Button(btn_box, text="删除选中", command=self.del_user).pack(side='left', padx=5)
        ttk.Button(btn_box, text="设为默认", command=self.set_default_user).pack(side='left', padx=5)
        
        self.refresh_user_list()

    def setup_rules_tab(self):
        p = ttk.Frame(self.frame_rules, padding=10)
        p.pack(fill='both', expand=True)

        # 供电所设置
        grp_station = ttk.LabelFrame(p, text="基本信息")
        grp_station.pack(fill='x', pady=5)
        
        ttk.Label(grp_station, text="供电所名:").grid(row=0, column=0)
        self.entry_st_name = ttk.Entry(grp_station)
        self.entry_st_name.insert(0, self.config['station_info']['name'])
        self.entry_st_name.grid(row=0, column=1)

        ttk.Label(grp_station, text="所属县城:").grid(row=0, column=2)
        self.entry_st_county = ttk.Entry(grp_station)
        self.entry_st_county.insert(0, self.config['station_info']['county'])
        self.entry_st_county.grid(row=0, column=3)

        ttk.Label(grp_station, text="所属城市:").grid(row=0, column=4)
        self.entry_st_city = ttk.Entry(grp_station)
        self.entry_st_city.insert(0, self.config['station_info']['city'])
        self.entry_st_city.grid(row=0, column=5)

        # 费用规则
        grp_rule = ttk.LabelFrame(p, text="费用规则 (元)")
        grp_rule.pack(fill='x', pady=5)

        # 辖区内
        ttk.Label(grp_rule, text="[辖区内] 伙食:").grid(row=0, column=0, pady=5)
        self.e_local_food = ttk.Entry(grp_rule, width=8)
        self.e_local_food.insert(0, self.config['rules']['local']['food'])
        self.e_local_food.grid(row=0, column=1)
        
        # 县城
        ttk.Label(grp_rule, text="[县城] 往返杂费:").grid(row=1, column=0, pady=5)
        self.e_county_round = ttk.Entry(grp_rule, width=8)
        self.e_county_round.insert(0, self.config['rules']['county']['misc_round_trip'])
        self.e_county_round.grid(row=1, column=1)

        ttk.Label(grp_rule, text="[县城] 单程杂费:").grid(row=1, column=2)
        self.e_county_single = ttk.Entry(grp_rule, width=8)
        self.e_county_single.insert(0, self.config['rules']['county']['misc_one_way'])
        self.e_county_single.grid(row=1, column=3)
        
        # 城市
        ttk.Label(grp_rule, text="[市区] 往返杂费:").grid(row=2, column=0, pady=5)
        self.e_city_round = ttk.Entry(grp_rule, width=8)
        self.e_city_round.insert(0, self.config['rules']['city']['misc_round_trip'])
        self.e_city_round.grid(row=2, column=1)

        ttk.Label(grp_rule, text="[市区] 单程杂费:").grid(row=2, column=2)
        self.e_city_single = ttk.Entry(grp_rule, width=8)
        self.e_city_single.insert(0, self.config['rules']['city']['misc_one_way'])
        self.e_city_single.grid(row=2, column=3)

        btn_save = ttk.Button(p, text="保存所有设置", command=self.save_all_settings)
        btn_save.pack(pady=20)

    # --- 逻辑功能 ---
    def refresh_user_list(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for u in self.config['users']:
            self.tree.insert('', 'end', values=(u['name'], u['phone'], u['bank'], u['card']))
    
    def update_user_combobox(self):
        names = [u['name'] for u in self.config['users']]
        self.cb_users['values'] = names
        if self.config['current_user_index'] >= 0 and self.config['current_user_index'] < len(names):
            self.cb_users.current(self.config['current_user_index'])

    def add_user(self):
        u = {k: v.get() for k, v in self.entries_user.items()}
        if not u["姓名"]: return
        self.config['users'].append({"name": u["姓名"], "phone": u["电话"], "bank": u["银行"], "card": u["卡号"]})
        self.save_config()
        self.refresh_user_list()
        self.update_user_combobox()
        # 清空输入框
        for e in self.entries_user.values(): e.delete(0, tk.END)

    def del_user(self):
        sel = self.tree.selection()
        if not sel: return
        item = self.tree.item(sel[0])
        name = item['values'][0]
        self.config['users'] = [u for u in self.config['users'] if u['name'] != name]
        self.config['current_user_index'] = -1
        self.save_config()
        self.refresh_user_list()
        self.update_user_combobox()

    def set_default_user(self):
        idx = self.cb_users.current()
        if idx != -1:
            self.config['current_user_index'] = idx
            self.save_config()
            messagebox.showinfo("成功", "已设为默认登录人")

    def save_all_settings(self):
        self.config['station_info']['name'] = self.entry_st_name.get()
        self.config['station_info']['county'] = self.entry_st_county.get()
        self.config['station_info']['city'] = self.entry_st_city.get()
        
        try:
            self.config['rules']['local']['food'] = float(self.e_local_food.get())
            self.config['rules']['county']['misc_round_trip'] = float(self.e_county_round.get())
            self.config['rules']['county']['misc_one_way'] = float(self.e_county_single.get())
            self.config['rules']['city']['misc_round_trip'] = float(self.e_city_round.get())
            self.config['rules']['city']['misc_one_way'] = float(self.e_city_single.get())
        except ValueError:
            return messagebox.showerror("错误", "费用必须是数字")

        self.save_config()
        # 更新下拉框
        c = self.config['station_info']['county']
        city = self.config['station_info']['city']
        self.cb_start['values'] = ["本所", c, city]
        self.cb_end['values'] = ["辖区线路", c, city]
        messagebox.showinfo("成功", "设置已保存")

    def generate_excel(self):
        user_idx = self.cb_users.current()
        if user_idx == -1: return messagebox.showerror("错误", "请选择用户")
        
        user = self.config['users'][user_idx]
        start_place = self.cb_start.get()
        end_place = self.cb_end.get()
        reason = self.entry_reason.get()
        
        try:
            fill_date = datetime.strptime(self.entry_fill_date.get(), "%Y-%m-%d")
            start_date = datetime.strptime(self.entry_start_date.get(), "%Y-%m-%d")
            
            if not self.var_same_day.get():
                end_date = datetime.strptime(self.entry_end_date.get(), "%Y-%m-%d")
            else:
                end_date = start_date
        except ValueError:
            return messagebox.showerror("错误", "日期格式错误，应为 YYYY-MM-DD")

        # --- 计算费用和行程 ---
        trips = [] # 存储每一行数据
        total_money = 0
        
        # 规则判断
        if end_place == "辖区线路":
            # 辖区内：只可能有当天往返
            rule = self.config['rules']['local']
            trips.append({
                "date": start_date,
                "start": self.config['station_info']['name'].replace("供电所",""), # 简写
                "end": "辖区",
                "food": rule['food'],
                "misc": rule['misc'],
                "days": 1
            })
        
        else:
            # 县城或市区
            if end_place == self.config['station_info']['county']:
                rule = self.config['rules']['county']
            else:
                rule = self.config['rules']['city']
            
            if self.var_same_day.get():
                # 当天往返
                trips.append({
                    "date": start_date,
                    "start": start_place.replace("本所", self.config['station_info']['name']),
                    "end": end_place,
                    "food": 0,
                    "misc": rule['misc_round_trip'],
                    "days": 1
                })
            else:
                # 非当天往返：拆分成去程和回程
                # 去程
                trips.append({
                    "date": start_date,
                    "start": start_place.replace("本所", self.config['station_info']['name']),
                    "end": end_place,
                    "food": 0,
                    "misc": rule['misc_one_way'],
                    "days": 1
                })
                # 回程
                trips.append({
                    "date": end_date,
                    "start": end_place,
                    "end": start_place.replace("本所", self.config['station_info']['name']),
                    "food": 0,
                    "misc": rule['misc_one_way'],
                    "days": 1
                })

        # 计算总金额
        for t in trips:
            total_money += t['food'] + t['misc']

        # --- 开始填表 1: 差旅费报销单 ---
        try:
            wb = openpyxl.load_workbook(self.config['template_paths']['expense'])
            ws = wb.active

            # 1. 顶部基础信息 (根据截图坐标)
            ws['K2'] = fill_date.year
            ws['M2'] = fill_date.month
            ws['O2'] = fill_date.day
            ws['B3'] = self.config['station_info']['name'] # 单位，截图显示B3可能是单位
            ws['G3'] = self.config['station_info']['name'] # 部门
            ws['B4'] = user['name'] # 姓名
            ws['E4'] = reason # 事由
            ws['G4'] = end_place # 地点
            
            # 出差日期说明 J4 (合并格)
            days_count = (end_date - start_date).days + 1
            date_desc = f"自 {start_date.year} 年 {start_date.month} 月 {start_date.day} 日 至 {end_date.year} 年 {end_date.month} 月 {end_date.day} 日 计 {days_count} 天"
            ws['J4'] = date_desc

            # 2. 填写行程列表 (从第8行开始)
            current_row = 8
            for t in trips:
                ws[f'A{current_row}'] = t['date'].year
                ws[f'B{current_row}'] = t['date'].month
                ws[f'C{current_row}'] = t['date'].day
                ws[f'D{current_row}'] = t['start']
                ws[f'E{current_row}'] = t['end']
                
                # 金额填入
                if t['food'] > 0:
                    ws[f'H{current_row}'] = 1 # 天数
                    ws[f'I{current_row}'] = t['food']
                
                if t['misc'] > 0:
                    ws[f'M{current_row}'] = t['misc']
                
                current_row += 1

            # 3. 底部信息
            ws['G14'] = num_to_cn_amount(total_money) # 总金额大写 (Row 14)
            ws['C15'] = user['name'] # 开户名称
            ws['F15'] = user['card'] # 账号
            ws['K15'] = user['bank'] # 银行
            ws['N15'] = user['phone'] # 电话

            save_name_1 = f"1_报销单_{user['name']}_{start_date.strftime('%m%d')}.xlsx"
            wb.save(save_name_1)

            # --- 开始填表 2: 报销审核单 ---
            wb2 = openpyxl.load_workbook(self.config['template_paths']['audit'])
            ws2 = wb2.active
            
            ws2['K4'] = fill_date.year
            ws2['M4'] = fill_date.month
            ws2['O4'] = fill_date.day
            ws2['E6'] = self.config['station_info']['name'] # 部门
            ws2['J10'] = total_money # 小写金额
            ws2['C11'] = num_to_cn_amount(total_money) # 大写金额
            
            # 银行信息 Row 12
            ws2['C12'] = user['name']
            ws2['F12'] = user['card']
            ws2['K12'] = user['bank']
            ws2['N12'] = user['phone']

            save_name_2 = f"2_审核单_{user['name']}_{start_date.strftime('%m%d')}.xlsx"
            wb2.save(save_name_2)

            # --- 开始填表 3: 未派车证明 (如果勾选) ---
            if self.var_need_nocar.get():
                wb3 = openpyxl.load_workbook(self.config['template_paths']['no_car'])
                ws3 = wb3.active
                
                # 证明日期 F3, H3, J3
                ws3['F3'] = fill_date.year
                ws3['H3'] = fill_date.month
                ws3['J3'] = fill_date.day
                
                ws3['B5'] = self.config['station_info']['name']
                ws3['E5'] = user['name']
                ws3['H5'] = end_place
                ws3['B7'] = reason
                
                # 起始日期 B8, D8
                ws3['B8'] = start_date.month
                ws3['D8'] = start_date.day
                # 截止日期 F8, H8
                ws3['F8'] = end_date.month
                ws3['H8'] = end_date.day

                save_name_3 = f"3_未派车_{user['name']}_{start_date.strftime('%m%d')}.xlsx"
                wb3.save(save_name_3)

            messagebox.showinfo("完成", f"已生成表格：\n{save_name_1}\n{save_name_2}")

        except Exception as e:
            messagebox.showerror("运行出错", f"错误信息: {str(e)}\n请检查模板文件是否存在且未被打开。")

if __name__ == "__main__":
    root = tk.Tk()
    app = TravelApp(root)
    root.mainloop()
