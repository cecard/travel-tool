import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Alignment
import copy

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
        self.root.title("供电所差旅费自动生成工具 V2.1 (修正版)")
        self.root.geometry("950x750")
        
        self.config = self.load_config()
        self.trip_list = [] 
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
        notebook.add(self.frame_gen, text="行程录入与生成")
        self.setup_gen_tab()

        self.frame_user = ttk.Frame(notebook)
        notebook.add(self.frame_user, text="人员管理")
        self.setup_user_tab()

        self.frame_rules = ttk.Frame(notebook)
        notebook.add(self.frame_rules, text="规则设置")
        self.setup_rules_tab()

    def setup_gen_tab(self):
        left_panel = ttk.Frame(self.frame_gen, padding=10)
        left_panel.pack(side='left', fill='y', expand=False)
        
        right_panel = ttk.Frame(self.frame_gen, padding=10)
        right_panel.pack(side='right', fill='both', expand=True)

        # --- 左侧控件 ---
        row = 0
        ttk.Label(left_panel, text="第一步：选择报销人").grid(row=row, column=0, columnspan=2, sticky='w', pady=(0,5))
        row += 1
        self.cb_users = ttk.Combobox(left_panel, state="readonly", width=25)
        self.cb_users.grid(row=row, column=0, columnspan=2, sticky='ew')
        self.update_user_combobox()
        
        row += 1
        ttk.Separator(left_panel, orient='horizontal').grid(row=row, column=0, columnspan=2, sticky='ew', pady=10)
        
        row += 1
        ttk.Label(left_panel, text="第二步：录入单次行程").grid(row=row, column=0, columnspan=2, sticky='w', pady=(0,5))

        row += 1
        ttk.Label(left_panel, text="出发日期:").grid(row=row, column=0, sticky='w')
        self.entry_start_date = ttk.Entry(left_panel)
        self.entry_start_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.entry_start_date.grid(row=row, column=1, sticky='ew')

        row += 1
        ttk.Label(left_panel, text="起点:").grid(row=row, column=0, sticky='w')
        self.cb_start = ttk.Combobox(left_panel, values=["本所", self.config['station_info']['county'], self.config['station_info']['city']])
        self.cb_start.current(0)
        self.cb_start.grid(row=row, column=1, sticky='ew')

        row += 1
        ttk.Label(left_panel, text="终点:").grid(row=row, column=0, sticky='w')
        self.cb_end = ttk.Combobox(left_panel, values=["辖区线路", self.config['station_info']['county'], self.config['station_info']['city']])
        self.cb_end.bind("<<ComboboxSelected>>", self.on_end_point_change)
        self.cb_end.grid(row=row, column=1, sticky='ew')

        row += 1
        self.var_same_day = tk.BooleanVar(value=True)
        self.chk_same_day = ttk.Checkbutton(left_panel, text="当天往返", variable=self.var_same_day, command=self.on_sameday_change)
        self.chk_same_day.grid(row=row, column=1, sticky='w')

        row += 1
        ttk.Label(left_panel, text="返回日期:").grid(row=row, column=0, sticky='w')
        self.entry_end_date = ttk.Entry(left_panel)
        self.entry_end_date.grid(row=row, column=1, sticky='ew')
        self.entry_end_date.config(state='disabled')

        row += 1
        self.var_need_nocar = tk.BooleanVar(value=False)
        self.chk_nocar = ttk.Checkbutton(left_panel, text="需未派车证明", variable=self.var_need_nocar)
        self.chk_nocar.grid(row=row, column=0, sticky='w')
        
        ttk.Label(left_panel, text="事由:").grid(row=row, column=1, sticky='w')
        self.entry_reason = ttk.Entry(left_panel)
        self.entry_reason.insert(0, "差旅")
        self.entry_reason.grid(row=row+1, column=1, sticky='ew')

        row += 2
        btn_add = ttk.Button(left_panel, text="⬇️ 添加到列表", command=self.add_trip_to_list)
        btn_add.grid(row=row, column=0, columnspan=2, pady=15, sticky='ew')

        # --- 右侧列表 ---
        ttk.Label(right_panel, text="待生成行程列表 (可多次添加):").pack(anchor='w')
        
        cols = ("日期", "地点", "金额", "未派车")
        self.tree_trips = ttk.Treeview(right_panel, columns=cols, show='headings', height=15)
        self.tree_trips.heading("日期", text="日期")
        self.tree_trips.heading("地点", text="行程")
        self.tree_trips.heading("金额", text="金额(元)")
        self.tree_trips.heading("未派车", text="未派车")
        
        self.tree_trips.column("日期", width=100)
        self.tree_trips.column("地点", width=200)
        self.tree_trips.column("金额", width=80)
        self.tree_trips.column("未派车", width=60)
        self.tree_trips.pack(fill='both', expand=True)

        btn_box = ttk.Frame(right_panel)
        btn_box.pack(fill='x', pady=5)
        ttk.Button(btn_box, text="删除选中行", command=self.del_trip_from_list).pack(side='left')
        ttk.Button(btn_box, text="清空列表", command=self.clear_trip_list).pack(side='left', padx=5)

        # --- 底部生成区 ---
        bottom_frame = ttk.LabelFrame(right_panel, text="第三步：填报设置与生成")
        bottom_frame.pack(fill='x', pady=10)
        
        ttk.Label(bottom_frame, text="填报日期:").pack(side='left', padx=5)
        self.entry_fill_date = ttk.Entry(bottom_frame, width=12)
        self.entry_fill_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.entry_fill_date.pack(side='left')

        btn_gen = ttk.Button(bottom_frame, text="🚀 生成所有文件 (自动增行)", command=self.generate_all_files)
        btn_gen.pack(side='right', padx=10, pady=5)
        
        self.lbl_total = ttk.Label(right_panel, text="当前总金额: 0 元")
        self.lbl_total.pack(anchor='e')

    # --- 交互逻辑 ---
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

    def add_trip_to_list(self):
        start_date_str = self.entry_start_date.get()
        start_place = self.cb_start.get()
        end_place = self.cb_end.get()
        is_same_day = self.var_same_day.get()
        reason = self.entry_reason.get()
        need_nocar = self.var_need_nocar.get()

        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            if not is_same_day:
                end_date = datetime.strptime(self.entry_end_date.get(), "%Y-%m-%d")
            else:
                end_date = start_date
        except ValueError:
            return messagebox.showerror("错误", "日期格式无效")

        new_trips = []
        
        if end_place == "辖区线路":
            rule = self.config['rules']['local']
            new_trips.append({
                "date": start_date,
                "start": self.config['station_info']['name'].replace("供电所",""), 
                "end": "辖区",
                "food": rule['food'],
                "misc": rule['misc'],
                "nocar": need_nocar,
                "reason": reason,
                "full_start_date": start_date, 
                "full_end_date": end_date
            })
        else:
            if end_place == self.config['station_info']['county']:
                rule = self.config['rules']['county']
            else:
                rule = self.config['rules']['city']
            
            clean_start = start_place.replace("本所", self.config['station_info']['name'].replace("供电所",""))
            
            if is_same_day:
                new_trips.append({
                    "date": start_date,
                    "start": clean_start,
                    "end": end_place,
                    "food": 0,
                    "misc": rule['misc_round_trip'],
                    "nocar": need_nocar,
                    "reason": reason,
                    "full_start_date": start_date,
                    "full_end_date": end_date
                })
            else:
                new_trips.append({
                    "date": start_date,
                    "start": clean_start,
                    "end": end_place,
                    "food": 0,
                    "misc": rule['misc_one_way'],
                    "nocar": need_nocar, 
                    "reason": reason,
                    "full_start_date": start_date,
                    "full_end_date": end_date,
                    "is_return_trip": False
                })
                new_trips.append({
                    "date": end_date,
                    "start": end_place,
                    "end": clean_start,
                    "food": 0,
                    "misc": rule['misc_one_way'],
                    "nocar": False, 
                    "reason": reason,
                    "is_return_trip": True
                })

        for t in new_trips:
            self.trip_list.append(t)
        
        self.refresh_trip_list_ui()

    def del_trip_from_list(self):
        sel = self.tree_trips.selection()
        if not sel: return
        idx = self.tree_trips.index(sel[0])
        del self.trip_list[idx]
        self.refresh_trip_list_ui()

    def clear_trip_list(self):
        self.trip_list = []
        self.refresh_trip_list_ui()

    def refresh_trip_list_ui(self):
        for i in self.tree_trips.get_children():
            self.tree_trips.delete(i)
        
        total = 0
        for t in self.trip_list:
            cost = t['food'] + t['misc']
            total += cost
            nocar_str = "是" if t.get('nocar') else "-"
            display_loc = f"{t['start']} -> {t['end']}"
            self.tree_trips.insert('', 'end', values=(t['date'].strftime("%m-%d"), display_loc, cost, nocar_str))
        
        self.lbl_total.config(text=f"当前总金额: {total} 元")

    def generate_all_files(self):
        if not self.trip_list:
            return messagebox.showerror("错误", "列表为空，请先添加行程")
        
        user_idx = self.cb_users.current()
        if user_idx == -1: return messagebox.showerror("错误", "请选择报销人")
        user = self.config['users'][user_idx]
        
        try:
            fill_date = datetime.strptime(self.entry_fill_date.get(), "%Y-%m-%d")
        except:
            return messagebox.showerror("错误", "填报日期格式错误")

        self.trip_list.sort(key=lambda x: x['date'])
        
        total_money = sum([t['food'] + t['misc'] for t in self.trip_list])
        min_date = self.trip_list[0]['date']
        max_date = self.trip_list[-1]['date']
        total_days = (max_date - min_date).days + 1
        date_desc = f"自 {min_date.year} 年 {min_date.month} 月 {min_date.day} 日 至 {max_date.year} 年 {max_date.month} 月 {max_date.day} 日 计 {total_days} 天"

        file_suffix = f"{user['name']}_{fill_date.strftime('%m%d')}"

        try:
            # 1. 生成报销单
            wb = openpyxl.load_workbook(self.config['template_paths']['expense'])
            ws = wb.active
            
            ws['K2'] = fill_date.year
            ws['M2'] = fill_date.month
            ws['O2'] = fill_date.day
            ws['B3'] = self.config['station_info']['name'] 
            ws['G3'] = self.config['station_info']['name'] 
            ws['B4'] = user['name']
            ws['E4'] = self.trip_list[0]['reason'] 
            ws['G4'] = "详见明细"
            ws['J4'] = date_desc

            start_row = 8
            original_empty_rows = 6 
            current_row = start_row
            
            for i, t in enumerate(self.trip_list):
                if i >= original_empty_rows:
                    ws.insert_rows(current_row)
                
                ws[f'A{current_row}'] = t['date'].year
                ws[f'B{current_row}'] = t['date'].month
                ws[f'C{current_row}'] = t['date'].day
                ws[f'D{current_row}'] = t['start']
                ws[f'E{current_row}'] = t['end']
                
                if t['food'] > 0:
                    ws[f'H{current_row}'] = 1
                    ws[f'I{current_row}'] = t['food']
                
                if t['misc'] > 0:
                    ws[f'M{current_row}'] = t['misc']
                
                current_row += 1

            added_rows = max(0, len(self.trip_list) - original_empty_rows)
            row_total = 14 + added_rows
            row_bank = 15 + added_rows

            ws[f'G{row_total}'] = num_to_cn_amount(total_money)
            ws[f'C{row_bank}'] = user['name']
            ws[f'F{row_bank}'] = user['card']
            ws[f'K{row_bank}'] = user['bank']
            ws[f'N{row_bank}'] = user['phone']

            wb.save(f"1_差旅费报销单_{file_suffix}.xlsx")

            # 2. 生成审核单
            wb2 = openpyxl.load_workbook(self.config['template_paths']['audit'])
            ws2 = wb2.active
            ws2['K4'] = fill_date.year
            ws2['M4'] = fill_date.month
            ws2['O4'] = fill_date.day
            ws2['E6'] = self.config['station_info']['name']
            ws2['J10'] = total_money
            ws2['C11'] = num_to_cn_amount(total_money)
            ws2['C12'] = user['name']
            ws2['F12'] = user['card']
            ws2['K12'] = user['bank']
            ws2['N12'] = user['phone']
            wb2.save(f"2_报销审核单_{file_suffix}.xlsx")

            # 3. 批量生成未派车证明
            nocar_trips = [t for t in self.trip_list if t.get('nocar')]
            count_nocar = 0
            
            for t in nocar_trips:
                wb3 = openpyxl.load_workbook(self.config['template_paths']['no_car'])
                ws3 = wb3.active
                
                proof_date = t['date']
                ws3['F3'] = proof_date.year
                ws3['H3'] = proof_date.month
                ws3['J3'] = proof_date.day
                
                ws3['B5'] = self.config['station_info']['name']
                ws3['E5'] = user['name']
                ws3['H5'] = t['end']
                ws3['B7'] = t['reason']
                
                fs = t.get('full_start_date', proof_date)
                fe = t.get('full_end_date', proof_date)
                
                ws3['B8'] = fs.month
                ws3['D8'] = fs.day
                ws3['F8'] = fe.month
                ws3['H8'] = fe.day
                
                fname = f"3_未派车_{user['name']}_{fs.strftime('%m%d')}_至_{t['end']}.xlsx"
                wb3.save(fname)
                count_nocar += 1

            messagebox.showinfo("成功", f"生成完毕！\n- 报销单: 1份\n- 审核单: 1份\n- 未派车证明: {count_nocar}份")

        except Exception as e:
            messagebox.showerror("运行出错", str(e))

    # --- 用户管理 (已修正) ---
    def setup_user_tab(self):
        p = ttk.Frame(self.frame_user, padding=10)
        p.pack(fill='both', expand=True)

        # 修正：表头现在明确显示“开户银行”
        cols = ("姓名", "联系电话", "开户银行", "银行卡号")
        self.tree = ttk.Treeview(p, columns=cols, show='headings', height=10)
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
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
        # 修正：读取字典时使用正确的 Key
        if not u["姓名"]: return
        self.config['users'].append({
            "name": u["姓名"], 
            "phone": u["联系电话"], 
            "bank": u["开户银行"], 
            "card": u["银行卡号"]
        })
        self.save_config()
        self.refresh_user_list()
        self.update_user_combobox()
        for e in self.entries_user.values(): e.delete(0, tk.END)

    def del_user(self):
        sel = self.tree.selection()
        if not sel: return
        name = self.tree.item(sel[0])['values'][0]
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
            messagebox.showinfo("成功", "已设为默认")

    def setup_rules_tab(self):
        p = ttk.Frame(self.frame_rules, padding=10)
        p.pack(fill='both', expand=True)
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
        grp_rule = ttk.LabelFrame(p, text="费用规则 (元)")
        grp_rule.pack(fill='x', pady=5)
        ttk.Label(grp_rule, text="[辖区内] 伙食:").grid(row=0, column=0, pady=5)
        self.e_local_food = ttk.Entry(grp_rule, width=8)
        self.e_local_food.insert(0, self.config['rules']['local']['food'])
        self.e_local_food.grid(row=0, column=1)
        ttk.Label(grp_rule, text="[县城] 往返杂费:").grid(row=1, column=0, pady=5)
        self.e_county_round = ttk.Entry(grp_rule, width=8)
        self.e_county_round.insert(0, self.config['rules']['county']['misc_round_trip'])
        self.e_county_round.grid(row=1, column=1)
        ttk.Label(grp_rule, text="[县城] 单程杂费:").grid(row=1, column=2)
        self.e_county_single = ttk.Entry(grp_rule, width=8)
        self.e_county_single.insert(0, self.config['rules']['county']['misc_one_way'])
        self.e_county_single.grid(row=1, column=3)
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
        c = self.config['station_info']['county']
        city = self.config['station_info']['city']
        self.cb_start['values'] = ["本所", c, city]
        self.cb_end['values'] = ["辖区线路", c, city]
        messagebox.showinfo("成功", "设置已保存")

if __name__ == "__main__":
    root = tk.Tk()
    app = TravelApp(root)
    root.mainloop()
