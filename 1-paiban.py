import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from calendar import monthrange
import json
import os

# 配置文件路径
CONFIG_FILE = "config.json"

# 创建主窗口
root = tk.Tk()
root.title("自动排班系统")

# 获取当前年份
current_year = 2024

# 创建年份和月份选择的下拉菜单
tk.Label(root, text="年份:").grid(row=0, column=0)
year_var = tk.IntVar(value=current_year)
year_menu = ttk.Combobox(root, textvariable=year_var)
year_menu['values'] = [current_year + i for i in range(-10, 11)]  # 前后10年范围
year_menu.grid(row=0, column=1)

tk.Label(root, text="月份:").grid(row=1, column=0)
month_var = tk.IntVar(value=1)
month_menu = ttk.Combobox(root, textvariable=month_var)
month_menu['values'] = list(range(1, 13))  # 1到12月
month_menu.grid(row=1, column=1)

# 定义医生名字和状态列表
doctor_data = []

# 读取配置文件
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as file:
            config = json.load(file)
            return config.get("doctor_data", [])
    return []

# 保存配置文件
def save_config():
    config = {"doctor_data": doctor_data}
    with open(CONFIG_FILE, "w", encoding="utf-8") as file:
        json.dump(config, file, ensure_ascii=False, indent=4)

# 定义添加医生名字、状态、公休和积休的函数
def add_doctor():
    doctor = doctor_entry.get().strip()
    status = status_var.get()
    public_holiday = public_holiday_var.get()
    accumulated_holiday = accumulated_holiday_var.get()
    
    if doctor:
        for doc in doctor_data:
            if doc['name'] == doctor:
                messagebox.showerror("错误", "不能添加同名的医生")
                return
        
        doctor_data.append({"name": doctor, "status": status, "public_holiday": public_holiday, "accumulated_holiday": accumulated_holiday})
        doctor_listbox.insert(tk.END, f"{doctor} - 状态: {status}, 公休: {public_holiday}, 积休: {accumulated_holiday}")
        doctor_entry.delete(0, tk.END)
        status_var.set("正常工作")
        public_holiday_var.set(0)
        accumulated_holiday_var.set(0)
        save_config()
    else:
        messagebox.showerror("错误", "医生名字不能为空")

# 加载医生名字到列表框
def load_doctors_to_listbox():
    for doctor in doctor_data:
        doctor_listbox.insert(tk.END, f"{doctor['name']} - 状态: {doctor['status']}, 公休: {doctor['public_holiday']}, 积休: {doctor['accumulated_holiday']}")

# 定义修改医生名字、状态、公休和积休的函数
def modify_doctor():
    selected_index = doctor_listbox.curselection()
    if not selected_index:
        messagebox.showerror("错误", "请选择一个医生进行修改")
        return

    index = selected_index[0]
    doctor = doctor_entry.get().strip()
    status = status_var.get()
    public_holiday = public_holiday_var.get()
    accumulated_holiday = accumulated_holiday_var.get()

    if doctor:
        for i, doc in enumerate(doctor_data):
            if doc['name'] == doctor and i != index:
                messagebox.showerror("错误", "不能添加同名的医生")
                return
        
        doctor_data[index] = {"name": doctor, "status": status, "public_holiday": public_holiday, "accumulated_holiday": accumulated_holiday}
        doctor_listbox.delete(index)
        doctor_listbox.insert(index, f"{doctor} - 状态: {status}, 公休: {public_holiday}, 积休: {accumulated_holiday}")
        doctor_entry.delete(0, tk.END)
        status_var.set("正常工作")
        public_holiday_var.set(0)
        accumulated_holiday_var.set(0)
        save_config()
    else:
        messagebox.showerror("错误", "医生名字不能为空")

# 定义选中医生时填充输入框的函数
def on_select(event):
    selected_index = doctor_listbox.curselection()
    if selected_index:
        index = selected_index[0]
        doctor = doctor_data[index]
        doctor_entry.delete(0, tk.END)
        doctor_entry.insert(0, doctor['name'])
        status_var.set(doctor['status'])
        public_holiday_var.set(doctor['public_holiday'])
        accumulated_holiday_var.set(doctor['accumulated_holiday'])

# 创建输入医生名字、状态、公休和积休的部分
tk.Label(root, text="医生名字:").grid(row=2, column=0)
doctor_entry = tk.Entry(root)
doctor_entry.grid(row=2, column=1)

tk.Label(root, text="状态:").grid(row=2, column=2)
status_var = tk.StringVar(value="正常工作")
status_menu = ttk.Combobox(root, textvariable=status_var)
status_menu['values'] = ["正常工作", "产假", "进修"]
status_menu.grid(row=2, column=3)

tk.Label(root, text="公休:").grid(row=3, column=0)
public_holiday_var = tk.IntVar(value=0)
public_holiday_entry = tk.Entry(root, textvariable=public_holiday_var)
public_holiday_entry.grid(row=3, column=1)

tk.Label(root, text="积休:").grid(row=3, column=2)
accumulated_holiday_var = tk.IntVar(value=0)
accumulated_holiday_entry = tk.Entry(root, textvariable=accumulated_holiday_var)
accumulated_holiday_entry.grid(row=3, column=3)

add_doctor_button = tk.Button(root, text="添加医生", command=add_doctor)
add_doctor_button.grid(row=3, column=4)

modify_doctor_button = tk.Button(root, text="修改医生", command=modify_doctor)
modify_doctor_button.grid(row=3, column=5)

doctor_listbox = tk.Listbox(root, width=50)
doctor_listbox.grid(row=4, column=0, columnspan=6)
doctor_listbox.bind('<<ListboxSelect>>', on_select)

# 定义生成表格的函数
def generate_schedule():
    year = year_var.get()
    month = month_var.get()
    
    if not year or not month:
        messagebox.showerror("错误", "请选择年份和月份")
        return
    
    days_in_month = monthrange(year, month)[1]
    days = list(range(1, days_in_month + 1))
    weekdays_cn = ["一", "二", "三", "四", "五", "六", "日"]
    weekdays = [weekdays_cn[pd.Timestamp(year, month, day).weekday()] for day in days]

    # 创建表格
    headers = ["", "公休", "积休"] + days
    second_row = ["", "", ""] + weekdays

    data = [second_row]
    
    for doctor in doctor_data:
        row = [doctor['name']] + [""] * (len(headers) - 1)
        data.append(row)
    
    df = pd.DataFrame(data, columns=headers)
    
    # 输出到Excel文件
    output_file = f"排班表_{year}_{month}.xlsx"
    df.to_excel(output_file, index=False, header=True)
    messagebox.showinfo("成功", f"排班表已生成: {output_file}")

# 创建生成按钮
generate_button = tk.Button(root, text="生成排班表", command=generate_schedule)
generate_button.grid(row=5, column=0, columnspan=6)

# 加载配置文件
doctor_data = load_config()
load_doctors_to_listbox()

# 运行主循环
root.mainloop()
