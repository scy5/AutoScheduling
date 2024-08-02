import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
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
            for doctor in config.get("doctor_data", []):
                if "public_holiday" not in doctor:
                    doctor["public_holiday"] = 0
                if "accumulated_holiday" not in doctor:
                    doctor["accumulated_holiday"] = 0
                if "next_shift_date" not in doctor:
                    doctor["next_shift_date"] = None
                if "shift_interval" not in doctor:
                    doctor["shift_interval"] = 0
                if "clinic_day" not in doctor:
                    doctor["clinic_day"] = ""
            return config.get("doctor_data", [])
    return []

# 保存配置文件
def save_config():
    with open(CONFIG_FILE, "w", encoding="utf-8") as file:
        json.dump({"doctor_data": doctor_data}, file, ensure_ascii=False, indent=4)

# 更新医生列表
def load_doctors_to_listbox():
    doctor_listbox.delete(0, tk.END)
    for doctor in doctor_data:
        doctor_listbox.insert(tk.END, f"{doctor['name']} - 状态: {doctor['status']}, 公休: {doctor['public_holiday']}, 积休: {doctor['accumulated_holiday']}, 门诊日: {doctor['clinic_day']}")

# 添加医生
def add_doctor():
    name = doctor_name_entry.get().strip()
    status = doctor_status_var.get()
    public_holiday = int(public_holiday_entry.get())
    accumulated_holiday = int(accumulated_holiday_entry.get())
    next_shift_date = next_shift_date_entry.get_date()
    shift_interval = int(shift_interval_entry.get().strip())
    clinic_day = doctor_clinic_day_var.get()

    if name and not any(d['name'] == name for d in doctor_data):
        doctor_data.append({
            "name": name,
            "status": status,
            "public_holiday": public_holiday,
            "accumulated_holiday": accumulated_holiday,
            "next_shift_date": next_shift_date.strftime('%Y-%m-%d'),
            "shift_interval": shift_interval,
            "clinic_day": clinic_day
        })
        load_doctors_to_listbox()
        save_config()
    else:
        messagebox.showerror("错误", "医生名字不能为空或已存在")

# 修改医生
def modify_doctor():
    selected_index = doctor_listbox.curselection()
    if not selected_index:
        messagebox.showerror("错误", "请选择一个医生进行修改")
        return
    
    index = selected_index[0]
    name = doctor_name_entry.get().strip()
    status = doctor_status_var.get()
    public_holiday = int(public_holiday_entry.get())
    accumulated_holiday = int(accumulated_holiday_entry.get())
    next_shift_date = next_shift_date_entry.get_date()
    shift_interval = int(shift_interval_entry.get().strip())
    clinic_day = doctor_clinic_day_var.get()

    if name and (name == doctor_data[index]['name'] or not any(d['name'] == name for d in doctor_data)):
        doctor_data[index] = {
            "name": name,
            "status": status,
            "public_holiday": public_holiday,
            "accumulated_holiday": accumulated_holiday,
            "next_shift_date": next_shift_date.strftime('%Y-%m-%d'),
            "shift_interval": shift_interval,
            "clinic_day": clinic_day
        }
        load_doctors_to_listbox()
        save_config()
    else:
        messagebox.showerror("错误", "医生名字不能为空或已存在")

# 生成排班表格
def generate_schedule():
    year = year_var.get()
    month = month_var.get()
    _, num_days = monthrange(year, month)

    days = list(range(1, num_days + 1))
    weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    weekday_headers = [weekdays[pd.Timestamp(year, month, day).dayofweek] for day in days]

    data = [
        ["", "公休", "积休"] + days,
        ["", "", ""] + weekday_headers
    ]

    for doctor in doctor_data:
        row = [doctor['name'], doctor['public_holiday'], doctor['accumulated_holiday']] + [''] * num_days

        # 1.处理非正常工作的情况
        if doctor['status'] != "正常工作":
            for day in days:
                row[2 + day] = doctor['status']
            data.append(row)
            continue

        if doctor['next_shift_date'] and doctor['shift_interval'] > 0:
            next_shift_date = pd.Timestamp(doctor['next_shift_date'])

            # 2.用工作日初始化所有日期
            for day in days:
                current_date = pd.Timestamp(year, month, day)
                if current_date.weekday() < 5:
                    row[2 + day] = "日"
                else:
                    row[2 + day] = "休"

            # 3.值班日处理
            for day in days:
                current_date = pd.Timestamp(year, month, day)
                if (current_date - next_shift_date).days % doctor['shift_interval'] == 0:
                    row[2 + day] = "值"
                    if day < num_days:
                        row[2 + day + 1] = "/"
                    if current_date.weekday() == 5:  # Saturday
                        if day + 2 <= num_days:
                            row[2 + day + 2] = "周六调休"
                    elif current_date.weekday() == 6:  # Sunday
                        if day + 3 <= num_days:
                            row[2 + day + 3] = "周日调休"

            # 4.门诊日处理
            for day in days:
                current_date = pd.Timestamp(year, month, day)
                if doctor['clinic_day'] == weekdays[current_date.weekday()]:
                    if row[2 + day] == "日":
                        row[2 + day] = "日/门"
                    elif row[2 + day] == "值":
                        row[2 + day] = "值/门"
                    elif row[2 + day] == "/":
                        row[2 + day] = "加/门"
                    elif row[2 + day] == "休":
                        row[2 + day] = "休/门"
                    elif row[2 + day] == "周六调休" or row[2 + day] == "周日调休":
                        row[2 + day] = "日/门"  # ？
                    else:
                        row[2 + day] += "/门"
        data.append(row)

    df = pd.DataFrame(data)
    df.to_excel(f"{year}-{month}_排班表.xlsx", index=False, header=False)

    messagebox.showinfo("成功", f"排班表已生成: {year}-{month}_排班表.xlsx")

# 初始化医生列表
doctor_data = load_config()

# 医生列表框
doctor_listbox = tk.Listbox(root, width=80)
doctor_listbox.grid(row=12, column=0, columnspan=2)

# 医生输入框和按钮
tk.Label(root, text="医生名字:").grid(row=2, column=0)
doctor_name_entry = tk.Entry(root, width=30)
doctor_name_entry.grid(row=2, column=1)

tk.Label(root, text="医生状态:").grid(row=3, column=0)
doctor_status_var = tk.StringVar()
doctor_status_menu = ttk.Combobox(root, textvariable=doctor_status_var, width=27)
doctor_status_menu['values'] = ["正常工作", "产假", "进修"]
doctor_status_menu.grid(row=3, column=1)

tk.Label(root, text="公休:").grid(row=4, column=0)
public_holiday_entry = tk.Entry(root, width=30)
public_holiday_entry.grid(row=4, column=1)

tk.Label(root, text="积休:").grid(row=5, column=0)
accumulated_holiday_entry = tk.Entry(root, width=30)
accumulated_holiday_entry.grid(row=5, column=1)

tk.Label(root, text="下次值班日期:").grid(row=6, column=0)
next_shift_date_entry = DateEntry(root, width=27, date_pattern='y-mm-dd')
next_shift_date_entry.grid(row=6, column=1)

tk.Label(root, text="值班间隔 (天):").grid(row=7, column=0)
shift_interval_entry = tk.Entry(root, width=30)
shift_interval_entry.grid(row=7, column=1)

tk.Label(root, text="门诊日:").grid(row=8, column=0)
doctor_clinic_day_var = tk.StringVar()
doctor_clinic_day_menu = ttk.Combobox(root, textvariable=doctor_clinic_day_var, width=27)
doctor_clinic_day_menu['values'] = ["一", "二", "三", "四", "五", "六", "日"]
doctor_clinic_day_menu.grid(row=8, column=1)

add_button = tk.Button(root, text="添加医生", command=add_doctor)
add_button.grid(row=9, column=0, columnspan=2)

modify_button = tk.Button(root, text="修改医生", command=modify_doctor)
modify_button.grid(row=10, column=0, columnspan=2)

generate_button = tk.Button(root, text="生成排班表", command=generate_schedule)
generate_button.grid(row=11, column=0, columnspan=2)

# 加载医生列表
load_doctors_to_listbox()

# 运行主窗口
root.mainloop()
