import tkinter as tk
from tkinter import messagebox
import pickle
import os
import sys
import win32com.client

data_file = "todo_list.pkl"

def save_data(todo_list):
    with open(data_file, 'wb') as f:
        pickle.dump(todo_list, f)
def load_data():
    if os.path.exists(data_file):
        with open(data_file, 'rb') as f:
            return pickle.load(f)
    return []

def add_todo():
    task = task_entry.get()
    time = time_entry.get()
    quantity = quantity_entry.get()

    if task and time and quantity:
        todos.append({"time": time, "task": task, "quantity": quantity, "completed": False})
        task_entry.delete(0, tk.END)
        time_entry.delete(0, tk.END)
        quantity_entry.delete(0, tk.END)
        update_todo_list()
        save_data(todos)
    else:
        messagebox.showwarning("警告", "请输入所有字段！")

def update_todo_list():
    listbox.delete(0, tk.END)

    # 显示表头
    listbox.insert(tk.END, f"{'时间':<10} | {'任务':<15} | {'数量':<5} | {'完成'}")

    # 显示每个待办事项
    for index, item in enumerate(todos):
        status = "✓" if item["completed"] else "✗"
        listbox.insert(tk.END, f"{item['time']:<10} | {item['task']:<15} | {item['quantity']:<5} | {status}")

def toggle_todo(event):
    selected_index = listbox.curselection()
    if selected_index:
        # 跳过表头
        if selected_index[0] == 0:
            return
        index = selected_index[0] - 1
        todos[index]["completed"] = not todos[index]["completed"]
        update_todo_list()
        save_data(todos)

def delete_completed():
    global todos
    todos = [item for item in todos if not item["completed"]]
    update_todo_list()
    save_data(todos)

def create_shortcut():
    shortcut_name = "MyToDoList.lnk"
    target = os.path.abspath(sys.argv[0])
    shell = win32com.client.Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(os.path.join(os.path.expanduser("~"), shortcut_name))
    shortcut.TargetPath = target
    shortcut.WorkingDirectory = os.path.dirname(target)
    shortcut.save()
    messagebox.showinfo("成功", "快捷方式已创建！")

todos = load_data()

root = tk.Tk()
root.title("待办事项列表")
root.geometry("400x500")

time_label = tk.Label(root, text="时间")
time_label.pack()
time_entry = tk.Entry(root, width=40)
time_entry.pack(pady=5)

task_label = tk.Label(root, text="事项")
task_label.pack()
task_entry = tk.Entry(root, width=40)
task_entry.pack(pady=5)

quantity_label = tk.Label(root, text="数量")
quantity_label.pack()
quantity_entry = tk.Entry(root, width=40)
quantity_entry.pack(pady=5)

add_button = tk.Button(root, text="添加", command=add_todo)
add_button.pack(pady=10)

listbox = tk.Listbox(root, width=50, height=15)
listbox.pack(pady=10)
listbox.bind('<Double-Button-1>', toggle_todo)  # 双击切换完成状态

delete_button = tk.Button(root, text="删除已完成任务", command=delete_completed)
delete_button.pack(pady=10)
shortcut_button = tk.Button(root, text="创建快捷方式", command=create_shortcut)
shortcut_button.pack(pady=10)
update_todo_list()
root.mainloop()
