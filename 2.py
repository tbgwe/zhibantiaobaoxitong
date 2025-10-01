import tkinter as tk
from tkinter import ttk, messagebox
from docxtpl import DocxTemplate
import datetime

# 模板路径（根据事件类型）
templates = {
    "安全": r"D:\值班制度\02\11.docx",
    "生产": "生产事故模板.docx",
    "信访": "信访投诉模板.docx"
}

root = tk.Tk()
root.title("突发事件报告生成器")
root.geometry("500x1000")

# -------------------- 顶部选择控件 --------------------
tk.Label(root, text="事件类型：").pack(anchor="w", padx=10, pady=5)
event_type_cb = ttk.Combobox(root, values=list(templates.keys()))
event_type_cb.configure(state="readonly")
event_type_cb.pack(fill="x", padx=10)

# 确定按钮
def sure():
    # 清空动态控件 Frame
    for widget in dynamic_frame.winfo_children():
        widget.destroy()
    generate_dynamic_widgets()

btn_sure = tk.Button(root, text="确定", command=sure, bg="lightblue")
btn_sure.pack(pady=20)

# -------------------- 动态控件容器 --------------------
dynamic_frame = tk.Frame(root)
dynamic_frame.pack(fill="x")

# -------------------- 动态生成控件函数 --------------------
def generate_dynamic_widgets():
    event_type = event_type_cb.get()
    if not event_type:
        messagebox.showwarning("提示", "请选择事件类型！")
        return

    widgets = {}  # 存放控件引用

    # 安全事件
    if event_type == '安全':
        tk.Label(dynamic_frame, text="事件时间：").pack(anchor="w", padx=10, pady=5)
        entry_time = tk.Entry(dynamic_frame)
        entry_time.insert(0, datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
        entry_time.pack(fill="x", padx=10)
        widgets['event_time'] = entry_time

        tk.Label(dynamic_frame, text="事件地点：").pack(anchor="w", padx=10, pady=5)
        entry_place = tk.Entry(dynamic_frame)
        entry_place.pack(fill="x", padx=10)
        widgets['event_place'] = entry_place

    # 信访事件
    if event_type == '信访':
        tk.Label(dynamic_frame, text="事件经过：").pack(anchor="w", padx=10, pady=5)
        text_detail = tk.Text(dynamic_frame, height=5)
        text_detail.pack(fill="x", padx=10)
        widgets['event_detail'] = text_detail

        tk.Label(dynamic_frame, text="处置情况：").pack(anchor="w", padx=10, pady=5)
        text_handle = tk.Text(dynamic_frame, height=5)
        text_handle.pack(fill="x", padx=10)
        widgets['event_handle'] = text_handle

        tk.Label(dynamic_frame, text="上报单位：").pack(anchor="w", padx=10, pady=5)
        entry_unit = tk.Entry(dynamic_frame)
        entry_unit.insert(0, "T型新科技")
        entry_unit.pack(fill="x", padx=10)
        widgets['report_unit'] = entry_unit

    # 生成报告按钮
    def generate_report():
        try:
            context = {}
            if 'event_time' in widgets:
                context["事件时间"] = widgets['event_time'].get()
            if 'event_place' in widgets:
                context["事件地点"] = widgets['event_place'].get()
            if 'event_detail' in widgets:
                context["事件经过"] = widgets['event_detail'].get("1.0", tk.END).strip()
            if 'event_handle' in widgets:
                context["处置情况"] = widgets['event_handle'].get("1.0", tk.END).strip()
            if 'report_unit' in widgets:
                context["上报单位"] = widgets['report_unit'].get()

            tpl = DocxTemplate(templates[event_type])
            tpl.render(context)
            filename = fr"D:\值班制度\02\{event_type}事件报告_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            tpl.save(filename)
            messagebox.showinfo("成功", f"报告已生成：{filename}")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败：{e}")

    btn_generate = tk.Button(dynamic_frame, text="生成报告", command=generate_report, bg="lightblue")
    btn_generate.pack(pady=20)


root.mainloop()
