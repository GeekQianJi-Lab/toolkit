import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import threading


class DtaToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("DTA 转 Excel 工具")
        self.root.geometry("500x320")
        self.root.resizable(False, False)

        self.dta_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        # 标题
        title_label = tk.Label(
            self.root,
            text="DTA 转 Excel 工具",
            font=("Arial", 16, "bold"),
            fg="#333333"
        )
        title_label.pack(pady=(20, 20))

        # 主框架
        main_frame = tk.Frame(self.root, padx=20)
        main_frame.pack(fill="x")

        # 选择 DTA 文件
        dta_frame = tk.Frame(main_frame)
        dta_frame.pack(fill="x", pady=10)

        tk.Label(dta_frame, text="DTA 文件:", font=("Arial", 10)).pack(side="left")

        dta_entry = tk.Entry(dta_frame, textvariable=self.dta_file_path, font=("Arial", 10))
        dta_entry.pack(side="left", fill="x", expand=True, padx=(10, 10))

        dta_btn = tk.Button(
            dta_frame,
            text="选择文件",
            command=self.select_dta_file,
            font=("Arial", 9),
            bg="#4A90E2",
            fg="white",
            width=10
        )
        dta_btn.pack(side="right")

        # 选择输出位置
        output_frame = tk.Frame(main_frame)
        output_frame.pack(fill="x", pady=10)

        tk.Label(output_frame, text="保存位置:", font=("Arial", 10)).pack(side="left")

        output_entry = tk.Entry(output_frame, textvariable=self.output_file_path, font=("Arial", 10))
        output_entry.pack(side="left", fill="x", expand=True, padx=(10, 10))

        output_btn = tk.Button(
            output_frame,
            text="选择位置",
            command=self.select_output_file,
            font=("Arial", 9),
            bg="#4A90E2",
            fg="white",
            width=10
        )
        output_btn.pack(side="right")

        # 状态显示
        status_frame = tk.Frame(main_frame)
        status_frame.pack(fill="x", pady=20)

        self.status_label = tk.Label(
            status_frame,
            text="就绪",
            font=("Arial", 10),
            fg="#666666"
        )
        self.status_label.pack()

        # 进度条
        self.progress = ttk.Progressbar(status_frame, mode="indeterminate", length=400)
        self.progress.pack(pady=(10, 0))

        # 转换按钮
        convert_btn = tk.Button(
            main_frame,
            text="开始转换",
            command=self.start_conversion,
            font=("Arial", 12, "bold"),
            bg="#50C878",
            fg="white",
            width=20,
            height=2,
            cursor="hand2"
        )
        convert_btn.pack(pady=20)

    def select_dta_file(self):
        file_path = filedialog.askopenfilename(
            title="选择 DTA 文件",
            filetypes=[("DTA 文件", "*.dta"), ("所有文件", "*.*")]
        )
        if file_path:
            self.dta_file_path.set(file_path)
            # 自动设置输出文件名
            if not self.output_file_path.get():
                import os
                output_path = file_path.rsplit(".", 1)[0] + ".xlsx"
                self.output_file_path.set(output_path)

    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="保存 Excel 文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.output_file_path.set(file_path)

    def start_conversion(self):
        dta_path = self.dta_file_path.get()
        output_path = self.output_file_path.get()

        if not dta_path:
            messagebox.showwarning("警告", "请选择 DTA 文件！")
            return
        if not output_path:
            messagebox.showwarning("警告", "请选择保存位置！")
            return

        # 使用线程避免界面卡顿
        thread = threading.Thread(target=self.convert_file, args=(dta_path, output_path))
        thread.daemon = True
        thread.start()

    def convert_file(self, dta_path, output_path):
        try:
            self.root.after(0, lambda: self.update_status("正在转换...", True))

            df = pd.read_stata(dta_path)
            df.to_excel(output_path, index=False)

            row_count = len(df)
            self.root.after(0, lambda: self.update_status(f"转换成功！共 {row_count} 行数据", False))
            self.root.after(0, lambda: messagebox.showinfo("完成", f"转换成功！\n共 {row_count} 行数据\n已保存至：\n{output_path}"))

        except Exception as e:
            self.root.after(0, lambda: self.update_status("转换失败！", False))
            self.root.after(0, lambda: messagebox.showerror("错误", f"转换失败：\n{str(e)}"))

    def update_status(self, message, loading):
        self.status_label.config(text=message)
        if loading:
            self.progress.start(10)
        else:
            self.progress.stop()


def main():
    root = tk.Tk()
    app = DtaToExcelConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
