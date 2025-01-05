import tkinter as tk
from tkinter import filedialog
from tkinter import ttk


class GUI:
    def __init__(self):
        self.file_path = None
        self.root = tk.Tk()
        self.root.title("Word 文档格式设置工具")
        self.root.geometry("650x450")  # 调整窗口大小为650x450

        # 文件选择按钮和文件路径显示
        self.upload_button = tk.Button(self.root, text="上传文件", command=self.upload_file)
        self.upload_button.grid(row=0, column=0, pady=10, padx=10, sticky="w")
        self.file_label = tk.Label(self.root, text="文件路径显示", fg="blue")
        self.file_label.grid(row=0, column=1, columnspan=4, sticky="w")

        # 自定义文字样式标题
        self.title_label = tk.Label(self.root, text="自定义各级文字样式", font=("Arial", 12, "bold"), fg="blue")
        self.title_label.grid(row=1, column=0, columnspan=5, pady=10)

        # 字体、字号和加粗标题
        self.font_label = tk.Label(self.root, text="字体", font=("Arial", 10, "bold"), fg="blue")
        self.font_label.grid(row=2, column=1, pady=5)
        self.size_label = tk.Label(self.root, text="字号", font=("Arial", 10, "bold"), fg="blue")
        self.size_label.grid(row=2, column=2, pady=5)
        self.bold_label = tk.Label(self.root, text="加粗", font=("Arial", 10, "bold"), fg="blue")
        self.bold_label.grid(row=2, column=3, pady=5)

        # 创建样式设置行的通用方法
        def create_style_row(row, label_text, default_font, default_size, default_bold, variable_prefix):
            label = tk.Label(self.root, text=label_text, font=("Arial", 10))
            label.grid(row=row, column=0, pady=5, sticky="w")

            font_var = tk.StringVar(value=default_font)
            font_entry = ttk.Combobox(self.root, textvariable=font_var, values=["黑体", "宋体", "楷体", "仿宋"])
            font_entry.grid(row=row, column=1, pady=5)

            size_var = tk.StringVar(value=default_size)
            size_entry = ttk.Combobox(self.root, textvariable=size_var, values=["小二", "三号", "小三", "四号", "小四", "五号", "小五"])
            size_entry.grid(row=row, column=2, pady=5)

            bold_var = tk.StringVar(value=default_bold)
            bold_entry = ttk.Combobox(self.root, textvariable=bold_var, values=["是", "否"])
            bold_entry.grid(row=row, column=3, pady=5)

            # 保存变量引用到实例
            setattr(self, f"{variable_prefix}_font", font_var)
            setattr(self, f"{variable_prefix}_size", size_var)
            setattr(self, f"{variable_prefix}_bold", bold_var)

        # 各级样式设置
        create_style_row(3, "一级标题", "黑体", "三号", "是", "heading1")
        create_style_row(4, "二级标题", "宋体", "四号", "是", "heading2")
        create_style_row(5, "三级标题", "宋体", "小四", "是", "heading3")
        create_style_row(6, "图片说明", "宋体", "五号", "否", "image")
        create_style_row(7, "表格说明", "宋体", "五号", "否", "table")
        create_style_row(8, "参考文献", "宋体", "小四", "否", "reference")

        # 继续按钮
        self.continue_button = tk.Button(self.root, text="继续", command=self.root.quit)
        self.continue_button.grid(row=9, column=0, columnspan=4, pady=10)

    def upload_file(self):
        file_path = filedialog.askopenfilename(title="选择Word文件", filetypes=[("Word Files", "*.doc *.docx")])
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=file_path)

    def get_user_settings(self):
        settings = {
            "heading1_font": self.heading1_font.get(),
            "heading1_size": self.heading1_size.get(),
            "heading1_bold": self.heading1_bold.get(),
            "heading2_font": self.heading2_font.get(),
            "heading2_size": self.heading2_size.get(),
            "heading2_bold": self.heading2_bold.get(),
            "heading3_font": self.heading3_font.get(),
            "heading3_size": self.heading3_size.get(),
            "heading3_bold": self.heading3_bold.get(),
            "image_font": self.image_font.get(),
            "image_size": self.image_size.get(),
            "image_bold": self.image_bold.get(),
            "table_font": self.table_font.get(),
            "table_size": self.table_size.get(),
            "table_bold": self.table_bold.get(),
            "reference_size": self.reference_size.get(),
            "reference_font": self.reference_font.get(),
            "reference_bold": False
        }
        return settings

    def run(self):
        self.root.mainloop()
        return self.file_path
