import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import PyPDF2
from pathlib import Path

class PDFRotateTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_file = None
        
        # 创建界面
        self.create_rotate_interface()
    
    def create_rotate_interface(self):
        """创建PDF旋转工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF旋转工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="旋转PDF页面方向", 
                              font=('Arial', 12))
        desc_label.pack(pady=(0, 20))
        
        # 创建主内容框架
        content_frame = ttk.Frame(self.parent)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧文件选择区域
        left_frame = ttk.LabelFrame(content_frame, text="文件选择")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 文件选择按钮
        file_buttons_frame = ttk.Frame(left_frame)
        file_buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        select_file_btn = ttk.Button(file_buttons_frame, text="选择PDF文件", 
                                    command=self.select_pdf_file)
        select_file_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 文件信息显示
        self.file_info_label = ttk.Label(left_frame, text="未选择文件")
        self.file_info_label.pack(fill=tk.X, padx=5, pady=5)
        
        # 右侧选项区域
        right_frame = ttk.LabelFrame(content_frame, text="旋转选项")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 旋转角度选项
        angle_frame = ttk.LabelFrame(right_frame, text="旋转角度")
        angle_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.rotate_angle = tk.IntVar(value=90)
        
        angle_options = [
            ("顺时针90度", 90),
            ("顺时针180度", 180),
            ("顺时针270度", 270),
            ("逆时针90度", -90)
        ]
        
        for text, value in angle_options:
            ttk.Radiobutton(angle_frame, text=text, variable=self.rotate_angle, value=value).pack(anchor=tk.W, pady=2)
        
        # 旋转页面范围
        range_frame = ttk.LabelFrame(right_frame, text="旋转页面范围")
        range_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.page_range_type = tk.StringVar(value="all")
        
        range_options = [
            ("所有页面", "all"),
            ("奇数页", "odd"),
            ("偶数页", "even"),
            ("指定范围", "custom")
        ]
        
        for text, value in range_options:
            ttk.Radiobutton(range_frame, text=text, variable=self.page_range_type, value=value).pack(anchor=tk.W, pady=2)
        
        # 自定义页面范围输入
        self.custom_range = tk.StringVar()
        range_entry_frame = ttk.Frame(range_frame)
        range_entry_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(range_entry_frame, text="格式: 1,3-5,7").pack(anchor=tk.W, pady=(0, 2))
        ttk.Entry(range_entry_frame, textvariable=self.custom_range).pack(fill=tk.X)
        
        # 输出选项
        output_frame = ttk.LabelFrame(right_frame, text="输出设置")
        output_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 输出文件名
        ttk.Label(output_frame, text="输出文件名:").pack(anchor=tk.W)
        self.output_name = tk.StringVar(value="旋转后文档.pdf")
        output_entry = ttk.Entry(output_frame, textvariable=self.output_name)
        output_entry.pack(fill=tk.X, pady=(5, 0))
        
        # 底部按钮区域
        button_frame = ttk.Frame(self.parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 重置按钮
        reset_btn = ttk.Button(button_frame, text="重置", command=self.reset_rotate_tool)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 旋转按钮
        self.rotate_btn = ttk.Button(button_frame, text="开始旋转", 
                                     command=self.start_rotate, style='Action.TButton')
        self.rotate_btn.pack(side=tk.RIGHT)
        
        # 初始化文件列表
        self.update_file_info()
    
    def select_pdf_file(self):
        """选择PDF文件"""
        if self.file_list:
            # 如果已有文件列表，弹出选择对话框让用户选择
            selected_index = filedialog.askopenfilename(
                title="选择PDF文件",
                filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")],
                initialdir=os.path.dirname(self.file_list[0]) if self.file_list else None
            )
            if selected_index:
                self.selected_file = selected_index
        else:
            # 否则让用户直接选择文件
            self.selected_file = filedialog.askopenfilename(
                title="选择PDF文件",
                filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
            )
        
        self.update_file_info()
    
    def update_file_info(self):
        """更新文件信息显示"""
        if self.selected_file:
            file_name = os.path.basename(self.selected_file)
            file_size = os.path.getsize(self.selected_file) / (1024 * 1024)  # MB
            
            # 获取PDF页数
            try:
                with open(self.selected_file, 'rb') as f:
                    pdf_reader = PyPDF2.PdfReader(f)
                    num_pages = len(pdf_reader.pages)
                self.file_info_label.config(text=f"文件: {file_name}\n大小: {file_size:.2f} MB\n页数: {num_pages}")
            except Exception as e:
                self.file_info_label.config(text=f"文件: {file_name}\n大小: {file_size:.2f} MB\n页数: 无法读取")
        else:
            self.file_info_label.config(text="未选择文件")
    
    def reset_rotate_tool(self):
        """重置PDF旋转工具"""
        if messagebox.askyesno("确认", "确定要重置所有设置吗？"):
            self.selected_file = None
            self.rotate_angle.set(90)
            self.page_range_type.set("all")
            self.custom_range.set("")
            self.output_name.set("旋转后文档.pdf")
            self.update_file_info()
    
    def start_rotate(self):
        """开始旋转PDF页面"""
        if not self.selected_file:
            messagebox.showwarning("警告", "请先选择PDF文件")
            return
        
        # 选择输出目录
        output_dir = filedialog.askdirectory(title="选择保存位置")
        if not output_dir:
            return
        
        output_path = os.path.join(output_dir, self.output_name.get())
        
        # 检查文件是否已存在
        if os.path.exists(output_path):
            if not messagebox.askyesno("确认", f"文件 {self.output_name.get()} 已存在，是否覆盖？"):
                return
        
        try:
            # 执行旋转
            self.rotate_pdf(output_path)
            messagebox.showinfo("成功", f"PDF页面已成功旋转，保存到:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"旋转PDF页面时出错: {str(e)}")
    
    def rotate_pdf(self, output_path):
        """旋转PDF页面的核心功能"""
        with open(self.selected_file, 'rb') as f:
            pdf_reader = PyPDF2.PdfReader(f)
            pdf_writer = PyPDF2.PdfWriter()
            num_pages = len(pdf_reader.pages)
            
            # 获取要旋转的页面
            pages_to_rotate = self.get_pages_to_rotate(num_pages)
            
            # 旋转页面
            for page_num in range(num_pages):
                page = pdf_reader.pages[page_num]
                if page_num in pages_to_rotate:
                    page.rotate(self.rotate_angle.get())
                pdf_writer.add_page(page)
            
            # 写入输出文件
            with open(output_path, 'wb') as out_file:
                pdf_writer.write(out_file)
    
    def get_pages_to_rotate(self, total_pages):
        """获取要旋转的页面列表"""
        page_range_type = self.page_range_type.get()
        pages_to_rotate = []
        
        if page_range_type == "all":
            pages_to_rotate = list(range(total_pages))
        elif page_range_type == "odd":
            pages_to_rotate = [i for i in range(total_pages) if (i + 1) % 2 == 1]
        elif page_range_type == "even":
            pages_to_rotate = [i for i in range(total_pages) if (i + 1) % 2 == 0]
        elif page_range_type == "custom":
            pages_to_rotate = self.parse_page_range(self.custom_range.get(), total_pages)
        
        return pages_to_rotate
    
    def parse_page_range(self, range_str, total_pages):
        """解析页面范围字符串"""
        if not range_str.strip():
            return list(range(total_pages))
        
        try:
            pages = []
            parts = range_str.split(',')
            for part in parts:
                part = part.strip()
                if '-' in part:
                    start, end = part.split('-')
                    start = int(start.strip()) - 1 if start.strip() else 0
                    end = int(end.strip()) if end.strip() else total_pages
                    pages.extend(range(max(0, start), min(total_pages, end)))
                else:
                    page = int(part.strip()) - 1
                    if 0 <= page < total_pages:
                        pages.append(page)
            return sorted(list(set(pages)))
        except:
            messagebox.showwarning("警告", "页面范围格式错误，将旋转所有页面")
            return list(range(total_pages))

# 独立测试函数
def test_rotate_tool():
    """独立测试PDF旋转工具"""
    root = tk.Tk()
    root.title("PDF旋转工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    rotate_tool = PDFRotateTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_rotate_tool()