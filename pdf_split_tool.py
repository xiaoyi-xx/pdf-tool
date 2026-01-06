import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import PyPDF2
from pathlib import Path

class PDFSplitTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_file = None
        
        # 创建界面
        self.create_split_interface()
    
    def create_split_interface(self):
        """创建分割工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF分割工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="将PDF文件分割为多个部分", 
                              font=('Arial', 12))
        desc_label.pack(pady=(0, 20))
        
        # 创建主内容框架
        content_frame = ttk.Frame(self.parent)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧文件选择和预览区域
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
        
        # 页面预览区域
        preview_frame = ttk.LabelFrame(left_frame, text="页面预览")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 简单的页面列表预览
        self.preview_text = tk.Text(preview_frame, height=10, wrap=tk.WORD)
        preview_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, 
                                         command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=preview_scrollbar.set)
        
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 右侧分割选项区域
        right_frame = ttk.LabelFrame(content_frame, text="分割选项")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 分割模式选择
        mode_frame = ttk.LabelFrame(right_frame, text="分割模式")
        mode_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.split_mode = tk.StringVar(value="every_page")
        
        ttk.Radiobutton(mode_frame, text="每页一个文件", 
                       variable=self.split_mode, value="every_page").pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(mode_frame, text="按页面范围", 
                       variable=self.split_mode, value="page_range").pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(mode_frame, text="按固定页数", 
                       variable=self.split_mode, value="fixed_pages").pack(anchor=tk.W, pady=2)
        
        # 页面范围输入
        range_frame = ttk.LabelFrame(right_frame, text="页面范围")
        range_frame.pack(fill=tk.X, padx=5, pady=10)
        
        ttk.Label(range_frame, text="格式: 1-3,5,7-9").pack(anchor=tk.W, pady=(5, 0))
        self.page_range = tk.StringVar()
        range_entry = ttk.Entry(range_frame, textvariable=self.page_range)
        range_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # 固定页数输入
        fixed_frame = ttk.LabelFrame(right_frame, text="每份页数")
        fixed_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.pages_per_split = tk.StringVar(value="1")
        fixed_entry = ttk.Entry(fixed_frame, textvariable=self.pages_per_split)
        fixed_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # 输出选项
        output_frame = ttk.LabelFrame(right_frame, text="输出选项")
        output_frame.pack(fill=tk.X, padx=5, pady=10)
        
        ttk.Label(output_frame, text="输出文件名前缀:").pack(anchor=tk.W)
        self.output_prefix = tk.StringVar(value="分割文档_")
        prefix_entry = ttk.Entry(output_frame, textvariable=self.output_prefix)
        prefix_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # 底部按钮区域
        button_frame = ttk.Frame(self.parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 重置按钮
        reset_btn = ttk.Button(button_frame, text="重置", command=self.reset_split_tool)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 分割按钮
        self.split_btn = ttk.Button(button_frame, text="开始分割", 
                                   command=self.start_split, style='Action.TButton')
        self.split_btn.pack(side=tk.RIGHT)
        
        # 初始化界面状态
        self.update_interface_state()
    
    def select_pdf_file(self):
        """选择PDF文件"""
        file = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if file:
            self.selected_file = file
            self.update_file_info()
            self.update_preview()
    
    def update_file_info(self):
        """更新文件信息显示"""
        if self.selected_file:
            try:
                with open(self.selected_file, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    page_count = len(pdf_reader.pages)
                    
                    file_name = os.path.basename(self.selected_file)
                    file_info = f"文件: {file_name}\n页数: {page_count}"
                    self.file_info_label.config(text=file_info)
                    
            except Exception as e:
                messagebox.showerror("错误", f"读取PDF文件时出错: {str(e)}")
        else:
            self.file_info_label.config(text="未选择文件")
    
    def update_preview(self):
        """更新页面预览"""
        self.preview_text.delete(1.0, tk.END)
        
        if not self.selected_file:
            self.preview_text.insert(tk.END, "请先选择PDF文件")
            return
        
        try:
            with open(self.selected_file, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                page_count = len(pdf_reader.pages)
                
                preview_text = f"PDF文件预览:\n\n"
                preview_text += f"总页数: {page_count}\n\n"
                preview_text += "页面列表:\n"
                
                # 显示前10页的预览信息
                for i in range(min(10, page_count)):
                    page = pdf_reader.pages[i]
                    text = page.extract_text()[:100]  # 只取前100个字符
                    preview_text += f"第{i+1}页: {text}...\n"
                
                if page_count > 10:
                    preview_text += f"... 还有 {page_count - 10} 页\n"
                
                self.preview_text.insert(tk.END, preview_text)
                
        except Exception as e:
            self.preview_text.insert(tk.END, f"预览时出错: {str(e)}")
    
    def update_interface_state(self):
        """根据分割模式更新界面状态"""
        mode = self.split_mode.get()
        
        # 根据模式启用/禁用相关控件
        if mode == "page_range":
            # 启用页面范围，禁用固定页数
            for widget in self.parent.winfo_children():
                if isinstance(widget, ttk.LabelFrame) and widget.cget("text") == "页面范围":
                    for child in widget.winfo_children():
                        child.configure(state="normal")
                elif isinstance(widget, ttk.LabelFrame) and widget.cget("text") == "每份页数":
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Entry):
                            child.configure(state="disabled")
        elif mode == "fixed_pages":
            # 启用固定页数，禁用页面范围
            for widget in self.parent.winfo_children():
                if isinstance(widget, ttk.LabelFrame) and widget.cget("text") == "页面范围":
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Entry):
                            child.configure(state="disabled")
                elif isinstance(widget, ttk.LabelFrame) and widget.cget("text") == "每份页数":
                    for child in widget.winfo_children():
                        child.configure(state="normal")
        else:  # every_page
            # 禁用页面范围和固定页数
            for widget in self.parent.winfo_children():
                if isinstance(widget, ttk.LabelFrame) and widget.cget("text") == "页面范围":
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Entry):
                            child.configure(state="disabled")
                elif isinstance(widget, ttk.LabelFrame) and widget.cget("text") == "每份页数":
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Entry):
                            child.configure(state="disabled")
    
    def reset_split_tool(self):
        """重置分割工具"""
        if messagebox.askyesno("确认", "确定要重置所有设置吗？"):
            self.selected_file = None
            self.split_mode.set("every_page")
            self.page_range.set("")
            self.pages_per_split.set("1")
            self.output_prefix.set("分割文档_")
            self.update_file_info()
            self.update_preview()
            self.update_interface_state()
    
    def start_split(self):
        """开始分割PDF文件"""
        if not self.selected_file:
            messagebox.showwarning("警告", "请先选择要分割的PDF文件")
            return
        
        # 选择输出目录
        output_dir = filedialog.askdirectory(title="选择保存位置")
        if not output_dir:
            return
        
        try:
            # 执行分割
            result = self.split_pdf(output_dir)
            messagebox.showinfo("成功", f"PDF文件已成功分割\n生成了 {result} 个文件")
            
        except Exception as e:
            messagebox.showerror("错误", f"分割PDF时出错: {str(e)}")
    
    def split_pdf(self, output_dir):
        """分割PDF文件的核心功能"""
        with open(self.selected_file, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            total_pages = len(pdf_reader.pages)
            
            mode = self.split_mode.get()
            
            if mode == "every_page":
                # 每页一个文件
                return self.split_every_page(pdf_reader, output_dir, total_pages)
            elif mode == "page_range":
                # 按页面范围分割
                return self.split_by_range(pdf_reader, output_dir, total_pages)
            elif mode == "fixed_pages":
                # 按固定页数分割
                return self.split_by_fixed_pages(pdf_reader, output_dir, total_pages)
    
    def split_every_page(self, pdf_reader, output_dir, total_pages):
        """每页一个文件的分割方式"""
        for page_num in range(total_pages):
            pdf_writer = PyPDF2.PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[page_num])
            
            output_filename = f"{self.output_prefix.get()}{page_num + 1}.pdf"
            output_path = os.path.join(output_dir, output_filename)
            
            with open(output_path, 'wb') as output_file:
                pdf_writer.write(output_file)
        
        return total_pages
    
    def split_by_range(self, pdf_reader, output_dir, total_pages):
        """按页面范围分割"""
        range_str = self.page_range.get().strip()
        if not range_str:
            messagebox.showwarning("警告", "请输入页面范围")
            return 0
        
        # 解析页面范围
        ranges = self.parse_page_ranges(range_str, total_pages)
        if not ranges:
            return 0
        
        file_count = 0
        for i, (start, end) in enumerate(ranges):
            pdf_writer = PyPDF2.PdfWriter()
            
            for page_num in range(start, end + 1):
                pdf_writer.add_page(pdf_reader.pages[page_num])
            
            output_filename = f"{self.output_prefix.get()}{i + 1}.pdf"
            output_path = os.path.join(output_dir, output_filename)
            
            with open(output_path, 'wb') as output_file:
                pdf_writer.write(output_file)
            
            file_count += 1
        
        return file_count
    
    def split_by_fixed_pages(self, pdf_reader, output_dir, total_pages):
        """按固定页数分割"""
        try:
            pages_per_split = int(self.pages_per_split.get())
            if pages_per_split <= 0:
                messagebox.showwarning("警告", "每份页数必须大于0")
                return 0
        except ValueError:
            messagebox.showwarning("警告", "请输入有效的页数")
            return 0
        
        file_count = 0
        for i in range(0, total_pages, pages_per_split):
            pdf_writer = PyPDF2.PdfWriter()
            
            end_page = min(i + pages_per_split, total_pages)
            for page_num in range(i, end_page):
                pdf_writer.add_page(pdf_reader.pages[page_num])
            
            output_filename = f"{self.output_prefix.get()}{file_count + 1}.pdf"
            output_path = os.path.join(output_dir, output_filename)
            
            with open(output_path, 'wb') as output_file:
                pdf_writer.write(output_file)
            
            file_count += 1
        
        return file_count
    
    def parse_page_ranges(self, range_str, total_pages):
        """解析页面范围字符串"""
        try:
            ranges = []
            parts = range_str.split(',')
            
            for part in parts:
                part = part.strip()
                if '-' in part:
                    start_end = part.split('-')
                    if len(start_end) != 2:
                        raise ValueError(f"无效的范围格式: {part}")
                    
                    start = int(start_end[0].strip()) - 1  # 转换为0-based索引
                    end = int(start_end[1].strip()) - 1   # 转换为0-based索引
                    
                    # 验证范围
                    if start < 0 or end >= total_pages or start > end:
                        raise ValueError(f"页面范围超出有效范围: {part}")
                    
                    ranges.append((start, end))
                else:
                    # 单个页面
                    page = int(part) - 1  # 转换为0-based索引
                    if page < 0 or page >= total_pages:
                        raise ValueError(f"页面超出有效范围: {part}")
                    
                    ranges.append((page, page))
            
            return ranges
        except Exception as e:
            messagebox.showerror("错误", f"解析页面范围时出错: {str(e)}")
            return None

# 独立测试函数
def test_split_tool():
    """独立测试分割工具"""
    root = tk.Tk()
    root.title("PDF分割工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    split_tool = PDFSplitTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_split_tool()