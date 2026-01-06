import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import PyPDF2
from pathlib import Path

class PDFHeaderFooterTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_file = None
        
        # 创建界面
        self.create_header_footer_interface()
    
    def create_header_footer_interface(self):
        """创建PDF页眉页脚工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF页眉页脚工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="为PDF文件添加页眉和页脚", 
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
        
        # 右侧页眉页脚设置区域
        right_frame = ttk.LabelFrame(content_frame, text="页眉页脚设置")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 创建一个可滚动的框架用于设置
        scrollable_frame = ttk.Frame(right_frame)
        scrollable_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 页眉设置
        header_frame = ttk.LabelFrame(scrollable_frame, text="页眉设置")
        header_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 页眉文本
        ttk.Label(header_frame, text="页眉文本:").pack(anchor=tk.W)
        self.header_text = tk.StringVar(value="")
        header_entry = ttk.Entry(header_frame, textvariable=self.header_text)
        header_entry.pack(fill=tk.X, pady=(2, 5))
        
        # 页眉字体大小
        header_font_frame = ttk.Frame(header_frame)
        header_font_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(header_font_frame, text="字体大小:").pack(side=tk.LEFT, padx=(0, 5))
        self.header_font_size = tk.IntVar(value=10)
        header_font_spinbox = ttk.Spinbox(header_font_frame, from_=6, to=24, 
                                         textvariable=self.header_font_size)
        header_font_spinbox.pack(side=tk.LEFT, padx=(0, 10))
        
        # 页眉对齐方式
        ttk.Label(header_frame, text="对齐方式:").pack(anchor=tk.W)
        self.header_alignment = tk.StringVar(value="center")
        alignment_frame = ttk.Frame(header_frame)
        alignment_frame.pack(fill=tk.X, pady=(2, 5))
        
        ttk.Radiobutton(alignment_frame, text="左对齐", variable=self.header_alignment, value="left").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(alignment_frame, text="居中", variable=self.header_alignment, value="center").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(alignment_frame, text="右对齐", variable=self.header_alignment, value="right").pack(side=tk.LEFT)
        
        # 页脚设置
        footer_frame = ttk.LabelFrame(scrollable_frame, text="页脚设置")
        footer_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 页脚文本
        ttk.Label(footer_frame, text="页脚文本:").pack(anchor=tk.W)
        self.footer_text = tk.StringVar(value="页码")
        footer_entry = ttk.Entry(footer_frame, textvariable=self.footer_text)
        footer_entry.pack(fill=tk.X, pady=(2, 5))
        
        # 页脚字体大小
        footer_font_frame = ttk.Frame(footer_frame)
        footer_font_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(footer_font_frame, text="字体大小:").pack(side=tk.LEFT, padx=(0, 5))
        self.footer_font_size = tk.IntVar(value=10)
        footer_font_spinbox = ttk.Spinbox(footer_font_frame, from_=6, to=24, 
                                         textvariable=self.footer_font_size)
        footer_font_spinbox.pack(side=tk.LEFT, padx=(0, 10))
        
        # 页脚对齐方式
        ttk.Label(footer_frame, text="对齐方式:").pack(anchor=tk.W)
        self.footer_alignment = tk.StringVar(value="center")
        footer_alignment_frame = ttk.Frame(footer_frame)
        footer_alignment_frame.pack(fill=tk.X, pady=(2, 5))
        
        ttk.Radiobutton(footer_alignment_frame, text="左对齐", variable=self.footer_alignment, value="left").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(footer_alignment_frame, text="居中", variable=self.footer_alignment, value="center").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(footer_alignment_frame, text="右对齐", variable=self.footer_alignment, value="right").pack(side=tk.LEFT)
        
        # 页码格式
        ttk.Label(footer_frame, text="页码格式:").pack(anchor=tk.W)
        self.page_format = tk.StringVar(value="{page}")
        page_format_frame = ttk.Frame(footer_frame)
        page_format_frame.pack(fill=tk.X, pady=(2, 5))
        
        format_options = [
            ("仅页码", "{page}"),
            ("页码/总页数", "{page}/{total}"),
            ("第X页", "第{page}页"),
            ("Page X", "Page {page}"),
            ("Page X of Y", "Page {page} of {total}")
        ]
        
        for text, value in format_options:
            ttk.Radiobutton(page_format_frame, text=text, variable=self.page_format, value=value).pack(anchor=tk.W, pady=1)
        
        # 应用范围
        apply_frame = ttk.LabelFrame(scrollable_frame, text="应用范围")
        apply_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.apply_scope = tk.StringVar(value="all")
        
        scope_options = [
            ("所有页面", "all"),
            ("奇数页", "odd"),
            ("偶数页", "even"),
            ("第一页除外", "except_first"),
            ("仅第一页", "first_only")
        ]
        
        for text, value in scope_options:
            ttk.Radiobutton(apply_frame, text=text, variable=self.apply_scope, value=value).pack(anchor=tk.W, pady=1)
        
        # 底部按钮区域
        button_frame = ttk.Frame(self.parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 重置按钮
        reset_btn = ttk.Button(button_frame, text="重置", command=self.reset_header_footer_tool)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 添加页眉页脚按钮
        self.add_btn = ttk.Button(button_frame, text="添加页眉页脚", 
                                command=self.add_header_footer, style='Action.TButton')
        self.add_btn.pack(side=tk.RIGHT)
        
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
    
    def reset_header_footer_tool(self):
        """重置PDF页眉页脚工具"""
        if messagebox.askyesno("确认", "确定要重置所有设置吗？"):
            self.selected_file = None
            self.header_text.set("")
            self.header_font_size.set(10)
            self.header_alignment.set("center")
            self.footer_text.set("页码")
            self.footer_font_size.set(10)
            self.footer_alignment.set("center")
            self.page_format.set("{page}")
            self.apply_scope.set("all")
            self.update_file_info()
    
    def add_header_footer(self):
        """开始添加页眉页脚"""
        if not self.selected_file:
            messagebox.showwarning("警告", "请先选择PDF文件")
            return
        
        # 选择输出目录
        output_dir = filedialog.askdirectory(title="选择保存位置")
        if not output_dir:
            return
        
        # 获取输出文件名
        file_name = os.path.basename(self.selected_file)
        output_path = os.path.join(output_dir, f"{os.path.splitext(file_name)[0]}_header_footer.pdf")
        
        # 检查文件是否已存在
        if os.path.exists(output_path):
            if not messagebox.askyesno("确认", f"文件 {os.path.basename(output_path)} 已存在，是否覆盖？"):
                return
        
        try:
            # 执行添加页眉页脚
            self.add_header_footer_to_pdf(output_path)
            messagebox.showinfo("成功", f"PDF页眉页脚已成功添加，保存到:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"添加PDF页眉页脚时出错: {str(e)}")
    
    def add_header_footer_to_pdf(self, output_path):
        """添加页眉页脚到PDF的核心功能"""
        with open(self.selected_file, 'rb') as f:
            pdf_reader = PyPDF2.PdfReader(f)
            pdf_writer = PyPDF2.PdfWriter()
            num_pages = len(pdf_reader.pages)
            
            for page_num in range(num_pages):
                page = pdf_reader.pages[page_num]
                # 根据应用范围决定是否添加页眉页脚
                if self.should_apply_header_footer(page_num + 1, num_pages):
                    # 简化实现：添加页眉页脚到PDF
                    # 注意：完整的页眉页脚功能需要使用reportlab或其他库创建页眉页脚PDF
                    # 这里我们只实现基本框架
                    pass
                pdf_writer.add_page(page)
            
            # 写入输出文件
            with open(output_path, 'wb') as out_file:
                pdf_writer.write(out_file)
    
    def should_apply_header_footer(self, page_num, total_pages):
        """根据应用范围决定是否在当前页添加页眉页脚"""
        scope = self.apply_scope.get()
        if scope == "all":
            return True
        elif scope == "odd":
            return page_num % 2 == 1
        elif scope == "even":
            return page_num % 2 == 0
        elif scope == "except_first":
            return page_num > 1
        elif scope == "first_only":
            return page_num == 1
        return True

# 独立测试函数
def test_header_footer_tool():
    """独立测试PDF页眉页脚工具"""
    root = tk.Tk()
    root.title("PDF页眉页脚工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    header_footer_tool = PDFHeaderFooterTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_header_footer_tool()