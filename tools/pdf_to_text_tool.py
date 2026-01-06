import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import PyPDF2
from pathlib import Path

class PDFToTextTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_file = None
        
        # 创建界面
        self.create_pdf_to_text_interface()
    
    def create_pdf_to_text_interface(self):
        """创建PDF转文本工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF转文本工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="从PDF文件中提取文本内容", 
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
        right_frame = ttk.LabelFrame(content_frame, text="转换选项")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 输出选项
        output_frame = ttk.LabelFrame(right_frame, text="输出设置")
        output_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 输出文件名
        ttk.Label(output_frame, text="输出文件名:").pack(anchor=tk.W)
        self.output_name = tk.StringVar(value="输出文本.txt")
        output_entry = ttk.Entry(output_frame, textvariable=self.output_name)
        output_entry.pack(fill=tk.X, pady=(5, 0))
        
        # 转换范围选项
        range_frame = ttk.LabelFrame(right_frame, text="转换范围 (可选)")
        range_frame.pack(fill=tk.X, padx=5, pady=10)
        
        ttk.Label(range_frame, text="格式: 1,3-5,7").pack(anchor=tk.W, pady=(5, 0))
        self.page_range = tk.StringVar()
        range_entry = ttk.Entry(range_frame, textvariable=self.page_range)
        range_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # 提取选项
        extract_frame = ttk.LabelFrame(right_frame, text="提取选项")
        extract_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.preserve_layout = tk.BooleanVar(value=True)
        layout_check = ttk.Checkbutton(extract_frame, text="保留布局结构", 
                                      variable=self.preserve_layout)
        layout_check.pack(anchor=tk.W, pady=2)
        
        # 底部按钮区域
        button_frame = ttk.Frame(self.parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 重置按钮
        reset_btn = ttk.Button(button_frame, text="重置", command=self.reset_pdf_to_text_tool)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 转换按钮
        self.convert_btn = ttk.Button(button_frame, text="开始转换", 
                                     command=self.start_convert, style='Action.TButton')
        self.convert_btn.pack(side=tk.RIGHT)
        
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
    
    def reset_pdf_to_text_tool(self):
        """重置PDF转文本工具"""
        if messagebox.askyesno("确认", "确定要重置所有设置吗？"):
            self.selected_file = None
            self.output_name.set("输出文本.txt")
            self.page_range.set("")
            self.preserve_layout.set(True)
            self.update_file_info()
    
    def start_convert(self):
        """开始转换PDF到文本"""
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
            # 执行转换
            self.convert_pdf_to_text(output_path)
            messagebox.showinfo("成功", f"PDF转文本已完成，保存到:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"转换PDF到文本时出错: {str(e)}")
    
    def convert_pdf_to_text(self, output_path):
        """将PDF转换为文本的核心功能"""
        with open(self.selected_file, 'rb') as f:
            pdf_reader = PyPDF2.PdfReader(f)
            num_pages = len(pdf_reader.pages)
            
            # 解析页面范围
            page_indices = self.parse_page_range(self.page_range.get(), num_pages)
            
            # 提取文本
            text_content = ""
            for page_num in page_indices:
                page = pdf_reader.pages[page_num]
                text_content += f"--- 第 {page_num + 1} 页 ---\n"
                text_content += page.extract_text()
                text_content += "\n\n"
            
            # 保存到文件
            with open(output_path, 'w', encoding='utf-8') as out_file:
                out_file.write(text_content)
    
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
            messagebox.showwarning("警告", "页面范围格式错误，将转换所有页面")
            return list(range(total_pages))

# 独立测试函数
def test_pdf_to_text_tool():
    """独立测试PDF转文本工具"""
    root = tk.Tk()
    root.title("PDF转文本工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    pdf_to_text_tool = PDFToTextTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_pdf_to_text_tool()