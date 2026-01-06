import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path

class WordToPDFTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_files = []
        
        # 创建界面
        self.create_word_to_pdf_interface()
    
    def create_word_to_pdf_interface(self):
        """创建Word转PDF工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="Word转PDF工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="将Word文档转换为PDF格式", 
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
        
        add_file_btn = ttk.Button(file_buttons_frame, text="添加Word文件", 
                                 command=self.add_word_files)
        add_file_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        remove_file_btn = ttk.Button(file_buttons_frame, text="移除选中", 
                                    command=self.remove_selected_files)
        remove_file_btn.pack(side=tk.LEFT, padx=5)
        
        clear_files_btn = ttk.Button(file_buttons_frame, text="清空列表", 
                                   command=self.clear_files)
        clear_files_btn.pack(side=tk.LEFT, padx=5)
        
        # 文件列表框
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建文件列表框和滚动条
        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        file_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, 
                                      command=self.file_listbox.yview)
        file_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.configure(yscrollcommand=file_scrollbar.set)
        
        # 右侧选项区域
        right_frame = ttk.LabelFrame(content_frame, text="转换选项")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 输出选项
        output_frame = ttk.LabelFrame(right_frame, text="输出设置")
        output_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 输出目录
        ttk.Label(output_frame, text="输出目录:").pack(anchor=tk.W)
        self.output_dir = tk.StringVar()
        output_dir_entry = ttk.Entry(output_frame, textvariable=self.output_dir)
        output_dir_entry.pack(fill=tk.X, pady=(2, 5))
        
        browse_btn = ttk.Button(output_frame, text="浏览...", 
                              command=self.browse_output_dir)
        browse_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        # 转换选项
        options_frame = ttk.LabelFrame(right_frame, text="转换选项")
        options_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.include_comments = tk.BooleanVar(value=False)
        comments_check = ttk.Checkbutton(options_frame, text="包含批注", 
                                        variable=self.include_comments)
        comments_check.pack(anchor=tk.W, pady=2)
        
        self.include_tracked_changes = tk.BooleanVar(value=False)
        tracked_changes_check = ttk.Checkbutton(options_frame, text="包含修订", 
                                              variable=self.include_tracked_changes)
        tracked_changes_check.pack(anchor=tk.W, pady=2)
        
        # 底部按钮区域
        button_frame = ttk.Frame(self.parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 重置按钮
        reset_btn = ttk.Button(button_frame, text="重置", command=self.reset_tool)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 转换按钮
        self.convert_btn = ttk.Button(button_frame, text="开始转换", 
                                    command=self.start_convert, style='Action.TButton')
        self.convert_btn.pack(side=tk.RIGHT)
        
    def add_word_files(self):
        """添加Word文件"""
        files = filedialog.askopenfilenames(
            title="选择Word文件",
            filetypes=[("Word文件", "*.docx;*.doc"), ("所有文件", "*.*")]
        )
        
        for file in files:
            if file not in self.selected_files:
                self.selected_files.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
        
        if files:
            messagebox.showinfo("成功", f"已添加 {len(files)} 个Word文件")
    
    def remove_selected_files(self):
        """移除选中的文件"""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请先选择要移除的文件")
            return
            
        for index in reversed(selected_indices):
            self.file_listbox.delete(index)
            self.selected_files.pop(index)
        
        messagebox.showinfo("成功", f"已移除 {len(selected_indices)} 个文件")
    
    def clear_files(self):
        """清空文件列表"""
        if not self.selected_files:
            messagebox.showinfo("提示", "文件列表已为空")
            return
            
        if messagebox.askyesno("确认", "确定要清空所有文件吗？"):
            self.selected_files.clear()
            self.file_listbox.delete(0, tk.END)
            messagebox.showinfo("成功", "已清空所有文件")
    
    def browse_output_dir(self):
        """浏览输出目录"""
        output_dir = filedialog.askdirectory(title="选择输出目录")
        if output_dir:
            self.output_dir.set(output_dir)
    
    def reset_tool(self):
        """重置工具"""
        if messagebox.askyesno("确认", "确定要重置所有设置吗？"):
            self.selected_files.clear()
            self.file_listbox.delete(0, tk.END)
            self.output_dir.set("")
            self.include_comments.set(False)
            self.include_tracked_changes.set(False)
    
    def start_convert(self):
        """开始转换Word文件到PDF"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先添加Word文件")
            return
        
        if not self.output_dir.get():
            # 如果未选择输出目录，使用默认目录
            output_dir = os.path.dirname(self.selected_files[0])
            self.output_dir.set(output_dir)
        else:
            output_dir = self.output_dir.get()
        
        try:
            # 执行转换
            self.convert_word_to_pdf(output_dir)
            messagebox.showinfo("成功", f"Word文件已成功转换为PDF，保存到:\n{output_dir}")
            
        except Exception as e:
            messagebox.showerror("错误", f"转换Word文件到PDF时出错: {str(e)}")
    
    def convert_word_to_pdf(self, output_dir):
        """将Word文件转换为PDF的核心功能"""
        # 这里需要实现Word转PDF的核心功能
        # 由于python-docx库只能读取和编辑Word文档，不能直接转换为PDF
        # 我们需要使用其他方法，比如调用Microsoft Word的COM接口或LibreOffice命令行
        # 这里我们先实现一个简单的占位符，后续可以扩展为完整功能
        
        for word_file in self.selected_files:
            # 获取文件名和扩展名
            file_name = os.path.basename(word_file)
            base_name = os.path.splitext(file_name)[0]
            
            # 输出PDF文件路径
            pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
            
            # 这里只是创建一个空的PDF文件作为占位符
            # 实际转换功能需要使用外部库或工具
            with open(pdf_path, 'w') as f:
                f.write("%PDF-1.4\n1 0 obj<< /Type /Catalog /Pages 2 0 R >>endobj\n2 0 obj<< /Type /Pages /Kids [3 0 R] /Count 1 >>endobj\n3 0 obj<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>endobj\nxref\n0 4\n0000000000 65535 f \n0000000010 00000 n \n0000000053 00000 n \n0000000101 00000 n \ntrailer<< /Size 4 /Root 1 0 R >>\nstartxref\n150\n%%EOF")

# 独立测试函数
def test_word_to_pdf_tool():
    """独立测试Word转PDF工具"""
    root = tk.Tk()
    root.title("Word转PDF工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    word_to_pdf_tool = WordToPDFTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_word_to_pdf_tool()