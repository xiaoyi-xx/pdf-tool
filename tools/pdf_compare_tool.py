import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
from PyPDF2 import PdfReader

class PDFCompareTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.pdf_file1 = ""
        self.pdf_file2 = ""
        
        # 创建界面
        self.create_compare_interface()
    
    def create_compare_interface(self):
        """创建PDF比较工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF比较工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="比较两个PDF文件的差异", 
                              font=('Arial', 12))
        desc_label.pack(pady=(0, 20))
        
        # 创建主内容框架
        content_frame = ttk.Frame(self.parent)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(content_frame, text="文件选择")
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 第一个PDF文件
        file1_frame = ttk.Frame(file_frame)
        file1_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(file1_frame, text="PDF文件1:").pack(side=tk.LEFT, padx=(0, 5))
        self.file1_var = tk.StringVar()
        file1_entry = ttk.Entry(file1_frame, textvariable=self.file1_var, width=50)
        file1_entry.pack(side=tk.LEFT, expand=True, padx=(0, 5))
        
        file1_btn = ttk.Button(file1_frame, text="浏览", command=self.select_file1)
        file1_btn.pack(side=tk.LEFT)
        
        # 第二个PDF文件
        file2_frame = ttk.Frame(file_frame)
        file2_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(file2_frame, text="PDF文件2:").pack(side=tk.LEFT, padx=(0, 5))
        self.file2_var = tk.StringVar()
        file2_entry = ttk.Entry(file2_frame, textvariable=self.file2_var, width=50)
        file2_entry.pack(side=tk.LEFT, expand=True, padx=(0, 5))
        
        file2_btn = ttk.Button(file2_frame, text="浏览", command=self.select_file2)
        file2_btn.pack(side=tk.LEFT)
        
        # 比较选项
        options_frame = ttk.LabelFrame(content_frame, text="比较选项")
        options_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.compare_text = tk.BooleanVar(value=True)
        text_check = ttk.Checkbutton(options_frame, text="比较文本内容", variable=self.compare_text)
        text_check.pack(anchor=tk.W, padx=5, pady=2)
        
        self.compare_pages = tk.BooleanVar(value=True)
        pages_check = ttk.Checkbutton(options_frame, text="比较页数", variable=self.compare_pages)
        pages_check.pack(anchor=tk.W, padx=5, pady=2)
        
        self.compare_metadata = tk.BooleanVar(value=True)
        metadata_check = ttk.Checkbutton(options_frame, text="比较元数据", variable=self.compare_metadata)
        metadata_check.pack(anchor=tk.W, padx=5, pady=2)
        
        # 比较结果显示区域
        result_frame = ttk.LabelFrame(content_frame, text="比较结果")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 结果文本框
        self.result_text = tk.Text(result_frame, wrap=tk.WORD, height=15)
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        result_scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, 
                                       command=self.result_text.yview)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        self.result_text.configure(yscrollcommand=result_scrollbar.set)
        
        # 底部按钮区域
        button_frame = ttk.Frame(self.parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 重置按钮
        reset_btn = ttk.Button(button_frame, text="重置", command=self.reset_tool)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 比较按钮
        compare_btn = ttk.Button(button_frame, text="开始比较", 
                               command=self.start_compare, style='Action.TButton')
        compare_btn.pack(side=tk.RIGHT)
        
    def select_file1(self):
        """选择第一个PDF文件"""
        file_path = filedialog.askopenfilename(
            title="选择第一个PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file_path:
            self.pdf_file1 = file_path
            self.file1_var.set(file_path)
    
    def select_file2(self):
        """选择第二个PDF文件"""
        file_path = filedialog.askopenfilename(
            title="选择第二个PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file_path:
            self.pdf_file2 = file_path
            self.file2_var.set(file_path)
    
    def reset_tool(self):
        """重置工具"""
        self.pdf_file1 = ""
        self.pdf_file2 = ""
        self.file1_var.set("")
        self.file2_var.set("")
        self.result_text.delete(1.0, tk.END)
        self.compare_text.set(True)
        self.compare_pages.set(True)
        self.compare_metadata.set(True)
    
    def start_compare(self):
        """开始比较两个PDF文件"""
        if not self.pdf_file1 or not self.pdf_file2:
            messagebox.showwarning("警告", "请选择两个PDF文件")
            return
        
        try:
            # 执行比较
            self.compare_pdfs()
        except Exception as e:
            messagebox.showerror("错误", f"比较PDF文件时出错: {str(e)}")
    
    def compare_pdfs(self):
        """比较两个PDF文件"""
        self.result_text.delete(1.0, tk.END)
        
        try:
            # 读取两个PDF文件
            with open(self.pdf_file1, 'rb') as f1, open(self.pdf_file2, 'rb') as f2:
                reader1 = PdfReader(f1)
                reader2 = PdfReader(f2)
                
                self.result_text.insert(tk.END, "PDF比较结果\n")
                self.result_text.insert(tk.END, "=" * 50 + "\n\n")
                
                # 比较页数
                if self.compare_pages.get():
                    pages1 = len(reader1.pages)
                    pages2 = len(reader2.pages)
                    self.result_text.insert(tk.END, f"页数比较: {pages1} 页 vs {pages2} 页")
                    if pages1 != pages2:
                        self.result_text.insert(tk.END, " → 页数不同\n")
                    else:
                        self.result_text.insert(tk.END, " → 页数相同\n")
                
                # 比较元数据
                if self.compare_metadata.get():
                    metadata1 = reader1.metadata
                    metadata2 = reader2.metadata
                    
                    self.result_text.insert(tk.END, "\n元数据比较:\n")
                    
                    # 获取所有元数据键
                    keys = set()
                    if metadata1:
                        keys.update(metadata1.keys())
                    if metadata2:
                        keys.update(metadata2.keys())
                    
                    for key in keys:
                        value1 = metadata1.get(key, "") if metadata1 else ""
                        value2 = metadata2.get(key, "") if metadata2 else ""
                        
                        self.result_text.insert(tk.END, f"  {key}: {value1} vs {value2}")
                        if value1 != value2:
                            self.result_text.insert(tk.END, " → 不同\n")
                        else:
                            self.result_text.insert(tk.END, " → 相同\n")
                
                # 比较文本内容
                if self.compare_text.get():
                    self.result_text.insert(tk.END, "\n文本内容比较:\n")
                    
                    # 比较每一页的文本
                    min_pages = min(len(reader1.pages), len(reader2.pages))
                    for i in range(min_pages):
                        page1 = reader1.pages[i]
                        page2 = reader2.pages[i]
                        
                        text1 = page1.extract_text() or ""
                        text2 = page2.extract_text() or ""
                        
                        if text1 != text2:
                            self.result_text.insert(tk.END, f"  第 {i+1} 页: 文本内容不同\n")
                    
                    # 如果页数不同，提示剩余页面
                    if len(reader1.pages) != len(reader2.pages):
                        self.result_text.insert(tk.END, f"  注意: 由于页数不同，只比较了前 {min_pages} 页\n")
                
                self.result_text.insert(tk.END, "\n" + "=" * 50 + "\n")
                self.result_text.insert(tk.END, "比较完成\n")
                
        except Exception as e:
            raise Exception(f"比较PDF文件时出错: {str(e)}")

# 独立测试函数
def test_compare_tool():
    """独立测试PDF比较工具"""
    root = tk.Tk()
    root.title("PDF比较工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    compare_tool = PDFCompareTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_compare_tool()