import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter

class PDFOptimizeTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_files = []
        
        # 创建界面
        self.create_optimize_interface()
    
    def create_optimize_interface(self):
        """创建PDF优化工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF优化工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="优化PDF文件结构和性能，减小文件大小", 
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
        
        add_file_btn = ttk.Button(file_buttons_frame, text="添加PDF文件", 
                                 command=self.add_pdf_files)
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
        
        # 优化选项
        options_frame = ttk.LabelFrame(left_frame, text="优化选项")
        options_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 优化级别
        ttk.Label(options_frame, text="优化级别:").pack(anchor=tk.W, padx=5, pady=2)
        self.optimize_level = tk.StringVar(value="medium")
        level_frame = ttk.Frame(options_frame)
        level_frame.pack(fill=tk.X, padx=5, pady=2)
        
        low_radio = ttk.Radiobutton(level_frame, text="低 (保留更多质量)", 
                                   variable=self.optimize_level, value="low")
        low_radio.pack(anchor=tk.W)
        
        medium_radio = ttk.Radiobutton(level_frame, text="中 (平衡质量和大小)", 
                                      variable=self.optimize_level, value="medium")
        medium_radio.pack(anchor=tk.W)
        
        high_radio = ttk.Radiobutton(level_frame, text="高 (更小文件大小)", 
                                    variable=self.optimize_level, value="high")
        high_radio.pack(anchor=tk.W)
        
        # 右侧输出信息区域
        right_frame = ttk.LabelFrame(content_frame, text="优化信息")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 输出信息文本框
        self.output_text = tk.Text(right_frame, wrap=tk.WORD, height=20)
        self.output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        output_scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, 
                                      command=self.output_text.yview)
        output_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        self.output_text.configure(yscrollcommand=output_scrollbar.set)
        
        # 底部输出和执行区域
        bottom_frame = ttk.Frame(self.parent)
        bottom_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 输出目录
        output_frame = ttk.LabelFrame(bottom_frame, text="输出设置")
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="输出目录:").pack(anchor=tk.W)
        self.output_dir = tk.StringVar()
        output_dir_entry = ttk.Entry(output_frame, textvariable=self.output_dir)
        output_dir_entry.pack(fill=tk.X, pady=(2, 5), padx=5)
        
        browse_btn = ttk.Button(output_frame, text="浏览...", 
                              command=self.browse_output_dir)
        browse_btn.pack(side=tk.RIGHT, padx=(5, 5))
        
        # 执行按钮
        button_frame = ttk.Frame(bottom_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        reset_btn = ttk.Button(button_frame, text="重置", command=self.reset_tool)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        optimize_btn = ttk.Button(button_frame, text="开始优化", 
                                command=self.start_optimize, style='Action.TButton')
        optimize_btn.pack(side=tk.RIGHT)
        
    def add_pdf_files(self):
        """添加PDF文件"""
        files = filedialog.askopenfilenames(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        for file in files:
            if file not in self.selected_files:
                self.selected_files.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
        
        if files:
            messagebox.showinfo("成功", f"已添加 {len(files)} 个PDF文件")
    
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
            # 清空文件列表
            self.selected_files.clear()
            self.file_listbox.delete(0, tk.END)
            
            # 清空输出信息
            self.output_text.delete(1.0, tk.END)
            
            # 清空输出目录
            self.output_dir.set("")
            
            # 重置优化级别
            self.optimize_level.set("medium")
    
    def start_optimize(self):
        """开始优化PDF文件"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先添加PDF文件")
            return
        
        if not self.output_dir.get():
            # 如果未选择输出目录，使用默认目录
            output_dir = os.path.dirname(self.selected_files[0])
            self.output_dir.set(output_dir)
        else:
            output_dir = self.output_dir.get()
        
        try:
            # 执行优化
            self.optimize_pdfs(output_dir)
            messagebox.showinfo("成功", f"PDF优化已完成，结果保存到:\n{output_dir}")
        except Exception as e:
            messagebox.showerror("错误", f"优化PDF文件时出错: {str(e)}")
    
    def optimize_pdfs(self, output_dir):
        """优化PDF文件"""
        self.output_text.delete(1.0, tk.END)
        
        self.output_text.insert(tk.END, "PDF优化开始\n")
        self.output_text.insert(tk.END, "=" * 50 + "\n\n")
        
        for pdf_file in self.selected_files:
            file_name = os.path.basename(pdf_file)
            self.output_text.insert(tk.END, f"正在优化: {file_name}\n")
            
            # 获取原始文件大小
            original_size = os.path.getsize(pdf_file)
            self.output_text.insert(tk.END, f"原始大小: {self.format_size(original_size)}\n")
            
            try:
                # 执行优化
                optimized_file = self.optimize_single_pdf(pdf_file, output_dir)
                
                # 获取优化后的文件大小
                if os.path.exists(optimized_file):
                    optimized_size = os.path.getsize(optimized_file)
                    self.output_text.insert(tk.END, f"优化后大小: {self.format_size(optimized_size)}\n")
                    
                    # 计算压缩率
                    compression_ratio = (1 - optimized_size / original_size) * 100
                    self.output_text.insert(tk.END, f"压缩率: {compression_ratio:.2f}%\n")
                    self.output_text.insert(tk.END, f"已保存: {self.format_size(original_size - optimized_size)}\n")
                
                self.output_text.insert(tk.END, "优化完成\n\n")
            except Exception as e:
                self.output_text.insert(tk.END, f"优化失败: {str(e)}\n\n")
        
        self.output_text.insert(tk.END, "=" * 50 + "\n")
        self.output_text.insert(tk.END, "所有PDF文件优化完成\n")
    
    def optimize_single_pdf(self, pdf_file, output_dir):
        """优化单个PDF文件"""
        # 获取文件名和扩展名
        file_name = os.path.basename(pdf_file)
        base_name = os.path.splitext(file_name)[0]
        
        # 输出PDF文件路径
        output_path = os.path.join(output_dir, f"{base_name}_optimized.pdf")
        
        try:
            with open(pdf_file, 'rb') as f:
                reader = PdfReader(f)
                writer = PdfWriter()
                
                # 复制所有页面，这本身就是一种简单的优化
                for page in reader.pages:
                    writer.add_page(page)
                
                # 注意：更高级的优化需要使用专业的PDF优化库
                # 这里我们使用PyPDF2的基本功能作为示例
                
                # 保存优化后的PDF文件
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
                    
            return output_path
        except Exception as e:
            raise Exception(f"优化文件时出错: {str(e)}")
    
    def format_size(self, size_bytes):
        """格式化文件大小"""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.2f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.2f} MB"
        else:
            return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"

# 独立测试函数
def test_optimize_tool():
    """独立测试PDF优化工具"""
    root = tk.Tk()
    root.title("PDF优化工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    optimize_tool = PDFOptimizeTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_optimize_tool()