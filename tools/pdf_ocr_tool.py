import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
from PyPDF2 import PdfReader

class PDFOCRTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_files = []
        self.ocr_output = ""
        
        # 创建界面
        self.create_ocr_interface()
    
    def create_ocr_interface(self):
        """创建OCR识别工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="OCR识别工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="识别扫描PDF中的文字内容", 
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
        
        # OCR选项
        ocr_frame = ttk.LabelFrame(left_frame, text="OCR选项")
        ocr_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 语言选择
        lang_frame = ttk.Frame(ocr_frame)
        lang_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(lang_frame, text="语言:").pack(side=tk.LEFT, padx=(0, 5))
        self.lang_var = tk.StringVar(value="chi_sim")
        lang_combobox = ttk.Combobox(lang_frame, textvariable=self.lang_var, 
                                   values=["chi_sim", "chi_tra", "eng", "jpn", "kor"], state="readonly")
        lang_combobox.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(lang_frame, text="(chi_sim=简体中文, chi_tra=繁体中文, eng=英文)").pack(side=tk.LEFT)
        
        # 右侧OCR结果区域
        right_frame = ttk.LabelFrame(content_frame, text="OCR结果")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # OCR结果文本框
        self.ocr_text = tk.Text(right_frame, wrap=tk.WORD, height=20)
        self.ocr_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        ocr_scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, 
                                    command=self.ocr_text.yview)
        ocr_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        self.ocr_text.configure(yscrollcommand=ocr_scrollbar.set)
        
        # OCR结果操作按钮
        ocr_buttons_frame = ttk.Frame(right_frame)
        ocr_buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        save_text_btn = ttk.Button(ocr_buttons_frame, text="保存文本", 
                                 command=self.save_ocr_text)
        save_text_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        copy_text_btn = ttk.Button(ocr_buttons_frame, text="复制文本", 
                                 command=self.copy_ocr_text)
        copy_text_btn.pack(side=tk.LEFT, padx=(0, 5))
        
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
        
        ocr_btn = ttk.Button(button_frame, text="开始OCR识别", 
                           command=self.start_ocr, style='Action.TButton')
        ocr_btn.pack(side=tk.RIGHT)
        
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
            
            # 清空OCR结果
            self.ocr_text.delete(1.0, tk.END)
            
            # 清空输出目录
            self.output_dir.set("")
            
            # 重置语言选择
            self.lang_var.set("chi_sim")
    
    def start_ocr(self):
        """开始OCR识别"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先添加PDF文件")
            return
        
        try:
            # 执行OCR识别
            self.perform_ocr()
            messagebox.showinfo("成功", "OCR识别已完成")
        except Exception as e:
            messagebox.showerror("错误", f"OCR识别时出错: {str(e)}")
    
    def perform_ocr(self):
        """执行OCR识别的核心功能"""
        # 清空之前的OCR结果
        self.ocr_text.delete(1.0, tk.END)
        
        # 注意：OCR功能需要安装额外的库，如pytesseract和PIL
        # 这里我们简化处理，只提取PDF中的文本内容
        # 实际项目中需要安装并配置Tesseract OCR引擎
        
        self.ocr_text.insert(tk.END, "OCR识别结果\n")
        self.ocr_text.insert(tk.END, "=" * 50 + "\n\n")
        
        for pdf_file in self.selected_files:
            file_name = os.path.basename(pdf_file)
            self.ocr_text.insert(tk.END, f"文件: {file_name}\n")
            self.ocr_text.insert(tk.END, "-" * 30 + "\n")
            
            try:
                with open(pdf_file, 'rb') as f:
                    reader = PdfReader(f)
                    
                    # 提取每一页的文本
                    for page_num, page in enumerate(reader.pages, 1):
                        self.ocr_text.insert(tk.END, f"第 {page_num} 页:\n")
                        
                        # 尝试直接提取文本（对于非扫描PDF）
                        text = page.extract_text()
                        if text:
                            self.ocr_text.insert(tk.END, text + "\n\n")
                        else:
                            self.ocr_text.insert(tk.END, "(无法直接提取文本，该PDF可能是扫描件)\n")
                            self.ocr_text.insert(tk.END, "注意: 完整的OCR功能需要安装Tesseract OCR引擎\n\n")
            except Exception as e:
                self.ocr_text.insert(tk.END, f"处理文件时出错: {str(e)}\n\n")
        
        self.ocr_text.insert(tk.END, "=" * 50 + "\n")
        self.ocr_text.insert(tk.END, "OCR识别完成\n")
    
    def save_ocr_text(self):
        """保存OCR识别结果到文本文件"""
        if not self.ocr_text.get(1.0, tk.END).strip():
            messagebox.showwarning("警告", "没有OCR结果可以保存")
            return
        
        # 如果未选择输出目录，使用默认目录
        if not self.output_dir.get():
            output_dir = os.path.expanduser("~")
        else:
            output_dir = self.output_dir.get()
        
        # 选择保存的文件名
        file_path = filedialog.asksaveasfilename(
            title="保存OCR结果",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            initialdir=output_dir
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.ocr_text.get(1.0, tk.END))
                messagebox.showinfo("成功", f"OCR结果已保存到:\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存OCR结果时出错: {str(e)}")
    
    def copy_ocr_text(self):
        """复制OCR识别结果到剪贴板"""
        if not self.ocr_text.get(1.0, tk.END).strip():
            messagebox.showwarning("警告", "没有OCR结果可以复制")
            return
        
        try:
            # 清除剪贴板
            self.parent.clipboard_clear()
            # 将OCR结果复制到剪贴板
            self.parent.clipboard_append(self.ocr_text.get(1.0, tk.END))
            messagebox.showinfo("成功", "OCR结果已复制到剪贴板")
        except Exception as e:
            messagebox.showerror("错误", f"复制OCR结果时出错: {str(e)}")

# 独立测试函数
def test_ocr_tool():
    """独立测试OCR识别工具"""
    root = tk.Tk()
    root.title("OCR识别工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    ocr_tool = PDFOCRTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_ocr_tool()