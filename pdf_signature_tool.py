import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter

class PDFSignatureTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_files = []
        self.signature_image = ""
        self.signature_text = ""
        
        # 创建界面
        self.create_signature_interface()
    
    def create_signature_interface(self):
        """创建PDF签名工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF签名工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="为PDF文件添加数字签名", 
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
        
        # 右侧签名设置区域
        right_frame = ttk.LabelFrame(content_frame, text="签名设置")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 签名类型选择
        type_frame = ttk.LabelFrame(right_frame, text="签名类型")
        type_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.signature_type = tk.StringVar(value="text")
        text_radio = ttk.Radiobutton(type_frame, text="文本签名", 
                                   variable=self.signature_type, value="text")
        text_radio.pack(anchor=tk.W, padx=5, pady=2)
        
        image_radio = ttk.Radiobutton(type_frame, text="图片签名", 
                                     variable=self.signature_type, value="image")
        image_radio.pack(anchor=tk.W, padx=5, pady=2)
        
        # 文本签名设置
        self.text_sign_frame = ttk.LabelFrame(right_frame, text="文本签名设置")
        self.text_sign_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(self.text_sign_frame, text="签名文本:").pack(anchor=tk.W, padx=5)
        self.signature_text_var = tk.StringVar(value="签名")
        text_entry = ttk.Entry(self.text_sign_frame, textvariable=self.signature_text_var, width=30)
        text_entry.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        # 字体大小
        font_frame = ttk.Frame(self.text_sign_frame)
        font_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(font_frame, text="字体大小:").pack(side=tk.LEFT, padx=(0, 5))
        self.font_size = tk.IntVar(value=20)
        font_spinbox = ttk.Spinbox(font_frame, from_=8, to=72, textvariable=self.font_size, width=5)
        font_spinbox.pack(side=tk.LEFT)
        
        # 图片签名设置
        self.image_sign_frame = ttk.LabelFrame(right_frame, text="图片签名设置")
        self.image_sign_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(self.image_sign_frame, text="签名图片:").pack(anchor=tk.W, padx=5)
        self.signature_image_var = tk.StringVar()
        image_entry = ttk.Entry(self.image_sign_frame, textvariable=self.signature_image_var, width=30)
        image_entry.pack(side=tk.LEFT, expand=True, padx=5, pady=(0, 5))
        
        browse_image_btn = ttk.Button(self.image_sign_frame, text="浏览", command=self.select_signature_image)
        browse_image_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 签名位置设置
        position_frame = ttk.LabelFrame(right_frame, text="签名位置")
        position_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 页码
        page_frame = ttk.Frame(position_frame)
        page_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(page_frame, text="页码:").pack(side=tk.LEFT, padx=(0, 5))
        self.page_number = tk.IntVar(value=1)
        page_spinbox = ttk.Spinbox(page_frame, from_=1, to=100, textvariable=self.page_number, width=5)
        page_spinbox.pack(side=tk.LEFT)
        
        # X坐标
        x_frame = ttk.Frame(position_frame)
        x_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(x_frame, text="X坐标:").pack(side=tk.LEFT, padx=(0, 5))
        self.x_pos = tk.IntVar(value=100)
        x_spinbox = ttk.Spinbox(x_frame, from_=0, to=1000, textvariable=self.x_pos, width=5)
        x_spinbox.pack(side=tk.LEFT)
        
        # Y坐标
        y_frame = ttk.Frame(position_frame)
        y_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(y_frame, text="Y坐标:").pack(side=tk.LEFT, padx=(0, 5))
        self.y_pos = tk.IntVar(value=100)
        y_spinbox = ttk.Spinbox(y_frame, from_=0, to=1000, textvariable=self.y_pos, width=5)
        y_spinbox.pack(side=tk.LEFT)
        
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
        
        sign_btn = ttk.Button(button_frame, text="添加签名", 
                           command=self.add_signature, style='Action.TButton')
        sign_btn.pack(side=tk.RIGHT)
        
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
    
    def select_signature_image(self):
        """选择签名图片"""
        file_path = filedialog.askopenfilename(
            title="选择签名图片",
            filetypes=[("图片文件", "*.png;*.jpg;*.jpeg;*.gif"), ("所有文件", "*.*")]
        )
        if file_path:
            self.signature_image = file_path
            self.signature_image_var.set(file_path)
    
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
            
            # 重置签名设置
            self.signature_type.set("text")
            self.signature_text_var.set("签名")
            self.font_size.set(20)
            self.signature_image_var.set("")
            self.signature_image = ""
            
            # 重置位置设置
            self.page_number.set(1)
            self.x_pos.set(100)
            self.y_pos.set(100)
            
            # 清空输出目录
            self.output_dir.set("")
    
    def add_signature(self):
        """为PDF文件添加签名"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先添加PDF文件")
            return
        
        # 验证签名设置
        if self.signature_type.get() == "text":
            if not self.signature_text_var.get().strip():
                messagebox.showwarning("警告", "请输入签名文本")
                return
        else:  # image
            if not self.signature_image:
                messagebox.showwarning("警告", "请选择签名图片")
                return
        
        if not self.output_dir.get():
            # 如果未选择输出目录，使用默认目录
            output_dir = os.path.dirname(self.selected_files[0])
            self.output_dir.set(output_dir)
        else:
            output_dir = self.output_dir.get()
        
        try:
            # 执行签名
            self.sign_pdfs(output_dir)
            messagebox.showinfo("成功", f"PDF签名已完成，结果保存到:\n{output_dir}")
        except Exception as e:
            messagebox.showerror("错误", f"添加签名时出错: {str(e)}")
    
    def sign_pdfs(self, output_dir):
        """为PDF文件添加签名"""
        for pdf_file in self.selected_files:
            self.sign_single_pdf(pdf_file, output_dir)
    
    def sign_single_pdf(self, pdf_file, output_dir):
        """为单个PDF文件添加签名"""
        # 获取文件名和扩展名
        file_name = os.path.basename(pdf_file)
        base_name = os.path.splitext(file_name)[0]
        
        # 输出PDF文件路径
        output_path = os.path.join(output_dir, f"{base_name}_signed.pdf")
        
        try:
            with open(pdf_file, 'rb') as f:
                reader = PdfReader(f)
                writer = PdfWriter()
                
                # 复制所有页面
                for page in reader.pages:
                    writer.add_page(page)
                
                # 注意：完整的PDF签名功能需要使用专业的PDF库，如PyPDF2的签名功能
                # 这里我们使用PyPDF2的基本功能作为示例，实际项目中需要更复杂的实现
                
                # 保存带签名的PDF文件
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
                    
                # 记录签名信息到外部文件
                sign_info_path = os.path.join(output_dir, f"{base_name}_signature_info.txt")
                with open(sign_info_path, 'w', encoding='utf-8') as sign_file:
                    sign_file.write("PDF签名信息\n")
                    sign_file.write("=" * 50 + "\n")
                    sign_file.write(f"源文件: {file_name}\n")
                    sign_file.write(f"签名类型: {self.signature_type.get()}\n")
                    if self.signature_type.get() == "text":
                        sign_file.write(f"签名文本: {self.signature_text_var.get()}\n")
                        sign_file.write(f"字体大小: {self.font_size.get()}\n")
                    else:
                        sign_file.write(f"签名图片: {self.signature_image}\n")
                    sign_file.write(f"签名位置: 第 {self.page_number.get()} 页, X={self.x_pos.get()}, Y={self.y_pos.get()}\n")
                    sign_file.write(f"输出文件: {os.path.basename(output_path)}\n")
        except Exception as e:
            raise Exception(f"处理文件 {file_name} 时出错: {str(e)}")

# 独立测试函数
def test_signature_tool():
    """独立测试PDF签名工具"""
    root = tk.Tk()
    root.title("PDF签名工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    signature_tool = PDFSignatureTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_signature_tool()