import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter

class PDFBatchTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_files = []
        self.batch_operations = []
        
        # 创建界面
        self.create_batch_interface()
    
    def create_batch_interface(self):
        """创建批量处理工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="批量处理工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="批量处理多个PDF文件，支持多种操作组合", 
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
        
        add_files_btn = ttk.Button(file_buttons_frame, text="添加文件", 
                                 command=self.add_pdf_files)
        add_files_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        add_folder_btn = ttk.Button(file_buttons_frame, text="添加文件夹", 
                                  command=self.add_folder)
        add_folder_btn.pack(side=tk.LEFT, padx=5)
        
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
        
        # 右侧操作选择区域
        right_frame = ttk.LabelFrame(content_frame, text="操作选择")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 操作列表
        operations_frame = ttk.LabelFrame(right_frame, text="可选操作")
        operations_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=10)
        
        # 操作选项
        self.operations = {
            "compress": tk.BooleanVar(value=False),
            "encrypt": tk.BooleanVar(value=False),
            "decrypt": tk.BooleanVar(value=False),
            "rotate": tk.BooleanVar(value=False),
            "add_watermark": tk.BooleanVar(value=False),
            "add_header_footer": tk.BooleanVar(value=False)
        }
        
        # 操作复选框
        for operation, var in self.operations.items():
            operation_name = self.get_operation_name(operation)
            check_btn = ttk.Checkbutton(operations_frame, text=operation_name, 
                                      variable=var, command=self.on_operation_select)
            check_btn.pack(anchor=tk.W, pady=2)
        
        # 操作参数区域
        self.params_frame = ttk.LabelFrame(right_frame, text="操作参数")
        self.params_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 初始时显示提示信息
        self.operation_params_label = ttk.Label(self.params_frame, text="请选择要执行的操作")
        self.operation_params_label.pack(padx=5, pady=5)
        
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
        
        execute_btn = ttk.Button(button_frame, text="执行批量处理", 
                               command=self.execute_batch, style='Action.TButton')
        execute_btn.pack(side=tk.RIGHT)
        
    def get_operation_name(self, operation):
        """获取操作的中文名称"""
        operation_names = {
            "compress": "PDF压缩",
            "encrypt": "PDF加密",
            "decrypt": "PDF解密",
            "rotate": "PDF旋转",
            "add_watermark": "添加水印",
            "add_header_footer": "添加页眉页脚"
        }
        return operation_names.get(operation, operation)
    
    def on_operation_select(self):
        """当选择操作时，显示对应的参数设置"""
        # 清除当前参数区域的内容
        for widget in self.params_frame.winfo_children():
            widget.destroy()
        
        # 统计已选择的操作
        selected_operations = [op for op, var in self.operations.items() if var.get()]
        
        if not selected_operations:
            self.operation_params_label = ttk.Label(self.params_frame, text="请选择要执行的操作")
            self.operation_params_label.pack(padx=5, pady=5)
            return
        
        # 根据选择的操作显示不同的参数设置
        for operation in selected_operations:
            operation_frame = ttk.LabelFrame(self.params_frame, text=self.get_operation_name(operation))
            operation_frame.pack(fill=tk.X, padx=5, pady=5)
            
            # 根据操作类型显示不同的参数
            if operation == "compress":
                # 压缩级别
                ttk.Label(operation_frame, text="压缩级别:").pack(anchor=tk.W)
                self.compress_level = tk.StringVar(value="medium")
                compress_combobox = ttk.Combobox(operation_frame, textvariable=self.compress_level, 
                                              values=["low", "medium", "high"], state="readonly")
                compress_combobox.pack(fill=tk.X, pady=(2, 5))
            
            elif operation == "encrypt":
                # 密码设置
                ttk.Label(operation_frame, text="用户密码:").pack(anchor=tk.W)
                self.user_password = tk.StringVar()
                user_pass_entry = ttk.Entry(operation_frame, textvariable=self.user_password, show="*")
                user_pass_entry.pack(fill=tk.X, pady=(2, 5))
                
                ttk.Label(operation_frame, text="所有者密码:").pack(anchor=tk.W)
                self.owner_password = tk.StringVar()
                owner_pass_entry = ttk.Entry(operation_frame, textvariable=self.owner_password, show="*")
                owner_pass_entry.pack(fill=tk.X, pady=(2, 5))
            
            elif operation == "decrypt":
                # 解密密码
                ttk.Label(operation_frame, text="解密密码:").pack(anchor=tk.W)
                self.decrypt_password = tk.StringVar()
                decrypt_pass_entry = ttk.Entry(operation_frame, textvariable=self.decrypt_password, show="*")
                decrypt_pass_entry.pack(fill=tk.X, pady=(2, 5))
            
            elif operation == "rotate":
                # 旋转角度
                ttk.Label(operation_frame, text="旋转角度:").pack(anchor=tk.W)
                self.rotate_angle = tk.IntVar(value=90)
                angle_combobox = ttk.Combobox(operation_frame, textvariable=self.rotate_angle, 
                                            values=[0, 90, 180, 270], state="readonly")
                angle_combobox.pack(fill=tk.X, pady=(2, 5))
            
            elif operation == "add_watermark":
                # 水印类型
                ttk.Label(operation_frame, text="水印类型:").pack(anchor=tk.W)
                self.watermark_type = tk.StringVar(value="text")
                watermark_combobox = ttk.Combobox(operation_frame, textvariable=self.watermark_type, 
                                               values=["text", "image"], state="readonly")
                watermark_combobox.pack(fill=tk.X, pady=(2, 5))
                
                # 水印文本
                ttk.Label(operation_frame, text="水印文本:").pack(anchor=tk.W)
                self.watermark_text = tk.StringVar(value="水印")
                watermark_entry = ttk.Entry(operation_frame, textvariable=self.watermark_text)
                watermark_entry.pack(fill=tk.X, pady=(2, 5))
            
            elif operation == "add_header_footer":
                # 页眉文本
                ttk.Label(operation_frame, text="页眉文本:").pack(anchor=tk.W)
                self.header_text = tk.StringVar()
                header_entry = ttk.Entry(operation_frame, textvariable=self.header_text)
                header_entry.pack(fill=tk.X, pady=(2, 5))
                
                # 页脚文本
                ttk.Label(operation_frame, text="页脚文本:").pack(anchor=tk.W)
                self.footer_text = tk.StringVar()
                footer_entry = ttk.Entry(operation_frame, textvariable=self.footer_text)
                footer_entry.pack(fill=tk.X, pady=(2, 5))
    
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
    
    def add_folder(self):
        """添加文件夹中的所有PDF文件"""
        folder_path = filedialog.askdirectory(title="选择包含PDF文件的文件夹")
        if folder_path:
            # 遍历文件夹中的所有PDF文件
            pdf_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) 
                        if f.lower().endswith('.pdf')]
            
            added_count = 0
            for file in pdf_files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    self.file_listbox.insert(tk.END, os.path.basename(file))
                    added_count += 1
            
            if added_count > 0:
                messagebox.showinfo("成功", f"已从文件夹添加 {added_count} 个PDF文件")
            else:
                messagebox.showinfo("提示", "所选文件夹中没有PDF文件")
    
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
            
            # 重置操作选择
            for var in self.operations.values():
                var.set(False)
            
            # 重置参数区域
            self.on_operation_select()
            
            # 清空输出目录
            self.output_dir.set("")
    
    def execute_batch(self):
        """执行批量处理"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先添加PDF文件")
            return
        
        # 统计已选择的操作
        selected_operations = [op for op, var in self.operations.items() if var.get()]
        if not selected_operations:
            messagebox.showwarning("警告", "请选择要执行的操作")
            return
        
        if not self.output_dir.get():
            # 如果未选择输出目录，使用默认目录
            output_dir = os.path.dirname(self.selected_files[0])
            self.output_dir.set(output_dir)
        else:
            output_dir = self.output_dir.get()
        
        try:
            # 执行批量处理
            self.process_batch_files(selected_operations, output_dir)
            messagebox.showinfo("成功", f"批量处理已完成，结果保存到:\n{output_dir}")
        except Exception as e:
            messagebox.showerror("错误", f"批量处理时出错: {str(e)}")
    
    def process_batch_files(self, operations, output_dir):
        """批量处理PDF文件"""
        for pdf_file in self.selected_files:
            try:
                # 复制原始文件到临时文件，然后依次应用操作
                temp_file = pdf_file
                
                # 执行每个操作
                for operation in operations:
                    output_file = self.apply_operation(temp_file, operation, output_dir)
                    temp_file = output_file
                
                # 最终输出文件
                file_name = os.path.basename(pdf_file)
                base_name = os.path.splitext(file_name)[0]
                final_output = os.path.join(output_dir, f"{base_name}_processed.pdf")
                
                # 复制最终结果
                os.replace(temp_file, final_output)
                
            except Exception as e:
                messagebox.showwarning("警告", f"处理文件 {os.path.basename(pdf_file)} 时出错: {str(e)}")
    
    def apply_operation(self, pdf_file, operation, output_dir):
        """应用单个操作到PDF文件"""
        # 获取文件名和扩展名
        file_name = os.path.basename(pdf_file)
        base_name = os.path.splitext(file_name)[0]
        
        # 输出PDF文件路径
        output_path = os.path.join(output_dir, f"{base_name}_{operation}.pdf")
        
        try:
            with open(pdf_file, 'rb') as f:
                reader = PdfReader(f)
                writer = PdfWriter()
                
                # 复制所有页面
                for page in reader.pages:
                    writer.add_page(page)
                
                # 根据操作类型执行不同的处理
                # 注意：这里只是简化的实现，实际项目中需要调用对应的工具模块
                
                # 保存处理后的PDF文件
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
                    
            return output_path
        except Exception as e:
            raise Exception(f"执行 {operation} 操作时出错: {str(e)}")

# 独立测试函数
def test_batch_tool():
    """独立测试批量处理工具"""
    root = tk.Tk()
    root.title("批量处理工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    batch_tool = PDFBatchTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_batch_tool()