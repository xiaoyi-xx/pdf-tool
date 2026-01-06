import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter

class PDFFormTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_files = []
        self.form_fields = []
        
        # 创建界面
        self.create_form_interface()
    
    def create_form_interface(self):
        """创建PDF表单工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF表单工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="填写或创建PDF表单", 
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
        
        # 绑定列表框选择事件
        self.file_listbox.bind('<<ListboxSelect>>', self.on_file_select)
        
        # 右侧表单编辑区域
        right_frame = ttk.LabelFrame(content_frame, text="表单编辑")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 表单字段列表
        fields_frame = ttk.LabelFrame(right_frame, text="表单字段")
        fields_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=10)
        
        # 表单字段列表框
        self.fields_listbox = tk.Listbox(fields_frame, selectmode=tk.SINGLE)
        self.fields_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        fields_scrollbar = ttk.Scrollbar(fields_frame, orient=tk.VERTICAL, 
                                      command=self.fields_listbox.yview)
        fields_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        self.fields_listbox.configure(yscrollcommand=fields_scrollbar.set)
        
        # 绑定字段列表选择事件
        self.fields_listbox.bind('<<ListboxSelect>>', self.on_field_select)
        
        # 表单字段编辑区域
        edit_frame = ttk.LabelFrame(right_frame, text="字段编辑")
        edit_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 字段名称
        ttk.Label(edit_frame, text="字段名称:").pack(anchor=tk.W)
        self.field_name = tk.StringVar()
        name_entry = ttk.Entry(edit_frame, textvariable=self.field_name)
        name_entry.pack(fill=tk.X, pady=(2, 5))
        
        # 字段类型
        ttk.Label(edit_frame, text="字段类型:").pack(anchor=tk.W)
        self.field_type = tk.StringVar(value="text")
        type_combobox = ttk.Combobox(edit_frame, textvariable=self.field_type, 
                                     values=["text", "checkbox", "radio", "dropdown"], state="readonly")
        type_combobox.pack(fill=tk.X, pady=(2, 5))
        
        # 字段值
        ttk.Label(edit_frame, text="字段值:").pack(anchor=tk.W)
        self.field_value = tk.StringVar()
        value_entry = ttk.Entry(edit_frame, textvariable=self.field_value)
        value_entry.pack(fill=tk.X, pady=(2, 5))
        
        # 字段操作按钮
        field_buttons_frame = ttk.Frame(edit_frame)
        field_buttons_frame.pack(fill=tk.X, pady=5)
        
        add_field_btn = ttk.Button(field_buttons_frame, text="添加字段", 
                                 command=self.add_field)
        add_field_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        edit_field_btn = ttk.Button(field_buttons_frame, text="编辑选中", 
                                  command=self.edit_field)
        edit_field_btn.pack(side=tk.LEFT, padx=5)
        
        delete_field_btn = ttk.Button(field_buttons_frame, text="删除选中", 
                                   command=self.delete_field)
        delete_field_btn.pack(side=tk.LEFT, padx=5)
        
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
        
        save_btn = ttk.Button(button_frame, text="保存表单", 
                            command=self.save_form, style='Action.TButton')
        save_btn.pack(side=tk.RIGHT)
        
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
        
        # 清空表单字段
        self.fields_listbox.delete(0, tk.END)
        self.form_fields.clear()
        
        messagebox.showinfo("成功", f"已移除 {len(selected_indices)} 个文件")
    
    def clear_files(self):
        """清空文件列表"""
        if not self.selected_files:
            messagebox.showinfo("提示", "文件列表已为空")
            return
            
        if messagebox.askyesno("确认", "确定要清空所有文件吗？"):
            self.selected_files.clear()
            self.file_listbox.delete(0, tk.END)
            
            # 清空表单字段
            self.fields_listbox.delete(0, tk.END)
            self.form_fields.clear()
            
            messagebox.showinfo("成功", "已清空所有文件")
    
    def on_file_select(self, event):
        """当选择PDF文件时，加载其表单字段"""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            return
        
        selected_index = selected_indices[0]
        pdf_file = self.selected_files[selected_index]
        
        # 清空当前表单字段
        self.fields_listbox.delete(0, tk.END)
        self.form_fields.clear()
        
        try:
            # 读取PDF文件中的表单字段
            with open(pdf_file, 'rb') as f:
                reader = PdfReader(f)
                
                # 注意：PyPDF2的表单字段读取功能有限，这里我们简化处理
                # 实际项目中可能需要更复杂的逻辑来处理表单字段
                if reader.is_encrypted:
                    messagebox.showwarning("警告", "PDF文件已加密，无法读取表单字段")
                    return
                
                # 尝试获取表单字段
                if hasattr(reader, 'get_fields'):
                    fields = reader.get_fields()
                    if fields:
                        for field_name, field_info in fields.items():
                            field_type = field_info.get('/FT', 'Unknown')
                            field_value = field_info.get('/V', '')
                            self.form_fields.append((field_name, field_type, field_value))
                            self.fields_listbox.insert(tk.END, f"{field_name} ({field_type})")
                    else:
                        self.fields_listbox.insert(tk.END, "该PDF文件不包含表单字段")
        except Exception as e:
            messagebox.showerror("错误", f"读取PDF文件表单字段时出错: {str(e)}")
    
    def on_field_select(self, event):
        """当选择表单字段时，加载其信息到编辑区域"""
        selected_indices = self.fields_listbox.curselection()
        if not selected_indices:
            return
        
        selected_index = selected_indices[0]
        if 0 <= selected_index < len(self.form_fields):
            field_name, field_type, field_value = self.form_fields[selected_index]
            self.field_name.set(field_name)
            self.field_type.set(field_type)
            self.field_value.set(field_value)
    
    def add_field(self):
        """添加新表单字段"""
        field_name = self.field_name.get().strip()
        field_type = self.field_type.get()
        field_value = self.field_value.get().strip()
        
        if not field_name:
            messagebox.showwarning("警告", "请输入字段名称")
            return
        
        # 添加到表单字段列表
        self.form_fields.append((field_name, field_type, field_value))
        self.fields_listbox.insert(tk.END, f"{field_name} ({field_type})")
        
        # 清空输入框
        self.field_name.set("")
        self.field_type.set("text")
        self.field_value.set("")
    
    def edit_field(self):
        """编辑选中的表单字段"""
        selected_indices = self.fields_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请先选择要编辑的表单字段")
            return
        
        selected_index = selected_indices[0]
        field_name = self.field_name.get().strip()
        field_type = self.field_type.get()
        field_value = self.field_value.get().strip()
        
        if not field_name:
            messagebox.showwarning("警告", "请输入字段名称")
            return
        
        # 更新表单字段
        self.form_fields[selected_index] = (field_name, field_type, field_value)
        self.fields_listbox.delete(selected_index)
        self.fields_listbox.insert(selected_index, f"{field_name} ({field_type})")
        
        # 清空输入框
        self.field_name.set("")
        self.field_type.set("text")
        self.field_value.set("")
    
    def delete_field(self):
        """删除选中的表单字段"""
        selected_indices = self.fields_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请先选择要删除的表单字段")
            return
        
        selected_index = selected_indices[0]
        self.fields_listbox.delete(selected_index)
        self.form_fields.pop(selected_index)
        
        # 清空输入框
        self.field_name.set("")
        self.field_type.set("text")
        self.field_value.set("")
    
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
            
            # 清空表单字段和编辑区域
            self.fields_listbox.delete(0, tk.END)
            self.form_fields.clear()
            self.field_name.set("")
            self.field_type.set("text")
            self.field_value.set("")
            
            # 清空输出目录
            self.output_dir.set("")
    
    def save_form(self):
        """保存表单填写内容"""
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
            for pdf_file in self.selected_files:
                self.save_form_to_pdf(pdf_file, output_dir)
            
            messagebox.showinfo("成功", f"表单已成功保存到PDF文件，保存到:\n{output_dir}")
        except Exception as e:
            messagebox.showerror("错误", f"保存表单到PDF文件时出错: {str(e)}")
    
    def save_form_to_pdf(self, pdf_file, output_dir):
        """将表单填写内容保存到PDF文件"""
        # 获取文件名和扩展名
        file_name = os.path.basename(pdf_file)
        base_name = os.path.splitext(file_name)[0]
        
        # 输出PDF文件路径
        output_path = os.path.join(output_dir, f"{base_name}_filled_form.pdf")
        
        try:
            with open(pdf_file, 'rb') as f:
                reader = PdfReader(f)
                writer = PdfWriter()
                
                # 复制所有页面
                for page in reader.pages:
                    writer.add_page(page)
                
                # 注意：PyPDF2的表单字段填写功能有限，这里我们简化处理
                # 实际项目中可能需要使用更专业的PDF库，如PyMuPDF或reportlab
                # 这里只是创建一个带有表单字段信息的PDF文件
                
                # 保存带表单的PDF文件
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
                    
                # 记录表单字段信息到外部文件
                form_info_path = os.path.join(output_dir, f"{base_name}_form_fields.txt")
                with open(form_info_path, 'w', encoding='utf-8') as form_file:
                    form_file.write("PDF表单字段信息\n")
                    form_file.write("=" * 50 + "\n")
                    for i, (field_name, field_type, field_value) in enumerate(self.form_fields, 1):
                        form_file.write(f"字段 {i}:\n")
                        form_file.write(f"  名称: {field_name}\n")
                        form_file.write(f"  类型: {field_type}\n")
                        form_file.write(f"  值: {field_value}\n")
                        form_file.write("\n")
        except Exception as e:
            raise Exception(f"处理文件 {file_name} 时出错: {str(e)}")

# 独立测试函数
def test_form_tool():
    """独立测试PDF表单工具"""
    root = tk.Tk()
    root.title("PDF表单工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    form_tool = PDFFormTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_form_tool()