import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import PyPDF2
from pathlib import Path

class PDFMetadataTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_file = None
        
        # 创建界面
        self.create_metadata_interface()
    
    def create_metadata_interface(self):
        """创建PDF元数据编辑工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF元数据编辑工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="编辑PDF文件的元数据信息", 
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
        
        # 右侧元数据编辑区域
        right_frame = ttk.LabelFrame(content_frame, text="元数据编辑")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 元数据编辑表单
        form_frame = ttk.Frame(right_frame)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建元数据编辑字段
        self.metadata_fields = {
            "Title": tk.StringVar(),
            "Author": tk.StringVar(),
            "Subject": tk.StringVar(),
            "Keywords": tk.StringVar(),
            "Creator": tk.StringVar(),
            "Producer": tk.StringVar(),
            "CreationDate": tk.StringVar(),
            "ModDate": tk.StringVar()
        }
        
        row = 0
        for field_name, var in self.metadata_fields.items():
            ttk.Label(form_frame, text=f"{field_name}:").grid(row=row, column=0, sticky=tk.W, pady=2, padx=5)
            entry = ttk.Entry(form_frame, textvariable=var, width=30)
            entry.grid(row=row, column=1, sticky=tk.EW, pady=2, padx=5)
            row += 1
        
        # 允许表单扩展
        form_frame.columnconfigure(1, weight=1)
        
        # 底部按钮区域
        button_frame = ttk.Frame(self.parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 重置按钮
        reset_btn = ttk.Button(button_frame, text="重置", command=self.reset_metadata_tool)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 保存按钮
        self.save_btn = ttk.Button(button_frame, text="保存修改", 
                                  command=self.save_metadata, style='Action.TButton')
        self.save_btn.pack(side=tk.RIGHT)
        
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
                self.load_metadata()
        else:
            # 否则让用户直接选择文件
            self.selected_file = filedialog.askopenfilename(
                title="选择PDF文件",
                filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
            )
            if self.selected_file:
                self.load_metadata()
        
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
    
    def load_metadata(self):
        """加载PDF文件的元数据"""
        try:
            with open(self.selected_file, 'rb') as f:
                pdf_reader = PyPDF2.PdfReader(f)
                metadata = pdf_reader.metadata
                
                # 填充元数据字段
                for field_name, var in self.metadata_fields.items():
                    if hasattr(metadata, field_name.lower()):
                        value = getattr(metadata, field_name.lower())
                        var.set(value) if value else var.set("")
        except Exception as e:
            messagebox.showerror("错误", f"加载PDF元数据时出错: {str(e)}")
    
    def reset_metadata_tool(self):
        """重置PDF元数据编辑工具"""
        if messagebox.askyesno("确认", "确定要重置所有设置吗？"):
            self.selected_file = None
            for var in self.metadata_fields.values():
                var.set("")
            self.update_file_info()
    
    def save_metadata(self):
        """保存修改后的元数据"""
        if not self.selected_file:
            messagebox.showwarning("警告", "请先选择PDF文件")
            return
        
        # 选择输出目录
        output_dir = filedialog.askdirectory(title="选择保存位置")
        if not output_dir:
            return
        
        # 获取输出文件名
        file_name = os.path.basename(self.selected_file)
        output_path = os.path.join(output_dir, f"{os.path.splitext(file_name)[0]}_metadata_edited.pdf")
        
        # 检查文件是否已存在
        if os.path.exists(output_path):
            if not messagebox.askyesno("确认", f"文件 {os.path.basename(output_path)} 已存在，是否覆盖？"):
                return
        
        try:
            # 执行元数据保存
            self.update_metadata(output_path)
            messagebox.showinfo("成功", f"PDF元数据已成功修改，保存到:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"保存PDF元数据时出错: {str(e)}")
    
    def update_metadata(self, output_path):
        """更新PDF元数据的核心功能"""
        with open(self.selected_file, 'rb') as f:
            pdf_reader = PyPDF2.PdfReader(f)
            pdf_writer = PyPDF2.PdfWriter()
            
            # 添加所有页面
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)
            
            # 创建新的元数据字典
            new_metadata = {}
            for field_name, var in self.metadata_fields.items():
                value = var.get().strip()
                if value:
                    new_metadata[field_name] = value
            
            # 更新元数据
            pdf_writer.add_metadata(new_metadata)
            
            # 写入输出文件
            with open(output_path, 'wb') as out_file:
                pdf_writer.write(out_file)

# 独立测试函数
def test_metadata_tool():
    """独立测试PDF元数据编辑工具"""
    root = tk.Tk()
    root.title("PDF元数据编辑工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    metadata_tool = PDFMetadataTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_metadata_tool()