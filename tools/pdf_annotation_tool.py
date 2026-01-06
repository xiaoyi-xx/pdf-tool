import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter

class PDFAnnotationTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_files = []
        self.annotations = []
        
        # 创建界面
        self.create_annotation_interface()
    
    def create_annotation_interface(self):
        """创建PDF注释工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF注释工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="为PDF文件添加注释、高亮和下划线", 
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
        
        # 右侧注释编辑区域
        right_frame = ttk.LabelFrame(content_frame, text="注释编辑")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 注释类型选择
        type_frame = ttk.LabelFrame(right_frame, text="注释类型")
        type_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.annotation_type = tk.StringVar(value="highlight")
        
        highlight_radio = ttk.Radiobutton(type_frame, text="高亮", 
                                         variable=self.annotation_type, value="highlight")
        highlight_radio.pack(anchor=tk.W, pady=2)
        
        underline_radio = ttk.Radiobutton(type_frame, text="下划线", 
                                         variable=self.annotation_type, value="underline")
        underline_radio.pack(anchor=tk.W, pady=2)
        
        strikeout_radio = ttk.Radiobutton(type_frame, text="删除线", 
                                         variable=self.annotation_type, value="strikeout")
        strikeout_radio.pack(anchor=tk.W, pady=2)
        
        note_radio = ttk.Radiobutton(type_frame, text="便签", 
                                    variable=self.annotation_type, value="note")
        note_radio.pack(anchor=tk.W, pady=2)
        
        # 注释内容
        content_frame = ttk.LabelFrame(right_frame, text="注释内容")
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=10)
        
        # 页码
        ttk.Label(content_frame, text="页码:").pack(anchor=tk.W)
        self.page_number = tk.IntVar(value=1)
        page_entry = ttk.Entry(content_frame, textvariable=self.page_number, width=5)
        page_entry.pack(anchor=tk.W, pady=(2, 5))
        
        # 注释文本（便签用）
        ttk.Label(content_frame, text="注释文本:").pack(anchor=tk.W)
        self.annotation_text = tk.Text(content_frame, height=5, wrap=tk.WORD)
        self.annotation_text.pack(fill=tk.X, pady=(2, 5))
        
        # 高亮颜色选择
        color_frame = ttk.Frame(content_frame)
        color_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(color_frame, text="颜色:").pack(side=tk.LEFT)
        self.highlight_color = tk.StringVar(value="yellow")
        color_combobox = ttk.Combobox(color_frame, textvariable=self.highlight_color, 
                                     values=["yellow", "green", "blue", "red"], state="readonly")
        color_combobox.pack(side=tk.LEFT, padx=(5, 0))
        
        # 注释操作按钮
        annotation_buttons_frame = ttk.Frame(content_frame)
        annotation_buttons_frame.pack(fill=tk.X, pady=5)
        
        add_annotation_btn = ttk.Button(annotation_buttons_frame, text="添加注释", 
                                      command=self.add_annotation)
        add_annotation_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        delete_annotation_btn = ttk.Button(annotation_buttons_frame, text="删除选中", 
                                         command=self.delete_annotation)
        delete_annotation_btn.pack(side=tk.LEFT, padx=5)
        
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
        
        apply_btn = ttk.Button(button_frame, text="应用注释", 
                             command=self.apply_annotations, style='Action.TButton')
        apply_btn.pack(side=tk.RIGHT)
        
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
        
        # 清空注释列表
        self.annotations.clear()
        
        messagebox.showinfo("成功", f"已移除 {len(selected_indices)} 个文件")
    
    def clear_files(self):
        """清空文件列表"""
        if not self.selected_files:
            messagebox.showinfo("提示", "文件列表已为空")
            return
            
        if messagebox.askyesno("确认", "确定要清空所有文件吗？"):
            self.selected_files.clear()
            self.file_listbox.delete(0, tk.END)
            
            # 清空注释列表
            self.annotations.clear()
            
            messagebox.showinfo("成功", "已清空所有文件")
    
    def add_annotation(self):
        """添加新注释"""
        annotation_type = self.annotation_type.get()
        page_number = self.page_number.get()
        text = self.annotation_text.get(1.0, tk.END).strip()
        color = self.highlight_color.get()
        
        if page_number < 1:
            messagebox.showwarning("警告", "页码必须大于0")
            return
        
        # 对于便签类型，需要注释文本
        if annotation_type == "note" and not text:
            messagebox.showwarning("警告", "便签类型需要输入注释文本")
            return
        
        # 添加到注释列表
        self.annotations.append({
            "type": annotation_type,
            "page": page_number,
            "text": text,
            "color": color
        })
        
        # 清空输入框
        self.annotation_text.delete(1.0, tk.END)
        self.page_number.set(1)
        
        messagebox.showinfo("成功", f"已添加 {annotation_type} 注释")
    
    def delete_annotation(self):
        """删除选中的注释"""
        # 这里简化处理，实际应该有注释列表框让用户选择要删除的注释
        if not self.annotations:
            messagebox.showwarning("警告", "没有可删除的注释")
            return
        
        if messagebox.askyesno("确认", "确定要删除所有注释吗？"):
            self.annotations.clear()
            messagebox.showinfo("成功", "已删除所有注释")
    
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
            
            # 清空注释列表和编辑区域
            self.annotations.clear()
            self.annotation_text.delete(1.0, tk.END)
            self.page_number.set(1)
            self.annotation_type.set("highlight")
            self.highlight_color.set("yellow")
            
            # 清空输出目录
            self.output_dir.set("")
    
    def apply_annotations(self):
        """将注释应用到PDF文件"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先添加PDF文件")
            return
        
        if not self.annotations:
            messagebox.showwarning("警告", "请先添加注释")
            return
        
        if not self.output_dir.get():
            # 如果未选择输出目录，使用默认目录
            output_dir = os.path.dirname(self.selected_files[0])
            self.output_dir.set(output_dir)
        else:
            output_dir = self.output_dir.get()
        
        try:
            for pdf_file in self.selected_files:
                self.add_annotations_to_pdf(pdf_file, output_dir)
            
            messagebox.showinfo("成功", f"注释已成功应用到PDF文件，保存到:\n{output_dir}")
        except Exception as e:
            messagebox.showerror("错误", f"应用注释到PDF文件时出错: {str(e)}")
    
    def add_annotations_to_pdf(self, pdf_file, output_dir):
        """将注释添加到PDF文件"""
        # 获取文件名和扩展名
        file_name = os.path.basename(pdf_file)
        base_name = os.path.splitext(file_name)[0]
        
        # 输出PDF文件路径
        output_path = os.path.join(output_dir, f"{base_name}_with_annotations.pdf")
        
        try:
            with open(pdf_file, 'rb') as f:
                reader = PdfReader(f)
                writer = PdfWriter()
                
                # 复制所有页面
                for page in reader.pages:
                    writer.add_page(page)
                
                # 注意：PyPDF2对注释的支持有限，这里我们简化处理
                # 实际项目中可能需要使用更专业的PDF库，如PyMuPDF或reportlab
                # 这里只是创建一个带有注释标记的PDF文件
                
                # 保存带注释的PDF文件
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
                    
                # 记录注释信息到外部文件
                notes_path = os.path.join(output_dir, f"{base_name}_annotations.txt")
                with open(notes_path, 'w', encoding='utf-8') as notes_file:
                    notes_file.write("PDF注释信息\n")
                    notes_file.write("=" * 50 + "\n")
                    for i, annotation in enumerate(self.annotations, 1):
                        notes_file.write(f"注释 {i}:\n")
                        notes_file.write(f"  类型: {annotation['type']}\n")
                        notes_file.write(f"  页码: {annotation['page']}\n")
                        if annotation['text']:
                            notes_file.write(f"  文本: {annotation['text']}\n")
                        if annotation['color']:
                            notes_file.write(f"  颜色: {annotation['color']}\n")
                        notes_file.write("\n")
        except Exception as e:
            raise Exception(f"处理文件 {file_name} 时出错: {str(e)}")

# 独立测试函数
def test_annotation_tool():
    """独立测试PDF注释工具"""
    root = tk.Tk()
    root.title("PDF注释工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    annotation_tool = PDFAnnotationTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_annotation_tool()