import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter

class PDFBookmarkTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_files = []
        self.bookmarks = []
        
        # 创建界面
        self.create_bookmark_interface()
    
    def create_bookmark_interface(self):
        """创建PDF书签工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF书签工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="为PDF文件添加或编辑书签", 
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
        
        # 右侧书签编辑区域
        right_frame = ttk.LabelFrame(content_frame, text="书签编辑")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 书签列表
        bookmark_list_frame = ttk.LabelFrame(right_frame, text="现有书签")
        bookmark_list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=10)
        
        # 书签列表框
        self.bookmark_listbox = tk.Listbox(bookmark_list_frame, selectmode=tk.SINGLE)
        self.bookmark_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        bookmark_scrollbar = ttk.Scrollbar(bookmark_list_frame, orient=tk.VERTICAL, 
                                         command=self.bookmark_listbox.yview)
        bookmark_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        self.bookmark_listbox.configure(yscrollcommand=bookmark_scrollbar.set)
        
        # 绑定书签列表选择事件
        self.bookmark_listbox.bind('<<ListboxSelect>>', self.on_bookmark_select)
        
        # 书签编辑区域
        edit_frame = ttk.LabelFrame(right_frame, text="添加/编辑书签")
        edit_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 书签标题
        ttk.Label(edit_frame, text="书签标题:").pack(anchor=tk.W)
        self.bookmark_title = tk.StringVar()
        title_entry = ttk.Entry(edit_frame, textvariable=self.bookmark_title)
        title_entry.pack(fill=tk.X, pady=(2, 5))
        
        # 页码
        ttk.Label(edit_frame, text="页码:").pack(anchor=tk.W)
        self.bookmark_page = tk.IntVar(value=1)
        page_entry = ttk.Entry(edit_frame, textvariable=self.bookmark_page)
        page_entry.pack(fill=tk.X, pady=(2, 5))
        
        # 书签操作按钮
        bookmark_buttons_frame = ttk.Frame(edit_frame)
        bookmark_buttons_frame.pack(fill=tk.X, pady=5)
        
        add_bookmark_btn = ttk.Button(bookmark_buttons_frame, text="添加书签", 
                                     command=self.add_bookmark)
        add_bookmark_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        edit_bookmark_btn = ttk.Button(bookmark_buttons_frame, text="编辑选中", 
                                      command=self.edit_bookmark)
        edit_bookmark_btn.pack(side=tk.LEFT, padx=5)
        
        delete_bookmark_btn = ttk.Button(bookmark_buttons_frame, text="删除选中", 
                                       command=self.delete_bookmark)
        delete_bookmark_btn.pack(side=tk.LEFT, padx=5)
        
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
        
        apply_btn = ttk.Button(button_frame, text="应用书签", 
                             command=self.apply_bookmarks, style='Action.TButton')
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
        
        # 清空书签列表
        self.bookmark_listbox.delete(0, tk.END)
        self.bookmarks.clear()
        self.bookmark_title.set("")
        self.bookmark_page.set(1)
        
        messagebox.showinfo("成功", f"已移除 {len(selected_indices)} 个文件")
    
    def clear_files(self):
        """清空文件列表"""
        if not self.selected_files:
            messagebox.showinfo("提示", "文件列表已为空")
            return
            
        if messagebox.askyesno("确认", "确定要清空所有文件吗？"):
            self.selected_files.clear()
            self.file_listbox.delete(0, tk.END)
            
            # 清空书签列表
            self.bookmark_listbox.delete(0, tk.END)
            self.bookmarks.clear()
            self.bookmark_title.set("")
            self.bookmark_page.set(1)
            
            messagebox.showinfo("成功", "已清空所有文件")
    
    def on_file_select(self, event):
        """当选择PDF文件时，加载其书签"""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            return
        
        selected_index = selected_indices[0]
        pdf_file = self.selected_files[selected_index]
        
        # 清空当前书签列表
        self.bookmark_listbox.delete(0, tk.END)
        self.bookmarks.clear()
        
        try:
            # 读取PDF文件中的书签
            with open(pdf_file, 'rb') as f:
                reader = PdfReader(f)
                
                # 注意：PyPDF2的书签读取功能有限，这里我们简化处理
                # 实际项目中可能需要更复杂的逻辑来处理嵌套书签
                if reader.outline:
                    # 处理大纲（书签）
                    self.load_outline(reader.outline)
        except Exception as e:
            messagebox.showerror("错误", f"读取PDF文件书签时出错: {str(e)}")
    
    def load_outline(self, outline, level=0):
        """递归加载PDF大纲（书签）"""
        for item in outline:
            if isinstance(item, list):
                # 嵌套书签
                self.load_outline(item, level + 1)
            else:
                # 单个书签
                title = item.title
                page_number = item.page_number + 1  # PyPDF2返回的是0-based页码，转为1-based
                indent = "  " * level
                self.bookmarks.append((title, page_number, level))
                self.bookmark_listbox.insert(tk.END, f"{indent}{title} (页 {page_number})")
    
    def on_bookmark_select(self, event):
        """当选择书签时，加载其信息到编辑区域"""
        selected_indices = self.bookmark_listbox.curselection()
        if not selected_indices:
            return
        
        selected_index = selected_indices[0]
        if 0 <= selected_index < len(self.bookmarks):
            title, page, level = self.bookmarks[selected_index]
            self.bookmark_title.set(title)
            self.bookmark_page.set(page)
    
    def add_bookmark(self):
        """添加新书签"""
        title = self.bookmark_title.get().strip()
        page = self.bookmark_page.get()
        
        if not title:
            messagebox.showwarning("警告", "请输入书签标题")
            return
        
        if page < 1:
            messagebox.showwarning("警告", "页码必须大于0")
            return
        
        # 添加到书签列表
        self.bookmarks.append((title, page, 0))
        self.bookmark_listbox.insert(tk.END, f"{title} (页 {page})")
        
        # 清空输入框
        self.bookmark_title.set("")
        self.bookmark_page.set(1)
    
    def edit_bookmark(self):
        """编辑选中的书签"""
        selected_indices = self.bookmark_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请先选择要编辑的书签")
            return
        
        selected_index = selected_indices[0]
        title = self.bookmark_title.get().strip()
        page = self.bookmark_page.get()
        
        if not title:
            messagebox.showwarning("警告", "请输入书签标题")
            return
        
        if page < 1:
            messagebox.showwarning("警告", "页码必须大于0")
            return
        
        # 更新书签
        old_title, old_page, level = self.bookmarks[selected_index]
        self.bookmarks[selected_index] = (title, page, level)
        self.bookmark_listbox.delete(selected_index)
        self.bookmark_listbox.insert(selected_index, f"{'  ' * level}{title} (页 {page})")
        
        # 清空输入框
        self.bookmark_title.set("")
        self.bookmark_page.set(1)
    
    def delete_bookmark(self):
        """删除选中的书签"""
        selected_indices = self.bookmark_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请先选择要删除的书签")
            return
        
        selected_index = selected_indices[0]
        self.bookmark_listbox.delete(selected_index)
        self.bookmarks.pop(selected_index)
        
        # 清空输入框
        self.bookmark_title.set("")
        self.bookmark_page.set(1)
    
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
            
            # 清空书签列表和编辑区域
            self.bookmark_listbox.delete(0, tk.END)
            self.bookmarks.clear()
            self.bookmark_title.set("")
            self.bookmark_page.set(1)
            
            # 清空输出目录
            self.output_dir.set("")
    
    def apply_bookmarks(self):
        """将书签应用到PDF文件"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先添加PDF文件")
            return
        
        if not self.bookmarks:
            messagebox.showwarning("警告", "请先添加书签")
            return
        
        if not self.output_dir.get():
            # 如果未选择输出目录，使用默认目录
            output_dir = os.path.dirname(self.selected_files[0])
            self.output_dir.set(output_dir)
        else:
            output_dir = self.output_dir.get()
        
        try:
            for pdf_file in self.selected_files:
                self.add_bookmarks_to_pdf(pdf_file, output_dir)
            
            messagebox.showinfo("成功", f"书签已成功应用到PDF文件，保存到:\n{output_dir}")
        except Exception as e:
            messagebox.showerror("错误", f"应用书签到PDF文件时出错: {str(e)}")
    
    def add_bookmarks_to_pdf(self, pdf_file, output_dir):
        """将书签添加到PDF文件"""
        # 获取文件名和扩展名
        file_name = os.path.basename(pdf_file)
        base_name = os.path.splitext(file_name)[0]
        
        # 输出PDF文件路径
        output_path = os.path.join(output_dir, f"{base_name}_with_bookmarks.pdf")
        
        try:
            with open(pdf_file, 'rb') as f:
                reader = PdfReader(f)
                writer = PdfWriter()
                
                # 复制所有页面
                for page in reader.pages:
                    writer.add_page(page)
                
                # 添加书签
                for title, page_number, level in self.bookmarks:
                    # PyPDF2需要0-based页码
                    writer.add_outline_item(title, page_number - 1)
                
                # 保存带书签的PDF文件
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
        except Exception as e:
            raise Exception(f"处理文件 {file_name} 时出错: {str(e)}")

# 独立测试函数
def test_bookmark_tool():
    """独立测试PDF书签工具"""
    root = tk.Tk()
    root.title("PDF书签工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    bookmark_tool = PDFBookmarkTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_bookmark_tool()