import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import PyPDF2
from pathlib import Path

class PDFMergeTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.merge_files = []  # 专门用于合并的文件列表
        
        # 创建界面
        self.create_merge_interface()
        
    def create_merge_interface(self):
        """创建合并工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF合并工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="将多个PDF文件合并为一个PDF文件", 
                              font=('Arial', 12))
        desc_label.pack(pady=(0, 20))
        
        # 创建主内容框架
        content_frame = ttk.Frame(self.parent)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧文件列表区域
        left_frame = ttk.LabelFrame(content_frame, text="待合并文件列表")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 文件列表和操作按钮
        list_buttons_frame = ttk.Frame(left_frame)
        list_buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 添加文件按钮
        add_file_btn = ttk.Button(list_buttons_frame, text="添加PDF文件", 
                                 command=self.add_merge_files)
        add_file_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 移除文件按钮
        remove_file_btn = ttk.Button(list_buttons_frame, text="移除选中", 
                                    command=self.remove_selected_files)
        remove_file_btn.pack(side=tk.LEFT, padx=5)
        
        # 清空列表按钮
        clear_list_btn = ttk.Button(list_buttons_frame, text="清空列表", 
                                   command=self.clear_merge_list)
        clear_list_btn.pack(side=tk.LEFT, padx=5)
        
        # 文件列表框架
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 文件列表框
        self.merge_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED)
        self.merge_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 列表滚动条
        list_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, 
                                      command=self.merge_listbox.yview)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.merge_listbox.configure(yscrollcommand=list_scrollbar.set)
        
        # 顺序调整按钮
        order_frame = ttk.Frame(left_frame)
        order_frame.pack(fill=tk.X, padx=5, pady=5)
        
        up_btn = ttk.Button(order_frame, text="上移", command=self.move_up)
        up_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        down_btn = ttk.Button(order_frame, text="下移", command=self.move_down)
        down_btn.pack(side=tk.LEFT, padx=5)
        
        # 右侧选项区域
        right_frame = ttk.LabelFrame(content_frame, text="合并选项")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0), ipadx=10)
        
        # 输出文件名
        output_frame = ttk.Frame(right_frame)
        output_frame.pack(fill=tk.X, padx=5, pady=10)
        
        ttk.Label(output_frame, text="输出文件名:").pack(anchor=tk.W)
        self.output_name = tk.StringVar(value="合并文档.pdf")
        output_entry = ttk.Entry(output_frame, textvariable=self.output_name)
        output_entry.pack(fill=tk.X, pady=(5, 0))
        
        # 页面范围选项
        range_frame = ttk.LabelFrame(right_frame, text="页面范围 (可选)")
        range_frame.pack(fill=tk.X, padx=5, pady=10)
        
        ttk.Label(range_frame, text="格式: 1,3-5,7").pack(anchor=tk.W, pady=(5, 0))
        self.page_range = tk.StringVar()
        range_entry = ttk.Entry(range_frame, textvariable=self.page_range)
        range_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # 合并选项
        options_frame = ttk.LabelFrame(right_frame, text="合并选项")
        options_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.add_bookmarks = tk.BooleanVar(value=True)
        bookmark_check = ttk.Checkbutton(options_frame, text="添加书签", 
                                        variable=self.add_bookmarks)
        bookmark_check.pack(anchor=tk.W, pady=2)
        
        self.preserve_metadata = tk.BooleanVar(value=True)
        metadata_check = ttk.Checkbutton(options_frame, text="保留元数据", 
                                        variable=self.preserve_metadata)
        metadata_check.pack(anchor=tk.W, pady=2)
        
        # 底部按钮区域
        button_frame = ttk.Frame(self.parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 重置按钮
        reset_btn = ttk.Button(button_frame, text="重置", command=self.reset_merge_tool)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 合并按钮
        self.merge_btn = ttk.Button(button_frame, text="开始合并", 
                                   command=self.start_merge, style='Action.TButton')
        self.merge_btn.pack(side=tk.RIGHT)
        
        # 初始化文件列表
        self.update_merge_list()
        
    def add_merge_files(self):
        """添加PDF文件到合并列表"""
        files = filedialog.askopenfilenames(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        for file in files:
            if file not in self.merge_files:
                self.merge_files.append(file)
        
        self.update_merge_list()
        
        if files:
            messagebox.showinfo("成功", f"已添加 {len(files)} 个PDF文件到合并列表")
    
    def remove_selected_files(self):
        """从合并列表中移除选中的文件"""
        selected_indices = self.merge_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请先选择要移除的文件")
            return
            
        for index in reversed(selected_indices):
            self.merge_listbox.delete(index)
            self.merge_files.pop(index)
    
    def clear_merge_list(self):
        """清空合并列表"""
        if not self.merge_files:
            return
            
        if messagebox.askyesno("确认", "确定要清空文件列表吗？"):
            self.merge_files.clear()
            self.merge_listbox.delete(0, tk.END)
    
    def update_merge_list(self):
        """更新合并列表显示"""
        self.merge_listbox.delete(0, tk.END)
        for file in self.merge_files:
            self.merge_listbox.insert(tk.END, os.path.basename(file))
    
    def move_up(self):
        """上移选中的文件"""
        selected_indices = self.merge_listbox.curselection()
        if not selected_indices or selected_indices[0] == 0:
            return
        
        for index in selected_indices:
            if index > 0:
                # 交换列表中的元素
                self.merge_files[index], self.merge_files[index-1] = self.merge_files[index-1], self.merge_files[index]
                # 更新列表框
                self.merge_listbox.delete(index)
                self.merge_listbox.insert(index-1, os.path.basename(self.merge_files[index-1]))
        
        # 重新选择移动后的项目
        for i in range(len(selected_indices)):
            self.merge_listbox.selection_set(selected_indices[0] - 1 + i)
    
    def move_down(self):
        """下移选中的文件"""
        selected_indices = self.merge_listbox.curselection()
        if not selected_indices or selected_indices[-1] == len(self.merge_files) - 1:
            return
        
        for index in reversed(selected_indices):
            if index < len(self.merge_files) - 1:
                # 交换列表中的元素
                self.merge_files[index], self.merge_files[index+1] = self.merge_files[index+1], self.merge_files[index]
                # 更新列表框
                self.merge_listbox.delete(index)
                self.merge_listbox.insert(index+1, os.path.basename(self.merge_files[index+1]))
        
        # 重新选择移动后的项目
        for i in range(len(selected_indices)):
            self.merge_listbox.selection_set(selected_indices[0] + 1 + i)
    
    def reset_merge_tool(self):
        """重置合并工具"""
        if messagebox.askyesno("确认", "确定要重置所有设置吗？"):
            self.merge_files.clear()
            self.update_merge_list()
            self.output_name.set("合并文档.pdf")
            self.page_range.set("")
            self.add_bookmarks.set(True)
            self.preserve_metadata.set(True)
    
    def start_merge(self):
        """开始合并PDF文件"""
        if len(self.merge_files) < 2:
            messagebox.showwarning("警告", "请至少选择两个PDF文件进行合并")
            return
        
        # 选择输出目录
        output_dir = filedialog.askdirectory(title="选择保存位置")
        if not output_dir:
            return
        
        output_path = os.path.join(output_dir, self.output_name.get())
        
        # 检查文件是否已存在
        if os.path.exists(output_path):
            if not messagebox.askyesno("确认", f"文件 {self.output_name.get()} 已存在，是否覆盖？"):
                return
        
        try:
            # 执行合并
            self.merge_pdfs(output_path)
            messagebox.showinfo("成功", f"PDF文件已成功合并到:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"合并PDF时出错: {str(e)}")
    
    def merge_pdfs(self, output_path):
        """合并PDF文件的核心功能"""
        pdf_merger = PyPDF2.PdfMerger()
        
        # 添加所有PDF文件
        for pdf_file in self.merge_files:
            # 如果有指定页面范围，则使用页面范围
            page_range = self.page_range.get().strip()
            if page_range:
                pdf_merger.append(pdf_file, pages=self.parse_page_range(page_range))
            else:
                pdf_merger.append(pdf_file)
        
        # 写入输出文件
        with open(output_path, 'wb') as output_file:
            pdf_merger.write(output_file)
        
        pdf_merger.close()
    
    def parse_page_range(self, range_str):
        """解析页面范围字符串"""
        # 这里简化处理，实际应用中需要更复杂的解析逻辑
        # 例如 "1,3-5,7" 应该解析为 [0, 2, 3, 4, 6] (0-based索引)
        try:
            pages = []
            parts = range_str.split(',')
            for part in parts:
                if '-' in part:
                    start, end = part.split('-')
                    pages.extend(range(int(start)-1, int(end)))
                else:
                    pages.append(int(part)-1)
            return pages
        except:
            # 如果解析失败，返回None表示使用所有页面
            return None

# 独立测试函数
def test_merge_tool():
    """独立测试合并工具"""
    root = tk.Tk()
    root.title("PDF合并工具测试")
    root.geometry("800x600")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    merge_tool = PDFMergeTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_merge_tool()