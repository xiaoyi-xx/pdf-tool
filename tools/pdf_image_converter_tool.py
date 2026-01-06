import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from pathlib import Path
import tempfile
from PIL import Image

class PDFImageConverterTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_files = []
        self.conversion_thread = None
        self.stop_conversion = False
        
        # 创建界面
        self.create_converter_interface()
    
    def create_converter_interface(self):
        """创建转换器界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF与图片互转工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="PDF转图片或将多张图片合并为PDF", 
                              font=('Arial', 12))
        desc_label.pack(pady=(0, 20))
        
        # 创建主内容框架
        content_frame = ttk.Frame(self.parent)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧文件选择区域
        left_frame = ttk.LabelFrame(content_frame, text="文件选择")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 操作模式选择
        mode_frame = ttk.Frame(left_frame)
        mode_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.conversion_mode = tk.StringVar(value="pdf_to_image")
        
        ttk.Radiobutton(mode_frame, text="PDF转图片", 
                       variable=self.conversion_mode, value="pdf_to_image",
                       command=self.update_interface_by_mode).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(mode_frame, text="图片转PDF", 
                       variable=self.conversion_mode, value="image_to_pdf",
                       command=self.update_interface_by_mode).pack(side=tk.LEFT)
        
        # 文件选择按钮
        file_buttons_frame = ttk.Frame(left_frame)
        file_buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        select_files_btn = ttk.Button(file_buttons_frame, text="选择文件", 
                                     command=self.select_files)
        select_files_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        clear_files_btn = ttk.Button(file_buttons_frame, text="清空列表", 
                                    command=self.clear_selected_files)
        clear_files_btn.pack(side=tk.LEFT, padx=5)
        
        # 文件列表
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.files_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, height=8)
        files_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, 
                                       command=self.files_listbox.yview)
        self.files_listbox.configure(yscrollcommand=files_scrollbar.set)
        
        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        files_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 文件信息显示
        info_frame = ttk.Frame(left_frame)
        info_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.file_info_text = scrolledtext.ScrolledText(info_frame, height=8, wrap=tk.WORD)
        self.file_info_text.pack(fill=tk.BOTH, expand=True)
        self.file_info_text.config(state=tk.DISABLED)
        
        # 右侧转换选项区域 - 使用Canvas添加滚动条
        right_container = ttk.Frame(content_frame)
        right_container.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0))
        
        # 创建Canvas和滚动条
        canvas = tk.Canvas(right_container, width=300)
        scrollbar = ttk.Scrollbar(right_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # 配置Canvas
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 布局Canvas和滚动条
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 绑定鼠标滚轮事件
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind("<MouseWheel>", _on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        
        # 右侧选项区域
        right_frame = ttk.LabelFrame(scrollable_frame, text="转换选项")
        right_frame.pack(fill=tk.BOTH, padx=5, pady=5, ipadx=10)
        
        # PDF转图片选项
        self.pdf_to_image_frame = ttk.LabelFrame(right_frame, text="PDF转图片选项")
        self.pdf_to_image_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 图片格式选择
        ttk.Label(self.pdf_to_image_frame, text="图片格式:").pack(anchor=tk.W, pady=(5, 0))
        self.image_format = tk.StringVar(value="png")
        
        format_frame = ttk.Frame(self.pdf_to_image_frame)
        format_frame.pack(fill=tk.X, padx=5, pady=5)
        
        formats = [("PNG", "png"), ("JPEG", "jpg"), ("TIFF", "tiff"), ("BMP", "bmp")]
        for text, value in formats:
            ttk.Radiobutton(format_frame, text=text, 
                           variable=self.image_format, value=value).pack(side=tk.LEFT)
        
        # 分辨率设置
        ttk.Label(self.pdf_to_image_frame, text="分辨率 (DPI):").pack(anchor=tk.W, pady=(10, 0))
        self.dpi_value = tk.IntVar(value=150)
        
        dpi_frame = ttk.Frame(self.pdf_to_image_frame)
        dpi_frame.pack(fill=tk.X, padx=5, pady=5)
        
        dpi_scale = ttk.Scale(dpi_frame, from_=72, to=300, 
                             variable=self.dpi_value, orient=tk.HORIZONTAL)
        dpi_scale.pack(fill=tk.X, side=tk.LEFT, expand=True)
        
        dpi_label = ttk.Label(dpi_frame, textvariable=self.dpi_value, width=4)
        dpi_label.pack(side=tk.RIGHT, padx=(5, 0))
        
        # 页面范围
        ttk.Label(self.pdf_to_image_frame, text="页面范围 (例如: 1-5, 8, 11-13):").pack(anchor=tk.W, pady=(10, 0))
        self.pages_range = tk.StringVar()
        pages_entry = ttk.Entry(self.pdf_to_image_frame, textvariable=self.pages_range)
        pages_entry.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(self.pdf_to_image_frame, text="留空表示转换所有页面", font=('Arial', 8), foreground='gray').pack(anchor=tk.W)
        
        # 输出选项 - PDF转图片
        self.pdf_output_frame = ttk.LabelFrame(self.pdf_to_image_frame, text="输出选项")
        self.pdf_output_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.output_prefix = tk.StringVar(value="page_")
        ttk.Label(self.pdf_output_frame, text="文件名前缀:").pack(anchor=tk.W)
        prefix_entry = ttk.Entry(self.pdf_output_frame, textvariable=self.output_prefix)
        prefix_entry.pack(fill=tk.X, padx=5, pady=5)
        
        self.single_folder = tk.BooleanVar(value=True)
        folder_check = ttk.Checkbutton(self.pdf_output_frame, text="所有图片放在同一文件夹", 
                                      variable=self.single_folder)
        folder_check.pack(anchor=tk.W, pady=2)
        
        # 图片转PDF选项
        self.image_to_pdf_frame = ttk.LabelFrame(right_frame, text="图片转PDF选项")
        self.image_to_pdf_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 页面大小
        ttk.Label(self.image_to_pdf_frame, text="页面大小:").pack(anchor=tk.W, pady=(5, 0))
        self.page_size = tk.StringVar(value="A4")
        
        size_frame = ttk.Frame(self.image_to_pdf_frame)
        size_frame.pack(fill=tk.X, padx=5, pady=5)
        
        sizes = [("A4", "A4"), ("A3", "A3"), ("Letter", "letter"), ("Legal", "legal")]
        for text, value in sizes:
            ttk.Radiobutton(size_frame, text=text, 
                           variable=self.page_size, value=value).pack(side=tk.LEFT)
        
        # 页面方向
        ttk.Label(self.image_to_pdf_frame, text="页面方向:").pack(anchor=tk.W, pady=(10, 0))
        self.page_orientation = tk.StringVar(value="portrait")
        
        orientation_frame = ttk.Frame(self.image_to_pdf_frame)
        orientation_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Radiobutton(orientation_frame, text="纵向", 
                       variable=self.page_orientation, value="portrait").pack(side=tk.LEFT)
        ttk.Radiobutton(orientation_frame, text="横向", 
                       variable=self.page_orientation, value="landscape").pack(side=tk.LEFT)
        
        # 图像质量
        ttk.Label(self.image_to_pdf_frame, text="图像质量:").pack(anchor=tk.W, pady=(10, 0))
        self.image_quality = tk.IntVar(value=90)
        
        quality_frame = ttk.Frame(self.image_to_pdf_frame)
        quality_frame.pack(fill=tk.X, padx=5, pady=5)
        
        quality_scale = ttk.Scale(quality_frame, from_=10, to=100, 
                                 variable=self.image_quality, orient=tk.HORIZONTAL)
        quality_scale.pack(fill=tk.X, side=tk.LEFT, expand=True)
        
        quality_label = ttk.Label(quality_frame, textvariable=self.image_quality, width=4)
        quality_label.pack(side=tk.RIGHT, padx=(5, 0))
        
        # 输出选项 - 图片转PDF
        self.image_output_frame = ttk.LabelFrame(self.image_to_pdf_frame, text="输出选项")
        self.image_output_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.pdf_output_name = tk.StringVar(value="合并文档.pdf")
        ttk.Label(self.image_output_frame, text="输出文件名:").pack(anchor=tk.W)
        pdf_name_entry = ttk.Entry(self.image_output_frame, textvariable=self.pdf_output_name)
        pdf_name_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # 排序选项
        sort_frame = ttk.Frame(self.image_output_frame)
        sort_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(sort_frame, text="上移", command=self.move_up).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(sort_frame, text="下移", command=self.move_down).pack(side=tk.LEFT)
        
        # 底部按钮区域
        button_frame = ttk.Frame(self.parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(button_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        self.progress_bar.pack_forget()  # 初始隐藏
        
        # 进度标签
        self.progress_label = ttk.Label(button_frame, text="")
        self.progress_label.pack(fill=tk.X, pady=(0, 10))
        self.progress_label.pack_forget()  # 初始隐藏
        
        # 按钮框架
        btn_subframe = ttk.Frame(button_frame)
        btn_subframe.pack(fill=tk.X)
        
        # 预览按钮
        preview_btn = ttk.Button(btn_subframe, text="预览文件", 
                                command=self.preview_files)
        preview_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 重置按钮
        reset_btn = ttk.Button(btn_subframe, text="重置", 
                              command=self.reset_tool)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 停止按钮
        self.stop_btn = ttk.Button(btn_subframe, text="停止转换", 
                                  command=self.stop_conversion_process)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.stop_btn.config(state=tk.DISABLED)
        
        # 转换按钮
        self.convert_btn = ttk.Button(btn_subframe, text="开始转换", 
                                     command=self.start_conversion, style='Action.TButton')
        self.convert_btn.pack(side=tk.RIGHT)
        
        # 初始化界面状态
        self.update_interface_by_mode()
    
    def update_interface_by_mode(self):
        """根据转换模式更新界面"""
        mode = self.conversion_mode.get()
        
        if mode == "pdf_to_image":
            # PDF转图片模式
            self.pdf_to_image_frame.pack(fill=tk.X, padx=5, pady=10)
            self.image_to_pdf_frame.pack_forget()
            self.convert_btn.config(text="开始转换 (PDF转图片)")
        else:
            # 图片转PDF模式
            self.pdf_to_image_frame.pack_forget()
            self.image_to_pdf_frame.pack(fill=tk.X, padx=5, pady=10)
            self.convert_btn.config(text="开始转换 (图片转PDF)")
    
    def select_files(self):
        """选择文件"""
        mode = self.conversion_mode.get()
        
        if mode == "pdf_to_image":
            # 选择PDF文件
            files = filedialog.askopenfilenames(
                title="选择PDF文件",
                filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
            )
        else:
            # 选择图片文件
            files = filedialog.askopenfilenames(
                title="选择图片文件",
                filetypes=[
                    ("图片文件", "*.png *.jpg *.jpeg *.bmp *.tiff *.tif"),
                    ("PNG文件", "*.png"),
                    ("JPEG文件", "*.jpg *.jpeg"),
                    ("BMP文件", "*.bmp"),
                    ("TIFF文件", "*.tiff *.tif"),
                    ("所有文件", "*.*")
                ]
            )
        
        for file in files:
            if file not in self.selected_files:
                self.selected_files.append(file)
                self.files_listbox.insert(tk.END, os.path.basename(file))
        
        if files:
            self.update_file_info()
    
    def clear_selected_files(self):
        """清空选择的文件"""
        if not self.selected_files:
            return
            
        if messagebox.askyesno("确认", "确定要清空文件列表吗？"):
            self.selected_files.clear()
            self.files_listbox.delete(0, tk.END)
            self.clear_file_info()
    
    def move_up(self):
        """上移选中的文件（图片转PDF时使用）"""
        if self.conversion_mode.get() != "image_to_pdf":
            return
            
        selected_indices = self.files_listbox.curselection()
        if not selected_indices or selected_indices[0] == 0:
            return
        
        for index in selected_indices:
            if index > 0:
                # 交换列表中的元素
                self.selected_files[index], self.selected_files[index-1] = self.selected_files[index-1], self.selected_files[index]
                # 更新列表框
                self.files_listbox.delete(index)
                self.files_listbox.insert(index-1, os.path.basename(self.selected_files[index-1]))
        
        # 重新选择移动后的项目
        for i in range(len(selected_indices)):
            self.files_listbox.selection_set(selected_indices[0] - 1 + i)
    
    def move_down(self):
        """下移选中的文件（图片转PDF时使用）"""
        if self.conversion_mode.get() != "image_to_pdf":
            return
            
        selected_indices = self.files_listbox.curselection()
        if not selected_indices or selected_indices[-1] == len(self.selected_files) - 1:
            return
        
        for index in reversed(selected_indices):
            if index < len(self.selected_files) - 1:
                # 交换列表中的元素
                self.selected_files[index], self.selected_files[index+1] = self.selected_files[index+1], self.selected_files[index]
                # 更新列表框
                self.files_listbox.delete(index)
                self.files_listbox.insert(index+1, os.path.basename(self.selected_files[index+1]))
        
        # 重新选择移动后的项目
        for i in range(len(selected_indices)):
            self.files_listbox.selection_set(selected_indices[0] + 1 + i)
    
    def update_file_info(self):
        """更新文件信息显示"""
        self.file_info_text.config(state=tk.NORMAL)
        self.file_info_text.delete(1.0, tk.END)
        
        if not self.selected_files:
            mode_text = "PDF" if self.conversion_mode.get() == "pdf_to_image" else "图片"
            self.file_info_text.insert(tk.END, f"请先选择{mode_text}文件")
            self.file_info_text.config(state=tk.DISABLED)
            return
        
        total_size = 0
        file_info = "已选文件信息:\n\n"
        
        for file in self.selected_files:
            file_name = os.path.basename(file)
            file_size = os.path.getsize(file)
            total_size += file_size
            
            if self.conversion_mode.get() == "pdf_to_image":
                # PDF文件信息
                page_count = self.get_pdf_page_count(file)
                file_info += f"• {file_name}\n"
                file_info += f"  大小: {self.format_file_size(file_size)}\n"
                file_info += f"  页数: {page_count}\n\n"
            else:
                # 图片文件信息
                image_info = self.get_image_info(file)
                file_info += f"• {file_name}\n"
                file_info += f"  大小: {self.format_file_size(file_size)}\n"
                file_info += f"  尺寸: {image_info}\n\n"
        
        file_info += f"总计: {len(self.selected_files)} 个文件, {self.format_file_size(total_size)}"
        
        self.file_info_text.insert(tk.END, file_info)
        self.file_info_text.config(state=tk.DISABLED)
    
    def get_pdf_page_count(self, file_path):
        """获取PDF文件的页数"""
        try:
            import PyPDF2
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                return len(pdf_reader.pages)
        except:
            return "未知"
    
    def get_image_info(self, file_path):
        """获取图片文件信息"""
        try:
            with Image.open(file_path) as img:
                return f"{img.width} × {img.height} 像素"
        except:
            return "未知"
    
    def clear_file_info(self):
        """清空文件信息"""
        self.file_info_text.config(state=tk.NORMAL)
        self.file_info_text.delete(1.0, tk.END)
        self.file_info_text.config(state=tk.DISABLED)
    
    def format_file_size(self, size_bytes):
        """格式化文件大小显示"""
        if size_bytes == 0:
            return "0 B"
        
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        
        return f"{size_bytes:.2f} {size_names[i]}"
    
    def preview_files(self):
        """预览文件"""
        if not self.selected_files:
            mode_text = "PDF" if self.conversion_mode.get() == "pdf_to_image" else "图片"
            messagebox.showwarning("警告", f"请先选择要预览的{mode_text}文件")
            return
        
        # 这里可以实现文件预览功能
        # 目前只是简单显示文件信息
        self.update_file_info()
        messagebox.showinfo("预览", "文件预览功能正在开发中")
    
    def reset_tool(self):
        """重置工具"""
        if messagebox.askyesno("确认", "确定要重置所有设置吗？"):
            self.selected_files.clear()
            self.files_listbox.delete(0, tk.END)
            self.clear_file_info()
            self.conversion_mode.set("pdf_to_image")
            self.image_format.set("png")
            self.dpi_value.set(150)
            self.pages_range.set("")
            self.output_prefix.set("page_")
            self.single_folder.set(True)
            self.page_size.set("A4")
            self.page_orientation.set("portrait")
            self.image_quality.set(90)
            self.pdf_output_name.set("合并文档.pdf")
            self.update_interface_by_mode()
    
    def validate_inputs(self):
        """验证输入"""
        if not self.selected_files:
            mode_text = "PDF" if self.conversion_mode.get() == "pdf_to_image" else "图片"
            return f"请先选择{mode_text}文件"
        
        return None
    
    def start_conversion(self):
        """开始转换过程"""
        # 验证输入
        error = self.validate_inputs()
        if error:
            messagebox.showwarning("输入错误", error)
            return
        
        # 检查是否安装了必要的库
        if not self.check_dependencies():
            return
        
        # 选择输出目录
        output_dir = filedialog.askdirectory(title="选择保存位置")
        if not output_dir:
            return
        
        # 禁用转换按钮，启用停止按钮
        self.convert_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))  # 显示进度条
        self.progress_label.pack(fill=tk.X, pady=(0, 10))  # 显示进度标签
        
        # 在后台线程中执行转换
        self.stop_conversion = False
        self.conversion_thread = threading.Thread(
            target=self.convert_files_thread,
            args=(output_dir,)
        )
        self.conversion_thread.daemon = True
        self.conversion_thread.start()
    
    def stop_conversion_process(self):
        """停止转换过程"""
        self.stop_conversion = True
        if self.conversion_thread and self.conversion_thread.is_alive():
            messagebox.showinfo("提示", "正在停止转换过程...")
    
    def convert_files_thread(self, output_dir):
        """在后台线程中转换文件"""
        try:
            mode = self.conversion_mode.get()
            
            if mode == "pdf_to_image":
                results = self.convert_pdfs_to_images(output_dir)
            else:
                results = self.convert_images_to_pdf(output_dir)
            
            # 在主线程中更新UI
            self.parent.after(0, self.conversion_complete, results, mode)
            
        except Exception as e:
            # 在主线程中显示错误
            self.parent.after(0, lambda: messagebox.showerror("错误", f"转换文件时出错: {str(e)}"))
            self.parent.after(0, self.reset_ui_after_conversion)
    
    def conversion_complete(self, results, mode):
        """转换完成后的回调"""
        # 重置UI状态
        self.reset_ui_after_conversion()
        
        # 显示转换结果
        success_count = len([r for r in results if r[0]])
        total_count = len(results)
        
        # 构建结果消息
        operation = "PDF转图片" if mode == "pdf_to_image" else "图片转PDF"
        result_message = f"{operation}完成:\n成功: {success_count}/{total_count} 个文件\n\n"
        
        # 添加每个文件的详细结果
        for success, file_path, output_info, error in results:
            file_name = os.path.basename(file_path)
            if success:
                if mode == "pdf_to_image":
                    result_message += f"✓ {file_name}: 转换成功 ({output_info} 张图片)\n"
                else:
                    result_message += f"✓ {file_name}: 已添加到PDF\n"
            else:
                result_message += f"✗ {file_name}: {error}\n"
        
        # 显示结果
        self.show_conversion_results(result_message)
    
    def reset_ui_after_conversion(self):
        """转换完成后重置UI状态"""
        self.convert_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.progress_bar.pack_forget()  # 隐藏进度条
        self.progress_label.pack_forget()  # 隐藏进度标签
        self.progress_var.set(0)
        self.progress_label.config(text="")
    
    def show_conversion_results(self, result_message):
        """显示转换结果"""
        # 创建一个新窗口显示详细结果
        result_window = tk.Toplevel(self.parent)
        result_window.title("转换结果")
        result_window.geometry("600x400")
        result_window.resizable(True, True)
        
        # 结果文本区域
        result_text = scrolledtext.ScrolledText(result_window, wrap=tk.WORD)
        result_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        result_text.insert(tk.END, result_message)
        result_text.config(state=tk.DISABLED)
        
        # 关闭按钮
        close_btn = ttk.Button(result_window, text="关闭", 
                              command=result_window.destroy)
        close_btn.pack(pady=10)
    
    def convert_pdfs_to_images(self, output_dir):
        """将PDF转换为图片"""
        try:
            import fitz  # PyMuPDF
        except ImportError:
            messagebox.showerror("错误", "未安装PyMuPDF库")
            return []
        
        results = []
        total_files = len(self.selected_files)
        
        for i, file_path in enumerate(self.selected_files):
            if self.stop_conversion:
                break
                
            try:
                # 更新进度
                progress = (i / total_files) * 100
                self.progress_var.set(progress)
                self.progress_label.config(text=f"正在转换: {os.path.basename(file_path)}")
                
                # 打开PDF文档
                doc = fitz.open(file_path)
                
                # 解析页面范围
                pages = self.parse_pages_range(self.pages_range.get().strip(), len(doc))
                
                # 确定输出目录
                if self.single_folder.get():
                    # 所有图片放在同一文件夹
                    base_output_dir = output_dir
                    file_prefix = f"{os.path.splitext(os.path.basename(file_path))[0]}_{self.output_prefix.get()}"
                else:
                    # 每个PDF创建单独文件夹
                    base_name = os.path.splitext(os.path.basename(file_path))[0]
                    base_output_dir = os.path.join(output_dir, base_name)
                    os.makedirs(base_output_dir, exist_ok=True)
                    file_prefix = self.output_prefix.get()
                
                # 转换页面
                image_count = 0
                for page_num in pages:
                    if page_num < len(doc):
                        page = doc[page_num]
                        
                        # 将页面转换为图片
                        pix = page.get_pixmap(matrix=fitz.Matrix(self.dpi_value.get()/72, self.dpi_value.get()/72))
                        
                        # 确定输出路径
                        output_path = os.path.join(
                            base_output_dir, 
                            f"{file_prefix}{page_num+1:04d}.{self.image_format.get()}"
                        )
                        
                        # 保存图片
                        pix.save(output_path)
                        image_count += 1
                
                # 关闭文档
                doc.close()
                
                results.append((True, file_path, image_count, ""))
                
            except Exception as e:
                results.append((False, file_path, 0, str(e)))
        
        # 完成进度
        self.progress_var.set(100)
        self.progress_label.config(text="转换完成")
        
        return results
    
    def convert_images_to_pdf(self, output_dir):
        """将图片转换为PDF"""
        try:
            from reportlab.lib.pagesizes import letter, A4, A3, legal
            from reportlab.pdfgen import canvas
            from reportlab.lib.utils import ImageReader
        except ImportError:
            messagebox.showerror("错误", "未安装reportlab库")
            return []
        
        results = []
        
        try:
            # 更新进度
            self.progress_label.config(text="正在创建PDF文档...")
            
            # 确定页面大小
            page_sizes = {
                "A4": A4,
                "A3": A3,
                "letter": letter,
                "legal": legal
            }
            page_size = page_sizes.get(self.page_size.get(), A4)
            
            # 确定输出路径
            output_path = os.path.join(output_dir, self.pdf_output_name.get())
            
            # 创建PDF文档
            c = canvas.Canvas(output_path, pagesize=page_size)
            
            # 处理方向
            if self.page_orientation.get() == "landscape":
                page_size = (page_size[1], page_size[0])  # 交换宽高
            
            # 添加图片到PDF
            total_images = len(self.selected_files)
            for i, image_path in enumerate(self.selected_files):
                if self.stop_conversion:
                    break
                    
                try:
                    # 更新进度
                    progress = (i / total_images) * 100
                    self.progress_var.set(progress)
                    self.progress_label.config(text=f"正在添加: {os.path.basename(image_path)}")
                    
                    # 打开图片
                    img = Image.open(image_path)
                    
                    # 调整图片大小以适应页面
                    img_width, img_height = img.size
                    page_width, page_height = page_size
                    
                    # 计算缩放比例
                    scale_x = page_width / img_width
                    scale_y = page_height / img_height
                    scale = min(scale_x, scale_y) * 0.9  # 留出边距
                    
                    # 计算图片在页面中的位置（居中）
                    scaled_width = img_width * scale
                    scaled_height = img_height * scale
                    x = (page_width - scaled_width) / 2
                    y = (page_height - scaled_height) / 2
                    
                    # 添加图片到PDF
                    c.drawImage(image_path, x, y, scaled_width, scaled_height)
                    
                    # 添加新页面（如果不是最后一张图片）
                    if i < total_images - 1:
                        c.showPage()
                    
                    results.append((True, image_path, "", ""))
                    
                except Exception as e:
                    results.append((False, image_path, "", str(e)))
            
            # 保存PDF
            c.save()
            
            # 完成进度
            self.progress_var.set(100)
            self.progress_label.config(text="转换完成")
            
        except Exception as e:
            results.append((False, "", "", f"创建PDF时出错: {str(e)}"))
        
        return results
    
    def parse_pages_range(self, range_str, total_pages):
        """解析页面范围字符串"""
        if not range_str:
            return list(range(total_pages))
        
        try:
            pages = []
            parts = range_str.split(',')
            
            for part in parts:
                part = part.strip()
                if '-' in part:
                    start_end = part.split('-')
                    if len(start_end) != 2:
                        raise ValueError(f"无效的范围格式: {part}")
                    
                    start = int(start_end[0].strip()) - 1  # 转换为0-based索引
                    end = int(start_end[1].strip()) - 1   # 转换为0-based索引
                    
                    # 验证范围
                    if start < 0 or end >= total_pages or start > end:
                        raise ValueError(f"页面范围超出有效范围: {part}")
                    
                    pages.extend(range(start, end + 1))
                else:
                    # 单个页面
                    page = int(part) - 1  # 转换为0-based索引
                    if page < 0 or page >= total_pages:
                        raise ValueError(f"页面超出有效范围: {part}")
                    
                    pages.append(page)
            
            return pages
        except Exception as e:
            raise ValueError(f"解析页面范围时出错: {str(e)}")
    
    def check_dependencies(self):
        """检查必要的依赖库是否已安装"""
        mode = self.conversion_mode.get()
        
        if mode == "pdf_to_image":
            try:
                import fitz  # PyMuPDF
                return True
            except ImportError:
                messagebox.showerror(
                    "缺少依赖", 
                    "未安装PyMuPDF库。\n\n请运行以下命令安装:\npip install PyMuPDF"
                )
                return False
        else:
            try:
                from reportlab.lib.pagesizes import letter, A4
                from reportlab.pdfgen import canvas
                return True
            except ImportError:
                messagebox.showerror(
                    "缺少依赖", 
                    "未安装reportlab库。\n\n请运行以下命令安装:\npip install reportlab"
                )
                return False

# 独立测试函数
def test_pdf_image_converter():
    """独立测试PDF与图片互转工具"""
    root = tk.Tk()
    root.title("PDF与图片互转工具测试")
    root.geometry("900x700")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    pdf_image_converter = PDFImageConverterTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_pdf_image_converter()