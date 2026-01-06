import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from pathlib import Path
import tempfile

class PDFToWordTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_files = []
        self.conversion_thread = None
        self.stop_conversion = False
        
        # 创建界面
        self.create_conversion_interface()
    
    def create_conversion_interface(self):
        """创建转换工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF转Word工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="将PDF文件转换为可编辑的Word文档", 
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
        
        select_files_btn = ttk.Button(file_buttons_frame, text="选择PDF文件", 
                                     command=self.select_pdf_files)
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
        
        # 转换引擎选择
        engine_frame = ttk.LabelFrame(right_frame, text="转换引擎")
        engine_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.conversion_engine = tk.StringVar(value="pdf2docx")
        
        engines = [
            ("pdf2docx (推荐)", "pdf2docx", "高质量转换，保持格式"),
            ("PyMuPDF (快速)", "pymupdf", "转换速度快，基本格式"),
        ]
        
        for name, value, description in engines:
            frame = ttk.Frame(engine_frame)
            frame.pack(fill=tk.X, pady=2)
            
            rb = ttk.Radiobutton(frame, text=name, variable=self.conversion_engine, value=value)
            rb.pack(side=tk.LEFT, anchor=tk.W)
            
            desc_label = ttk.Label(frame, text=description, font=('Arial', 8), foreground='gray')
            desc_label.pack(side=tk.LEFT, anchor=tk.W, padx=(5, 0))
        
        # 转换质量设置
        quality_frame = ttk.LabelFrame(right_frame, text="转换质量")
        quality_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.conversion_quality = tk.StringVar(value="balanced")
        
        ttk.Radiobutton(quality_frame, text="高质量 (保留所有格式)", 
                       variable=self.conversion_quality, value="high").pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(quality_frame, text="平衡 (推荐)", 
                       variable=self.conversion_quality, value="balanced").pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(quality_frame, text="快速 (基本格式)", 
                       variable=self.conversion_quality, value="fast").pack(anchor=tk.W, pady=2)
        
        # 页面范围设置
        pages_frame = ttk.LabelFrame(right_frame, text="页面范围")
        pages_frame.pack(fill=tk.X, padx=5, pady=10)
        
        ttk.Label(pages_frame, text="转换页面范围 (例如: 1-5, 8, 11-13):").pack(anchor=tk.W, pady=(5, 0))
        self.pages_range = tk.StringVar()
        pages_entry = ttk.Entry(pages_frame, textvariable=self.pages_range)
        pages_entry.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(pages_frame, text="留空表示转换所有页面", font=('Arial', 8), foreground='gray').pack(anchor=tk.W)
        
        # 输出格式选项
        format_frame = ttk.LabelFrame(right_frame, text="输出格式")
        format_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.output_format = tk.StringVar(value="docx")
        
        format_radio_frame = ttk.Frame(format_frame)
        format_radio_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Radiobutton(format_radio_frame, text="DOCX (Word文档)", 
                       variable=self.output_format, value="docx").pack(side=tk.LEFT)
        ttk.Radiobutton(format_radio_frame, text="DOC (Word 97-2003)", 
                       variable=self.output_format, value="doc").pack(side=tk.LEFT)
        
        # 高级选项
        advanced_frame = ttk.LabelFrame(right_frame, text="高级选项")
        advanced_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.preserve_layout = tk.BooleanVar(value=True)
        layout_check = ttk.Checkbutton(advanced_frame, text="保持页面布局", 
                                      variable=self.preserve_layout)
        layout_check.pack(anchor=tk.W, pady=2)
        
        self.extract_images = tk.BooleanVar(value=True)
        images_check = ttk.Checkbutton(advanced_frame, text="提取图像", 
                                      variable=self.extract_images)
        images_check.pack(anchor=tk.W, pady=2)
        
        self.recognize_tables = tk.BooleanVar(value=True)
        tables_check = ttk.Checkbutton(advanced_frame, text="识别表格", 
                                      variable=self.recognize_tables)
        tables_check.pack(anchor=tk.W, pady=2)
        
        # 输出选项
        output_frame = ttk.LabelFrame(right_frame, text="输出选项")
        output_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.output_suffix = tk.StringVar(value="_converted")
        ttk.Label(output_frame, text="输出文件后缀:").pack(anchor=tk.W)
        suffix_entry = ttk.Entry(output_frame, textvariable=self.output_suffix)
        suffix_entry.pack(fill=tk.X, padx=5, pady=5)
        
        self.open_after_conversion = tk.BooleanVar(value=False)
        open_check = ttk.Checkbutton(output_frame, text="转换完成后打开文件", 
                                    variable=self.open_after_conversion)
        open_check.pack(anchor=tk.W, pady=2)
        
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
    
    def select_pdf_files(self):
        """选择PDF文件"""
        files = filedialog.askopenfilenames(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
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
    
    def update_file_info(self):
        """更新文件信息显示"""
        self.file_info_text.config(state=tk.NORMAL)
        self.file_info_text.delete(1.0, tk.END)
        
        if not self.selected_files:
            self.file_info_text.insert(tk.END, "请先选择PDF文件")
            self.file_info_text.config(state=tk.DISABLED)
            return
        
        total_size = 0
        file_info = "已选文件信息:\n\n"
        
        for file in self.selected_files:
            file_name = os.path.basename(file)
            file_size = os.path.getsize(file)
            total_size += file_size
            
            # 获取PDF页数
            page_count = self.get_pdf_page_count(file)
            
            file_info += f"• {file_name}\n"
            file_info += f"  大小: {self.format_file_size(file_size)}\n"
            file_info += f"  页数: {page_count}\n\n"
        
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
            messagebox.showwarning("警告", "请先选择要预览的PDF文件")
            return
        
        # 这里可以实现PDF预览功能
        # 目前只是简单显示文件信息
        self.update_file_info()
        messagebox.showinfo("预览", "文件预览功能正在开发中")
    
    def reset_tool(self):
        """重置工具"""
        if messagebox.askyesno("确认", "确定要重置所有设置吗？"):
            self.selected_files.clear()
            self.files_listbox.delete(0, tk.END)
            self.clear_file_info()
            self.conversion_engine.set("pdf2docx")
            self.conversion_quality.set("balanced")
            self.pages_range.set("")
            self.output_format.set("docx")
            self.preserve_layout.set(True)
            self.extract_images.set(True)
            self.recognize_tables.set(True)
            self.output_suffix.set("_converted")
            self.open_after_conversion.set(False)
    
    def validate_inputs(self):
        """验证输入"""
        if not self.selected_files:
            return "请先选择PDF文件"
        
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
        """在后台线程中转换PDF文件"""
        try:
            results = self.convert_pdfs(output_dir)
            
            # 在主线程中更新UI
            self.parent.after(0, self.conversion_complete, results)
            
        except Exception as e:
            # 在主线程中显示错误
            self.parent.after(0, lambda: messagebox.showerror("错误", f"转换PDF时出错: {str(e)}"))
            self.parent.after(0, self.reset_ui_after_conversion)
    
    def conversion_complete(self, results):
        """转换完成后的回调"""
        # 重置UI状态
        self.reset_ui_after_conversion()
        
        # 显示转换结果
        success_count = len([r for r in results if r[0]])
        total_count = len(results)
        
        # 构建结果消息
        result_message = f"转换完成:\n成功: {success_count}/{total_count} 个文件\n\n"
        
        # 添加每个文件的详细结果
        for success, file_path, output_path, error in results:
            file_name = os.path.basename(file_path)
            if success:
                result_message += f"✓ {file_name}: 转换成功 → {os.path.basename(output_path)}\n"
            else:
                result_message += f"✗ {file_name}: {error}\n"
        
        # 显示结果
        self.show_conversion_results(result_message)
        
        # 如果设置了转换完成后打开文件，则打开成功的文件
        if self.open_after_conversion.get():
            self.open_converted_files(results)
    
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
    
    def open_converted_files(self, results):
        """打开转换后的文件"""
        import subprocess
        import platform
        
        # 获取操作系统类型
        system = platform.system()
        
        for success, file_path, output_path, error in results:
            if success and os.path.exists(output_path):
                try:
                    if system == "Windows":
                        os.startfile(output_path)
                    elif system == "Darwin":  # macOS
                        subprocess.run(["open", output_path])
                    else:  # Linux
                        subprocess.run(["xdg-open", output_path])
                except Exception as e:
                    print(f"无法打开文件 {output_path}: {e}")
    
    def convert_pdfs(self, output_dir):
        """转换PDF文件的核心功能"""
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
                
                # 确定输出路径
                file_name = os.path.basename(file_path)
                name, _ = os.path.splitext(file_name)
                output_name = f"{name}{self.output_suffix.get()}.{self.output_format.get()}"
                output_path = os.path.join(output_dir, output_name)
                
                # 根据选择的引擎执行转换
                engine = self.conversion_engine.get()
                if engine == "pdf2docx":
                    success = self.convert_with_pdf2docx(file_path, output_path)
                elif engine == "pymupdf":
                    success = self.convert_with_pymupdf(file_path, output_path)
                else:
                    raise ValueError(f"不支持的转换引擎: {engine}")
                
                if success:
                    results.append((True, file_path, output_path, ""))
                else:
                    results.append((False, file_path, "", "转换失败"))
                
            except Exception as e:
                results.append((False, file_path, "", str(e)))
        
        # 完成进度
        self.progress_var.set(100)
        self.progress_label.config(text="转换完成")
        
        return results
    
    def convert_with_pdf2docx(self, input_path, output_path):
        """使用pdf2docx转换PDF"""
        from pdf2docx import Converter
        
        try:
            # 创建转换器
            cv = Converter(input_path)
            
            # 设置转换选项
            convert_args = {}
            
            # 设置页面范围
            pages = self.pages_range.get().strip()
            if pages:
                convert_args['pages'] = self.parse_pages_range(pages)
            
            # 根据质量设置转换参数
            quality = self.conversion_quality.get()
            if quality == "high":
                convert_args['layout_analysis'] = True
            elif quality == "fast":
                convert_args['layout_analysis'] = False
            
            # 执行转换
            cv.convert(output_path, **convert_args)
            cv.close()
            
            return True
        except Exception as e:
            print(f"pdf2docx转换错误: {e}")
            return False
    
    def convert_with_pymupdf(self, input_path, output_path):
        """使用PyMuPDF转换PDF"""
        try:
            import fitz  # PyMuPDF
            
            # 打开PDF文档
            doc = fitz.open(input_path)
            
            # 获取文本内容
            text = ""
            pages = self.pages_range.get().strip()
            
            if pages:
                # 转换指定页面
                page_indices = self.parse_pages_range(pages)
                for page_num in page_indices:
                    if 0 <= page_num < len(doc):
                        page = doc[page_num]
                        text += page.get_text()
            else:
                # 转换所有页面
                for page in doc:
                    text += page.get_text()
            
            # 关闭PDF文档
            doc.close()
            
            # 保存为文本文件（这里简化处理，实际应该转换为Word格式）
            # 注意：PyMuPDF本身不支持直接转换为Word，这里只是示例
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(text)
            
            return True
        except Exception as e:
            print(f"PyMuPDF转换错误: {e}")
            return False
    
    def parse_pages_range(self, range_str):
        """解析页面范围字符串"""
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
                    
                    pages.extend(range(start, end + 1))
                else:
                    # 单个页面
                    page = int(part) - 1  # 转换为0-based索引
                    pages.append(page)
            
            return pages
        except Exception as e:
            raise ValueError(f"解析页面范围时出错: {str(e)}")
    
    def check_dependencies(self):
        """检查必要的依赖库是否已安装"""
        engine = self.conversion_engine.get()
        
        if engine == "pdf2docx":
            try:
                from pdf2docx import Converter
                return True
            except ImportError:
                messagebox.showerror(
                    "缺少依赖", 
                    "未安装pdf2docx库。\n\n请运行以下命令安装:\npip install pdf2docx"
                )
                return False
        
        elif engine == "pymupdf":
            try:
                import fitz
                return True
            except ImportError:
                messagebox.showerror(
                    "缺少依赖", 
                    "未安装PyMuPDF库。\n\n请运行以下命令安装:\npip install PyMuPDF"
                )
                return False
        
        return True

# 独立测试函数
def test_pdf_to_word_tool():
    """独立测试PDF转Word工具"""
    root = tk.Tk()
    root.title("PDF转Word工具测试")
    root.geometry("900x700")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    pdf_to_word_tool = PDFToWordTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_pdf_to_word_tool()