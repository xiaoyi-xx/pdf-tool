import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from pathlib import Path
import tempfile
import shutil

class PDFCompressTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_files = []
        self.compression_thread = None
        self.stop_compression = False
        
        # 创建界面
        self.create_compress_interface()
    
    def create_compress_interface(self):
        """创建压缩工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF压缩工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="使用多种算法减小PDF文件大小，优化存储和传输", 
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
        
        # 右侧压缩选项区域 - 使用Canvas添加滚动条
        right_container = ttk.Frame(content_frame)
        right_container.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0))
        
        # 创建Canvas和滚动条
        canvas = tk.Canvas(right_container, width=300)  # 设置固定宽度
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
        right_frame = ttk.LabelFrame(scrollable_frame, text="压缩选项")
        right_frame.pack(fill=tk.BOTH, padx=5, pady=5, ipadx=10)
        
        # 压缩算法选择
        algorithm_frame = ttk.LabelFrame(right_frame, text="压缩算法")
        algorithm_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.compression_algorithm = tk.StringVar(value="pikepdf")
        
        algorithms = [
            ("pikepdf (推荐)", "pikepdf", "使用QPDF引擎，压缩效果好"),
            ("pypdfium2 (快速)", "pypdfium2", "基于Google PDFium，处理速度快"),
            ("Ghostscript (强力)", "ghostscript", "使用Ghostscript，压缩率最高")
        ]
        
        for name, value, description in algorithms:
            frame = ttk.Frame(algorithm_frame)
            frame.pack(fill=tk.X, pady=2)
            
            rb = ttk.Radiobutton(frame, text=name, variable=self.compression_algorithm, value=value)
            rb.pack(side=tk.LEFT, anchor=tk.W)
            
            desc_label = ttk.Label(frame, text=description, font=('Arial', 8), foreground='gray')
            desc_label.pack(side=tk.LEFT, anchor=tk.W, padx=(5, 0))
        
        # 压缩级别选择
        level_frame = ttk.LabelFrame(right_frame, text="压缩级别")
        level_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.compression_level = tk.StringVar(value="medium")
        
        ttk.Radiobutton(level_frame, text="轻度压缩 (质量优先)", 
                       variable=self.compression_level, value="low").pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(level_frame, text="中等压缩 (平衡)", 
                       variable=self.compression_level, value="medium").pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(level_frame, text="强力压缩 (大小优先)", 
                       variable=self.compression_level, value="high").pack(anchor=tk.W, pady=2)
        
        # 图像压缩选项
        image_frame = ttk.LabelFrame(right_frame, text="图像压缩")
        image_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 图像质量
        ttk.Label(image_frame, text="图像质量:").pack(anchor=tk.W, pady=(5, 0))
        self.image_quality = tk.IntVar(value=75)
        quality_scale = ttk.Scale(image_frame, from_=10, to=100, 
                                 variable=self.image_quality, orient=tk.HORIZONTAL)
        quality_scale.pack(fill=tk.X, padx=5, pady=5)
        
        quality_value_frame = ttk.Frame(image_frame)
        quality_value_frame.pack(fill=tk.X, padx=5)
        ttk.Label(quality_value_frame, text="低").pack(side=tk.LEFT)
        ttk.Label(quality_value_frame, textvariable=self.image_quality).pack(side=tk.LEFT, expand=True)
        ttk.Label(quality_value_frame, text="高").pack(side=tk.RIGHT)
        
        # 图像格式
        ttk.Label(image_frame, text="图像格式:").pack(anchor=tk.W, pady=(10, 0))
        self.image_format = tk.StringVar(value="jpeg")
        format_frame = ttk.Frame(image_frame)
        format_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Radiobutton(format_frame, text="JPEG", 
                       variable=self.image_format, value="jpeg").pack(side=tk.LEFT)
        ttk.Radiobutton(format_frame, text="PNG", 
                       variable=self.image_format, value="png").pack(side=tk.LEFT)
        
        # 高级选项
        advanced_frame = ttk.LabelFrame(right_frame, text="高级选项")
        advanced_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.remove_metadata = tk.BooleanVar(value=True)
        metadata_check = ttk.Checkbutton(advanced_frame, text="删除元数据", 
                                        variable=self.remove_metadata)
        metadata_check.pack(anchor=tk.W, pady=2)
        
        self.optimize_images = tk.BooleanVar(value=True)
        images_check = ttk.Checkbutton(advanced_frame, text="优化图像", 
                                      variable=self.optimize_images)
        images_check.pack(anchor=tk.W, pady=2)
        
        self.remove_bookmarks = tk.BooleanVar(value=False)
        bookmarks_check = ttk.Checkbutton(advanced_frame, text="删除书签", 
                                         variable=self.remove_bookmarks)
        bookmarks_check.pack(anchor=tk.W, pady=2)
        
        # 输出选项
        output_frame = ttk.LabelFrame(right_frame, text="输出选项")
        output_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.output_suffix = tk.StringVar(value="_compressed")
        ttk.Label(output_frame, text="输出文件后缀:").pack(anchor=tk.W)
        suffix_entry = ttk.Entry(output_frame, textvariable=self.output_suffix)
        suffix_entry.pack(fill=tk.X, padx=5, pady=5)
        
        self.overwrite_original = tk.BooleanVar(value=False)
        overwrite_check = ttk.Checkbutton(output_frame, text="覆盖原文件", 
                                         variable=self.overwrite_original)
        overwrite_check.pack(anchor=tk.W, pady=2)
        
        # 底部按钮区域
        button_frame = ttk.Frame(self.parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(button_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        self.progress_bar.pack_forget()  # 初始隐藏
        
        # 按钮框架
        btn_subframe = ttk.Frame(button_frame)
        btn_subframe.pack(fill=tk.X)
        
        # 分析按钮
        analyze_btn = ttk.Button(btn_subframe, text="分析文件", 
                                command=self.analyze_files)
        analyze_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 重置按钮
        reset_btn = ttk.Button(btn_subframe, text="重置", 
                              command=self.reset_compress_tool)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 停止按钮
        self.stop_btn = ttk.Button(btn_subframe, text="停止压缩", 
                                  command=self.stop_compression_process)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.stop_btn.config(state=tk.DISABLED)
        
        # 压缩按钮
        self.compress_btn = ttk.Button(btn_subframe, text="开始压缩", 
                                      command=self.start_compression, style='Action.TButton')
        self.compress_btn.pack(side=tk.RIGHT)
    
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
            
            file_info += f"• {file_name}\n"
            file_info += f"  大小: {self.format_file_size(file_size)}\n\n"
        
        file_info += f"总计: {len(self.selected_files)} 个文件, {self.format_file_size(total_size)}"
        
        self.file_info_text.insert(tk.END, file_info)
        self.file_info_text.config(state=tk.DISABLED)
    
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
    
    def analyze_files(self):
        """分析文件"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先选择要分析的PDF文件")
            return
        
        # 这里可以实现更复杂的PDF分析功能
        # 目前只是简单显示文件信息
        self.update_file_info()
        messagebox.showinfo("分析完成", "文件分析完成，请查看文件信息区域")
    
    def reset_compress_tool(self):
        """重置压缩工具"""
        if messagebox.askyesno("确认", "确定要重置所有设置吗？"):
            self.selected_files.clear()
            self.files_listbox.delete(0, tk.END)
            self.clear_file_info()
            self.compression_algorithm.set("pikepdf")
            self.compression_level.set("medium")
            self.image_quality.set(75)
            self.image_format.set("jpeg")
            self.remove_metadata.set(True)
            self.optimize_images.set(True)
            self.remove_bookmarks.set(False)
            self.output_suffix.set("_compressed")
            self.overwrite_original.set(False)
    
    def start_compression(self):
        """开始压缩PDF文件"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先选择要压缩的PDF文件")
            return
        
        # 检查是否安装了必要的库
        if not self.check_dependencies():
            return
        
        # 选择输出目录（如果不覆盖原文件）
        output_dir = None
        if not self.overwrite_original.get():
            output_dir = filedialog.askdirectory(title="选择保存位置")
            if not output_dir:
                return
        
        # 禁用压缩按钮，启用停止按钮
        self.compress_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))  # 显示进度条
        
        # 在后台线程中执行压缩
        self.stop_compression = False
        self.compression_thread = threading.Thread(
            target=self.compress_pdfs_thread,
            args=(output_dir,)
        )
        self.compression_thread.daemon = True
        self.compression_thread.start()
    
    def stop_compression_process(self):
        """停止压缩过程"""
        self.stop_compression = True
        if self.compression_thread and self.compression_thread.is_alive():
            messagebox.showinfo("提示", "正在停止压缩过程...")
    
    def compress_pdfs_thread(self, output_dir):
        """在后台线程中压缩PDF文件"""
        try:
            results = self.compress_pdfs(output_dir)
            
            # 在主线程中更新UI
            self.parent.after(0, self.compression_complete, results)
            
        except Exception as e:
            # 在主线程中显示错误
            self.parent.after(0, lambda: messagebox.showerror("错误", f"压缩PDF时出错: {str(e)}"))
            self.parent.after(0, self.reset_ui_after_compression)
    
    def compression_complete(self, results):
        """压缩完成后的回调"""
        # 重置UI状态
        self.reset_ui_after_compression()
        
        # 显示压缩结果
        success_count = len([r for r in results if r[0]])
        total_count = len(results)
        
        # 构建结果消息
        result_message = f"压缩完成:\n成功: {success_count}/{total_count} 个文件\n\n"
        
        # 添加每个文件的详细结果
        for success, file_path, original_size, compressed_size, error in results:
            file_name = os.path.basename(file_path)
            if success:
                compression_ratio = (1 - compressed_size / original_size) * 100
                result_message += f"✓ {file_name}: {self.format_file_size(original_size)} → {self.format_file_size(compressed_size)} (减少 {compression_ratio:.1f}%)\n"
            else:
                result_message += f"✗ {file_name}: {error}\n"
        
        # 显示结果
        self.show_compression_results(result_message)
    
    def reset_ui_after_compression(self):
        """压缩完成后重置UI状态"""
        self.compress_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.progress_bar.pack_forget()  # 隐藏进度条
        self.progress_var.set(0)
    
    def show_compression_results(self, result_message):
        """显示压缩结果"""
        # 创建一个新窗口显示详细结果
        result_window = tk.Toplevel(self.parent)
        result_window.title("压缩结果")
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
    
    def compress_pdfs(self, output_dir):
        """压缩PDF文件的核心功能"""
        results = []
        total_files = len(self.selected_files)
        
        for i, file_path in enumerate(self.selected_files):
            if self.stop_compression:
                break
                
            try:
                # 更新进度
                progress = (i / total_files) * 100
                self.progress_var.set(progress)
                
                # 确定输出路径
                if self.overwrite_original.get():
                    output_path = file_path
                else:
                    file_name = os.path.basename(file_path)
                    name, ext = os.path.splitext(file_name)
                    output_name = f"{name}{self.output_suffix.get()}{ext}"
                    output_path = os.path.join(output_dir, output_name)
                
                # 获取原始文件大小
                original_size = os.path.getsize(file_path)
                
                # 根据选择的算法执行压缩
                algorithm = self.compression_algorithm.get()
                if algorithm == "pikepdf":
                    compressed_size = self.compress_with_pikepdf(file_path, output_path)
                elif algorithm == "pypdfium2":
                    compressed_size = self.compress_with_pypdfium2(file_path, output_path)
                elif algorithm == "ghostscript":
                    compressed_size = self.compress_with_ghostscript(file_path, output_path)
                else:
                    raise ValueError(f"不支持的压缩算法: {algorithm}")
                
                # 检查压缩结果
                if compressed_size < 0:
                    raise Exception("压缩失败")
                
                results.append((True, file_path, original_size, compressed_size, ""))
                
            except Exception as e:
                results.append((False, file_path, 0, 0, str(e)))
        
        # 完成进度
        self.progress_var.set(100)
        
        return results
    
    def check_dependencies(self):
        """检查必要的依赖库是否已安装"""
        algorithm = self.compression_algorithm.get()
        
        if algorithm == "pikepdf":
            try:
                import pikepdf
                return True
            except ImportError:
                messagebox.showerror(
                    "缺少依赖", 
                    "未安装pikepdf库。\n\n请运行以下命令安装:\npip install pikepdf"
                )
                return False
        
        elif algorithm == "pypdfium2":
            try:
                import pypdfium2
                return True
            except ImportError:
                messagebox.showerror(
                    "缺少依赖", 
                    "未安装pypdfium2库。\n\n请运行以下命令安装:\npip install pypdfium2"
                )
                return False
        
        elif algorithm == "ghostscript":
            try:
                import subprocess
                # 检查Ghostscript是否在系统路径中
                result = subprocess.run(["gs", "--version"], capture_output=True, text=True)
                if result.returncode == 0:
                    return True
                else:
                    messagebox.showerror(
                        "缺少依赖", 
                        "未安装Ghostscript。\n\n请从以下网址下载并安装:\nhttps://www.ghostscript.com/"
                    )
                    return False
            except:
                messagebox.showerror(
                    "缺少依赖", 
                    "未安装Ghostscript。\n\n请从以下网址下载并安装:\nhttps://www.ghostscript.com/"
                )
                return False
        
        return True
    
    def compress_with_pikepdf(self, input_path, output_path):
        """使用pikepdf压缩PDF"""
        import pikepdf
        
        with pikepdf.open(input_path) as pdf:
            # 根据压缩级别设置选项
            level = self.compression_level.get()
            
            # 设置压缩选项
            if level == "high":
                # 强力压缩
                pdf.save(output_path, 
                        compress_streams=True,
                        object_stream_mode=pikepdf.ObjectStreamMode.generate,
                        stream_decode_level=pikepdf.StreamDecodeLevel.generalized)
            elif level == "medium":
                # 中等压缩
                pdf.save(output_path, 
                        compress_streams=True,
                        object_stream_mode=pikepdf.ObjectStreamMode.preserve,
                        stream_decode_level=pikepdf.StreamDecodeLevel.specialized)
            else:
                # 轻度压缩
                pdf.save(output_path, 
                        compress_streams=True,
                        object_stream_mode=pikepdf.ObjectStreamMode.preserve)
        
        return os.path.getsize(output_path)
    
    def compress_with_pypdfium2(self, input_path, output_path):
        """使用pypdfium2压缩PDF"""
        import pypdfium2 as pdfium
        
        # 打开PDF
        pdf = pdfium.PdfDocument(input_path)
        
        # 根据压缩级别设置选项
        level = self.compression_level.get()
        
        # 设置保存选项
        save_args = {}
        if level == "high":
            save_args["compress"] = pdfium.PdfCompressMode.ALL
        elif level == "medium":
            save_args["compress"] = pdfium.PdfCompressMode.IMAGE
        else:
            save_args["compress"] = pdfium.PdfCompressMode.NONE
        
        # 保存PDF
        pdf.save(output_path, **save_args)
        
        # 关闭PDF
        pdf.close()
        
        return os.path.getsize(output_path)
    
    def compress_with_ghostscript(self, input_path, output_path):
        """使用Ghostscript压缩PDF"""
        import subprocess
        
        # 根据压缩级别设置Ghostscript参数
        level = self.compression_level.get()
        
        if level == "high":
            # 强力压缩
            gs_args = [
                "gs", "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                "-dPDFSETTINGS=/screen", "-dNOPAUSE", "-dQUIET", "-dBATCH",
                f"-sOutputFile={output_path}", input_path
            ]
        elif level == "medium":
            # 中等压缩
            gs_args = [
                "gs", "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                "-dPDFSETTINGS=/ebook", "-dNOPAUSE", "-dQUIET", "-dBATCH",
                f"-sOutputFile={output_path}", input_path
            ]
        else:
            # 轻度压缩
            gs_args = [
                "gs", "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                "-dPDFSETTINGS=/printer", "-dNOPAUSE", "-dQUIET", "-dBATCH",
                f"-sOutputFile={output_path}", input_path
            ]
        
        # 执行Ghostscript命令
        result = subprocess.run(gs_args, capture_output=True, text=True)
        
        if result.returncode != 0:
            raise Exception(f"Ghostscript执行失败: {result.stderr}")
        
        return os.path.getsize(output_path)

# 独立测试函数
def test_compress_tool():
    """独立测试压缩工具"""
    root = tk.Tk()
    root.title("PDF压缩工具测试")
    root.geometry("900x700")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    compress_tool = PDFCompressTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_compress_tool()