import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from pathlib import Path

class PDFEncryptDecryptTool:
    def __init__(self, parent_frame, file_list=None):
        self.parent = parent_frame
        self.file_list = file_list if file_list else []
        self.selected_files = []
        self.process_thread = None
        self.stop_process = False
        
        # 创建界面
        self.create_encrypt_decrypt_interface()
    
    def create_encrypt_decrypt_interface(self):
        """创建加密解密工具界面"""
        # 清除父框架中的所有内容
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # 标题
        title_label = ttk.Label(self.parent, text="PDF加密/解密工具", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(self.parent, text="为PDF文件添加密码保护或移除密码保护", 
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
        
        # 右侧选项区域 - 使用Canvas添加滚动条
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
        right_frame = ttk.LabelFrame(scrollable_frame, text="加密/解密选项")
        right_frame.pack(fill=tk.BOTH, padx=5, pady=5, ipadx=10)
        
        # 操作模式选择
        mode_frame = ttk.LabelFrame(right_frame, text="操作模式")
        mode_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.operation_mode = tk.StringVar(value="encrypt")
        
        ttk.Radiobutton(mode_frame, text="加密PDF (添加密码)", 
                       variable=self.operation_mode, value="encrypt",
                       command=self.update_interface_by_mode).pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(mode_frame, text="解密PDF (移除密码)", 
                       variable=self.operation_mode, value="decrypt",
                       command=self.update_interface_by_mode).pack(anchor=tk.W, pady=2)
        
        # 密码设置区域
        self.password_frame = ttk.LabelFrame(right_frame, text="密码设置")
        self.password_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 密码输入
        ttk.Label(self.password_frame, text="密码:").pack(anchor=tk.W, pady=(5, 0))
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(self.password_frame, textvariable=self.password_var, show="*")
        password_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # 确认密码（仅加密时显示）
        self.confirm_password_frame = ttk.Frame(self.password_frame)
        self.confirm_password_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(self.confirm_password_frame, text="确认密码:").pack(anchor=tk.W)
        self.confirm_password_var = tk.StringVar()
        self.confirm_password_entry = ttk.Entry(self.confirm_password_frame, 
                                               textvariable=self.confirm_password_var, show="*")
        self.confirm_password_entry.pack(fill=tk.X, pady=(5, 0))
        
        # 显示密码复选框
        self.show_password_var = tk.BooleanVar(value=False)
        show_password_check = ttk.Checkbutton(self.password_frame, text="显示密码",
                                             variable=self.show_password_var,
                                             command=self.toggle_password_visibility)
        show_password_check.pack(anchor=tk.W, pady=5)
        
        # 加密选项区域（仅加密时显示）
        self.encryption_options_frame = ttk.LabelFrame(right_frame, text="加密选项")
        self.encryption_options_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 加密算法
        ttk.Label(self.encryption_options_frame, text="加密算法:").pack(anchor=tk.W, pady=(5, 0))
        self.encryption_algorithm = tk.StringVar(value="AES-256")
        
        algorithm_frame = ttk.Frame(self.encryption_options_frame)
        algorithm_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Radiobutton(algorithm_frame, text="AES-256 (推荐)", 
                       variable=self.encryption_algorithm, value="AES-256").pack(side=tk.LEFT)
        ttk.Radiobutton(algorithm_frame, text="AES-128", 
                       variable=self.encryption_algorithm, value="AES-128").pack(side=tk.LEFT)
        
        # 权限设置
        permissions_frame = ttk.LabelFrame(self.encryption_options_frame, text="权限设置")
        permissions_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.allow_printing = tk.BooleanVar(value=True)
        printing_check = ttk.Checkbutton(permissions_frame, text="允许打印", 
                                        variable=self.allow_printing)
        printing_check.pack(anchor=tk.W, pady=2)
        
        self.allow_copying = tk.BooleanVar(value=True)
        copying_check = ttk.Checkbutton(permissions_frame, text="允许复制内容", 
                                       variable=self.allow_copying)
        copying_check.pack(anchor=tk.W, pady=2)
        
        self.allow_modification = tk.BooleanVar(value=False)
        modification_check = ttk.Checkbutton(permissions_frame, text="允许修改文档", 
                                            variable=self.allow_modification)
        modification_check.pack(anchor=tk.W, pady=2)
        
        self.allow_annotations = tk.BooleanVar(value=True)
        annotations_check = ttk.Checkbutton(permissions_frame, text="允许添加注释", 
                                           variable=self.allow_annotations)
        annotations_check.pack(anchor=tk.W, pady=2)
        
        # 输出选项
        output_frame = ttk.LabelFrame(right_frame, text="输出选项")
        output_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.output_suffix = tk.StringVar(value="_encrypted")
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
                              command=self.reset_tool)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 停止按钮
        self.stop_btn = ttk.Button(btn_subframe, text="停止处理", 
                                  command=self.stop_process_execution)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.stop_btn.config(state=tk.DISABLED)
        
        # 执行按钮
        self.execute_btn = ttk.Button(btn_subframe, text="开始处理", 
                                     command=self.start_process, style='Action.TButton')
        self.execute_btn.pack(side=tk.RIGHT)
        
        # 初始化界面状态
        self.update_interface_by_mode()
    
    def update_interface_by_mode(self):
        """根据操作模式更新界面"""
        mode = self.operation_mode.get()
        
        if mode == "encrypt":
            # 加密模式
            self.confirm_password_frame.pack(fill=tk.X, padx=5, pady=5)
            self.encryption_options_frame.pack(fill=tk.X, padx=5, pady=10)
            self.output_suffix.set("_encrypted")
            self.execute_btn.config(text="开始加密")
        else:
            # 解密模式
            self.confirm_password_frame.pack_forget()
            self.encryption_options_frame.pack_forget()
            self.output_suffix.set("_decrypted")
            self.execute_btn.config(text="开始解密")
    
    def toggle_password_visibility(self):
        """切换密码显示/隐藏"""
        show_password = self.show_password_var.get()
        
        # 获取所有密码输入框
        password_widgets = []
        for widget in self.password_frame.winfo_children():
            if isinstance(widget, ttk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Entry):
                        password_widgets.append(child)
            elif isinstance(widget, ttk.Entry):
                password_widgets.append(widget)
        
        # 切换显示模式
        show_char = "" if show_password else "*"
        for widget in password_widgets:
            widget.config(show=show_char)
    
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
            
            # 尝试检测文件是否已加密
            encryption_status = self.check_file_encryption(file)
            
            file_info += f"• {file_name}\n"
            file_info += f"  大小: {self.format_file_size(file_size)}\n"
            file_info += f"  加密状态: {encryption_status}\n\n"
        
        file_info += f"总计: {len(self.selected_files)} 个文件, {self.format_file_size(total_size)}"
        
        self.file_info_text.insert(tk.END, file_info)
        self.file_info_text.config(state=tk.DISABLED)
    
    def check_file_encryption(self, file_path):
        """检查PDF文件是否已加密"""
        try:
            import pikepdf
            with pikepdf.open(file_path) as pdf:
                return "未加密" if not pdf.is_encrypted else "已加密"
        except pikepdf.PasswordError:
            return "已加密 (需要密码)"
        except Exception:
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
    
    def analyze_files(self):
        """分析文件"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先选择要分析的PDF文件")
            return
        
        self.update_file_info()
        messagebox.showinfo("分析完成", "文件分析完成，请查看文件信息区域")
    
    def reset_tool(self):
        """重置工具"""
        if messagebox.askyesno("确认", "确定要重置所有设置吗？"):
            self.selected_files.clear()
            self.files_listbox.delete(0, tk.END)
            self.clear_file_info()
            self.operation_mode.set("encrypt")
            self.password_var.set("")
            self.confirm_password_var.set("")
            self.show_password_var.set(False)
            self.encryption_algorithm.set("AES-256")
            self.allow_printing.set(True)
            self.allow_copying.set(True)
            self.allow_modification.set(False)
            self.allow_annotations.set(True)
            self.output_suffix.set("_encrypted")
            self.overwrite_original.set(False)
            self.update_interface_by_mode()
            self.toggle_password_visibility()
    
    def validate_inputs(self):
        """验证输入"""
        if not self.selected_files:
            return "请先选择PDF文件"
        
        mode = self.operation_mode.get()
        password = self.password_var.get().strip()
        
        if mode == "encrypt":
            if not password:
                return "请设置加密密码"
            
            if len(password) < 4:
                return "密码长度至少为4个字符"
            
            confirm_password = self.confirm_password_var.get().strip()
            if password != confirm_password:
                return "密码和确认密码不一致"
        
        else:  # 解密模式
            if not password:
                return "请输入PDF文件的密码"
        
        return None
    
    def start_process(self):
        """开始加密或解密过程"""
        # 验证输入
        error = self.validate_inputs()
        if error:
            messagebox.showwarning("输入错误", error)
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
        
        # 禁用执行按钮，启用停止按钮
        self.execute_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))  # 显示进度条
        
        # 在后台线程中执行处理
        self.stop_process = False
        self.process_thread = threading.Thread(
            target=self.process_files_thread,
            args=(output_dir,)
        )
        self.process_thread.daemon = True
        self.process_thread.start()
    
    def stop_process_execution(self):
        """停止处理过程"""
        self.stop_process = True
        if self.process_thread and self.process_thread.is_alive():
            messagebox.showinfo("提示", "正在停止处理过程...")
    
    def process_files_thread(self, output_dir):
        """在后台线程中处理PDF文件"""
        try:
            mode = self.operation_mode.get()
            password = self.password_var.get()
            
            if mode == "encrypt":
                results = self.encrypt_pdfs(output_dir, password)
            else:
                results = self.decrypt_pdfs(output_dir, password)
            
            # 在主线程中更新UI
            self.parent.after(0, self.process_complete, results, mode)
            
        except Exception as e:
            # 在主线程中显示错误
            self.parent.after(0, lambda: messagebox.showerror("错误", f"处理PDF时出错: {str(e)}"))
            self.parent.after(0, self.reset_ui_after_process)
    
    def process_complete(self, results, mode):
        """处理完成后的回调"""
        # 重置UI状态
        self.reset_ui_after_process()
        
        # 显示处理结果
        success_count = len([r for r in results if r[0]])
        total_count = len(results)
        
        # 构建结果消息
        operation = "加密" if mode == "encrypt" else "解密"
        result_message = f"{operation}完成:\n成功: {success_count}/{total_count} 个文件\n\n"
        
        # 添加每个文件的详细结果
        for success, file_path, error in results:
            file_name = os.path.basename(file_path)
            if success:
                result_message += f"✓ {file_name}: {operation}成功\n"
            else:
                result_message += f"✗ {file_name}: {error}\n"
        
        # 显示结果
        self.show_process_results(result_message)
    
    def reset_ui_after_process(self):
        """处理完成后重置UI状态"""
        self.execute_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.progress_bar.pack_forget()  # 隐藏进度条
        self.progress_var.set(0)
    
    def show_process_results(self, result_message):
        """显示处理结果"""
        # 创建一个新窗口显示详细结果
        result_window = tk.Toplevel(self.parent)
        result_window.title("处理结果")
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
    
    def encrypt_pdfs(self, output_dir, password):
        """加密PDF文件"""
        import pikepdf
        
        results = []
        total_files = len(self.selected_files)
        
        for i, file_path in enumerate(self.selected_files):
            if self.stop_process:
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
                
                # 打开PDF文件
                with pikepdf.open(file_path) as pdf:
                    # 设置加密选项
                    encryption_args = {
                        'owner': password,
                        'user': password,
                        'allow': self.get_permissions_flags()
                    }
                    
                    # 根据选择的算法设置加密
                    if self.encryption_algorithm.get() == "AES-256":
                        encryption_args['aes'] = True
                    
                    # 保存加密文件
                    pdf.save(output_path, encryption=encryption_args)
                
                results.append((True, file_path, ""))
                
            except Exception as e:
                results.append((False, file_path, str(e)))
        
        # 完成进度
        self.progress_var.set(100)
        
        return results
    
    def decrypt_pdfs(self, output_dir, password):
        """解密PDF文件"""
        import pikepdf
        
        results = []
        total_files = len(self.selected_files)
        
        for i, file_path in enumerate(self.selected_files):
            if self.stop_process:
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
                
                # 尝试使用密码打开加密的PDF
                try:
                    with pikepdf.open(file_path, password=password) as pdf:
                        # 保存未加密的文件
                        pdf.save(output_path)
                    
                    results.append((True, file_path, ""))
                    
                except pikepdf.PasswordError:
                    results.append((False, file_path, "密码错误或文件未加密"))
                except Exception as e:
                    results.append((False, file_path, str(e)))
                
            except Exception as e:
                results.append((False, file_path, str(e)))
        
        # 完成进度
        self.progress_var.set(100)
        
        return results
    
    def get_permissions_flags(self):
        """获取权限标志"""
        import pikepdf
        
        permissions = []
        
        if self.allow_printing.get():
            permissions.append(pikepdf.Permissions.print_lowres)
            permissions.append(pikepdf.Permissions.print_highres)
        
        if self.allow_copying.get():
            permissions.append(pikepdf.Permissions.extract)
        
        if self.allow_modification.get():
            permissions.append(pikepdf.Permissions.modify)
            permissions.append(pikepdf.Permissions.modify_annotations)
            permissions.append(pikepdf.Permissions.modify_forms)
            permissions.append(pikepdf.Permissions.modify_assembly)
        
        if self.allow_annotations.get():
            permissions.append(pikepdf.Permissions.annotate)
        
        # 如果没有选择任何权限，则返回空权限
        if not permissions:
            return pikepdf.Permissions()
        
        # 合并权限
        result = permissions[0]
        for perm in permissions[1:]:
            result |= perm
        
        return result
    
    def check_dependencies(self):
        """检查必要的依赖库是否已安装"""
        try:
            import pikepdf
            return True
        except ImportError:
            messagebox.showerror(
                "缺少依赖", 
                "未安装pikepdf库。\n\n请运行以下命令安装:\npip install pikepdf"
            )
            return False

# 独立测试函数
def test_encrypt_decrypt_tool():
    """独立测试加密解密工具"""
    root = tk.Tk()
    root.title("PDF加密/解密工具测试")
    root.geometry("900x700")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    encrypt_decrypt_tool = PDFEncryptDecryptTool(frame)
    
    root.mainloop()

if __name__ == "__main__":
    test_encrypt_decrypt_tool()